"""MySQL connection, SQL compatibility, and asynchronous log persistence."""

import atexit
import hashlib
import logging
import os
import queue
import re
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pymysql
from pymysql.cursors import DictCursor

from mysql_migrations import apply_mysql_migrations


_ON_CONFLICT_RE = re.compile(r"ON\s+CONFLICT\s*\([^)]*\)\s+DO\s+UPDATE\s+SET", re.IGNORECASE)
_EXCLUDED_RE = re.compile(r"\bexcluded\.([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
_FIELD_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)=([^ ]*)")


@dataclass(frozen=True)
class MySQLConfig:
    host: str
    port: int
    user: str
    password: str
    database: str = "baoguan_data"
    pool_size: int = 8
    connect_timeout: int = 8
    read_timeout: int = 30
    write_timeout: int = 30
    ssl_ca: str = ""
    require_tls: bool = True

    @classmethod
    def from_mapping(cls, values):
        def pick(*names, default=""):
            for name in names:
                value = str(values.get(name, "") or "").strip()
                if value:
                    return value
            return default

        def pick_bool(name, default):
            value = pick(name)
            return default if not value else value.lower() not in {"0", "false", "no", "off"}

        config = cls(
            host=pick("MYSQL_HOST", "host"),
            port=int(pick("MYSQL_PORT", "port", default="3306")),
            user=pick("MYSQL_USER", "user"),
            password=pick("MYSQL_PASSWORD", "password"),
            database=pick("MYSQL_DATABASE", default="baoguan_data"),
            pool_size=max(2, int(pick("MYSQL_POOL_SIZE", default="8"))),
            connect_timeout=max(1, int(pick("MYSQL_CONNECT_TIMEOUT_SEC", default="8"))),
            read_timeout=max(1, int(pick("MYSQL_READ_TIMEOUT_SEC", default="30"))),
            write_timeout=max(1, int(pick("MYSQL_WRITE_TIMEOUT_SEC", default="30"))),
            ssl_ca=pick("MYSQL_SSL_CA"),
            require_tls=pick_bool("MYSQL_REQUIRE_TLS", True),
        )
        missing = [
            name
            for name, value in (
                ("MYSQL_HOST", config.host),
                ("MYSQL_USER", config.user),
                ("MYSQL_PASSWORD", config.password),
            )
            if not value
        ]
        if missing:
            raise ValueError("missing MySQL configuration: " + ", ".join(missing))
        if config.database != "baoguan_data":
            raise ValueError("MYSQL_DATABASE must be baoguan_data")
        return config


def _translate_sql(sql):
    translated = _ON_CONFLICT_RE.sub("ON DUPLICATE KEY UPDATE", sql)
    translated = _EXCLUDED_RE.sub(r"VALUES(\1)", translated)
    return translated.replace("?", "%s")


class _ConnectionPool:
    def __init__(self, config):
        self.config = config
        self._available = queue.LifoQueue(maxsize=config.pool_size)
        self._created = 0
        self._lock = threading.Lock()

    def _connect(self):
        kwargs = {
            "host": self.config.host,
            "port": self.config.port,
            "user": self.config.user,
            "password": self.config.password,
            "database": self.config.database,
            "charset": "utf8mb4",
            "cursorclass": DictCursor,
            "autocommit": False,
            "connect_timeout": self.config.connect_timeout,
            "read_timeout": self.config.read_timeout,
            "write_timeout": self.config.write_timeout,
        }
        if self.config.ssl_ca:
            ca_path = Path(self.config.ssl_ca).expanduser().resolve()
            if not ca_path.is_file():
                raise ValueError("MYSQL_SSL_CA does not exist: %s" % ca_path)
            kwargs["ssl"] = {"ca": str(ca_path), "check_hostname": True}
        connection = pymysql.connect(**kwargs)
        if self.config.require_tls:
            with connection.cursor() as cursor:
                cursor.execute("SHOW STATUS LIKE 'Ssl_cipher'")
                row = cursor.fetchone()
            cipher = str(row.get("Value", "") if isinstance(row, dict) else row[1] if row else "")
            if not cipher:
                connection.close()
                raise RuntimeError("MySQL connection is not encrypted with TLS")
        return connection

    def acquire(self):
        try:
            connection = self._available.get_nowait()
        except queue.Empty:
            with self._lock:
                if self._created < self.config.pool_size:
                    self._created += 1
                    try:
                        return self._connect()
                    except Exception:
                        self._created -= 1
                        raise
            connection = self._available.get(timeout=self.config.connect_timeout)
        try:
            connection.ping(reconnect=True)
            return connection
        except Exception:
            self.discard(connection)
            return self.acquire()

    def release(self, connection):
        try:
            connection.rollback()
            self._available.put_nowait(connection)
        except Exception:
            self.discard(connection)

    def discard(self, connection):
        try:
            connection.close()
        finally:
            with self._lock:
                self._created = max(0, self._created - 1)


class MySQLConnectionAdapter:
    def __init__(self, pool):
        self._pool = pool
        self._connection = pool.acquire()
        self._cursor = None
        self.rowcount = -1
        self._closed = False

    def execute(self, sql, params=()):
        if self._cursor is not None:
            self._cursor.close()
        self._cursor = self._connection.cursor()
        self._cursor.execute(_translate_sql(sql), tuple(params))
        self.rowcount = int(self._cursor.rowcount)
        return self

    def fetchone(self):
        return self._cursor.fetchone() if self._cursor is not None else None

    def fetchall(self):
        return list(self._cursor.fetchall()) if self._cursor is not None else []

    def commit(self):
        self._connection.commit()

    def rollback(self):
        self._connection.rollback()

    def close(self):
        if self._closed:
            return
        self._closed = True
        if self._cursor is not None:
            self._cursor.close()
        self._pool.release(self._connection)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, _exc, _tb):
        try:
            if exc_type is None:
                self.commit()
            else:
                self.rollback()
        finally:
            self.close()


class MySQLDatabase:
    def __init__(self, config):
        self.config = config
        self._pool = _ConnectionPool(config)

    @classmethod
    def from_env(cls):
        return cls(MySQLConfig.from_mapping(os.environ))

    def initialize(self):
        connection = self._pool.acquire()
        try:
            apply_mysql_migrations(connection)
        except Exception:
            connection.rollback()
            raise
        finally:
            self._pool.release(connection)

    def connect(self):
        return MySQLConnectionAdapter(self._pool)

    def ping(self):
        with self.connect() as connection:
            row = connection.execute("SELECT 1 AS ok").fetchone()
        return bool(row and int(row["ok"]) == 1)

    def insert_application_logs(self, rows):
        if not rows:
            return 0
        sql = """
            INSERT IGNORE INTO application_logs(
                occurred_at, level, event_type, trace_id, method, path, status,
                took_ms, contract_no, token, filename, row_count, message,
                source, source_fingerprint
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        values = [
            (
                row["occurred_at"], row["level"], row["event_type"], row["trace_id"],
                row["method"], row["path"], row["status"], row["took_ms"],
                row["contract_no"], row["token"], row["filename"], row["row_count"],
                row["message"], row["source"], row["source_fingerprint"],
            )
            for row in rows
        ]
        connection = self._pool.acquire()
        try:
            with connection.cursor() as cursor:
                inserted = cursor.executemany(sql, values)
            connection.commit()
            return int(inserted)
        except Exception:
            connection.rollback()
            raise
        finally:
            self._pool.release(connection)

    def upsert_request_cache(self, rows):
        sql = """
            INSERT INTO request_cache(token, trace_id, payload, created_at, expires_at)
            VALUES (%s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                trace_id=VALUES(trace_id), payload=VALUES(payload),
                created_at=VALUES(created_at), expires_at=VALUES(expires_at)
        """
        values = [
            (row["token"], row["trace_id"], row["payload"], row["created_at"], row["expires_at"])
            for row in rows
        ]
        return self._executemany(sql, values)

    def upsert_generated_files(self, rows):
        sql = """
            INSERT INTO generated_files(token, trace_id, filename, created_at, last_access_at)
            VALUES (%s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                trace_id=VALUES(trace_id), filename=VALUES(filename),
                created_at=VALUES(created_at), last_access_at=VALUES(last_access_at)
        """
        values = [
            (row["token"], row["trace_id"], row["filename"], row["created_at"], row["last_access_at"])
            for row in rows
        ]
        return self._executemany(sql, values)

    def _executemany(self, sql, values):
        if not values:
            return 0
        connection = self._pool.acquire()
        try:
            with connection.cursor() as cursor:
                affected = cursor.executemany(sql, values)
            connection.commit()
            return int(affected)
        except Exception:
            connection.rollback()
            raise
        finally:
            self._pool.release(connection)


def _optional_int(value):
    try:
        return int(value) if value not in (None, "") else None
    except (TypeError, ValueError):
        return None


def application_log_row(message, level, occurred_at, source, fingerprint_seed):
    fields = dict(_FIELD_RE.findall(message))
    first = message.split(" ", 1)[0] if message else ""
    event_type = first if first in {"REQ", "RES", "CACHE", "BIZ", "ERR", "UNHANDLED", "CLEANUP", "HTTP_ACCESS"} else "SYSTEM"
    fingerprint = hashlib.sha256(fingerprint_seed.encode("utf-8", errors="replace")).hexdigest()
    return {
        "occurred_at": occurred_at,
        "level": str(level or "INFO")[:16],
        "event_type": event_type,
        "trace_id": fields.get("trace_id", "")[:191],
        "method": fields.get("method", "")[:16],
        "path": fields.get("path", "")[:512],
        "status": _optional_int(fields.get("status")),
        "took_ms": _optional_int(fields.get("took_ms")),
        "contract_no": fields.get("contract_no", "")[:255],
        "token": fields.get("token", "")[:64],
        "filename": fields.get("filename", "")[:512],
        "row_count": _optional_int(fields.get("row_count")),
        "message": message,
        "source": source[:64],
        "source_fingerprint": fingerprint,
    }


def application_log_from_record(record, source="runtime"):
    occurred_at = datetime.fromtimestamp(record.created)
    seed = "%s|%.6f|%s|%s|%s" % (
        source,
        record.created,
        record.process,
        record.thread,
        record.getMessage(),
    )
    return application_log_row(record.getMessage(), record.levelname, occurred_at, source, seed)


class AsyncMySQLLogHandler(logging.Handler):
    def __init__(self, database, batch_size=100, flush_interval=0.5, queue_size=10000):
        super().__init__()
        self.database = database
        self.batch_size = batch_size
        self.flush_interval = flush_interval
        self.queue = queue.Queue(maxsize=queue_size)
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, name="mysql-log-writer", daemon=True)
        self._thread.start()
        atexit.register(self.close)

    def emit(self, record):
        try:
            self.queue.put_nowait(application_log_from_record(record))
        except Exception:
            self.handleError(record)

    def _run(self):
        while not self._stop.is_set() or not self.queue.empty():
            try:
                first = self.queue.get(timeout=self.flush_interval)
            except queue.Empty:
                continue
            batch = [first]
            while len(batch) < self.batch_size:
                try:
                    batch.append(self.queue.get_nowait())
                except queue.Empty:
                    break
            try:
                self.database.insert_application_logs(batch)
            except Exception:
                time.sleep(min(1.0, self.flush_interval * 2))
            finally:
                for _ in batch:
                    self.queue.task_done()

    def close(self):
        if getattr(self, "_stop", None) is None or self._stop.is_set():
            return
        self._stop.set()
        self._thread.join(timeout=5)
        super().close()


MYSQL_INTEGRITY_ERRORS = (pymysql.IntegrityError,)
