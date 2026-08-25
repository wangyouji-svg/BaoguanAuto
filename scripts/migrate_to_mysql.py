"""Migrate customs SQLite business data and text logs to MySQL."""

import argparse
import hashlib
import json
import os
import re
import sqlite3
import tempfile
from datetime import datetime
from pathlib import Path

from dotenv import dotenv_values

from mysql_database import MySQLConfig, MySQLDatabase, application_log_row


APP_LOG_PATTERNS = (
    re.compile(r"^(?P<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3}) (?P<level>[A-Z]+) (?P<message>.*)$"),
    re.compile(r"^\[(?P<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3})\] (?P<level>[A-Z]+).*?: (?P<message>.*)$"),
)
HTTP_ACCESS_RE = re.compile(
    r'^.*?\[(?P<ts>\d{2}/[A-Za-z]{3}/\d{4} \d{2}:\d{2}:\d{2})\] '
    r'"(?P<method>[A-Z]+) (?P<path>[^ ]+).*?" (?P<status>\d{3})'
)


def _load_config(env_path):
    values = dict(dotenv_values(env_path))
    values.update({key: value for key, value in os.environ.items() if str(value).strip()})
    values["MYSQL_DATABASE"] = "baoguan_data"
    return MySQLConfig.from_mapping(values)


def _consistent_copy(source_path, destination_path):
    source = sqlite3.connect("file:%s?mode=ro" % source_path.resolve(), uri=True)
    destination = sqlite3.connect(destination_path)
    try:
        source.backup(destination)
    finally:
        destination.close()
        source.close()


def parse_log_line(raw_line, source, line_number, fallback_time):
    raw_line = raw_line.rstrip("\r\n")
    occurred_at = fallback_time
    level = "INFO"
    message = raw_line
    for pattern in APP_LOG_PATTERNS:
        match = pattern.match(raw_line)
        if match:
            occurred_at = datetime.strptime(match.group("ts"), "%Y-%m-%d %H:%M:%S,%f")
            level = match.group("level")
            message = match.group("message")
            break
    else:
        match = HTTP_ACCESS_RE.match(raw_line)
        if match:
            occurred_at = datetime.strptime(match.group("ts"), "%d/%b/%Y %H:%M:%S")
            message = "HTTP_ACCESS method=%s path=%s status=%s" % (
                match.group("method"),
                match.group("path"),
                match.group("status"),
            )

    seed = "%s|%s|%s" % (source, line_number, raw_line)
    return application_log_row(message, level, occurred_at, source, seed)


def _migrate_sqlite(snapshot_path, database):
    source = sqlite3.connect(snapshot_path)
    source.row_factory = sqlite3.Row
    try:
        cache_rows = [dict(row) for row in source.execute(
            "SELECT token, trace_id, payload, created_at, expires_at FROM request_cache ORDER BY token"
        )]
        generated_rows = [dict(row) for row in source.execute(
            "SELECT token, trace_id, filename, created_at, last_access_at FROM generated_files ORDER BY token"
        )]
    finally:
        source.close()

    database.upsert_request_cache(cache_rows)
    database.upsert_generated_files(generated_rows)
    return {
        "request_cache": _verify_tokens(database, "request_cache", cache_rows),
        "generated_files": _verify_tokens(database, "generated_files", generated_rows),
    }


def _verify_tokens(database, table, source_rows):
    missing = []
    with database.connect() as connection:
        for row in source_rows:
            found = connection.execute(
                "SELECT 1 AS ok FROM `%s` WHERE token = ?" % table,
                (row["token"],),
            ).fetchone()
            if found is None:
                missing.append(row["token"])
    if missing:
        raise RuntimeError("verification failed for %s: %s missing tokens" % (table, len(missing)))
    return {"source": len(source_rows), "verified": len(source_rows)}


def _migrate_log_file(path, database, batch_size=500):
    if not path.is_file():
        return {"source": 0, "inserted": 0}
    fallback_time = datetime.fromtimestamp(path.stat().st_mtime)
    source_name = path.name
    source_count = 0
    inserted = 0
    batch = []
    with path.open("r", encoding="utf-8", errors="replace") as handle:
        for line_number, line in enumerate(handle, start=1):
            source_count += 1
            batch.append(parse_log_line(line, source_name, line_number, fallback_time))
            if len(batch) >= batch_size:
                inserted += database.insert_application_logs(batch)
                batch = []
    inserted += database.insert_application_logs(batch)
    return {"source": source_count, "inserted": inserted}


def migrate(source_db, log_paths, database):
    started_at = datetime.now()
    run_seed = "%s|%s|%s" % (source_db.resolve(), source_db.stat().st_mtime_ns, started_at.isoformat())
    run_id = hashlib.sha256(run_seed.encode("utf-8")).hexdigest()
    report = {"runId": run_id, "startedAt": started_at.isoformat(timespec="milliseconds")}
    with tempfile.TemporaryDirectory(prefix="baoguan-mysql-migration-") as temp_dir:
        snapshot = Path(temp_dir) / "request_cache.snapshot.sqlite3"
        _consistent_copy(source_db, snapshot)
        report["tables"] = _migrate_sqlite(snapshot, database)
    report["logs"] = {
        path.name: _migrate_log_file(path, database)
        for path in log_paths
    }
    report["completedAt"] = datetime.now().isoformat(timespec="milliseconds")
    with database.connect() as connection:
        connection.execute(
            """
            INSERT INTO data_migration_runs(
                run_id, source_database, status, report_json, started_at, completed_at
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                str(source_db.resolve()),
                "completed",
                json.dumps(report, ensure_ascii=False, separators=(",", ":")),
                started_at,
                datetime.now(),
            ),
        )
    return report


def main():
    parser = argparse.ArgumentParser(description="Migrate customs data and logs to baoguan_data")
    parser.add_argument("--source", required=True)
    parser.add_argument("--env", required=True)
    parser.add_argument("--log", action="append", default=[])
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    source = Path(args.source).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    database = MySQLDatabase(_load_config(Path(args.env).expanduser().resolve()))
    database.initialize()
    report = migrate(source, [Path(value).expanduser().resolve() for value in args.log], database)
    report_path = Path(args.report).expanduser().resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
