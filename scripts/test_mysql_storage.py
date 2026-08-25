import logging
import unittest
from datetime import datetime

from migrate_to_mysql import parse_log_line
from mysql_database import MySQLConfig, _translate_sql, application_log_from_record
from mysql_migrations import apply_mysql_migrations


class MySQLConfigTests(unittest.TestCase):
    def test_requires_connection_settings_and_uses_baoguan_database(self):
        config = MySQLConfig.from_mapping(
            {
                "MYSQL_HOST": "db.example",
                "MYSQL_USER": "baoguan",
                "MYSQL_PASSWORD": "secret",
            }
        )

        self.assertEqual(config.database, "baoguan_data")
        self.assertTrue(config.require_tls)

    def test_missing_password_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "MYSQL_PASSWORD"):
            MySQLConfig.from_mapping({"MYSQL_HOST": "db.example", "MYSQL_USER": "baoguan"})


class SQLTranslationTests(unittest.TestCase):
    def test_translates_parameters_and_sqlite_upsert(self):
        sql = """
        INSERT INTO generated_files(token, filename) VALUES (?, ?)
        ON CONFLICT(token) DO UPDATE SET filename=excluded.filename
        """

        translated = _translate_sql(sql)

        self.assertIn("VALUES (%s, %s)", translated)
        self.assertIn("ON DUPLICATE KEY UPDATE", translated)
        self.assertIn("filename=VALUES(filename)", translated)


class _Cursor:
    def __init__(self):
        self.executed = []
        self.selecting = False

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return None

    def execute(self, sql, params=()):
        self.executed.append((sql, params))
        self.selecting = sql.strip().startswith("SELECT version")

    def fetchall(self):
        return []


class _Connection:
    def __init__(self):
        self.cursor_instance = _Cursor()
        self.committed = False

    def cursor(self):
        return self.cursor_instance

    def commit(self):
        self.committed = True


class MigrationTests(unittest.TestCase):
    def test_initial_migration_creates_business_and_log_tables(self):
        connection = _Connection()

        apply_mysql_migrations(connection)

        sql = "\n".join(statement for statement, _ in connection.cursor_instance.executed)
        self.assertIn("CREATE TABLE IF NOT EXISTS request_cache", sql)
        self.assertIn("CREATE TABLE IF NOT EXISTS generated_files", sql)
        self.assertIn("CREATE TABLE IF NOT EXISTS application_logs", sql)
        self.assertIn("INSERT INTO schema_migrations", sql)
        self.assertTrue(connection.committed)


class ApplicationLogTests(unittest.TestCase):
    def test_structured_fields_are_extracted_from_existing_log_message(self):
        record = logging.LogRecord(
            "backend_server",
            logging.INFO,
            __file__,
            1,
            "RES trace_id=BG-TEST method=GET path=/generate status=200 took_ms=1086",
            (),
            None,
        )
        record.created = 1787648400.125

        row = application_log_from_record(record, source="runtime")

        self.assertEqual(row["event_type"], "RES")
        self.assertEqual(row["trace_id"], "BG-TEST")
        self.assertEqual(row["method"], "GET")
        self.assertEqual(row["path"], "/generate")
        self.assertEqual(row["status"], 200)
        self.assertEqual(row["took_ms"], 1086)
        self.assertEqual(len(row["source_fingerprint"]), 64)

    def test_historical_file_log_has_stable_fingerprint(self):
        line = "2026-08-25 17:57:22,478 INFO RES trace_id=BG-HISTORY method=GET path=/generate status=200 took_ms=3\n"

        first = parse_log_line(line, "backend_access.log", 42, datetime(2026, 8, 25))
        second = parse_log_line(line, "backend_access.log", 42, datetime(2026, 8, 25))

        self.assertEqual(first["occurred_at"], datetime(2026, 8, 25, 17, 57, 22, 478000))
        self.assertEqual(first["trace_id"], "BG-HISTORY")
        self.assertEqual(first["source_fingerprint"], second["source_fingerprint"])

    def test_werkzeug_access_log_is_structured(self):
        line = '127.0.0.1 - - [25/Aug/2026 17:57:22] "GET /generate?t=x HTTP/1.0" 200 -\n'

        row = parse_log_line(line, "server.log", 7, datetime(2026, 8, 25))

        self.assertEqual(row["event_type"], "HTTP_ACCESS")
        self.assertEqual(row["method"], "GET")
        self.assertEqual(row["path"], "/generate?t=x")
        self.assertEqual(row["status"], 200)


if __name__ == "__main__":
    unittest.main()
