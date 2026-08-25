"""Ordered and checksum-protected MySQL schema migrations."""

import hashlib
from pathlib import Path


def _load_statements(filename):
    body = (Path(__file__).resolve().parent / filename).read_text(encoding="utf-8")
    return tuple(part.strip() for part in body.split("-- statement --") if part.strip())


MIGRATIONS = (
    (1, "initial_mysql_storage", _load_statements("V001__initial_mysql_storage.sql")),
)


def _checksum(statements):
    body = "\n-- statement --\n".join(statement.strip() for statement in statements)
    return hashlib.sha256(body.encode("utf-8")).hexdigest()


def apply_mysql_migrations(connection):
    with connection.cursor() as cursor:
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS schema_migrations (
                version INT PRIMARY KEY,
                name VARCHAR(191) NOT NULL,
                checksum CHAR(64) NOT NULL,
                applied_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
        )
        cursor.execute("SELECT version, checksum FROM schema_migrations ORDER BY version")
        applied = {}
        for row in cursor.fetchall():
            if isinstance(row, dict):
                applied[int(row["version"])] = str(row["checksum"])
            else:
                applied[int(row[0])] = str(row[1])

        for version, name, statements in MIGRATIONS:
            checksum = _checksum(statements)
            existing = applied.get(version)
            if existing is not None:
                if existing != checksum:
                    raise RuntimeError("MySQL migration v%s checksum mismatch" % version)
                continue
            for statement in statements:
                cursor.execute(statement)
            cursor.execute(
                "INSERT INTO schema_migrations(version, name, checksum) VALUES (%s, %s, %s)",
                (version, name, checksum),
            )
    connection.commit()
