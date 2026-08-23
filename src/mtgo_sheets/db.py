import re
import sqlite3
from collections.abc import Generator
from contextlib import contextmanager
from pathlib import Path

from mtgo_sheets.config import CONFIG


def regexp(expr: str, item: str) -> bool:
    """Allow queries to use regex."""
    reg = re.compile(expr)
    return reg.search(item) is not None

def init_db(conn: sqlite3.Connection) -> None:
    """Initialize DB tables if needed. Does nothing if the DB is already set up."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS prices (
        id INTEGER PRIMARY KEY,
        price_list TEXT)
        """,
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS defs (
        id INTEGER PRIMARY KEY,
        name TEXT,
        set_ TEXT,
        foil BOOLEAN,
        ver TEXT)
        """,
    )

    conn.commit()

@contextmanager
def get_db_connection(db_path: Path | None = None) -> Generator[sqlite3.Connection]:
    """Yield a database connection and ensures it is safely closed afterwards."""
    target_path = db_path or CONFIG.paths.db
    conn = sqlite3.connect(target_path, autocommit=False)
    conn.create_function("REGEXP", 2, regexp)

    try:
        init_db(conn)
        yield conn
    finally:
        conn.close()
