import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest

from app import database


@pytest.fixture()
def sqlite_connection(tmp_path):
    db_path = tmp_path / "test.db"
    database.init_db(db_path)
    conn = database.get_connection(db_path)
    try:
        yield conn
    finally:
        conn.close()
