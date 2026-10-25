from __future__ import annotations
from collections.abc import Callable, Generator
from pathlib import Path
from typing import Any

import pytest
import sqlite3

from mtgo_sheets import config, db, tracker
from mtgo_sheets.models import CardDef
from tests.helpers import make_card_def
from tests.xlsx_fakes import FakeApp

# Override CONFIG in each module
mock_cfg = config.load_config(Path(__file__).with_name("mock0.toml"))
config.CONFIG = mock_cfg
tracker.CONFIG = mock_cfg


@pytest.fixture
def fake_app() -> Any:
    return FakeApp()

@pytest.fixture
def xlsx_cfg(tmp_path: Path) -> Generator[config.XlsxFile]:
    """Provide an active XlsxFile config with all file paths routed to pytest's tmp_path."""
    base_file = config.CONFIG.xlsx["main"]
    file_cfg = base_file.model_copy(
        update={
            "path": tmp_path / "sheet.xlsx",
            "log_path": tmp_path / "log.xlsx",
            "wishlist_dir": tmp_path / "wishlists",
            "suggestion_card_blacklist": tmp_path / "blacklist.json",
        }
    )
    # Set the active sheet context required by tracker helper functions
    current_xlsx_var = getattr(tracker, "_CURRENT_XLSX")
    token = current_xlsx_var.set(file_cfg)
    yield file_cfg
    current_xlsx_var.reset(token)


@pytest.fixture
def sample_card_def() -> CardDef:
    """Default CardDef instance for single-card unit tests."""
    return make_card_def()


@pytest.fixture
def db_conn() -> Generator[sqlite3.Connection]:
    conn = sqlite3.connect(":memory:", autocommit=False)
    conn.create_function("REGEXP", 2, db.regexp)
    db.init_db(conn)
    yield conn
    conn.close()
