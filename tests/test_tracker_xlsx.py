# type: ignore[arg-type]

import sqlite3
from typing import Any
from unittest import mock

import pytest

from mtgo_sheets.config import CONFIG, XlsxFile
from mtgo_sheets.tracker import (
    TransactionLogger,
    WishlistHandler,
    create_entries_from_rows,
    display_alerts,
    parse_cards_from_sheet,
    process_entries,
    process_order,
    update_xlsx,
    write_back_card_rows,
)
from mtgo_sheets.updater import update_db_with_prices, update_defs_from_zip
from tests.helpers import make_card_def
from tests.xlsx_fakes import FakeBook, FakeSheet


def _seed_tracker_db(conn: sqlite3.Connection) -> None:
    defs = [
        make_card_def(id_=101, name="opt", set_code="fdn"),
        make_card_def(id_=102, name="lightning bolt", set_code="m10"),
        make_card_def(id_=103, name="reanimate", set_code="tmp"),
        make_card_def(id_=104, name="regrowth", set_code="lea")
    ]
    prices = {
        101: [0.04, 0.06],
        102: [0.05, 0.08],
        103: [1.0, 2.0],
        104: [0.12, 0.14]
    }

    update_defs_from_zip(conn, defs)
    update_db_with_prices(conn, prices, len(prices))


def _dict_to_row(cols: list[str], data: dict[str, Any]) -> list[Any]:
    return [data.get(col) for col in cols]


def test_parse_cards_from_sheet(db_conn: sqlite3.Connection, xlsx_cfg: XlsxFile) -> None:
    """Verify parse_cards_from_sheet handles all cases: empty sheet, blank rows, name/id given."""
    _seed_tracker_db(db_conn)
    ws = FakeSheet()

    assert parse_cards_from_sheet(db_conn, ws) == [] # Empty sheet

    ws.range((xlsx_cfg.first_data_row, 1)).value = [
        _dict_to_row(xlsx_cfg.cols, {"id": None, "name": None, "qty": 1}), # Skip - no id/name

        # No id - resolve it using the name
        _dict_to_row(xlsx_cfg.cols, {"id": None, "name": "Lightning Bolt", "set": "m10", "foil": False, "qty": 4}),

        # Id is given
        _dict_to_row(xlsx_cfg.cols, {"id": 101, "name": "Opt", "set": "FDN", "qty": 2}),
    ]

    parsed = parse_cards_from_sheet(db_conn, ws)
    assert len(parsed) == 2
    assert parsed[0]["id"] == 102
    assert parsed[1]["id"] == 101


def test_create_entries_from_rows(db_conn: sqlite3.Connection) -> None:
    """Verify WatchlistEntries are constructed from rows. Bad rows are added to failed_rows."""
    _seed_tracker_db(db_conn)
    rows = [
        {"id": 101, "name": "Opt", "qty": 2, "purchase_price": 0.04},      # Valid
        {"id": 999, "name": "Fake Card", "qty": 1, "notes": "note text"},   # Invalid
    ]

    entries, failed = create_entries_from_rows(db_conn, rows)
    assert len(entries) == 1
    assert entries[0].id_ == 101

    assert len(failed) == 1
    assert failed[0]["id"] == 999
    assert "note text" in failed[0]["notes"]


def test_process_entries_and_orders(db_conn: sqlite3.Connection, xlsx_cfg: XlsxFile, fake_app: Any) -> None:
    """Verify process_entries performs all order types.

    Entry attributes and alerts are updated. Trigger alerts are captured."""
    _seed_tracker_db(db_conn)
    logger = TransactionLogger(xlsx_cfg.log_path, xlsx_cfg.log_cols, fake_app)
    wishlist_handler = WishlistHandler()
    buy_wishlist_path = xlsx_cfg.wishlist_dir / f"wishlist_{CONFIG.today}_buy_goat.dek"
    sell_wishlist_path = xlsx_cfg.wishlist_dir / f"wishlist_{CONFIG.today}_sell_goat.dek"

    rows = [
        # Buy order - triggers drop - clears drop thresh (bought at/below the drop threshold)
        {"id": 101, "name": "Opt", "qty": 2, "purchase_price": 0.04, "drop": 0.10, "order": "b2;0.05"},

        # Sell order - triggers rise - doesn't clear rise thresh (3 copies still held)
        {"id": 102, "name": "Lightning Bolt", "qty": 4, "purchase_price": 0.05, "rise": 0.08, "order": "s1;0.09"},

        # Suggestion - ignored (no thresholds) - will be considered "suggested last time"
        {"id": 103, "name": "Reanimate", "qty": -1},

        {"id": 104, "name": "Regrowth", "qty": 1, "purchase_price": 0.25, "drop": 0.20},

    ]
    entries, _ = create_entries_from_rows(db_conn, rows)
    output_rows, suggested_last_time, alerts = process_entries(logger, wishlist_handler, entries)

    # Output rows and suggestion tracking
    assert len(output_rows) == 3

    # Opt buy row
    assert output_rows[0]["qty"] == 4
    assert output_rows[0]["drop"] == ""
    assert output_rows[0]["hit"] == ""

    # Bolt sell row
    assert output_rows[1]["qty"] == 3
    assert output_rows[1]["rise"] == 0.08
    assert output_rows[1]["hit"] == ""

    # Reanimate suggestion row
    assert suggested_last_time == [103]

    # Regrowth row
    assert output_rows[2]["qty"] == 1
    assert output_rows[2]["hit"] == "dropped"

    # Alerts
    assert list(alerts.keys()) == ["dropped"]
    assert len(alerts["dropped"]) == 1
    assert "Regrowth" in alerts["dropped"][0]

    # Logger
    assert len(logger.transactions) == 2

    # Wishlist Hadler
    assert (101, 4) in wishlist_handler.wishlists[buy_wishlist_path]
    assert (102, 1) in wishlist_handler.wishlists[sell_wishlist_path]

    # Commit transactions, relaunch logger, then make more transactions
    logger.close()
    fake_app.books.by_path[xlsx_cfg.log_path] = logger.wb
    logger = TransactionLogger(xlsx_cfg.log_path, xlsx_cfg.log_cols, fake_app)

    # Arbitrage and Revision order types
    new_rows = [
        # Revise the buy order from above
        {"id": 101, "name": "Opt", "qty": 4, "purchase_price": 0.045, "order": "r1;0.04"},

        # Arbitrage order
        {"id": 102, "name": "Lightning Bolt", "qty": 3, "purchase_price": 0.05, "order": "a1;0.04;0.09"},
    ]
    new_entries, _ = create_entries_from_rows(db_conn, new_rows)
    process_entries(logger, wishlist_handler, new_entries)
    assert len(logger.revisions) == 1
    assert len(logger.transactions) == 1

    # Revise an order for a card that has no existing transactions - Throws error
    invalid_revision_rows = [{"id": 103, "name": "Reanimate", "qty": 1, "purchase_price": 0.01, "order": "r2"}]
    invalid_revision_entries, _ = create_entries_from_rows(db_conn, invalid_revision_rows)
    with pytest.raises(ValueError, match="Unable to undo a past transaction"):
        process_order(invalid_revision_entries[0], logger, wishlist_handler)

    # Invalid command - Throws error
    invalid_command_rows = [{"id": 103, "name": "Reanimate", "qty": 1, "purchase_price": 0.01, "order": "z2"}]
    invalid_command_entries, _ = create_entries_from_rows(db_conn, invalid_command_rows)
    with pytest.raises(ValueError, match="Unknown action z"):
        process_order(invalid_command_entries[0], logger, wishlist_handler)

    logger.close()


def test_write_back_card_rows(xlsx_cfg: XlsxFile) -> None:
    """Verify write_back_card_rows updates sheet, clears old data, and checks if the changes were made succesfully."""
    ws = FakeSheet()
    wb = FakeBook(xlsx_cfg.path)

    ws.cells(xlsx_cfg.first_data_row + 1, 1).value = "should be cleared"
    new_rows = [{"id": 101, "name": "Opt", "qty": 4}]

    # Fake mtime changing, indicating successful write
    write_back_card_rows(ws, wb, new_rows)

    assert str(CONFIG.today) in ws.cells(1, 2).value # "Last Updated" date cell
    assert ws.cells(xlsx_cfg.first_data_row, 1).value == 101
    assert ws.cells(xlsx_cfg.first_data_row + 1, 1).value is None # old data cleared

    # Fake no mtime change indicating locked file (no data saved)
    wb.simulate_locked = True
    with pytest.raises(OSError, match="Unable to save the Excel file"):
        write_back_card_rows(ws, wb, new_rows)


def test_display_alerts(capsys: pytest.CaptureFixture[str]) -> None:
    """Verify display_alerts prints all alerts in the expected order (arb, rise, drop)."""
    alerts = {
        "drop": ["Drop alert 1"],
        "arb": ["Arb alert 1"],
        "rise": ["Rise alert 1", "Rise alert 2"],
    }
    expected_display_order = [
        "Arb",
        "Arb alert 1",
        "Rise",
        "Rise alert 1",
        "Rise alert 2",
        "Drop",
        "Drop alert 1"
    ]

    display_alerts(alerts)
    out = capsys.readouterr().out
    pos = 0

    # Check that each expected substring is present in the correct order
    for expected in expected_display_order:
        pos = out.index(expected, pos) + len(expected)


def test_update_xlsx(
    db_conn: sqlite3.Connection, xlsx_cfg: XlsxFile, fake_app: Any) -> None:
    """Verify update_xlsx results in a correctly updated sheet.

    Each intermediate function call is not directly tested here."""
    _seed_tracker_db(db_conn)

    wb = fake_app.books.open(xlsx_cfg.path)
    ws = wb.sheets[0]
    ws.range((xlsx_cfg.first_data_row, 1)).value = [
        _dict_to_row(xlsx_cfg.cols, {"id": 102, "name": "Lightning Bolt", "set": "M10", "qty": 1}),
        _dict_to_row(xlsx_cfg.cols, {"id": 101, "name": "Opt", "set": "FDN", "qty": 2, "order": "b2;0.05"}),
    ]

    with mock.patch("mtgo_sheets.tracker.get_db_connection") as mock_get_conn:
        mock_get_conn.return_value.__enter__.return_value = db_conn
        update_xlsx(xlsx_cfg)

    ws = fake_app.books.by_path[xlsx_cfg.path].sheets[0]

    # Verify rows are sorted using the default sort in mock0.toml (alphabetically)
    first_row = xlsx_cfg.first_data_row
    assert ws.cells(first_row, 1).value == 102 # id
    assert ws.cells(first_row, 5).value == 1 # qty

    assert ws.cells(first_row + 1, 1).value == 101
    assert ws.cells(first_row + 1, 5).value == 4 # Opt's quantity increased after the buy order

    # Verify the tracker sheet, log sheet, and wishlist dumps are all updated and saved
    assert ws.cells(1, 2).value == f"Last updated: {CONFIG.today}"
    assert xlsx_cfg.log_path.exists()

    buy_wishlist = xlsx_cfg.wishlist_dir / f"wishlist_{CONFIG.today}_buy_goat.dek"
    assert buy_wishlist.exists()

    assert fake_app.quitted is True # Verify a graceful exit to prevent locking up any sheets
