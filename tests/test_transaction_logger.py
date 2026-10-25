import datetime
from pathlib import Path

import pytest
from pydantic import ValidationError

from mtgo_sheets.config import CONFIG
from mtgo_sheets.models import Transaction
from mtgo_sheets.tracker import TransactionLogger
from tests.xlsx_fakes import FakeApp, FakeBook

LOG_COLS = [
    "date", "action", "id", "name", "price_each", "price_total",
    "net_each", "net_total", "qty", "vendor", "holdings", "notes",
]


def test_transaction_required_fields() -> None:
    """Verify Transaction object validates all required fields for each action type."""
    buy_tx = Transaction(action="buy", id_=123, name="Opt", price_each=0.004, qty=4, vendor="goat", holdings=4)
    assert buy_tx.action == "buy"

    with pytest.raises(ValidationError, match="Invalid action unknown"):
        Transaction(action="unknown", id_=123, name="Opt", price_each=0.004, qty=1, vendor="goat", holdings=1)

    with pytest.raises(ValidationError, match="Missing required fields for action 'sell': 'net_each'"):
        Transaction(action="sell", id_=123, name="Opt", price_each=0.004, qty=1, vendor="goat", holdings=0)

    with pytest.raises(ValidationError, match="Missing required fields for action 'arb': 'net_each'"):
        Transaction(action="arb", id_=123, name="Opt", price_each=0.004, qty=1, vendor="goat", holdings=0)

    with pytest.raises(ValidationError, match="Missing required fields for action 'revise': 'old_action', 'old_cmd'"):
        Transaction(action="revise", id_=123, name="Opt", price_each=0.004, qty=1, vendor="goat", holdings=1)


def test_logger_initializes_fresh_workbook_and_appends_transactions(
    tmp_path: Path, fake_app: FakeApp, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify Logger initializes a new log file with headers and writes calculated totals on close."""
    monkeypatch.setattr("mtgo_sheets.tracker.xw.Book", FakeBook)
    log_path = tmp_path / "new_log.xlsx"

    logger = TransactionLogger(log_path, LOG_COLS, fake_app)  # type: ignore[arg-type]
    assert log_path.exists()
    assert logger.stored_data == []

    buy_tx = Transaction(action="buy", id_=123, name="Lightning Bolt", price_each=0.003, qty=4, vendor="goat", holdings=4)
    sell_tx = Transaction(action="sell", id_=123, name="Lightning Bolt", price_each=0.007, net_each=0.004, qty=2, vendor="goat", holdings=2)
    logger.add_transaction(buy_tx)
    logger.add_transaction(sell_tx)
    logger.close()

    ws = logger.ws
    assert ws.range((1, 1), (1, len(LOG_COLS))).value == LOG_COLS
    assert ws.cells(2, 3).value == 123
    assert ws.cells(2, 6).value == pytest.approx(0.012)
    assert ws.cells(3, 2).value == "sell"
    assert ws.cells(3, 6).value == pytest.approx(0.014)
    assert ws.cells(3, 8).value == pytest.approx(0.008)


def test_logger_loads_existing_workbook_and_finds_last_transaction(tmp_path: Path, fake_app: FakeApp) -> None:
    """Verify Logger reads existing log file and returns the most recent transaction for a card id."""
    log_path = tmp_path / "existing_log.xlsx"
    wb = fake_app.books.open(log_path)
    ws = wb.sheets[0]
    ws.range((1, 1)).value = LOG_COLS
    ws.range((2, 1)).value = [
        [datetime.date(2026, 1, 20), "buy", 123, "Lightning Bolt", 0.002, 0.004, None, None, 2, "goat", 2, "old_123"],
        [datetime.date(2026, 1, 21), "buy", 456, "Brainstorm", 0.010, 0.010, None, None, 1, "goat", 1, "456"],
        [datetime.date(2026, 1, 22), "sell", 123, "Lightning Bolt", 0.005, 0.005, 0.003, 0.003, 1, "goat", 1, "new_123"],
    ]

    logger = TransactionLogger(log_path, LOG_COLS, fake_app)  # type: ignore[arg-type]

    assert len(logger.stored_data) == 3

    # Test most recent card with multiple entries
    last_tx_123 = logger.get_last_transaction(123)
    assert last_tx_123 is not None
    assert last_tx_123.notes == "new_123"

    # Test tx that isn't the most recent
    tx_456 = logger.get_last_transaction(456)
    assert tx_456 is not None
    assert tx_456.notes == "456"

    # Test non-existent id
    assert logger.get_last_transaction(555) is None
    logger.close()


def test_logger_revises_existing_transaction(tmp_path: Path, fake_app: FakeApp) -> None:
    """Verify revising a transaction overwrites only the latest row for that card and updates its notes."""
    log_path = tmp_path / "revise_log.xlsx"
    wb = fake_app.books.open(log_path)
    ws = wb.sheets[0]
    ws.range((1, 1)).value = LOG_COLS
    ws.range((2, 1)).value = [
        [datetime.date(2026, 10, 20), "sell", 123, "Lightning Bolt", 0.008, 0.016, 0.005, 0.010, 2, "goat", 0, "older bolt tx"],
        [datetime.date(2026, 10, 22), "buy", 123, "Lightning Bolt", 0.004, 0.016, None, None, 4, "goat", 4, "old note"],
        [datetime.date(2026, 10, 23), "buy", 456, "Brainstorm", 0.010, 0.020, None, None, 2, "goat", 2, "other card"],
    ]

    logger = TransactionLogger(log_path, LOG_COLS, fake_app)  # type: ignore[arg-type]
    old_tx = logger.get_last_transaction(123)
    assert old_tx is not None
    assert old_tx.action == "buy"

    revised_tx = old_tx.model_copy(
        update={
            "action": "revise",
            "old_action": "buy",
            "old_cmd": "b4;0.004",
            "price_each": 0.003,
            "qty": 2,
            "holdings": 2,
        }
    )
    logger.add_transaction(revised_tx)

    # Test revising an id with no associated tx - should be ignored
    unmatched_rev = revised_tx.model_copy(update={"id_": 999, "action": "revise"})
    logger.add_transaction(unmatched_rev)
    logger.close()

    # Old lightning bolt entry (row 2) is unmodified
    assert ws.cells(2, 2).value == "sell"
    assert ws.cells(2, 5).value == 0.008
    assert ws.cells(2, 9).value == 2
    assert ws.cells(2, 12).value == "older bolt tx"

    # Newest lightning bolt entry (row 3) is revised
    assert ws.cells(3, 2).value == "buy"
    assert ws.cells(3, 5).value == 0.003
    assert ws.cells(3, 6).value == pytest.approx(0.006)
    assert ws.cells(3, 9).value == 2
    assert ws.cells(3, 12).value == f"old note - Revised {CONFIG.today} - original cmd: b4;0.004"

    # No other entries are revised
    assert ws.cells(4, 3).value == 456
    assert ws.cells(4, 12).value == "other card"


def test_logger_rejects_operations_after_close(tmp_path: Path, fake_app: FakeApp) -> None:
    """Verify calling Logger methods after close() raises RuntimeError to protect the integrity of the log file."""
    log_path = tmp_path / "closed_log.xlsx"
    logger = TransactionLogger(log_path, LOG_COLS, fake_app) # type: ignore[arg-type]
    logger.close()

    with pytest.raises(RuntimeError, match="The logger is closed"):
        logger.get_last_transaction(1)

    with pytest.raises(RuntimeError, match="The logger is closed"):
        logger.close()
