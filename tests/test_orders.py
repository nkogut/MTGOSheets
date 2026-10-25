"""
Unit tests covering the following functionality in tracker.py
- Vendor/price selection and processing for new entries
- Alert triggering
- Order processing (generated transactions and updates to the entry)
    - Note: transaction storage is not yet tested

Usage: pytest tests/test_orders.py
"""


from pathlib import Path
from statistics import stdev

import orjson as json
import pytest

from mtgo_sheets.config import CONFIG
from mtgo_sheets.models import CardDef, Transaction
from mtgo_sheets.tracker import (
    DEFAULT_VENDOR,
    WatchlistEntry,
    parse_order,
    process_arb_order,
    process_revise_order,
)


def test_vendor_price_and_stats_selection(
    sample_card_def: CardDef,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Verify optimal vendors and prices are selected, and correct price stats are computed."""
    ext_file = tmp_path / "ext_prices.json"
    ext_data = {
        "goat": {"123": {"buy": 0.004, "sell": 0.008}},
        "ext_vendor_1": {"123": {"buy": 0.006, "sell": 0.010}},
        "ext_vendor_2": {"123": {"buy": 0.002, "sell": 0.005}},
    }
    ext_file.write_bytes(json.dumps(ext_data))
    monkeypatch.setattr(CONFIG.paths, "external_prices", ext_file)

    # Leading 0s should not be included in statistics
    prices = [0.0, 0.002, 0.002, 0.003, 0.003, 0.004, 0.004, 0.005, 0.006, 0.006, 0.007]
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 1, "purchase_price": 0.004},
        card_def=sample_card_def,
        price_hist=prices,
        ext_prices=ext_data,
    )

    assert entry.age == 10
    assert entry.sell_to == "ext_vendor_1"
    assert entry.buy_price == 0.006
    assert entry.buy_from == "ext_vendor_2"
    assert entry.sell_price == 0.005
    assert entry.price_data["min"] == min(prices[1:])
    assert entry.price_data["max"] == max(prices)
    assert entry.price_data["7d_delta"] == prices[-1] - prices[-7]


def test_fallback_to_non_external_vendors(sample_card_def: CardDef) -> None:
    """Test fallback to DEFAULT_VENDOR when external prices are enabled but not available."""
    prices = [0.005]
    ext_prices = {"ext_vendor_1": {"123": {"buy": -1.0, "sell": 999.0}}}

    entry = WatchlistEntry(
        vals={"id": 123, "qty": 1},
        card_def=sample_card_def,
        price_hist=prices,
        ext_prices=ext_prices,
    )

    assert entry.sell_to == DEFAULT_VENDOR
    assert entry.buy_price == 0.0
    assert entry.buy_from == DEFAULT_VENDOR
    assert entry.sell_price == 0.005


def test_skip_first_n_days(sample_card_def: CardDef, monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify that skip_first_n_days is applied to all stats"""
    monkeypatch.setattr(CONFIG, "skip_first_n_days", 3)

    prices = [0.10, 0.09, 0.08, 0.003, 0.004, 0.005, 0.006, 0.007]
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 1, "purchase_price": 0.005},
        card_def=sample_card_def,
        price_hist=prices,
    )

    assert entry.age == 8
    assert entry.price_data["min"] == 0.003
    assert entry.price_data["max"] == 0.007
    assert entry.price_data["today_price"] == 0.007


def test_zero_volatility(sample_card_def: CardDef) -> None:
    """Verify 0.0 volatility when daily returns have zero variance."""
    flat_prices = [0.005] * 7
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 1},
        card_def=sample_card_def,
        price_hist=flat_prices,
    )
    assert entry.price_data["7d_volatility"] == 0.0


def test_nonzero_price_volatility(sample_card_def: CardDef) -> None:
    """Verify volatility calculation is correct and uses the correct days."""
    # 8 daily prices - the first should be ignored
    prices = [100, 0.004, 0.005, 0.004, 0.006, 0.005, 0.007, 0.008]
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 1},
        card_def=sample_card_def,
        price_hist=prices,
    )

    last_7 = prices[-7:]
    expected_returns = [(last_7[i] - last_7[i - 1]) / last_7[i - 1] for i in range(1, len(last_7))]
    assert entry.price_data["7d_volatility"] == pytest.approx(stdev(expected_returns))


def test_short_history_card_handling(sample_card_def: CardDef) -> None:
    """Verify 7d stats are 'N/A' when card age is shorter than 7."""
    prices = [0.002, 0.002, 0.003]
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 1},
        card_def=sample_card_def,
        price_hist=prices,
    )
    assert entry.age == 3
    assert entry.price_data["today_price"] == 0.003
    assert entry.price_data["7d_delta"] == "N/A"
    assert entry.price_data["7d_volatility"] == "N/A"


def test_alert_triggers(sample_card_def: CardDef) -> None:
    """Verify that arb, dropped, and rose alerts trigger with correct priority."""
    prices = [0.25]

    # Arbitrage alert should have priority over buy/sell
    arb_prices = {"goat": {"123": {"buy": 0.08, "sell": 0.04}}}
    arb_entry = WatchlistEntry(
        vals={"id": 123, "qty": 0, "drop": 0.10, "rise": 0.02},
        card_def=sample_card_def,
        price_hist=prices,
        ext_prices=arb_prices,
    )
    assert arb_entry.hit == "arb"

    norm_prices = {"goat": {"123": {"buy": 0.20, "sell": 0.20}}}
    drop_entry = WatchlistEntry(
        vals={"id": 123, "qty": 0, "drop": 0.201},
        card_def=sample_card_def,
        price_hist=prices,
        ext_prices=norm_prices,
    )
    assert drop_entry.hit == "dropped"

    rise_entry = WatchlistEntry(
        vals={"id": 123, "qty": 1, "rise": 0.199},
        card_def=sample_card_def,
        price_hist=prices,
        ext_prices=norm_prices,
    )
    assert rise_entry.hit == "rose"

    no_alert_entry = WatchlistEntry(
        vals={"id": 123, "qty": 1, "drop": 0.15, "rise": 0.40},
        card_def=sample_card_def,
        price_hist=prices,
        ext_prices=norm_prices,
    )
    assert no_alert_entry.hit == ""


# Test orders

def test_buy_order(sample_card_def: CardDef) -> None:
    """Test the generated tx and all side effects of a buy order (cost basis and threshold updated if appropriate)."""
    entry = WatchlistEntry(
        vals={"id": 123, "qty": -1, "purchase_price": 0.002, "drop": 0.004},
        card_def=sample_card_def,
        price_hist=[0.006],
    )
    assert entry.active is True

    # Buying above drop_thresh (0.006 > 0.004) does not disable the alert
    cmd1 = "b2" # @ $0.006
    _, qty, prices = parse_order(cmd1)
    price = prices[0] if prices else None
    tx_above = entry.buy(buy_qty=qty, price=price)
    assert entry.qty == 2
    assert entry.purchase_price == 0.006
    assert entry.drop_thresh == 0.004
    assert tx_above.price_each == 0.006
    assert tx_above.holdings == 2

    # Buying at or below drop_thresh (0.002 <= 0.004) satisfies the target and the alert threshold should be removed
    cmd2 = "b2;0.002"
    _, qty, prices = parse_order(cmd2)
    tx_below = entry.buy(buy_qty=qty, price=prices[0])
    assert entry.qty == 4
    assert entry.purchase_price == 0.004
    assert entry.drop_thresh is None
    assert tx_below.holdings == 4


def test_sell_order(sample_card_def: CardDef) -> None:
    """Test the generated tx and all side effects of a sell order (cost basis and threshold updated if appropriate)."""
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 4, "purchase_price": 0.003, "rise": 0.008},
        card_def=sample_card_def,
        price_hist=[0.005],
        ext_prices={"goat": {"123": {"buy": 0.007, "sell": 0.009}}},
    )

    # Inventory remains, so the alert threshold is not changed
    cmd1 = "s2" # @ $0.007
    _, qty, prices = parse_order(cmd1)
    price = prices[0] if prices else None
    partial_tx = entry.sell(sell_qty=qty, price=price)
    assert entry.qty == 2
    assert entry.rise_thresh == 0.008
    assert partial_tx.price_each == 0.007
    assert partial_tx.net_each == pytest.approx(0.004)

    # All inventory is sold, so the alert threshold is removed
    cmd2 = "s2;0.008"
    _, qty, prices = parse_order(cmd2)
    full_tx = entry.sell(sell_qty=qty, price=prices[0])
    assert entry.qty == 0
    assert entry.rise_thresh is None
    assert full_tx.holdings == 0

    # Selling more copies than are known to be held should clear the alert
    cmd3 = "s1;0.008"
    _, qty, prices = parse_order(cmd3)
    oversell_tx = entry.sell(sell_qty=qty, price=prices[0])
    assert entry.qty == -1
    assert entry.rise_thresh is None
    assert oversell_tx.holdings == -1


def test_arb_order(sample_card_def: CardDef) -> None:
    """Test the generated tx and all side effects of an arbitrage order."""
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 1},
        card_def=sample_card_def,
        price_hist=[0.004],
        ext_prices={"goat": {"123": {"buy": 0.009, "sell": 0.004}}},
    )

    # Profit is calculated based on the difference of both given prices, regardless of order
    cmd1 = "a2;0.005;0.002"
    _, qty, prices = parse_order(cmd1)
    explicit_tx = process_arb_order(entry, qty=qty, prices=prices)
    assert explicit_tx.action == "arb"
    assert explicit_tx.price_each == 0.002
    assert explicit_tx.net_each == 0.003

    # Profit is calculated based on the difference of both given prices, regardless of order
    cmd2 = "a2;0.002;0.005"
    _, qty, prices = parse_order(cmd2)
    explicit_tx = process_arb_order(entry, qty=qty, prices=prices)
    assert explicit_tx.action == "arb"
    assert explicit_tx.price_each == 0.002
    assert explicit_tx.net_each == 0.003

    # Profit is calculated using today's prices ($0.009 - $0.004)
    cmd3 = "a3"
    _, qty, prices = parse_order(cmd3)
    default_tx = process_arb_order(entry, qty=qty, prices=prices)
    assert default_tx.price_each == 0.004
    assert default_tx.net_each == pytest.approx(0.005)


def test_revise_buy_order(sample_card_def: CardDef) -> None:
    """Test the generated tx and all side effects of revising a buy order (cost basis and threshold updated if appropriate)."""
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 2, "purchase_price": 0.003},
        card_def=sample_card_def,
        price_hist=[0.004],
    )
    old_buy = entry.buy(buy_qty=2, price=0.007)
    assert entry.qty == 4
    assert entry.purchase_price == 0.005

    # Reduced qty, same price: holdings and cost basis update, drop_thresh stays None
    cmd1 = "r1;0.006"
    _, qty, prices = parse_order(cmd1)
    rev_partial = process_revise_order(entry, qty=qty, prices=prices, old=old_buy)
    assert rev_partial.action == "revise"
    assert rev_partial.old_action == "buy"
    assert entry.qty == 3
    assert entry.purchase_price == pytest.approx(0.004)
    assert entry.drop_thresh is None


    entry_full = WatchlistEntry(
        vals={"id": 123, "qty": 2, "purchase_price": 0.003},
        card_def=sample_card_def,
        price_hist=[0.004],
    )
    old_buy_full = entry_full.buy(buy_qty=2, price=0.007)

    # Same qty, new price: update cost basis, restore drop_thresh to the previous price paid
    cmd2 = "r2;0.005"
    _, qty, prices = parse_order(cmd2)
    process_revise_order(entry_full, qty=qty, prices=prices, old=old_buy_full)
    assert entry_full.qty == 4
    assert entry_full.purchase_price == pytest.approx(0.004)
    assert entry_full.drop_thresh == 0.007


def test_revise_sell_order(sample_card_def: CardDef) -> None:
    """Test the generated tx and all side effects of revising a sell order (cost basis and threshold updated if appropriate)."""
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 4, "purchase_price": 0.003, "rise": 0.008},
        card_def=sample_card_def,
        price_hist=[0.004],
    )
    old_sell = entry.sell(sell_qty=4, price=0.009)
    assert entry.qty == 0
    assert entry.rise_thresh is None

    # Changed qty, changed price: restore rise_thresh to the previous sell price and update cost basis
    cmd1 = "r1;0.007"
    _, qty, prices = parse_order(cmd1)
    rev_sell = process_revise_order(entry, qty=qty, prices=prices, old=old_sell)
    assert rev_sell.action == "revise"
    assert rev_sell.old_action == "sell"
    assert entry.qty == 3
    assert entry.rise_thresh == 0.009
    assert rev_sell.net_each == pytest.approx(0.004)
    assert rev_sell.holdings == 3


def test_revise_arb_order(sample_card_def: CardDef) -> None:
    """Test the generated tx and all side effects of revising an arbitrage order."""
    entry = WatchlistEntry(
        vals={"id": 123, "qty": 2},
        card_def=sample_card_def,
        price_hist=[0.004],
    )
    old_arb = Transaction(
        action="arb",
        id_=123,
        name="Lightning Bolt",
        price_each=0.002,
        net_each=0.003,
        qty=2,
        vendor="goat -> goat",
        holdings=2,
    )

    # Different qty and prices: update calculated prices
    cmd1 = "r1;0.003;0.007"
    _, qty, prices = parse_order(cmd1)
    rev_explicit = process_revise_order(entry, qty=qty, prices=prices, old=old_arb)
    assert rev_explicit.action == "revise"
    assert rev_explicit.price_each == 0.003
    assert rev_explicit.net_each == pytest.approx(0.004)
    assert rev_explicit.qty == 1
