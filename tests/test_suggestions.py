from pathlib import Path
import sqlite3
from typing import Any

import orjson as json
import pytest

from mtgo_sheets.config import CONFIG, XlsxFile
from mtgo_sheets.tracker import get_suggested_cards, parse_suggestion_opts_from_sheet, update_blacklist
from mtgo_sheets.updater import update_db_with_prices, update_defs_from_zip
from tests.helpers import make_card_def
from tests.xlsx_fakes import FakeSheet

# ID numbering scheme (ABCD):
# A:    Set Legality (1: Standard, 2: Modern, 3: Eternal, 9: Misc)
# B/C:  Unique Card Identifier 00-99
# D:    Printing Info (0: base nonfoil, 1: base foil, 2: non-base nonfoil, 3: non-base foil)

@pytest.fixture
def seeded_db(db_conn: sqlite3.Connection, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> sqlite3.Connection:
    """Seed representative cards and set legalities for suggestion tests."""
    legality_file = tmp_path / "legalities.json"
    legality_file.write_bytes(json.dumps({"modern": ["fdn", "mh1", "mh2", "prm"], "standard": ["fdn"]}))
    monkeypatch.setattr(CONFIG.paths, "set_legalities", legality_file)

    defs = [
        # 1xxx: standard legal
        make_card_def(id_=1010, name="1010_std", set_code="fdn"),
        make_card_def(id_=1011, name="1011_std_foil", set_code="fdn", foil=True),

        # 2xxx: modern legal
        make_card_def(id_=2010, name="2010_mod", set_code="mh1"),
        make_card_def(
            id_=2012, name="2012_mod_showcase", set_code="mh1",
            base_printing=False, variant_data={"treatment": "showcase"},
        ),

        make_card_def(id_=2910, name="29xx_dup", set_code="mh1"),
        make_card_def(id_=2920, name="29xx_dup", set_code="mh2"),
        make_card_def(id_=2921, name="29xx_dup", set_code="mh2", foil=True),
        make_card_def(id_=2922, name="29xx_dup", set_code="mh2", base_printing=False),
        make_card_def(id_=2923, name="29xx_dup", set_code="mh2", foil=True, base_printing=False),

        # 3xxx: Eternal legal
        make_card_def(id_=3010, name="3010_ete", set_code="leb"),

        # 9xxx: Misc (should always be excluded)
        make_card_def(id_=9010, name="9010_prm", set_code="prm"),
        make_card_def(id_=9020, name="MH2 Booster Pack", set_code="mh2"),
        make_card_def(id_=9030, name="9030_no_prices", set_code="mh1"),

    ]
    update_defs_from_zip(db_conn, defs)

    # Price characteristics for modern cards used in tests:
    # ID    age min     max     sell    dist    1d      2d
    # 2010  3   0.010   0.020   0.015   1.5     -25%    +0.005
    # 2910  4   0.002   0.006   0.005   2.5     -16.7%  +0.001
    # 2920  4   0.004   0.008   0.008   2.0     0%      +0.002
    prices = {
        1010: [0.010, 0.012, 0.015, 0.014],
        1011: [0.0, 0.005, 0.008, 0.006],
        2010: [0.0, 0.010, 0.020, 0.015],
        2012: [0.0, 0.020, 0.030, 0.025],
        3010: [0.050, 0.055, 0.060, 0.058],
        9010: [0.005, 0.005, 0.005, 0.005],
        9020: [0.040, 0.040, 0.040, 0.040],
        2910: [0.002, 0.004, 0.006, 0.005],
        2920: [0.004, 0.006, 0.008, 0.008],
        2921: [0.010, 0.012, 0.014, 0.014],
        2922: [0.004, 0.006, 0.008, 0.008],
        2923: [0.015, 0.018, 0.020, 0.020],
    }
    update_db_with_prices(db_conn, prices, total_days=4)
    return db_conn


def _base_opts(**overrides: Any) -> dict[str, Any]:
    opts: dict[str, Any] = {
        "skip_names": set(),
        "modern_only": True,
        "allow_standard": False,
        "price": None,
        "age": None,
        "lowest_max": None,
        "dist_from_min": None,
        "price_change": None,
    }
    opts.update(overrides)
    return opts


def test_suggestions_legality_blacklist_variant_filters(seeded_db: sqlite3.Connection, fake_app) -> None:
    """Verify set legality, blacklisted sets/names, foil, base_printing, and duplicate name filters."""
    # Default opts: modern, nonfoil, base printing
    rows = get_suggested_cards(seeded_db, fake_app, _base_opts(), suggested_last_time=[])
    assert {r["id"] for r in rows} == {2010, 2910, 2920}

    # Blacklist (all cards with the specified names)
    rows_skipped = get_suggested_cards(seeded_db, fake_app, _base_opts(skip_names={"29xx_dup"}), suggested_last_time=[])
    assert [r["id"] for r in rows_skipped] == [2010]

    # Allow standard
    rows_std = get_suggested_cards(seeded_db, fake_app, _base_opts(allow_standard=True), suggested_last_time=[])
    assert {r["id"] for r in rows_std} == {1010, 2010, 2910, 2920}

    # Allow non-modern only
    rows_eternal_no_std = get_suggested_cards(
        seeded_db, fake_app, _base_opts(modern_only=False, allow_standard=False), suggested_last_time=[]
    )
    assert {r["id"] for r in rows_eternal_no_std} == {2010, 2910, 2920, 3010}

    # Allow any legality
    rows_all_legal = get_suggested_cards(
        seeded_db, fake_app, _base_opts(modern_only=False, allow_standard=True), suggested_last_time=[]
    )
    assert {r["id"] for r in rows_all_legal} == {1010, 2010, 2910, 2920, 3010}


def test_suggestions_numerical_filters(
    seeded_db: sqlite3.Connection, fake_app
) -> None:
    """Verify numerical suggestion filters (based on age or price stats)."""

    # Price range filter is agnostic of operand order
    price_filtered = get_suggested_cards(seeded_db, fake_app, _base_opts(price="0.007;0.004"), suggested_last_time=[])
    assert [r["id"] for r in price_filtered] == [2910]

    # Age
    age_filtered = get_suggested_cards(seeded_db, fake_app, _base_opts(age=4), suggested_last_time=[])
    assert {r["id"] for r in age_filtered} == {2910, 2920}

    # Lowest max
    max_filtered = get_suggested_cards(seeded_db, fake_app, _base_opts(lowest_max=0.015), suggested_last_time=[])
    assert [r["id"] for r in max_filtered] == [2010]

    # Dist from min
    dist_filtered = get_suggested_cards(seeded_db, fake_app, _base_opts(dist_from_min=1.8), suggested_last_time=[])
    assert [r["id"] for r in dist_filtered] == [2010]


def test_suggestions_price_change_targets(seeded_db: sqlite3.Connection, fake_app) -> None:
    """Verify % and absolute $ price_change target filters."""

    # 1 day % drop
    drop_rows = get_suggested_cards(seeded_db, fake_app, _base_opts(price_change="1;-20%"), suggested_last_time=[])
    assert [r["id"] for r in drop_rows] == [2010]

    # Multi-day $ and % with a period beyond the price history start date (ignored)
    compound_rows = get_suggested_cards(
        seeded_db, fake_app, _base_opts(price_change="2;0.004 | 1;-20% | 30;500%"), suggested_last_time=[]
    )
    assert [r["id"] for r in compound_rows] == [2010]


def test_suggestions_sorting(
    seeded_db: sqlite3.Connection, fake_app
) -> None:
    """Verify suggestions are sorted with new at the top, and then by the config's sort_fn."""

    # New suggestions before existing ones
    rows = get_suggested_cards(seeded_db, fake_app, _base_opts(allow_standard=True), suggested_last_time=[1010, 2010])
    assert [r["id"] for r in rows] == [2910, 2920, 1010, 2010]

    # Sort_fn from config is applied
    # Names sorted alphabetically according to mock0.toml' sort_fn
    sort_fn = CONFIG.xlsx["main"].sort_fn
    assert sort_fn is not None
    sorted_by_cfg = sorted(rows, key=sort_fn)
    assert [r["id"] for r in sorted_by_cfg] == [1010, 2010, 2910, 2920]


def test_external_suggestions_source_bypassses_filters(seeded_db: sqlite3.Connection, tmp_path: Path, fake_app) -> None:
    """Verify external_suggestions_path is the source of truth - that exactly its IDs are considered."""
    missing_path = tmp_path / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        get_suggested_cards(seeded_db, fake_app, _base_opts(), [], external_suggestions_path=missing_path)

    ext_path = tmp_path / "ext_suggestions.xlsx"
    wb = fake_app.books.open(ext_path)
    ws = wb.sheets[0]

    # Verify columns - Missing ID column raises ValueError
    ws.range((1, 1)).value = [["name", "notes"]]
    with pytest.raises(ValueError, match="missing a header row or the 'id' column"):
        get_suggested_cards(seeded_db, fake_app, _base_opts(), [], external_suggestions_path=ext_path)

    ws.used_range.clear_contents()
    ws.range((1, 1)).value = [
        ["id", "notes", "tag"],
        [2921, "foil base", "foil-spec"],
        [2922, "non-base target", "variant"],
        [2923, "foil variant", "foil-var"],
    ]

    rows = get_suggested_cards(seeded_db, fake_app, _base_opts(), [], external_suggestions_path=ext_path)
    by_id = {r["id"]: r for r in rows}
    assert set(by_id.keys()) == {2921, 2922, 2923}
    assert by_id[2921]["foil"] is True
    assert by_id[2921]["notes"] == "foil base"
    assert by_id[2922]["foil"] is False
    assert by_id[2922]["tag"] == "variant"
    assert by_id[2923]["foil"] is True
    assert by_id[2923]["tag"] == "foil-var"


def test_parse_and_initialize_sheet_for_opts(xlsx_cfg: XlsxFile) -> None:
    """Verify suggestion options table and blacklist cells are initialized/parsed."""
    ws = FakeSheet()

    # Initialize suggestion opts
    initial_opts = parse_suggestion_opts_from_sheet(ws)  # type: ignore[arg-type]
    assert ws.cells(xlsx_cfg.suggestions_top_row - 1, xlsx_cfg.suggestions_left_col).value == "Suggestion Args"
    assert all(v is None for v in initial_opts.values())

    # Set opt values on the sheet and parse
    ws.cells(xlsx_cfg.suggestions_top_row, xlsx_cfg.suggestions_left_col + 1).value = True
    ws.cells(xlsx_cfg.suggestions_top_row + 3, xlsx_cfg.suggestions_left_col + 1).value = "0.005;0.020"
    parsed_opts = parse_suggestion_opts_from_sheet(ws)  # type: ignore[arg-type]
    assert parsed_opts["enabled"] is True
    assert parsed_opts["price"] == "0.005;0.020"

    # Seed blacklist file
    assert xlsx_cfg.suggestion_card_blacklist is not None
    xlsx_cfg.suggestion_card_blacklist.write_bytes(json.dumps(["Old Blacklisted Card"]))

    # Add to blacklist by typing the name
    ws.cells(xlsx_cfg.blacklist_row, xlsx_cfg.blacklist_col).value = "Grief; Fury"

    # Add to blacklist by checking the blacklist column
    card_rows = [
        {"id": 2010, "name": "2010_Mod", "add_to_blacklist": False},
        {"id": 2910, "name": "29xx_Dup", "add_to_blacklist": True},
    ]
    blacklist, remaining = update_blacklist(ws, card_rows)  # type: ignore[arg-type]

    assert set(blacklist) == {"Grief", "Fury", "29xx_Dup", "Old Blacklisted Card"}
    assert [r["id"] for r in remaining] == [2010]
    assert json.loads(xlsx_cfg.suggestion_card_blacklist.read_bytes()) == blacklist
