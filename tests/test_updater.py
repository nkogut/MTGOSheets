import datetime
import gzip
import orjson as json
import pytest
import sqlite3

from io import BytesIO
from pathlib import Path
from unittest import mock
from zipfile import ZipFile


from mtgo_sheets.config import CONFIG
from mtgo_sheets.updater import (
    collector_sort_key,
    extract_variant_data,
    fetch_defs_data,
    update_db_with_prices,
    update_defs_from_zip,
    update_legal_sets,
    update_prices,
    get_price_history_url,
)
from mtgo_sheets.util import clean_collector_no


def _make_price_zip(files: dict[str, dict[str, float]]) -> bytes:
    """Mocks a daily price history zip (in memory) using file={filename: {id_str: price}}

    Usage: zip = _make_price_zip({"price-history-2026-01-02.txt": {"123": 0.050}})"""
    buf = BytesIO()
    with ZipFile(buf, "w") as zf:
        for name, data in files.items():
            zf.writestr(name, json.dumps(data))
    return buf.getvalue()

def test_extract_variant_data() -> None:
    """Verify card metadata is extracted."""
    assert clean_collector_no("0042a/350*") == "42a"
    assert clean_collector_no("*001/002") == "1"
    assert clean_collector_no("") == ""

    assert extract_variant_data(None, is_foil=False) == {
        "collector_num": "",
        "treatment": "normal",
        "frame": "regular",
        "finish": "normal",
    }

    scryfall_card = {
        "collector_number": "0199*",
        "frame_effects": ["showcase"],
        "promo_types": [],
        "finishes": ["foil", "etched"],
        "frame": "1997",
        "border_color": "black",
    }
    variant = extract_variant_data(scryfall_card, is_foil=True)
    assert variant == {
        "collector_num": "199",
        "treatment": "showcase",
        "frame": "retro",
        "finish": "etched",
    }


def test_update_db_with_prices_pads_missing_days(db_conn: sqlite3.Connection) -> None:
    """Verify update_db_with_prices maintains exactly 1 price per day by left padding 0s for missing days."""
    update_db_with_prices(db_conn, {100: [1.5, 1.6]}, total_days=2)
    update_db_with_prices(db_conn, {100: [2.0], 200: [5.0]}, total_days=5)

    rows = dict(db_conn.execute("SELECT id, price_list FROM prices").fetchall())
    assert json.loads(rows[100]) == [1.5, 1.6, 0, 0, 2.0]
    assert json.loads(rows[200]) == [0, 0, 0, 0, 5.0]


def test_fetch_defs_retrieves_correct_records(db_conn: sqlite3.Connection) -> None:
    """Tests whether fetch_defs_data accesses and parses correct data from the source API."""
    # Mock downloaded defs file from Goatbots
    zip_buf = BytesIO()
    goat_defs = {
        "123": {"name": "Opt", "cardset": "FDN", "foil": 0, "version": "111", "rarity": "uncommon"},
        "456": {"name": "Opt", "cardset": "FDN", "foil": 0, "version": "222", "rarity": "uncommon"},
        "789": {"name": "Opt", "cardset": "FDN", "foil": 1, "version": "222", "rarity": "uncommon"},
        "790": {"name": "Opt", "cardset": "FDN", "foil": 1, "version": "999", "rarity": "uncommon"},
        "999": {"name": "Unmatched", "cardset": "FDN", "foil": 0, "version": "333", "rarity": "common"},
    }
    with ZipFile(zip_buf, "w") as zf:
        zf.writestr("card-definitions.txt", json.dumps(goat_defs))

    # Mock downloaded bulk data file from Scryfall
    scryfall_lines = [
        json.dumps({
            "mtgo_id": 123, "mtgo_foil_id": 124, "name": "Opt", "set": "fdn", "collector_number": "111",
            "border_color": "borderless", "frame": "2015",
        }),
        json.dumps({
            "mtgo_id": 456, "name": "Opt", "set": "fdn", "collector_number": "222",
            "border_color": "black", "frame": "2015",
        }),
    ]
    gz_buf = BytesIO()
    with gzip.GzipFile(fileobj=gz_buf, mode="wb") as gz:
        gz.write(b"\n".join(scryfall_lines))
    gz_buf.seek(0)

    resp_defs = mock.MagicMock(content=zip_buf.getvalue())
    resp_cat = mock.MagicMock()
    resp_cat.json.return_value = {
        "data": [{"type": "default_cards", "jsonl_download_uri": "https://scryfall.testapi/cards.jsonl.gz"}]
    }
    resp_stream = mock.MagicMock(raw=BytesIO(gz_buf.getvalue()))

    with mock.patch("mtgo_sheets.updater.requests.get", side_effect=[resp_defs, resp_cat, resp_stream]):
        cards = fetch_defs_data(db_conn)

    by_id = {c.id_: c for c in cards}

    assert by_id[123].base_printing is False
    assert by_id[123].variant_data["treatment"] == "borderless"

    assert by_id[456].base_printing is True
    assert collector_sort_key(by_id[456]) == 222

    assert by_id[789].base_printing is True
    assert by_id[789].variant_data["collector_num"] == "222"

    # Foil falls back to printing attributes of non-foil entry
    assert by_id[790].variant_data["collector_num"] == "111"

    # Card in Goatbot data but not Scryfall data uses available information
    assert by_id[999].variant_data["collector_num"] == "333"

    # Updating defs returns an empty list (all defs already added)
    update_defs_from_zip(db_conn, cards)
    with mock.patch("mtgo_sheets.updater.requests.get", return_value=resp_defs):
        assert fetch_defs_data(db_conn) == []


def test_get_set_legalities(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify set legalities are updated for modern and standard."""
    legality_file = tmp_path / "legalities.json"
    monkeypatch.setattr(CONFIG.paths, "set_legalities", legality_file)

    # Mock Scryfall sets endpoint
    modern_resp = mock.MagicMock()
    modern_resp.json.return_value = {
        "data": [
            # Pre-8th Edition set (not legal)
            {"mtgo_code": "old", "set_type": "core", "released_at": "2001-01-01", "name": "SET_OLD"},

            # Modern legal sets
            {"mtgo_code": "8ed", "set_type": "core", "released_at": "2003-01-01", "name": "SET_8ed"},
            {"mtgo_code": "mod", "set_type": "expansion", "released_at": "2022-01-01", "name": "SET_MODERN"},

            # Commander set (not legal - due to "commander" in name)
            {"mtgo_code": "cmd", "set_type": "draft_innovation", "released_at": "2020-01-01", "name": "SET_COMMANDER"},
        ]
    }

    # Mock WhatsInStandard API (https://whatsinstandard.com/api/v6/standard.json)
    standard_resp = mock.MagicMock()
    standard_resp.json.return_value = {
        "sets": [
            # Future rotation date (legal)
            {"code": "STD_1", "exitDate": {"exact": "2999-01-01T00:00:00.0"}},

            # No rotation date yet (legal)
            {"code": "STD_2", "exitDate": {"exact": None}},

            # Rotated (not legal)
            {"code": "ROTATED", "exitDate": {"exact": "2000-01-01T00:00:00.0"}},

            # Missing set code (not legal)
            {"code": None, "exitDate": {"exact": None}},
        ]
    }

    # Mock requests.get twice:
    # 1st call is to get_modern_sets
    # 2nd call is to get_standard_sets
    with mock.patch("mtgo_sheets.updater.requests.get", side_effect=[modern_resp, standard_resp]):
        update_legal_sets()

    assert json.loads(legality_file.read_bytes()) == {
        "modern": ["8ed", "mod"],
        "standard": ["std_1", "std_2"],
    }


def test_update_missing_prices(db_conn: sqlite3.Connection, monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify update_prices backfills missing dates before adding today's prices."""
    monkeypatch.setattr(CONFIG, "start_date", datetime.date(2025, 1, 1))
    monkeypatch.setattr(CONFIG, "today", datetime.date(2025, 1, 5))

    # Seed 1 day of existing price history (for 2025-01-02)
    update_db_with_prices(db_conn, {123: [0.010]}, total_days=1)

    # Backfill 2025-01-03 and add 2025-01-04
    zip_1 = _make_price_zip({"price-history-2025-01-04.txt": {"123": 0.030}})

    # Skip 2025-01-02 (already added) and add 2025-01-03
    zip_2 = _make_price_zip({
        "price-history-2025-01-02.txt": {"123": 0.010},
        "price-history-2025-01-03.txt": {"123": 0.020},
    })

    resp_zip_1_1 = mock.MagicMock(content=zip_1)
    resp_zip_2 = mock.MagicMock(content=zip_2)
    resp_zip_1_2 = mock.MagicMock(content=zip_1)

    @mock.patch("mtgo_sheets.updater.get_db_connection")
    def _run(mock_get_conn: mock.MagicMock) -> None:
        mock_get_conn.return_value.__enter__.return_value = db_conn
        with mock.patch("mtgo_sheets.updater.requests.get", side_effect=[resp_zip_1_1, resp_zip_2, resp_zip_1_2]):
            update_prices()
            update_prices() # No effect - all prices already added

    _run()

    row = db_conn.execute("SELECT price_list FROM prices WHERE id = 123").fetchone()
    assert json.loads(row[0]) == [0.010, 0.020, 0.030]

    with pytest.raises(ValueError, match="Year must be >="):
        get_price_history_url(2024) # Start date is 2025

