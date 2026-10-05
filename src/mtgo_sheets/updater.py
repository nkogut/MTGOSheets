import sqlite3
import tempfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import orjson as json
import requests

from mtgo_sheets.config import CONFIG, XlsxFile
from mtgo_sheets.db import get_db_connection
from mtgo_sheets.models import CardDef
from mtgo_sheets.tracker import dump_ids


def fetch_defs_data() -> list[CardDef]:
    """Download and parse card definitions. Return a CardDef for every ID available."""
    url = "https://www.goatbots.com/download/prices/card-definitions.zip"
    res = requests.get(url)
    res.raise_for_status()

    with tempfile.TemporaryDirectory() as temp_dir:
        zip_dir = Path(temp_dir)
        with ZipFile(BytesIO(res.content)) as zip_file:
            zip_file.extractall(zip_dir)

        defs_path = zip_dir / "card-definitions.txt"
        all_defs = json.loads(defs_path.read_bytes())

        return [
            CardDef(
                id_=int(id_),
                name=attrs["name"],
                set_code=attrs["cardset"],
                foil=bool(attrs["foil"]),
                ver=attrs["version"] or "",
                rarity=attrs["rarity"] or "",
            )
            for id_, attrs in all_defs.items()
        ]

def update_defs_from_zip(conn: sqlite3.Connection, defs: list[CardDef]) -> None:
    """Initialize the definitions of new card printings."""
    conn.executemany(
        """
        INSERT INTO defs (id, name, set_, foil, ver, rarity)
        VALUES (:id_, :name, :set_code, :foil, :ver, :rarity)
        ON CONFLICT(id) DO NOTHING
        """,
        [card.model_dump() for card in defs],
    )

def update_legal_sets() -> None:
    """Get up-to-date set legality for each format."""
    set_legalities = {
        "modern": get_modern_sets(),
        "standard": get_standard_sets(),
    }

    CONFIG.paths.set_legalities.write_bytes(json.dumps(set_legalities))

def get_modern_sets() -> list[str]:
    """Return all modern legal set codes."""
    headers = {
        "User-Agent": "MTGSetLegalityApp/1.0",
        "Accept": "*/*",
    }
    modern_legality_req = "https://api.scryfall.com/sets"
    res = requests.get(modern_legality_req, headers=headers)
    res.raise_for_status()
    modern_data = res.json()["data"]

    allowed_set_types = ["core", "expansion", "draft_innovation"]
    name_blacklist = ["commander"]

    sets = [
        s for s in modern_data if "mtgo_code" in s and
        s["set_type"] in allowed_set_types
    ]
    release_8ed = next(s["released_at"] for s in sets if s["mtgo_code"] == "8ed")
    return [
        s["mtgo_code"] for s in sets
        if s["released_at"] >= release_8ed and
        all(name not in s["name"].lower() for name in name_blacklist)
    ]

def get_standard_sets() -> list[str]:
    """Return all standard legal set codes."""
    standard_legality_req = "https://whatsinstandard.com/api/v6/standard.json"
    res = requests.get(standard_legality_req)
    res.raise_for_status()

    standard_data = res.json()["sets"]
    return [
        s["code"].lower() for s in standard_data if
        s["code"] is not None and (
        s["exitDate"]["exact"] is None or
        datetime.strptime(standard_data[0]["exitDate"]["exact"], '%Y-%m-%dT%H:%M:%S.%f').date() > CONFIG.today)
    ]

def get_price_history_url(year: int | None) -> str:
    """Return the year-long goatbots price history zip url for the year (if provided) or the url for just today."""
    if year is None:
        return "https://www.goatbots.com/download/prices/price-history.zip"
    if year < CONFIG.start_date.year:
        raise ValueError(f"Year must be >= {CONFIG.start_date.year}")

    return f"https://www.goatbots.com/download/prices/price-history-{year}.zip"

def fetch_missing_price_history(year: int | None) -> None:
    """Update price lists for all dates since the last update."""
    print("Missing older prices. Retrieving them first.")

    if year is not None:
        # Go back one year first, then retry current year
        update_prices(year - 1)
        update_prices(year)
    else:
        # Get all prices up to today's, then get today's
        update_prices(CONFIG.today.year)
        update_prices()

def update_db_with_prices(conn: sqlite3.Connection, new_prices: dict[int, list[float]], total_days: int) -> None:
    """Update all ids in the database with the prices from this update."""
    if not new_prices:
        return

    all_prices = {}
    price_entries = conn.execute("SELECT id, price_list FROM prices").fetchall()
    existing_prices = {id_: json.loads(price_list) for id_, price_list in price_entries}

    for id_, new in new_prices.items():
        old = existing_prices.get(id_, [])
        padding = [0] * (total_days - len(old) - len(new))
        all_prices[id_] = old + padding + new

    with conn:
        conn.executemany(
            """
            INSERT INTO prices (id, price_list)
            VALUES (?, ?)
            ON CONFLICT(id) DO UPDATE
            SET price_list = excluded.price_list
            """,
            [(id_, json.dumps(prices)) for id_, prices in all_prices.items()],
        )

def update_prices(year: int | None = None) -> None:
    """Get all price updates since the last update and append them to each id's price history."""
    # Get the number of dates already in the database by selecting an arbitrary price list
    with get_db_connection() as conn:
        price_list = conn.execute("SELECT price_list FROM prices LIMIT 1").fetchone()

    if price_list is not None:
        days_recorded = len(json.loads(price_list[0]))
    else:
        defs_to_insert = fetch_defs_data()
        with get_db_connection() as conn:
            update_defs_from_zip(conn, defs_to_insert)

        days_recorded = 0

    next_date = CONFIG.start_date + timedelta(days=days_recorded + 1)

    # The most recent price sheet is always 1 day behind today
    if next_date >= CONFIG.today:
        print("All dates have already been added")
        return

    # Download and process missing days
    url = get_price_history_url(year)
    print(f"Getting newest prices from {url}")
    res = requests.get(url)
    res.raise_for_status()

    new_prices = defaultdict(list)
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_dir = Path(temp_dir)
        with ZipFile(BytesIO(res.content)) as zip_file:
            zip_file.extractall(zip_dir)

        file_paths = sorted(zip_dir.iterdir(), key=lambda p: p.name)

        for file_path in file_paths:
            file_date = date.fromisoformat(file_path.stem.removeprefix("price-history-"))

            if (file_date > next_date):
                fetch_missing_price_history(year)
                return

            if file_date < next_date:
                continue

            print(f"Getting prices from {file_path.name}")
            next_date += timedelta(days=1)
            days_recorded += 1

            contents = json.loads(file_path.read_bytes())
            for key, price in contents.items():
                new_prices[int(key)].append(price)

    with get_db_connection() as conn:
        update_db_with_prices(conn, new_prices, days_recorded)

    print(f"Successfully updated db with {url}")

def update_external(id_sources: list[XlsxFile]) -> None:
    """Save live prices from external sources."""
    try:
        import mtgo_sheets.external_price_updater as ext_updater  # noqa: PLC0415
    except ImportError as e:
        raise NotImplementedError(
            "External price updating is disabled in the public release. "
            "Implement mtgo_sheets.external_price_updater.update_external_prices()."
        ) from e

    ext_wishlist_path = dump_ids(id_sources)
    try:
        ext_prices = ext_updater.update_external_prices(str(ext_wishlist_path))

        CONFIG.paths.external_prices.write_bytes(json.dumps(ext_prices))

        hist_path = CONFIG.paths.external_prices_hist
        try:
            ext_prices_hist = json.loads(hist_path.read_bytes())
        except FileNotFoundError:
            ext_prices_hist = {}

        # Only storing the final snapshot taken per day
        ext_prices_hist[CONFIG.today.strftime("%Y-%m-%d")] = ext_prices
        hist_path.write_bytes(json.dumps(ext_prices_hist))

    except Exception as e:
        print(f"\n\nWARNING: Failed to update external prices.\nExternal prices may be stale.\n{e}\n\n")

if __name__ == "__main__":
    update_prices()
    update_legal_sets()

    files = [file for file in CONFIG.xlsx.values()  if not file.ignore_for_ext_prices]
    update_external(files)
