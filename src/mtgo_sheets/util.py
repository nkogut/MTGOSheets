import re
import sqlite3
from functools import lru_cache
from pathlib import Path

import orjson as json

from mtgo_sheets.models import CardDef

DEFAULT_COLLECTOR_NUMBER = "999"

def name_to_id(conn: sqlite3.Connection, name: str, set_: str | None, foil: bool = False) -> int | None:
    """Return an id associated with the name.

    Prefer to pick an id with a standard frame/art printing.
    """
    if set_ is None:
        # Fall back to a broader search
        return get_cheapest_variant(conn, name, foil)
    set_ = set_.upper()

    defs = conn.execute(
        """
        SELECT *
        FROM defs
        WHERE name LIKE ? AND set_ LIKE ? AND foil = ?
        ORDER BY name
        """,
        (name, set_, foil),
    ).fetchall()

    if not defs:
        return None

    # Try to pick the card whose ver holds the lowest collector number (most likely to be the base art)
    # Note: version schemas are inconsistent
    return min(defs, key=lambda def_: int(clean_def_version(def_[4])))[0]

def get_def(conn: sqlite3.Connection, id_: int) -> CardDef:
    """Find and return the definition for the id's printing."""
    row = conn.execute("SELECT id, name, set_, foil, ver, rarity FROM defs WHERE id = ?", (id_,)).fetchone()
    if row is None:
        raise ValueError(f"No def found for id {id_}")
    id_, name, set_code, foil, ver, rarity = row
    # if ver is None:
    #     ver = ""
    return CardDef(id_=id_, name=name, set_code=set_code, foil=bool(foil), ver=ver, rarity=rarity)

def get_price_list(conn: sqlite3.Connection, id_: int) -> list[float]:
    """Return the price list associated with the id."""
    prices = conn.execute(
        """
        SELECT price_list
        FROM prices
        WHERE id = ?
        """,
        (id_,),
    ).fetchone()
    if prices is None:
        raise ValueError(f"Failed to find price list for {id_}")

    return json.loads(prices[0])

def clean_def_version(ver: str) -> str:
    """Normalize a definition's version."""
    if not ver or ver.startswith("/"):
        ver = DEFAULT_COLLECTOR_NUMBER
    else:
        ver = ver.replace("*", "")
        ver = ver.split("/")[0]
        ver = re.sub(r'[a-zA-Z]', DEFAULT_COLLECTOR_NUMBER, ver)
    return ver

@lru_cache(maxsize=1024)
def get_cheapest_variant(conn: sqlite3.Connection, name: str, foil: bool = False) -> int | None:
    """Return the cheapest printing with the given name and foil status."""
    row = conn.execute(
        """
        SELECT p.id
        FROM prices p
        JOIN defs d ON p.id = d.id
        WHERE d.name LIKE ? AND d.foil = ?
        ORDER BY CAST(p.price_list ->> -1 AS REAL)
        LIMIT 1
        """,
        (name, foil),
    ).fetchone()

    if row is None:
        return None
    return row[0]

def get_price_lists(conn: sqlite3.Connection, ids: list[int] | None = None) -> dict[int, list[float]]:
    """Return a map of each id to its price list. Default to all ids if none are specified."""
    prices_map: dict[int, list[float]] = {}
    if ids is None:
        rows = conn.execute("SELECT id, price_list FROM prices").fetchall()

    else:
        rows = []
        chunk_size = 30_000
        for i in range(0, len(ids), chunk_size):
            chunk = ids[i:i + chunk_size]
            placeholders = ",".join("?" for _ in chunk)
            rows.extend(
                conn.execute(
                    f"SELECT id, price_list FROM prices WHERE id IN ({placeholders})",
                    chunk,
                ).fetchall(),
            )

    for row in rows:
        prices_map[row[0]] = json.loads(row[1])

    return prices_map

def assert_editable(path: Path) -> None:
    """Raise runtimeerror if the file is not editable."""
    try:
        with path.open("r+b"):
            pass
    except PermissionError as e:
        raise RuntimeError(
            f"Workbook '{path}' is not writable. It may be open in Excel or locked by another process"
        ) from e
