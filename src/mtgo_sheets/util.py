import orjson as json
import re
import sqlite3
from functools import lru_cache
from pathlib import Path

import orjson as json

from mtgo_sheets.models import CardDef

def name_to_id(conn: sqlite3.Connection, name: str, set_: str | None, foil: bool = False) -> int | None:
    """Return an id associated with the name.

    Prefer to pick an id with a standard frame/art printing.
    """
    query = """
        SELECT id
        FROM defs
        WHERE name LIKE ? AND foil = ?
    """
    params = [name, foil]

    if set_ is not None:
        query += " AND set_ LIKE ?"
        params.append(set_.upper())

    query += " ORDER BY base_printing DESC, id ASC LIMIT 1"
    row = conn.execute(query, params).fetchone()
    return row[0] if row else None

def get_def(conn: sqlite3.Connection, id_: int) -> CardDef:
    """Find and return the definition for the id's printing."""
    row = conn.execute(
        "SELECT id, name, set_, foil, base_printing, variant_data, rarity FROM defs WHERE id = ?",
        (id_,)
    ).fetchone()

    if row is None:
        raise ValueError(f"No def found for id {id_}")

    id_, name, set_code, foil, base_printing, variant_data_json, rarity = row

    return CardDef(
        id_=id_,
        name=name,
        set_code=set_code,
        foil=bool(foil),
        base_printing=bool(base_printing),
        variant_data=json.loads(variant_data_json),
        rarity=rarity
    )

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

def clean_collector_no(collector_no: str) -> str:
    """Strip extraneous chars from collector_no. Result looks like 123a or 23."""
    collector_no = collector_no.split("/")[0]
    collector_no = collector_no.replace("*", "")
    collector_no = collector_no.strip().lstrip("0")
    return collector_no
