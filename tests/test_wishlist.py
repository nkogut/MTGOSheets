from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

import pytest

from mtgo_sheets.config import CONFIG, XlsxFile
from mtgo_sheets.tracker import WishlistHandler, dump_ids


def test_wishlist_add_and_flush_valid_xml(tmp_path: Path) -> None:
    """Verify WishlistHandler generates correct XML and writes back and clears state on flush."""
    handler = WishlistHandler()
    wishlist_dir = tmp_path / "wishlists"

    handler.add_to_wishlist(wishlist_dir, id_=123, action="buy", vendor="goat", qty=4)
    handler.add_to_wishlist(wishlist_dir, id_=456, action="dump", vendor=None, qty=1)

    modified = handler.flush()
    assert len(modified) == 2
    assert handler.wishlists == {}

    buy_path = wishlist_dir / f"wishlist_{CONFIG.today}_buy_goat.dek"
    dump_path = wishlist_dir / f"wishlist_{CONFIG.today}_dump.dek"
    assert set(modified) == {buy_path, dump_path}

    root = ET.parse(buy_path).getroot()
    cards = root.findall("Cards")
    assert len(cards) == 1
    assert cards[0].attrib == {"CatID": "123", "Quantity": "4"}


def test_wishlist_appends_existing_file(tmp_path: Path) -> None:
    """Verify the wishlist file is extended (not overwritten) if it exists."""
    handler = WishlistHandler()
    target_path = tmp_path / "existing_wishlist.dek"

    handler.add_to_wishlist_path(id_=123, qty=2, path=target_path)
    handler.flush()

    handler.add_to_wishlist_path(id_=456, qty=3, path=target_path)
    handler.flush()

    raw_xml = target_path.read_text(encoding="utf-8")

    # Verify that the closing tag was removed, data was added, then closing deck was replaced
    assert raw_xml.count("</Deck>") == 1

    cards = ET.parse(target_path).getroot().findall("Cards")
    assert [(c.get("CatID"), c.get("Quantity")) for c in cards] == [
        ("123", "2"),
        ("456", "3"),
    ]


def test_dump_ids(xlsx_cfg: XlsxFile, fake_app: Any) -> None:
    """Verify dump_ids includes all ids in designated sheets (except those belonging to suggestions)."""
    wb = fake_app.books.open(xlsx_cfg.path)
    ws = wb.sheets[0]

    # Dumping from a sheet with no valid cards raises a RuntimeError
    ws.range((1, 1)).value = [xlsx_cfg.cols]
    with pytest.raises(RuntimeError, match="Failed to dump ids"):
        dump_ids([xlsx_cfg])

    def _make_row(id_: int | None, qty: int, drop: float | None = None, rise: float | None = None) -> list[Any]:
        row_map = {"id": id_, "qty": qty, "drop": drop, "rise": rise}
        return [row_map.get(col) for col in xlsx_cfg.cols]

    # Verify that the correct ids from the file are included in the dump
    ws.range((xlsx_cfg.first_data_row, 1)).value = [
        # Included in dump
        _make_row(id_=101, qty=2),              # qty > 0
        _make_row(id_=102, qty=0, drop=10.0),   # qty == 0 but has an alert configured
        _make_row(id_=103, qty=-1, rise=10.0),  # suggestion, but has an alert configured

        # Not included in dump
        _make_row(id_=104, qty=-1),             # suggestion with no alert configured
        _make_row(id_=105, qty=0),              # qty == 0
        _make_row(id_=None, qty=4),             # missing id
    ]

    dek_path = dump_ids([xlsx_cfg])
    cards = ET.parse(dek_path).getroot().findall("Cards")
    assert [c.get("CatID") for c in cards] == ["101", "102", "103"]
