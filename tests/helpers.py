from typing import Any
from mtgo_sheets.models import CardDef


def make_card_def(
    id_: int = 123,
    name: str = "default card",
    set_code: str = "fdn", # legal for all sets
    foil: bool = False,
    base_printing: bool = True,
    variant_data: dict[str, Any] | None = None,
    rarity: str = "common",
) -> CardDef:
    """Build a CardDef with default test values."""
    if variant_data is None:
        variant_data = {"collector_num": "1", "treatment": "normal", "frame": "regular", "finish": "normal"}
    return CardDef(
        id_=id_,
        name=name,
        set_code=set_code,
        foil=foil,
        base_printing=base_printing,
        variant_data=variant_data,
        rarity=rarity,
    )
