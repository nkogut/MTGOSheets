from __future__ import annotations

import datetime
from typing import Any

from pydantic import BaseModel, Field, model_validator

type CardRow = dict[str, Any]
type ExtPriceMap = dict[str, dict[str, dict[str, float]]] # {id_str: {vendor: {"buy": price, "sell": price}}}

class CardDef(BaseModel):
    """Metadata for a particular card printing."""

    id_: int
    name: str
    set_code: str
    foil: bool
    ver: str

TX_REQUIRED_EXTRAS: dict[str, list[str]] = {
    "buy": [],
    "revise": ["old_action", "old_cmd"],
    "sell": ["net_each"],
    "arb": ["net_each"],
}

class Transaction(BaseModel):
    """Representation of one order entered on the sheet. Used to update Logger state."""

    action: str
    id_: int
    name: str
    price_each: float
    qty: int
    vendor: str
    holdings: int

    notes: str | None = None
    date: datetime.date = Field(default_factory=datetime.date.today)
    price_total: float | None = None

    # Sell / Arb fields
    net_each: float | None = None
    net_total: float | None = None

    # Revision fields
    old_action: str | None = None
    old_cmd: str | None = None

    @model_validator(mode="after")
    def validate_action_and_extras(self) -> Transaction:
        """Validate required fields based on transaction action."""
        if self.action not in TX_REQUIRED_EXTRAS:
            raise ValueError(f"Invalid action {self.action}")

        missing_fields = [
            field for field in TX_REQUIRED_EXTRAS[self.action]
            if getattr(self, field) is None
        ]

        if missing_fields:
            fields_str = ", ".join(f"'{f}'" for f in missing_fields)
            raise ValueError(f"Missing required fields for action '{self.action}': {fields_str}")

        return self
