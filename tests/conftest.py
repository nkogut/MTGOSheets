from pathlib import Path

import pytest

from mtgo_sheets import config

# Override CONFIG before importing other modules so they bind to the mock config
config.CONFIG = config.load_config(Path(__file__).with_name("mock0.toml"))

from mtgo_sheets.models import CardDef


@pytest.fixture
def sample_card_def() -> CardDef:
    """Baseline card definition for unit tests."""
    return CardDef(
        id_=123,
        name="lightning bolt",
        set_code="lea",
        foil=False,
        base_printing=True,
        variant_data={"collector_num": "1", "treatment": "normal", "frame": "regular", "finish": "normal"},
        rarity="common",
    )
