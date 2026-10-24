import datetime
from decimal import Decimal
from pathlib import Path

import pytest
from pydantic import ValidationError

from mtgo_sheets.config import compile_sort_fn, load_config, PathsConfig, ROOT, SortClause


def test_load_config_toml() -> None:
    """Verify loading all settings from mock0.toml."""
    cfg = load_config(Path(__file__).with_name("mock0.toml"))

    assert cfg.debug_mode is True
    assert cfg.start_date == datetime.date(2025, 1, 10)
    assert cfg.min_arb_thresh == 0.03
    assert cfg.skip_first_n_days == 0

    assert cfg.paths.db == (ROOT / "test.db").resolve()
    assert cfg.paths.set_legalities == (ROOT / "legalities.json").resolve()
    assert "price_change" in cfg.lists.suggestion_option_keys
    assert cfg.lists.suggestion_set_blacklist == ["prm"]

    main_sheet = cfg.xlsx["main"]
    assert main_sheet.path == (ROOT / "sheet.xlsx").resolve()
    assert main_sheet.suggestion_card_blacklist == (ROOT / "blacklist.json").resolve()
    assert main_sheet.first_data_row == 3
    assert main_sheet.auto_update is True
    assert main_sheet.ignore_for_ext_prices is False
    assert main_sheet.sort_fn is not None


def test_xlsx_inheritance_and_overrides() -> None:
    """Verify multi-sheet TOML config inherits global settings and applies per-sheet overrides."""
    cfg = load_config(Path(__file__).with_name("mock1.toml"))
    assert cfg.lists.suggestion_option_keys == []
    assert cfg.lists.suggestion_set_blacklist == []

    # default_sheet inherits global suggestion_card_blacklist and all xlsx_defaults
    default_sh = cfg.xlsx["default_sheet"]
    assert default_sh.suggestion_card_blacklist == (ROOT / "global_bl.json").resolve()
    assert default_sh.first_data_row == 4
    assert default_sh.auto_update is True
    assert default_sh.ignore_for_ext_prices is False
    assert default_sh.external_suggestions_path is None

    # custom_sheet overrides defaults and clears blacklist by using the 'none' sentinel
    custom_sh = cfg.xlsx["custom_sheet"]
    assert custom_sh.suggestion_card_blacklist is None
    assert custom_sh.external_suggestions_path == (ROOT / "ext_sugg.xlsx").resolve()
    assert custom_sh.first_data_row == 6
    assert custom_sh.auto_update is False
    assert custom_sh.ignore_for_ext_prices is True


def test_relative_and_absolute_config_paths(tmp_path: Path) -> None:
    """Verify PathsConfig uses correct base for both relative and absolute paths."""
    abs_db = (tmp_path / "abs.db").resolve()
    paths = PathsConfig(
        set_legalities="rel_sets.json",
        suggestion_card_blacklist="rel_bl.json",
        external_prices="rel_ext.json",
        external_prices_hist="rel_hist.json",
        db=abs_db,
        custom_extra="extra_file.txt",  # type: ignore[call-arg]
    )
    assert paths.set_legalities == (ROOT / "rel_sets.json").resolve()
    assert paths.db == abs_db
    assert getattr(paths, "custom_extra") == (ROOT / "extra_file.txt").resolve()


def test_sort_clause_parsing() -> None:
    """Verify SortClause parses clauses given in any valid input format:
    - "col": sort ascending by column.
    - ["col", True/False]: sort by column and descending for True, ascending for False.
    - ["col", "op", val]: sort by applying an op to val ()"==", "!=", ">", "<", ">=", "<=").
    - dict: explicit mapping of all SortClause fields (col, desc, op, val)."""
    assert SortClause.model_validate("name").col == "name"

    desc_clause = SortClause.model_validate(["qty", True])
    assert (desc_clause.col, desc_clause.desc) == ("qty", True)

    op_clause = SortClause.model_validate(["qty", ">", 0])
    assert (op_clause.col, op_clause.op, op_clause.val) == ("qty", ">", 0)

    dict_clause = SortClause.model_validate({"col": "price", "desc": True})
    assert (dict_clause.col, dict_clause.desc) == ("price", True)

    with pytest.raises(ValidationError, match="No options supplied"):
        SortClause.model_validate("")


def test_compile_sort_fn_all_types_and_directions() -> None:
    """Verify the compiled sort_fn handles all operators, data types, and sort direction."""
    clauses = [
        SortClause.model_validate(["qty", "==", 0]),
        SortClause.model_validate(["foil", True]),
        SortClause.model_validate(["price", True]),
        SortClause.model_validate(["name", True]),
        SortClause.model_validate("date_created"),
    ]
    valid_cols = ["qty", "foil", "price", "name", "date_created"]
    sort_fn = compile_sort_fn(clauses, valid_cols)

    row_a = {"qty": 2, "foil": True, "price": Decimal("1.50"), "name": "Abc", "date_created": datetime.date(2026, 1, 2)}
    row_b = {"qty": 2, "foil": False, "price": Decimal("2.00"), "name": "Zyx", "date_created": datetime.date(2026, 1, 1)}

    sorted_rows = sorted([row_b, row_a], key=sort_fn)
    assert sorted_rows[0] == row_a

    bad_col_fn = compile_sort_fn([SortClause(col="missing")], valid_cols)
    with pytest.raises(ValueError, match="Invalid sort column"):
        bad_col_fn(row_a)

    bad_op_fn = compile_sort_fn([SortClause(col="qty", op="~=", val=1)], valid_cols)
    with pytest.raises(ValueError, match="Unsupported sort operator"):
        bad_op_fn(row_a)

    bad_desc_fn = compile_sort_fn([SortClause(col="name", desc=True)], valid_cols)
    with pytest.raises(ValueError, match="Unable to apply descending sort"):
        bad_desc_fn({**row_a, "name": None})
