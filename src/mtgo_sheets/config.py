from __future__ import annotations

import numbers
import operator
import tomllib
from collections.abc import Callable
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field, model_validator

from mtgo_sheets.models import CardRow

ROOT = Path(__file__).resolve().parents[2]
TOML_PATH = ROOT / "config.toml"

RAW_TOML = tomllib.loads(TOML_PATH.read_text())

OPERATOR_MAP = {
    "==": operator.eq,
    "!=": operator.ne,
    ">": operator.gt,
    "<": operator.lt,
    ">=": operator.ge,
    "<=": operator.le,
}

class SortClause(BaseModel):
    """One rule used when sorting CardRows."""

    col: str
    desc: bool = False
    op: str | None = None
    val: Any | None = None

    @model_validator(mode="before")
    @classmethod
    def parse_shortcut(cls, data: str | list | tuple | dict[str, Any]) -> dict[str, Any]:
        """If the data is not provided as a dict, assume the parameters based on length."""
        if not data:
            raise ValueError("No options supplied for the sort clause")

        parsed: dict[str, Any] = {}
        # col only
        if isinstance(data, str):
            parsed = {"col": data}

        elif isinstance(data, (list, tuple)):
            # col + desc
            if len(data) == 2:
                parsed = {"col": data[0], "desc": data[1]}

            # col + operator + value
            elif len(data) == 3:
                parsed = {"col": data[0], "op": data[1], "val": data[2]}
        else:
            parsed = data

        return parsed

class PathsConfig(BaseModel):
    """Collection of paths required in config.toml. Additional paths are also allowed."""

    set_legalities: Path
    suggestion_card_blacklist: Path
    external_prices: Path
    external_prices_hist: Path
    db: Path

    model_config = {"extra": "allow"}
    @model_validator(mode="before")
    @classmethod
    def resolve_absolute_paths(cls, data: dict[str, Any]) -> dict[str, Path]:
        """Anchor all input strings into path objects."""
        return {key: ROOT / Path(val) for key, val in data.items()}


class ListsConfig(BaseModel):
    """Collection of optional lists. Default to [] if not provided. Additional lists are allowed."""

    suggestion_option_keys: list[str] = Field(default_factory=list)
    suggestion_set_blacklist: list[str] = Field(default_factory=list)
    model_config = {"extra": "allow"}

class XlsxConfigDefaults(BaseModel):
    """Optional config options available for each xlsx file: [xlsx.<name>]."""

    log_path: Path | None = None
    log_cols: list[str] | None = None
    wishlist_dir: Path | None = None
    sort: list[SortClause] | None = None

    # Sheet layout
    cols: list[str] | None = None
    first_data_row: int | None = None
    blacklist_row: int | None = None
    blacklist_col: int | None = None
    suggestions_top_row: int | None = None
    suggestions_left_col: int | None = None

    # Misc Options
    auto_update: bool | None = None
    ignore_for_ext_prices: bool | None = None


class XlsxFile(BaseModel):
    """All args for one xlsx file configuration.

    Fall back to defaults from [xlsx_defaults] for any args not provided.
    """

    path: Path
    log_path: Path
    log_cols: list[str]
    wishlist_dir: Path
    cols: list[str]
    sort: list[SortClause]

    first_data_row: int
    blacklist_row: int
    blacklist_col: int
    suggestions_top_row: int
    suggestions_left_col: int

    auto_update: bool
    ignore_for_ext_prices: bool

    sort_fn: Callable[[CardRow], tuple] | None = None

    @model_validator(mode="before")
    @classmethod
    def apply_defaults(cls, data: dict[str, Any]) -> dict[str, Any]:
        """Unless an option is specified for this file, apply the default that the user defined for all files."""
        defaults: XlsxConfigDefaults | None = data.pop("__defaults__", None)
        if not defaults:
            return data

        fields_to_check = [
            "path", "log_path", "log_cols", "cols", "sort", "first_data_row", "wishlist_dir",
            "blacklist_row", "blacklist_col", "suggestions_top_row", "suggestions_left_col",
            "auto_update", "ignore_for_ext_prices",
        ]

        for field in fields_to_check:
            if data.get(field) is None:
                default_val = getattr(defaults, field)
                if default_val is not None:
                    data[field] = default_val

        return data

    @model_validator(mode="after")
    def resolve_all_paths(self) -> XlsxFile:
        """Guarantee all Paths are absolute."""
        path_fields = ["path", "log_path", "wishlist_dir"]
        for field in path_fields:
            current_path: Path = getattr(self, field)
            if current_path and not current_path.is_absolute():
                setattr(self, field, (ROOT / current_path).resolve())
            else:
                setattr(self, field, current_path.resolve())
        return self


class AppConfig(BaseModel):
    """Process the config toml file. Store all configuration options."""

    debug_mode: bool = False
    start_date: date
    today: date = Field(default_factory=date.today)

    paths: PathsConfig
    lists: ListsConfig

    min_arb_thresh: float = 0.03
    skip_first_n_days: int = 14

    xlsx: dict[str, XlsxFile] = {}

    @model_validator(mode="before")
    @classmethod
    def build_xlsx_cfgs(cls, data: dict[str, Any]) -> dict[str, Any]:
        """Process each xslx file described in the config file.

        Return the updated internal state with the new xlsx file objects.
        """
        raw_data = data.copy()

        # Build file structures using default + file-specific values
        defaults_raw = raw_data.get("xlsx_defaults", {})
        defaults_container = XlsxConfigDefaults(**defaults_raw)

        processed_files: dict[str, XlsxFile] = {}
        xlsx_raw_map = raw_data.get("xlsx", {})

        for name, cfg in xlsx_raw_map.items():
            file_data = cfg.copy()
            file_data["path"] = ROOT / cfg["path"]
            file_data["__defaults__"] = defaults_container

            file_instance = XlsxFile(**file_data)
            file_instance.sort_fn = compile_sort_fn(file_instance.sort, file_instance.cols)
            processed_files[name] = file_instance

        raw_data["xlsx"] = processed_files
        return raw_data


def compile_sort_fn(
    sort_spec: list[SortClause],
    valid_cols: list[str],
) -> Callable[[CardRow], tuple]:
    """Compile a list of conditional clauses into one sort function."""
    def sort_fn(row: CardRow) -> tuple:
        """Process a single row and returns a tuple indicating sort priority."""
        result = []
        for clause in sort_spec:
            if clause.col not in valid_cols:
                raise ValueError(f"Invalid sort column: {clause.col!r}")

            val = row[clause.col]
            if clause.op is not None:
                if clause.op not in OPERATOR_MAP:
                    raise ValueError(f"Unsupported sort operator: {clause.op!r}")

                # Handle boolean results by applying the operator
                condition_met = OPERATOR_MAP[clause.op](val, clause.val)
                result.append(condition_met)

            # Handle non-boolean results
            elif clause.desc:
                # Invert the output
                if isinstance(val, (numbers.Real, Decimal)):
                    result.append(-val)

                elif isinstance(val, bool):
                    result.append(not val)

                elif isinstance(val, str):
                    inverted_string_tuple = tuple(-ord(char) for char in val)
                    result.append(inverted_string_tuple)

                else:
                    raise ValueError(f"Unable to apply descending sort to {val!r} of type {type(val)}")
            else:
                # Ascending
                result.append(val)

        return tuple(result)

    return sort_fn

CONFIG = AppConfig(**RAW_TOML)
