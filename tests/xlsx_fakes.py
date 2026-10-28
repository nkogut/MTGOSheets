from __future__ import annotations

import os

from pathlib import Path
from typing import Any


class FakeRange:
    """Mimics xw.Range."""
    def __init__(
        self,
        sheet: FakeSheet,
        start_coord: tuple[int, int],
        end_coord: tuple[int, int] | None = None,
    ) -> None:
        self.sheet = sheet
        self.start_coord = start_coord
        self.end_coord = end_coord or start_coord
        self.ndim: int | None = None

    @property
    def row(self) -> int:
        return self.start_coord[0]

    def options(self, ndim: int | None = None) -> FakeRange:
        self.ndim = ndim
        return self

    def end(self, direction: str) -> FakeRange:
        col = self.start_coord[1]
        populated_rows = [r for (r, c), v in self.sheet.grid.items() if c == col and v is not None and v != ""]
        last_row = max(populated_rows) if populated_rows else 1
        return FakeRange(self.sheet, (last_row, col))

    def clear_contents(self) -> None:
        r1, c1 = self.start_coord
        r2, c2 = self.end_coord
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self.sheet.grid.pop((r, c), None)

    @property
    def value(self) -> Any:
        r1, c1 = self.start_coord
        r2, c2 = self.end_coord
        if r1 == r2 and c1 == c2 and self.ndim != 2:
            return self.sheet.grid.get((r1, c1))
        rows = []
        for r in range(r1, r2 + 1):
            row_vals = [self.sheet.grid.get((r, c)) for c in range(c1, c2 + 1)]
            rows.append(row_vals)
        if r1 == r2 and self.ndim != 2:
            return rows[0]
        return rows

    @value.setter
    def value(self, val: Any) -> None:
        r1, c1 = self.start_coord
        if isinstance(val, list):
            if val and isinstance(val[0], list):
                for dr, row_vals in enumerate(val):
                    for dc, cell_val in enumerate(row_vals):
                        self.sheet.grid[(r1 + dr, c1 + dc)] = cell_val
            else:
                for dc, cell_val in enumerate(val):
                    self.sheet.grid[(r1, c1 + dc)] = cell_val
        else:
            self.sheet.grid[(r1, c1)] = val


class FakeCells:
    """Mimics xw.Sheet.cells."""
    def __init__(self, sheet: FakeSheet) -> None:
        self.sheet = sheet
        self.last_cell = FakeRange(sheet, (1000, 1000))

    def __call__(self, row: int, col: int) -> FakeRange:
        return FakeRange(self.sheet, (row, col), (row, col))


class FakeSheet:
    """Mimics xw.Sheet."""
    def __init__(self) -> None:
        self.grid: dict[tuple[int, int], Any] = {}
        self.cells = FakeCells(self)

    @property
    def used_range(self) -> FakeRange:
        if not self.grid:
            return FakeRange(self, (1, 1), (1, 1))
        max_r = max(r for r, _ in self.grid)
        max_c = max(c for _, c in self.grid)
        return FakeRange(self, (1, 1), (max_r, max_c))

    def range(self, start: tuple[int, int], end: tuple[int, int] | None = None) -> FakeRange:
        return FakeRange(self, start, end)


class FakeBook:
    """Mimics xw.Book. Touches actual workbook path on creation and save."""
    def __init__(self, path: Path | None = None) -> None:
        self.path = path
        if self.path:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.touch()
        self.sheets = [FakeSheet()]
        self.saved = False
        self.closed = False

    def save(self, path: Path | None = None) -> None:
        target = path or self.path
        if target:
            self.path = Path(target)
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_bytes(b"fake-xlsx-content")
            mtime = self.path.stat().st_mtime + 1.0 if self.path.exists() else 999.9
            os.utime(self.path, (mtime, mtime))
        self.saved = True

    def close(self) -> None:
        self.closed = True


class FakeBooks:
    """Mimics xw.App.books. Caches FakeBooks by path. Touches actual workbook on open."""
    def __init__(self) -> None:
        self.by_path: dict[Path, FakeBook] = {}

    def open(self, path: Path) -> FakeBook:
        if path not in self.by_path:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.touch()
            self.by_path[path] = FakeBook(path)
        return self.by_path[path]

    def add(self) -> FakeBook:
        return FakeBook()


class FakeApp:
    """Mimics xw.App."""
    def __init__(self) -> None:
        self.books = FakeBooks()
        self.quitted = False

    def quit(self) -> None:
        self.quitted = True
