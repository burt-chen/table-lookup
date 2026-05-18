"""Data processing core — no pandas / numpy dependency.

Reads xlsx/csv/json into a lightweight Table (columns + rows-of-lists),
runs the filter/merge, and writes the output back out. Uses only the
stdlib plus openpyxl (for xlsx).
"""
from __future__ import annotations

import csv
import json
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator


TEMPLATE_PATTERN = re.compile(r"\{([^{}]+)\}")


# ============================================================================
#  Table — minimal tabular container (replaces pandas.DataFrame)
# ============================================================================

class Row:
    """Pandas-Series-shaped view over one Table row.

    Provides just the bits the UI needs: ``tolist()``, ``__getitem__``, ``get``,
    ``to_dict``.
    """

    __slots__ = ("_values", "_cols", "_idx_of")

    def __init__(self, values: list, columns: list[str], idx_of: dict[str, int]):
        self._values = values
        self._cols = columns
        self._idx_of = idx_of

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._values[key]
        i = self._idx_of.get(key)
        return None if i is None else self._values[i]

    def get(self, key, default=None):
        i = self._idx_of.get(key)
        if i is None:
            return default
        v = self._values[i] if i < len(self._values) else None
        return default if v is None else v

    def tolist(self) -> list:
        return list(self._values)

    def to_dict(self) -> dict[str, Any]:
        out = {}
        for c, i in self._idx_of.items():
            out[c] = self._values[i] if i < len(self._values) else None
        return out


@dataclass
class Table:
    columns: list[str] = field(default_factory=list)
    rows: list[list] = field(default_factory=list)

    def __len__(self) -> int:
        return len(self.rows)

    @property
    def shape(self) -> tuple[int, int]:
        return (len(self.rows), len(self.columns))

    def head(self, n: int) -> "Table":
        return Table(columns=list(self.columns), rows=list(self.rows[:n]))

    def column_index(self, name: str) -> int:
        return self.columns.index(name)

    def column_values(self, name: str) -> list:
        idx = self.column_index(name)
        return [r[idx] if idx < len(r) else None for r in self.rows]

    def row_dict(self, i: int) -> dict[str, Any]:
        row = self.rows[i]
        return {c: (row[j] if j < len(row) else None) for j, c in enumerate(self.columns)}

    def iter_dicts(self) -> Iterator[dict[str, Any]]:
        for i in range(len(self.rows)):
            yield self.row_dict(i)

    def iterrows(self) -> Iterator[tuple[int, Row]]:
        """pandas-style row iterator yielding (index, Row)."""
        idx_of = {c: j for j, c in enumerate(self.columns)}
        for i, r in enumerate(self.rows):
            yield i, Row(r, self.columns, idx_of)


def _is_nan(v: Any) -> bool:
    return isinstance(v, float) and math.isnan(v)


def _safe_str(v: Any) -> str:
    if v is None or _is_nan(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ============================================================================
#  Readers
# ============================================================================

def list_sheets(path: str | Path) -> list[str]:
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".xls":
        book = _open_xls(p)
        try:
            return list(book.sheet_names())
        finally:
            book.release_resources()
    if ext not in (".xlsx", ".xlsm"):
        return []
    from openpyxl import load_workbook
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def read_table(path: str | Path, sheet: str | int | None = None) -> Table:
    p = Path(path)
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(p, sheet)
    if ext == ".xls":
        return _read_xls(p, sheet)
    if ext == ".csv":
        return _read_csv(p)
    if ext == ".json":
        return _read_json(p)
    raise ValueError(f"不支援的檔案格式:{ext}")


def _read_xlsx(p: Path, sheet: str | int | None) -> Table:
    from openpyxl import load_workbook
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        if sheet is None or sheet == "":
            ws = wb.worksheets[0]
        elif isinstance(sheet, int):
            ws = wb.worksheets[sheet]
        else:
            if sheet not in wb.sheetnames:
                raise KeyError(f"工作表不存在:{sheet}")
            ws = wb[sheet]

        rows_iter = ws.iter_rows(values_only=True)
        header_tuple = next(rows_iter, None)
        if header_tuple is None:
            return Table(columns=[], rows=[])
        columns = [_safe_str(c) if c is not None else f"欄位{i + 1}"
                   for i, c in enumerate(header_tuple)]
        rows: list[list] = []
        for r in rows_iter:
            row = list(r)
            # Skip wholly empty rows (openpyxl returns trailing blanks sometimes)
            if all(v is None or (isinstance(v, str) and v == "") for v in row):
                continue
            # Pad / trim to header width
            if len(row) < len(columns):
                row = row + [None] * (len(columns) - len(row))
            elif len(row) > len(columns):
                row = row[:len(columns)]
            rows.append(row)
        return Table(columns=columns, rows=rows)
    finally:
        wb.close()


def _open_xls(p: Path):
    """開啟舊版二進位 .xls(需 xlrd 套件)。"""
    try:
        import xlrd
    except ImportError as e:
        raise ImportError(
            "讀取 .xls 檔需要 xlrd 套件,請執行:pip install xlrd"
        ) from e
    return xlrd.open_workbook(str(p), on_demand=True)


def _read_xls(p: Path, sheet: str | int | None) -> Table:
    import xlrd

    book = _open_xls(p)
    try:
        if sheet is None or sheet == "":
            ws = book.sheet_by_index(0)
        elif isinstance(sheet, int):
            ws = book.sheet_by_index(sheet)
        else:
            if sheet not in book.sheet_names():
                raise KeyError(f"工作表不存在:{sheet}")
            ws = book.sheet_by_name(sheet)

        def _cell(c, ctype):
            if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                return None
            if ctype == xlrd.XL_CELL_BOOLEAN:
                return bool(c)
            if ctype == xlrd.XL_CELL_DATE:
                try:
                    return xlrd.xldate.xldate_as_datetime(c, book.datemode)
                except Exception:
                    return c
            if ctype == xlrd.XL_CELL_NUMBER and float(c).is_integer():
                return int(c)
            return c

        if ws.nrows == 0:
            return Table(columns=[], rows=[])
        header = [
            _safe_str(v) if v not in (None, "") else f"欄位{i + 1}"
            for i, v in enumerate(ws.row_values(0))
        ]
        rows: list[list] = []
        for r in range(1, ws.nrows):
            row = [_cell(c, t) for c, t in zip(ws.row_values(r), ws.row_types(r))]
            if all(v is None or (isinstance(v, str) and v == "") for v in row):
                continue
            if len(row) < len(header):
                row = row + [None] * (len(header) - len(row))
            elif len(row) > len(header):
                row = row[:len(header)]
            rows.append(row)
        return Table(columns=header, rows=rows)
    finally:
        book.release_resources()


def _read_csv(p: Path) -> Table:
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
        try:
            with open(p, encoding=enc, newline="") as f:
                reader = csv.reader(f)
                all_rows = list(reader)
            break
        except UnicodeDecodeError as e:
            last_err = e
            continue
    else:
        raise ValueError(f"無法解碼 CSV:{p}({last_err})")

    if not all_rows:
        return Table(columns=[], rows=[])
    header = [c if c else f"欄位{i + 1}" for i, c in enumerate(all_rows[0])]
    rows: list[list] = []
    for r in all_rows[1:]:
        if len(r) < len(header):
            r = list(r) + [None] * (len(header) - len(r))
        elif len(r) > len(header):
            r = r[:len(header)]
        rows.append(list(r))
    return Table(columns=header, rows=rows)


def _read_json(p: Path) -> Table:
    data = json.loads(p.read_text(encoding="utf-8"))
    if isinstance(data, dict):
        # accept wrappers like {"data": [...]} / {"outputs": [...]}
        for key in ("data", "records", "rows", "outputs"):
            if key in data and isinstance(data[key], list):
                data = data[key]
                break

    if isinstance(data, list):
        # list of dicts → records
        columns: list[str] = []
        seen: set[str] = set()
        for row in data:
            if not isinstance(row, dict):
                raise ValueError("JSON 結構不正確:應為物件陣列或欄位字典")
            for k in row.keys():
                if k not in seen:
                    seen.add(k)
                    columns.append(k)
        rows = [[row.get(c) for c in columns] for row in data]
        return Table(columns=columns, rows=rows)

    if isinstance(data, dict):
        # column-oriented: {col: [values...]}
        columns = list(data.keys())
        n = max((len(v) for v in data.values() if isinstance(v, list)), default=0)
        rows = []
        for i in range(n):
            row = []
            for c in columns:
                v = data[c]
                row.append(v[i] if isinstance(v, list) and i < len(v) else None)
            rows.append(row)
        return Table(columns=columns, rows=rows)

    raise ValueError("JSON 結構不正確")


# ============================================================================
#  Template / reference resolution
# ============================================================================

def resolve_column_ref(ref: str, src1: dict, src2: dict):
    ref = (ref or "").strip()
    if ref.startswith("1."):
        return src1.get(ref[2:], "")
    if ref.startswith("2."):
        return src2.get(ref[2:], "")
    if ref in src1:
        return src1[ref]
    return src2.get(ref, "")


def render_template(template: str, src1: dict, src2: dict | None = None) -> str:
    src2 = src2 or {}

    def repl(m: re.Match) -> str:
        key = m.group(1).strip()
        val = resolve_column_ref(key, src1, src2)
        if val is None or _is_nan(val):
            return ""
        return _safe_str(val)

    return TEMPLATE_PATTERN.sub(repl, template)


def template_columns(template: str) -> list[str]:
    return [m.group(1).strip() for m in TEMPLATE_PATTERN.finditer(template)]


# ============================================================================
#  Merge config / result
# ============================================================================

@dataclass
class OutputColumn:
    """One output column.

    mode = "copy"     → value is a source column ref (e.g. "1.Town")
    mode = "template" → value is a template string (e.g. "{圖號}_114LU")
    """

    name: str
    mode: str = "copy"
    value: str = ""

    def to_dict(self) -> dict:
        return {"name": self.name, "mode": self.mode, "value": self.value}

    @classmethod
    def from_dict(cls, d: dict) -> "OutputColumn":
        return cls(
            name=d.get("name", ""),
            mode=d.get("mode", "copy"),
            value=d.get("value", ""),
        )


@dataclass
class MergeConfig:
    file1: str = ""
    sheet1: str = ""
    file2: str = ""
    sheet2: str = ""
    match_col1: str = ""
    filter_col2: str = ""
    drop_duplicates: bool = True
    keep_unmatched: bool = False
    outputs: list[OutputColumn] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "file1": self.file1,
            "sheet1": self.sheet1,
            "file2": self.file2,
            "sheet2": self.sheet2,
            "match_col1": self.match_col1,
            "filter_col2": self.filter_col2,
            "drop_duplicates": self.drop_duplicates,
            "keep_unmatched": self.keep_unmatched,
            "outputs": [o.to_dict() for o in self.outputs],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "MergeConfig":
        c = cls()
        c.file1 = d.get("file1", "")
        c.sheet1 = d.get("sheet1", "")
        c.file2 = d.get("file2", "")
        c.sheet2 = d.get("sheet2", "")
        c.match_col1 = d.get("match_col1", "")
        c.filter_col2 = d.get("filter_col2", "")
        c.drop_duplicates = d.get("drop_duplicates", True)
        c.keep_unmatched = d.get("keep_unmatched", False)
        c.outputs = [OutputColumn.from_dict(o) for o in d.get("outputs", [])]
        return c


@dataclass
class MergeResult:
    table: Table
    missing_keys: list[str]
    total_keys: int

    # Backwards-compat alias so callers expecting .df still work.
    @property
    def df(self) -> Table:
        return self.table


def _normalize_key(v: Any) -> str:
    if v is None or _is_nan(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


# ============================================================================
#  Run merge
# ============================================================================

def run_merge(
    cfg: MergeConfig,
    t1: Table | None = None,
    t2: Table | None = None,
    progress: Callable[[int, int, str], None] | None = None,
) -> MergeResult:
    """Execute the filter/merge. t1/t2 can be passed in to avoid re-reading."""

    def _p(i: int, n: int, msg: str) -> None:
        if progress is not None:
            progress(i, n, msg)

    _p(0, 100, "讀取來源 1…")
    if t1 is None:
        t1 = read_table(cfg.file1, cfg.sheet1 or None)
    _p(20, 100, "讀取來源 2…")
    if t2 is None:
        t2 = read_table(cfg.file2, cfg.sheet2 or None)

    if cfg.match_col1 not in t1.columns:
        raise KeyError(f"來源 1 找不到欄位:{cfg.match_col1}")
    if cfg.filter_col2 not in t2.columns:
        raise KeyError(f"來源 2 找不到欄位:{cfg.filter_col2}")

    _p(35, 100, "建立索引…")
    match_idx = t1.column_index(cfg.match_col1)
    filter_idx = t2.column_index(cfg.filter_col2)

    # Index of source 1: first occurrence wins when drop_duplicates is on
    lookup: dict[str, list] = {}
    for row in t1.rows:
        key = _normalize_key(row[match_idx] if match_idx < len(row) else None)
        if not key:
            continue
        if key in lookup and cfg.drop_duplicates:
            continue
        lookup[key] = row

    keys: list[str] = []
    for row in t2.rows:
        k = _normalize_key(row[filter_idx] if filter_idx < len(row) else None)
        if k:
            keys.append(k)
    total = len(keys)
    _p(45, 100, f"開始比對 {total} 筆…")

    out_columns = [c.name for c in cfg.outputs]
    out_rows: list[list] = []
    missing: list[str] = []
    idx = 0
    progress_every = max(1, total // 50) if total else 1

    for r2 in t2.rows:
        key = _normalize_key(r2[filter_idx] if filter_idx < len(r2) else None)
        if not key:
            continue
        src2_dict = {c: (r2[j] if j < len(r2) else None) for j, c in enumerate(t2.columns)}

        if key in lookup:
            r1 = lookup[key]
            src1_dict = {c: (r1[j] if j < len(r1) else None) for j, c in enumerate(t1.columns)}
        else:
            missing.append(key)
            if not cfg.keep_unmatched:
                idx += 1
                if progress and total and idx % progress_every == 0:
                    _p(45 + int(45 * idx / total), 100, f"比對中 {idx}/{total}")
                continue
            src1_dict = {cfg.match_col1: key}

        row_out: list = []
        for col in cfg.outputs:
            if col.mode == "copy":
                v = resolve_column_ref(col.value, src1_dict, src2_dict)
                row_out.append("" if v is None or _is_nan(v) else v)
            else:
                row_out.append(render_template(col.value, src1_dict, src2_dict))
        out_rows.append(row_out)

        idx += 1
        if progress and total and idx % progress_every == 0:
            _p(45 + int(45 * idx / total), 100, f"比對中 {idx}/{total}")

    _p(95, 100, "建立輸出表…")
    out_table = Table(columns=out_columns, rows=out_rows)
    _p(100, 100, "完成")
    return MergeResult(table=out_table, missing_keys=missing, total_keys=total)


# ============================================================================
#  Writers
# ============================================================================

def export(table: Table, path: str | Path) -> None:
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".csv":
        _write_csv(table, p)
    elif ext in (".xlsx", ".xlsm"):
        _write_xlsx(table, p)
    elif ext == ".json":
        _write_json(table, p)
    else:
        raise ValueError(f"不支援的輸出格式:{ext}")


def _write_csv(table: Table, p: Path) -> None:
    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(table.columns)
        for row in table.rows:
            w.writerow(["" if v is None or _is_nan(v) else v for v in row])


def _write_xlsx(table: Table, p: Path) -> None:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(list(table.columns))
    for row in table.rows:
        ws.append([None if (v is None or _is_nan(v)) else v for v in row])
    wb.save(p)


def _write_json(table: Table, p: Path) -> None:
    records = []
    for row in table.rows:
        d = {}
        for j, c in enumerate(table.columns):
            v = row[j] if j < len(row) else None
            d[c] = None if v is None or _is_nan(v) else v
        records.append(d)
    p.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
