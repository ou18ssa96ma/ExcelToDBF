"""Core Excel to DBF conversion engine for v2.

Independent of PyQt so it can be used programmatically or via CLI.
"""

from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Callable, Sequence

import dbf

SUPPORTED_SUFFIXES = (".xlsx", ".xlsm", ".xls")

MAX_DBF_FIELDS = 255
MAX_DBF_CHAR_WIDTH = 254
MAX_FIELD_NAME = 10
MAX_NUMERIC_WIDTH = 20
MAX_SAFE_SIGNIFICANT_DIGITS = 15

_WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_NON_FIELD_CHARS = re.compile(r"[^A-Za-z0-9_]+")
_MULTI_UNDERSCORE = re.compile(r"_+")


class ConversionError(Exception):
    """Raised when conversion cannot proceed."""


class _Cancelled(Exception):
    """Internal signal for user cancellation."""


@dataclass(frozen=True)
class ConversionOptions:
    first_row_header: bool = True
    overwrite: bool = True
    encoding: str = "utf-8"     # "utf-8", "cp1256", "cp1252"
    dbf_format: str = "db3"     # "db3", "db4", "foxpro"


@dataclass
class SheetInfo:
    name: str
    row_count: int
    col_count: int


@dataclass
class SheetResult:
    sheet: str
    status: str  # "ok", "skipped", "failed", "cancelled"
    rows: int = 0
    fields: int = 0
    path: Path | None = None
    message: str = ""
    warnings: list[str] = field(default_factory=list)


@dataclass
class ConversionSummary:
    source: Path
    output_dir: Path
    results: list[SheetResult] = field(default_factory=list)
    cancelled: bool = False

    @property
    def created(self) -> int:
        return sum(1 for r in self.results if r.status == "ok")

    @property
    def skipped(self) -> int:
        return sum(1 for r in self.results if r.status == "skipped")

    @property
    def failed(self) -> int:
        return sum(1 for r in self.results if r.status == "failed")

    @property
    def total_rows(self) -> int:
        return sum(r.rows for r in self.results if r.status == "ok")


@dataclass
class _ColumnPlan:
    name: str
    kind: str  # C, N, D, L
    width: int
    decimals: int = 0
    mode: str = "text"

    def spec(self) -> str:
        if self.kind == "N":
            return f"{self.name} N({self.width},{self.decimals})"
        if self.kind == "C":
            return f"{self.name} C({self.width})"
        return f"{self.name} {self.kind}"

    def to_dbf(self, value: Any, *, encoding: str) -> Any:
        if value is None:
            return None

        if self.mode == "logical":
            return bool(value)

        if self.mode == "date":
            if isinstance(value, dt.datetime):
                return value.date()
            if isinstance(value, dt.date):
                return value
            return None

        if self.mode == "number":
            try:
                number = Decimal(str(value))
                if self.decimals == 0:
                    return int(number.to_integral_value(rounding=ROUND_HALF_UP))
                quantum = Decimal(1).scaleb(-self.decimals)
                return float(number.quantize(quantum, rounding=ROUND_HALF_UP))
            except (InvalidOperation, ValueError):
                return None

        # Text mode (C)
        text = _format_as_text(value)
        if not text.strip():
            return None
        text = _make_encodable(text, encoding)
        return _truncate_to_bytes(text, self.width, encoding)


def default_output_dir(source: str | Path) -> Path:
    source = Path(source)
    return source.parent / f"{source.stem}_dbf"


def inspect_workbook(source: str | Path) -> list[SheetInfo]:
    """Inspects workbook worksheets to quickly return names and dimensions."""
    source = Path(source)
    if not source.is_file():
        raise ConversionError(f"File not found: {source}")

    suffix = source.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise ConversionError(f"Unsupported format '{suffix}'. Supported: {', '.join(SUPPORTED_SUFFIXES)}")

    sheets: list[SheetInfo] = []
    if suffix in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(str(source), read_only=True, data_only=True)
        try:
            for name in wb.sheetnames:
                ws = wb[name]
                r_count = max(0, (ws.max_row or 1) - 1)
                c_count = ws.max_column or 0
                sheets.append(SheetInfo(name=name, row_count=r_count, col_count=c_count))
        finally:
            wb.close()
    else:  # .xls
        import xlrd
        wb = xlrd.open_workbook(str(source), on_demand=True)
        try:
            for name in wb.sheet_names():
                sheet = wb.sheet_by_name(name)
                r_count = max(0, sheet.nrows - 1)
                c_count = sheet.ncols
                sheets.append(SheetInfo(name=name, row_count=r_count, col_count=c_count))
        finally:
            wb.release_resources()

    return sheets


def convert_workbook(
    source: str | Path,
    output_dir: str | Path,
    options: ConversionOptions | None = None,
    *,
    selected_sheets: Sequence[str] | None = None,
    log: Callable[[str], None] | None = None,
    progress: Callable[[int, int, str], None] | None = None,
    is_cancelled: Callable[[], bool] | None = None,
) -> ConversionSummary:
    """Convert worksheets from source into DBF files in output_dir."""
    source = Path(source)
    output_dir = Path(output_dir)
    options = options or ConversionOptions()
    encoding = options.encoding.lower()
    dbf_codepage = "utf8" if encoding == "utf-8" else encoding

    if not source.is_file():
        raise ConversionError(f"Workbook not found: {source}")
    if source.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ConversionError(f"Unsupported file type '{source.suffix}'. Supported: {', '.join(SUPPORTED_SUFFIXES)}")

    _emit_log(log, f"Opening workbook: {source.name}")
    output_dir.mkdir(parents=True, exist_ok=True)

    summary = ConversionSummary(source=source, output_dir=output_dir)
    used_filenames: set[str] = set()

    suffix = source.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        all_sheets, read_sheet, close_reader = _open_xlsx(source)
    else:
        all_sheets, read_sheet, close_reader = _open_xls(source)

    try:
        targets = [s for s in all_sheets if selected_sheets is None or s in selected_sheets]
        total_sheets = len(targets)
        _emit_log(log, f"Targeting {total_sheets} worksheet(s) for conversion.")

        for idx, sheet_name in enumerate(targets, start=1):
            if is_cancelled and is_cancelled():
                summary.cancelled = True
                _emit_log(log, "Conversion cancelled by user.")
                break

            try:
                rows = read_sheet(sheet_name)
            except Exception as exc:
                summary.results.append(SheetResult(sheet_name, "failed", message=f"Read error: {exc}"))
                _emit_log(log, f"[{sheet_name}] Read error: {exc}")
                if progress:
                    progress(idx, total_sheets, sheet_name)
                continue

            try:
                result = _convert_sheet(
                    name=sheet_name,
                    rows=rows,
                    output_dir=output_dir,
                    used_filenames=used_filenames,
                    options=options,
                    encoding=encoding,
                    dbf_codepage=dbf_codepage,
                    is_cancelled=is_cancelled,
                )
            except _Cancelled:
                summary.cancelled = True
                _emit_log(log, f"[{sheet_name}] Cancelled mid-conversion.")
                if progress:
                    progress(idx, total_sheets, sheet_name)
                break
            except Exception as exc:
                msg = str(exc).strip() or exc.__class__.__name__
                summary.results.append(SheetResult(sheet_name, "failed", message=msg))
                _emit_log(log, f"[{sheet_name}] Failed: {msg}")
            else:
                summary.results.append(result)
                if result.status == "ok":
                    _emit_log(log, f"[{sheet_name}] OK: {result.rows} row(s) -> {result.path.name if result.path else ''}")
                elif result.status == "skipped":
                    _emit_log(log, f"[{sheet_name}] Skipped: {result.message}")
                for w in result.warnings:
                    _emit_log(log, f"[{sheet_name}] Warning: {w}")

            if progress:
                progress(idx, total_sheets, sheet_name)

    finally:
        try:
            close_reader()
        except Exception:
            pass

    _emit_log(log, f"Finished: {summary.created} created, {summary.failed} failed, {summary.skipped} skipped.")
    return summary


def _convert_sheet(
    *,
    name: str,
    rows: list[list[Any]],
    output_dir: Path,
    used_filenames: set[str],
    options: ConversionOptions,
    encoding: str,
    dbf_codepage: str,
    is_cancelled: Callable[[], bool] | None,
) -> SheetResult:
    grid = _trim_grid(rows)
    if not grid:
        return SheetResult(name, "skipped", message="Worksheet is empty")

    if options.first_row_header:
        headers = grid[0]
        data_rows = [r for r in grid[1:] if any(not _is_blank(v) for v in r)]
    else:
        headers = []
        data_rows = [r for r in grid if any(not _is_blank(v) for v in r)]

    field_count = len(grid[0])
    if field_count > MAX_DBF_FIELDS:
        raise ConversionError(f"Sheet has {field_count} columns; DBF supports at most {MAX_DBF_FIELDS}")

    names = _unique_field_names(headers, field_count)
    warnings: list[str] = []
    plans = [
        _build_column_plan(
            names[i],
            [r[i] for r in data_rows],
            encoding=encoding,
            warnings=warnings,
        )
        for i in range(field_count)
    ]

    filename = _unique_filename(name, used_filenames)
    path = output_dir / filename

    if path.exists() and not options.overwrite:
        return SheetResult(name, "skipped", fields=field_count, message=f"{filename} exists (overwrite disabled)")

    if is_cancelled and is_cancelled():
        raise _Cancelled

    _cleanup_file(path)

    spec = "; ".join(p.spec() for p in plans)
    table = None
    try:
        table = dbf.Table(
            str(path),
            spec,
            dbf_type=options.dbf_format,
            codepage=dbf_codepage,
        )
        table.open(dbf.READ_WRITE)
    except Exception as exc:
        if table is not None:
            try:
                table.close()
            except Exception:
                pass
        _cleanup_file(path)
        raise ConversionError(f"Could not initialize DBF file: {exc}") from exc

    written = 0
    try:
        for idx, row in enumerate(data_rows, start=1):
            if idx % 200 == 0 and is_cancelled and is_cancelled():
                raise _Cancelled

            record = tuple(
                plan.to_dbf(row[col_idx], encoding=encoding)
                for col_idx, plan in enumerate(plans)
            )
            table.append(record)
            written += 1
    except Exception:
        try:
            table.close()
        finally:
            _cleanup_file(path)
        raise
    else:
        table.close()

    # If UTF-8, write modern .cpg sidecar
    return SheetResult(
        sheet=name,
        status="ok",
        rows=written,
        fields=field_count,
        path=path,
        warnings=warnings,
    )


# ── Internal readers ────────────────────────────────────────────────────────

def _open_xlsx(source: Path):
    from openpyxl import load_workbook
    wb = load_workbook(filename=str(source), read_only=True, data_only=True)
    names = list(wb.sheetnames)

    def read_sheet(sheet_name: str) -> list[list[Any]]:
        ws = wb[sheet_name]
        out: list[list[Any]] = []
        for row in ws.iter_rows(values_only=True):
            r = list(row)
            while r and _is_blank(r[-1]):
                r.pop()
            if r:
                out.append(r)
        return out

    def close():
        wb.close()

    return names, read_sheet, close


def _open_xls(source: Path):
    import xlrd
    wb = xlrd.open_workbook(str(source), on_demand=True)
    names = list(wb.sheet_names())

    def read_sheet(sheet_name: str) -> list[list[Any]]:
        sh = wb.sheet_by_name(sheet_name)
        out: list[list[Any]] = []
        for r_idx in range(sh.nrows):
            row: list[Any] = []
            for c_idx in range(sh.ncols):
                cell = sh.cell(r_idx, c_idx)
                ctype = cell.ctype
                val = cell.value
                if ctype == xlrd.XL_CELL_DATE:
                    try:
                        val = xlrd.xldate_as_datetime(val, wb.datemode)
                    except Exception:
                        val = None
                elif ctype == xlrd.XL_CELL_BOOLEAN:
                    val = bool(val)
                elif ctype == xlrd.XL_CELL_ERROR:
                    val = xlrd.error_text_from_code.get(int(val), "#ERROR")
                elif ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    val = None
                row.append(val)
            while row and _is_blank(row[-1]):
                row.pop()
            if row:
                out.append(row)
        return out

    def close():
        wb.release_resources()

    return names, read_sheet, close


# ── Grid & Column Planning ──────────────────────────────────────────────────

def _trim_grid(rows: Sequence[Sequence[Any]]) -> list[list[Any]]:
    cleaned: list[list[Any]] = []
    for r in rows:
        row = list(r)
        while row and _is_blank(row[-1]):
            row.pop()
        cleaned.append(row)

    start = 0
    while start < len(cleaned) and not cleaned[start]:
        start += 1
    if start == len(cleaned):
        return []

    end = len(cleaned)
    while end > start and not cleaned[end - 1]:
        end -= 1
    cleaned = cleaned[start:end]

    width = max(len(r) for r in cleaned)
    if width == 0:
        return []

    padded = [r + [None] * (width - len(r)) for r in cleaned]
    keep = [c for c in range(width) if any(not _is_blank(padded[r_i][c]) for r_i in range(len(padded)))]
    if not keep:
        return []

    return [[row[c] for c in keep] for row in padded]


def _build_column_plan(
    name: str,
    values: Sequence[Any],
    *,
    encoding: str,
    warnings: list[str],
) -> _ColumnPlan:
    present = [v for v in values if not _is_blank(v)]
    if not present:
        return _ColumnPlan(name, "C", 1, mode="text")

    if all(isinstance(v, bool) for v in present):
        return _ColumnPlan(name, "L", 1, mode="logical")

    if all(_is_number(v) for v in present):
        plan = _numeric_plan(name, present)
        if plan is not None:
            return plan

    if all(isinstance(v, (dt.datetime, dt.date)) for v in present):
        if all(_is_midnight(v) for v in present):
            return _ColumnPlan(name, "D", 8, mode="date")
        # Has timestamps: format as text
        return _text_plan(
            name,
            [_format_datetime(v) for v in present],
            encoding=encoding,
            warnings=warnings,
            mode="datetime",
        )

    return _text_plan(
        name,
        [_format_as_text(v) for v in present],
        encoding=encoding,
        warnings=warnings,
        mode="text",
    )


def _numeric_plan(name: str, values: Sequence[Any]) -> _ColumnPlan | None:
    max_int_digits = 0
    max_decimals = 0
    max_significant = 0
    has_negative = False

    for v in values:
        try:
            num = Decimal(str(v))
        except (InvalidOperation, ValueError):
            return None
        if not num.is_finite():
            return None

        if num < 0:
            has_negative = True

        txt = format(num, "f")
        int_part, _, frac_part = txt.partition(".")
        int_digits = len(int_part.lstrip("-").lstrip("0")) or 1
        frac_digits = len(frac_part.rstrip("0"))

        max_int_digits = max(max_int_digits, int_digits)
        max_decimals = max(max_decimals, frac_digits)
        sig = len(num.normalize().as_tuple().digits)
        max_significant = max(max_significant, sig)

    # Protect IDs / barcodes from float precision loss:
    if max_significant > MAX_SAFE_SIGNIFICANT_DIGITS:
        return None

    decimals = max_decimals
    width = max_int_digits + (1 if decimals else 0) + decimals
    if has_negative:
        width += 1
    width = max(width, decimals + 2 if decimals else 1, 1)

    if decimals > MAX_NUMERIC_WIDTH - 2 or width > MAX_NUMERIC_WIDTH:
        return None

    return _ColumnPlan(name, "N", width, decimals, mode="number")


def _text_plan(
    name: str,
    values: Sequence[str],
    *,
    encoding: str,
    warnings: list[str],
    mode: str,
) -> _ColumnPlan:
    lengths = [_encoded_byte_length(v, encoding) for v in values]
    max_bytes = max(lengths, default=1)

    if max_bytes > MAX_DBF_CHAR_WIDTH:
        warnings.append(
            f"Column '{name}' has values exceeding {MAX_DBF_CHAR_WIDTH} bytes. Values truncated."
        )
        max_bytes = MAX_DBF_CHAR_WIDTH

    return _ColumnPlan(name, "C", max(max_bytes, 1), mode=mode)


# ── String & Naming Utilities ───────────────────────────────────────────────

def _unique_field_names(headers: Sequence[Any], count: int) -> list[str]:
    used: set[str] = set()
    names: list[str] = []
    for idx in range(count):
        raw = headers[idx] if idx < len(headers) else None
        base = _sanitize_field_name(raw, fallback=f"FIELD_{idx + 1}")
        candidate = base
        seq = 2
        while candidate.upper() in used:
            suffix = f"_{seq}"
            candidate = f"{base[:MAX_FIELD_NAME - len(suffix)]}{suffix}"
            seq += 1
        used.add(candidate.upper())
        names.append(candidate)
    return names


def _sanitize_field_name(value: Any, *, fallback: str) -> str:
    if value is None:
        return fallback

    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.strip().upper()
    text = _NON_FIELD_CHARS.sub("_", text)
    text = _MULTI_UNDERSCORE.sub("_", text).strip("_")

    if not text:
        return fallback
    if not text[0].isalpha():
        text = f"F_{text}"

    if len(text) > MAX_FIELD_NAME:
        compact = text.replace("_", "")
        text = compact if len(compact) <= MAX_FIELD_NAME else text[:MAX_FIELD_NAME]

    text = text.rstrip("_")
    if not text:
        return fallback
    if not text[0].isalpha():
        text = f"F{text}"[:MAX_FIELD_NAME]
    return text


def _unique_filename(sheet_name: str, used: set[str]) -> str:
    base = _sanitize_filename(sheet_name)
    candidate = base
    counter = 2
    while f"{candidate}.dbf".casefold() in used:
        candidate = f"{base} ({counter})"
        counter += 1
    fn = f"{candidate}.dbf"
    used.add(fn.casefold())
    return fn


def _sanitize_filename(name: str) -> str:
    text = _INVALID_FILENAME_CHARS.sub("_", str(name)).strip().rstrip(". ")
    if not text:
        text = "Sheet"
    if text.split(".")[0].upper() in _WINDOWS_RESERVED_NAMES:
        text = f"_{text}"
    if len(text) > 100:
        text = text[:100].rstrip(". ")
    return text or "Sheet"


def _cleanup_file(path: Path):
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


def _is_blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _is_number(v: Any) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float, Decimal)):
        return True
    return False


def _is_midnight(v: Any) -> bool:
    if isinstance(v, dt.datetime):
        return v.time() == dt.time(0, 0, 0)
    return isinstance(v, dt.date)


def _format_datetime(v: Any) -> str:
    if isinstance(v, dt.datetime):
        if v.time() == dt.time(0, 0, 0):
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    return str(v)


def _format_as_text(v: Any) -> str:
    if isinstance(v, (dt.datetime, dt.date)):
        return _format_datetime(v)
    if isinstance(v, float):
        try:
            d = Decimal(str(v))
            txt = format(d, "f")
            if "." in txt:
                txt = txt.rstrip("0").rstrip(".")
            return txt
        except Exception:
            return str(v)
    return str(v)


def _make_encodable(text: str, encoding: str) -> str:
    try:
        text.encode(encoding)
        return text
    except (UnicodeEncodeError, LookupError):
        return text.encode(encoding, errors="replace").decode(encoding)


def _encoded_byte_length(text: str, encoding: str) -> int:
    try:
        return len(text.encode(encoding, errors="replace"))
    except Exception:
        return len(text)


def _truncate_to_bytes(text: str, max_bytes: int, encoding: str) -> str:
    raw = text.encode(encoding, errors="replace")
    if len(raw) <= max_bytes:
        return text
    return raw[:max_bytes].decode(encoding, errors="ignore")


def _emit_log(log: Callable[[str], None] | None, msg: str):
    if log:
        log(msg)
