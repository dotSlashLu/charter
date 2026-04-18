"""Convert a DSL action tree into Excel sheets populated with formulas."""

from __future__ import annotations

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import log
from .dsl import (
    AggregationValue,
    AnyAction,
    FilterAction,
    GroupByAction,
    PivotAction,
    SelectAction,
    SortAction,
)
from .reader import SheetSchema

_log = log.get("executor")


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def execute_dsl(
    dsl: AnyAction,
    schema: SheetSchema,
    ws_source: Worksheet,
    wb: "Workbook",  # noqa: F821 – forward ref to openpyxl
    sheet_title: str,
) -> Worksheet:
    """Create a new sheet in *wb* named *sheet_title* driven by *dsl* formulas.

    Returns the newly created worksheet.
    """
    _log.debug("execute_dsl: action=%s  sheet='%s'", dsl.action, sheet_title)
    source_info = _resolve_source(dsl, schema, ws_source, wb)
    _log.debug(
        "  Source: sheet='%s'  columns=%s  data_rows=%d",
        source_info.sheet_name,
        source_info.columns,
        source_info.data_rows,
    )
    ws_out = wb.create_sheet(title=sheet_title[:31])

    match dsl:
        case SelectAction():
            _write_select(dsl, source_info, ws_out)
        case FilterAction():
            _write_filter(dsl, source_info, ws_out)
        case GroupByAction():
            _write_groupby(dsl, source_info, ws_out, ws_source)
        case PivotAction():
            _write_pivot(dsl, source_info, ws_out, ws_source)
        case SortAction():
            _write_sort(dsl, source_info, ws_out)

    out_rows = ws_out.max_row - 1 if ws_out.max_row else 0
    _log.debug("  Sheet '%s' written: %d data rows", ws_out.title, out_rows)
    return ws_out


# ---------------------------------------------------------------------------
# Source resolution helpers
# ---------------------------------------------------------------------------

class _SourceInfo:
    """Describes where source data lives (sheet name, columns, row range)."""

    def __init__(self, sheet_name: str, columns: list[str], col_indices: dict[str, int], data_rows: int):
        self.sheet_name = sheet_name
        self.columns = columns
        self.col_indices = col_indices  # column name -> 1-based index
        self.data_rows = data_rows      # number of data rows (excl header)

    def col_letter(self, col_name: str) -> str:
        return get_column_letter(self.col_indices[col_name])

    def range(self, col_name: str) -> str:
        """Absolute data range like ``Sheet1!A$2:A$1500``."""
        letter = self.col_letter(col_name)
        end = self.data_rows + 1  # +1 because row 1 is header
        return f"'{self.sheet_name}'!{letter}$2:{letter}${end}"

    def cell(self, col_name: str, data_row: int) -> str:
        """Absolute cell ref for a specific data row (1-based among data)."""
        letter = self.col_letter(col_name)
        return f"'{self.sheet_name}'!{letter}{data_row + 1}"


def _resolve_source(
    dsl: AnyAction,
    schema: SheetSchema,
    ws_source: Worksheet,
    wb: "Workbook",
) -> _SourceInfo:
    """If the DSL has a nested ``source``, materialise it first as a helper sheet."""
    inner = getattr(dsl, "source", None)
    if inner is None:
        _log.debug("  No nested source – using original sheet '%s'", schema.sheet_name)
        return _source_info_from_schema(schema)

    helper_title = f"_helper_{id(inner)}"[:31]
    _log.debug("  Resolving nested source (action=%s) -> helper sheet '%s'", inner.action, helper_title)
    helper_ws = execute_dsl(inner, schema, ws_source, wb, helper_title)
    helper_ws.sheet_state = "hidden"

    cols: list[str] = []
    col_map: dict[str, int] = {}
    for cell in next(helper_ws.iter_rows(min_row=1, max_row=1)):
        name = str(cell.value)
        cols.append(name)
        col_map[name] = cell.column

    data_rows = max(helper_ws.max_row - 1, 0)
    return _SourceInfo(helper_ws.title, cols, col_map, data_rows)


def _source_info_from_schema(schema: SheetSchema) -> _SourceInfo:
    cols = [c.name for c in schema.columns]
    col_map = {c.name: c.index for c in schema.columns}
    return _SourceInfo(schema.sheet_name, cols, col_map, schema.total_rows)


# ---------------------------------------------------------------------------
# Action writers
# ---------------------------------------------------------------------------

def _write_select(dsl: SelectAction, src: _SourceInfo, ws: Worksheet) -> None:
    _log.debug("  [select] columns=%s  distinct=%s", dsl.columns, dsl.distinct)
    for out_col_idx, col_name in enumerate(dsl.columns, start=1):
        out_letter = get_column_letter(out_col_idx)
        ws.cell(row=1, column=out_col_idx, value=col_name)

        if dsl.distinct:
            src_range = src.range(col_name)
            for data_row in range(1, src.data_rows + 1):
                excel_row = data_row + 1
                prev = f"${out_letter}$1:{out_letter}{excel_row - 1}"
                formula = (
                    f'=IFERROR(INDEX({src_range},'
                    f'MATCH(0,COUNTIF({prev},{src_range}),0)),"")'
                )
                ws.cell(row=excel_row, column=out_col_idx, value=formula)
            _log.debug("    '%s': %d distinct-formula rows", col_name, src.data_rows)
        else:
            for data_row in range(1, src.data_rows + 1):
                ref = src.cell(col_name, data_row)
                ws.cell(row=data_row + 1, column=out_col_idx, value=f"={ref}")
            _log.debug("    '%s': %d ref rows", col_name, src.data_rows)


def _write_filter(dsl: FilterAction, src: _SourceInfo, ws: Worksheet) -> None:
    _log.debug(
        "  [filter] conditions=%s  logic=%s",
        [(c.column, c.operator, c.value) for c in dsl.conditions],
        dsl.logic,
    )
    all_cols = src.columns
    for out_col_idx, col_name in enumerate(all_cols, start=1):
        ws.cell(row=1, column=out_col_idx, value=col_name)

    cond_parts: list[str] = []
    for cond in dsl.conditions:
        src_range = src.range(cond.column)
        val = f'"{cond.value}"' if isinstance(cond.value, str) else str(cond.value)
        op = cond.operator
        if op == "=":
            cond_parts.append(f"({src_range}={val})")
        elif op == "!=":
            cond_parts.append(f"({src_range}<>{val})")
        else:
            cond_parts.append(f"({src_range}{op}{val})")

    if dsl.logic == "AND":
        combined = "*".join(cond_parts)
    else:
        combined = "SIGN(" + "+".join(cond_parts) + ")"

    _log.debug("    Combined condition expression: %s", combined)

    first_col = all_cols[0]
    row_range = src.range(first_col)

    for out_col_idx, col_name in enumerate(all_cols, start=1):
        col_range = src.range(col_name)
        for data_row in range(1, src.data_rows + 1):
            excel_row = data_row + 1
            formula = (
                f'=IFERROR(INDEX({col_range},'
                f'SMALL(IF({combined},'
                f'ROW({row_range})-ROW(INDIRECT("'
                f"'{src.sheet_name}'!{src.col_letter(first_col)}$2\"))+1),"
                f'ROW(1:{data_row})),"")'
            )
            ws.cell(row=excel_row, column=out_col_idx, value=formula)

    _log.debug("    Wrote %d x %d filter formula cells", src.data_rows, len(all_cols))


def _write_groupby(
    dsl: GroupByAction,
    src: _SourceInfo,
    ws: Worksheet,
    ws_raw: Worksheet,
) -> None:
    _log.debug(
        "  [groupby] rows=%s  values=%s",
        dsl.rows,
        [(v.aggregation, v.column) for v in dsl.values],
    )
    unique_keys = _unique_values(ws_raw, [src.col_indices[r] for r in dsl.rows])
    _log.debug("    Found %d unique key combinations", len(unique_keys))
    if unique_keys:
        _log.debug("    First 5 keys: %s", unique_keys[:5])

    headers: list[str] = list(dsl.rows)
    for v in dsl.values:
        headers.append(f"{v.aggregation}_{v.column}")
    for idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=idx, value=h)

    dim_count = len(dsl.rows)

    for row_offset, key_tuple in enumerate(unique_keys):
        excel_row = row_offset + 2
        for d, val in enumerate(key_tuple):
            ws.cell(row=excel_row, column=d + 1, value=val)

        for v_idx, agg in enumerate(dsl.values):
            out_col = dim_count + v_idx + 1
            formula = _agg_formula(agg, dsl.rows, excel_row, src, dim_count)
            ws.cell(row=excel_row, column=out_col, value=formula)

    if unique_keys:
        sample_row = 2
        sample_col = dim_count + 1
        _log.debug("    Sample formula [row %d]: %s", sample_row, ws.cell(row=sample_row, column=sample_col).value)


def _write_pivot(
    dsl: PivotAction,
    src: _SourceInfo,
    ws: Worksheet,
    ws_raw: Worksheet,
) -> None:
    row_dims = dsl.rows
    col_dims = dsl.columns
    _log.debug(
        "  [pivot] row_dims=%s  col_dims=%s  values=%s",
        row_dims,
        col_dims,
        [(v.aggregation, v.column) for v in dsl.values],
    )

    row_keys = _unique_values(ws_raw, [src.col_indices[r] for r in row_dims])
    col_keys = _unique_values(ws_raw, [src.col_indices[c] for c in col_dims])
    _log.debug("    Row keys: %d  Col keys: %d", len(row_keys), len(col_keys))

    header_offset = len(row_dims)
    for d, name in enumerate(row_dims):
        ws.cell(row=1, column=d + 1, value=name)

    for ck_idx, ck in enumerate(col_keys):
        for v_idx, agg in enumerate(dsl.values):
            out_col = header_offset + ck_idx * len(dsl.values) + v_idx + 1
            label = " / ".join(str(x) for x in ck) + f" ({agg.aggregation}_{agg.column})"
            ws.cell(row=1, column=out_col, value=label)

    for rk_idx, rk in enumerate(row_keys):
        excel_row = rk_idx + 2
        for d, val in enumerate(rk):
            ws.cell(row=excel_row, column=d + 1, value=val)

        for ck_idx, ck in enumerate(col_keys):
            for v_idx, agg in enumerate(dsl.values):
                out_col = header_offset + ck_idx * len(dsl.values) + v_idx + 1
                formula = _pivot_formula(agg, row_dims, col_dims, rk, ck, src)
                ws.cell(row=excel_row, column=out_col, value=formula)

    total_cells = len(row_keys) * len(col_keys) * len(dsl.values)
    _log.debug("    Wrote %d pivot formula cells (%d rows x %d cols)", total_cells, len(row_keys), len(col_keys))


def _write_sort(dsl: SortAction, src: _SourceInfo, ws: Worksheet) -> None:
    _log.debug(
        "  [sort] orderBy=%s  direction=%s  limit=%s",
        dsl.order_by, dsl.direction, dsl.limit,
    )
    all_cols = src.columns
    for out_col_idx, col_name in enumerate(all_cols, start=1):
        ws.cell(row=1, column=out_col_idx, value=col_name)

    order_range = src.range(dsl.order_by)
    n_rows = dsl.limit if dsl.limit else src.data_rows

    first_col = all_cols[0]
    row_range = src.range(first_col)

    rank_fn = "SMALL" if dsl.direction == "ASC" else "LARGE"
    _log.debug("    rank_fn=%s  output_rows=%d", rank_fn, n_rows)

    for data_row in range(1, n_rows + 1):
        excel_row = data_row + 1
        for out_col_idx, col_name in enumerate(all_cols, start=1):
            col_range = src.range(col_name)
            formula = (
                f"=IFERROR(INDEX({col_range},"
                f"MATCH({rank_fn}({order_range},{data_row}),"
                f'{order_range},0)),"")'
            )
            ws.cell(row=excel_row, column=out_col_idx, value=formula)

    if n_rows > 0:
        _log.debug("    Sample formula [row 2]: %s", ws.cell(row=2, column=1).value)


# ---------------------------------------------------------------------------
# Formula helpers
# ---------------------------------------------------------------------------

_AGG_FN = {
    "SUM": "SUMIFS",
    "AVERAGE": "AVERAGEIFS",
    "COUNT": "COUNTIFS",
    "MAX": "MAXIFS",
    "MIN": "MINIFS",
}


def _agg_formula(
    agg: AggregationValue,
    dim_names: list[str],
    excel_row: int,
    src: _SourceInfo,
    dim_count: int,
) -> str:
    fn = _AGG_FN[agg.aggregation]
    val_range = src.range(agg.column)

    if fn == "COUNTIFS":
        parts = []
        for d, dim in enumerate(dim_names):
            crit_range = src.range(dim)
            crit_cell = f"${get_column_letter(d + 1)}${excel_row}"
            parts.append(f"{crit_range},{crit_cell}")
        return f"={fn}({','.join(parts)})"

    parts = [val_range]
    for d, dim in enumerate(dim_names):
        crit_range = src.range(dim)
        crit_cell = f"${get_column_letter(d + 1)}${excel_row}"
        parts.append(f"{crit_range},{crit_cell}")
    return f"={fn}({','.join(parts)})"


def _pivot_formula(
    agg: AggregationValue,
    row_dims: list[str],
    col_dims: list[str],
    row_key: tuple,
    col_key: tuple,
    src: _SourceInfo,
) -> str:
    fn = _AGG_FN[agg.aggregation]
    val_range = src.range(agg.column)

    criteria_parts: list[str] = []
    for dim, val in list(zip(row_dims, row_key)) + list(zip(col_dims, col_key)):
        crit_range = src.range(dim)
        crit_val = f'"{val}"' if isinstance(val, str) else str(val)
        criteria_parts.append(f"{crit_range},{crit_val}")

    if fn == "COUNTIFS":
        return f"={fn}({','.join(criteria_parts)})"

    return f"={fn}({val_range},{','.join(criteria_parts)})"


def _unique_values(ws: Worksheet, col_indices: list[int]) -> list[tuple]:
    """Extract unique tuples from the given column indices (1-based)."""
    seen: set[tuple] = set()
    result: list[tuple] = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        key = tuple(row[i - 1].value for i in col_indices)
        if key not in seen and any(v is not None for v in key):
            seen.add(key)
            result.append(key)
    result.sort(key=lambda t: tuple(str(x) for x in t))
    return result
