#!/usr/bin/env python3
"""
date created: 2026-08-08
date updated: 2026-08-08
date surroundings last checked: 2026-08-08

Throwaway practicality test for wsnSearch4.md Item 2 section 2 (peek_row/peek_col).
Reuses harness.py's actual _detect_blocks/_merge_blocks (not a reimplementation) so
"containing block" here means exactly what map_block would report to the model.
"""
import sys
import warnings
warnings.filterwarnings("ignore")

from harness import Workbook, _detect_blocks, _merge_blocks

WINDOW_RADIUS = 20
CHAR_LIMIT = 1000


def find_entry(entries, row, col):
    for e in entries:
        r0, r1 = e["rows"]
        c0, c1 = e["cols"]
        if r0 <= row <= r1 and c0 <= col <= c1:
            return e
    return None


def _trim_to_char_limit(values_out, anchor_index, char_limit):
    rendered = repr(values_out)
    if len(rendered) <= char_limit:
        return values_out, None
    order = sorted(values_out.keys(), key=lambda k: abs(k - anchor_index))
    trimmed = {}
    total = 0
    for k in order:
        piece_len = len(repr({k: values_out[k]}))
        if total + piece_len > char_limit:
            break
        trimmed[k] = values_out[k]
        total += piece_len
    if not trimmed:
        return trimmed, "showing none (single value exceeds char limit)"
    shown = sorted(trimmed.keys())
    return trimmed, f"showing {shown[0]}-{shown[-1]}"


def peek_axis(values, n_rows, n_cols, row, col, fixed_is_col):
    """fixed_is_col=True -> peek_col (anchor's column fixed, scan its row range).
    fixed_is_col=False -> peek_row (anchor's row fixed, scan its col range)."""
    entries = _merge_blocks(_detect_blocks(values, n_rows, n_cols))
    entry = find_entry(entries, row, col)
    note = None
    if entry is None:
        if fixed_is_col:
            lo = max(1, row - WINDOW_RADIUS)
            hi = min(n_rows, row + WINDOW_RADIUS)
        else:
            lo = max(1, col - WINDOW_RADIUS)
            hi = min(n_cols, col + WINDOW_RADIUS)
        note = f"no block detected, showing {lo}-{hi}"
        block_span = None
    else:
        r0, r1 = entry["rows"]
        c0, c1 = entry["cols"]
        if fixed_is_col:
            span = r1 - r0 + 1
            block_span = span
            if span <= 2 * WINDOW_RADIUS + 1:
                lo, hi = r0, r1
            else:
                lo = max(r0, row - WINDOW_RADIUS)
                hi = min(r1, row + WINDOW_RADIUS)
                note = f"block spans {span} rows, showing {lo}-{hi} centered on the requested cell"
        else:
            span = c1 - c0 + 1
            block_span = span
            if span <= 2 * WINDOW_RADIUS + 1:
                lo, hi = c0, c1
            else:
                lo = max(c0, col - WINDOW_RADIUS)
                hi = min(c1, col + WINDOW_RADIUS)
                note = f"block spans {span} cols, showing {lo}-{hi} centered on the requested cell"

    values_out = {}
    if fixed_is_col:
        for r in range(lo, hi + 1):
            v = values[r - 1][col - 1]
            if v is not None:
                values_out[r] = v
        anchor_index = row
    else:
        for c in range(lo, hi + 1):
            v = values[row - 1][c - 1]
            if v is not None:
                values_out[c] = v
        anchor_index = col

    values_out, trim_note = _trim_to_char_limit(values_out, anchor_index, CHAR_LIMIT)
    if trim_note:
        note = trim_note
    elif not values_out:
        axis_word = "column" if fixed_is_col else "row"
        note = f"{axis_word} {col if fixed_is_col else row} is blank throughout the {lo}-{hi} range checked"

    result = {"values": values_out}
    if note:
        result["note"] = note
    result["_block_span_found"] = block_span
    result["_entry"] = entry
    return result


CASES = [
    dict(kind="peek_col", label="Advantest BOE lineitem block (confirm siblings) -- col=1, WRONG anchor",
         file="[Quarter added]Advantest 6857 JP BOE.xlsx", sheet="Advantest BOE", row=86, col=1),
    dict(kind="peek_col", label="Advantest BOE lineitem block (confirm siblings) -- col=2, real label col",
         file="[Quarter added]Advantest 6857 JP BOE.xlsx", sheet="Advantest BOE", row=86, col=2),
    dict(kind="peek_row", label="Wuxi Consolidate period row (confirm date col, from C-test)",
         file="300450_Wuxi Lead_20210422_client.xlsx", sheet="Consolidate", row=2, col=13),
    dict(kind="peek_row", label="Parade Shipment period row (confirm date col, from C-test)",
         file="20220404_Parade_BOE.xlsx", sheet="Shipment", row=2, col=7),
    dict(kind="peek_row", label="Advantest period row (confirm date col, from C-test)",
         file="[Quarter added]Advantest 6857 JP BOE.xlsx", sheet="Advantest BOE", row=3, col=29),
    dict(kind="peek_col", label="99-00 Study Raw big contiguous table (expect >41 windowing)",
         file="99-00 Study Raw.xlsx", sheet="Existing", row=100, col=7),
    dict(kind="peek_col", label="Chinese cash-burn lineitem block",
         file="2.元戎启行资金消耗统计 2019.2-2021.4.xlsx", sheet="元戎启行资金消耗统计", row=4, col=2),
    dict(kind="peek_col", label="WaferSupplyDemand Disclaimer megablock (expect char-limit trim)",
         file="20220204_WaferSupplyDemandModel.xlsx", sheet="Disclaimer", row=5, col=1),
    dict(kind="peek_row", label="99-00 Study Raw header row, full block <=41 cols but long datetimes (expect char-limit trim)",
         file="99-00 Study Raw.xlsx", sheet="Existing", row=4, col=20),
]

wb_cache = {}

def load(fname):
    if fname not in wb_cache:
        wb_cache[fname] = Workbook(fname)
    return wb_cache[fname]


for case in CASES:
    wb = load(case["file"])
    values, n_rows, n_cols, true_max_row, true_max_col = wb._scan(case["sheet"])
    fixed_is_col = case["kind"] == "peek_col"
    res = peek_axis(values, n_rows, n_cols, case["row"], case["col"], fixed_is_col)
    print("=" * 100)
    print(f"[{case['kind']}] {case['label']}")
    print(f"  {case['file']} | {case['sheet']} | anchor row={case['row']} col={case['col']}  (scanned {n_rows}x{n_cols}, true {true_max_row}x{true_max_col})")
    entry = res["_entry"]
    print(f"  containing block: {entry['rows'] if entry else None} rows x {entry['cols'] if entry else None} cols  (span found: {res['_block_span_found']})")
    print(f"  note: {res.get('note')}")
    vals = res["values"]
    print(f"  values returned: {len(vals)} entries, char length {len(repr(vals))}")
    shown = sorted(vals.items())[:6]
    for k, v in shown:
        print(f"    {k}: {v!r}")
    if len(vals) > 6:
        print(f"    ... ({len(vals)-6} more)")
