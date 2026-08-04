#!/usr/bin/env python3
"""
date created: 2026-08-04
date updated: 2026-08-04
date surroundings last checked: 2026-08-04
"""
import json
import os
import sys

import openpyxl
import requests

HARNESS_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HARNESS_DIR, ".env")
PROTOCOL_PATH = os.path.join(HARNESS_DIR, "BunNav1.md")

DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1/chat/completions"
DEFAULT_MODEL = "deepseek-v4-pro"

MAX_TOOL_CALLS = 40
PEEK_MAX_CELLS = 50
ROW_SCAN_CAP = 500
COL_SCAN_CAP = 300
MAX_TOOL_OUTPUT_CHARS = 6000

EXCEL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}


def is_excel_error(v):
    return isinstance(v, str) and v in EXCEL_ERRORS

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "list_sheets",
            "description": "List visible sheet names (hidden sheets are not accessible — same as what an analyst opening this file in Excel would see by default).",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_dims",
            "description": "Get max_row/max_col for a sheet. Upper bound only, not the real extent — confirm with map_structure.",
            "parameters": {
                "type": "object",
                "properties": {"sheet": {"type": "string"}},
                "required": ["sheet"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "map_structure",
            "description": "Find every table zone on a sheet in one call: column bands split by blank columns, then row bands split by blank rows within each band. Returns bounding boxes. Call before peek.",
            "parameters": {
                "type": "object",
                "properties": {"sheet": {"type": "string"}},
                "required": ["sheet"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "peek",
            "description": "Read real values in a small window. HARD CAP: (r1-r0+1)*(c1-c0+1) must be <= 50 — compute this before calling. Fails with a ready-to-use suggested window if you exceed it.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet": {"type": "string"},
                    "r0": {"type": "integer"},
                    "r1": {"type": "integer"},
                    "c0": {"type": "integer"},
                    "c1": {"type": "integer"},
                },
                "required": ["sheet", "r0", "r1", "c0", "c1"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "read_axis",
            "description": "Dump every non-blank value along one full row or column. kind is 'row' or 'col'. Use once peek has confirmed which row/column is the axis you need.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet": {"type": "string"},
                    "kind": {"type": "string", "enum": ["row", "col"]},
                    "index": {"type": "integer"},
                },
                "required": ["sheet", "kind", "index"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "read_cell",
            "description": "Read a single cell value.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet": {"type": "string"},
                    "row": {"type": "integer"},
                    "col": {"type": "integer"},
                },
                "required": ["sheet", "row", "col"],
            },
        },
    },
]


def load_env(path):
    env = {}
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
                value = value[1:-1]
            env[key.strip()] = value
    return env


class Workbook:
    def __init__(self, path):
        self.wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    def _check_visible(self, sheet):
        if sheet not in self.wb.sheetnames:
            raise ValueError(f"no sheet named '{sheet}'")
        if self.wb[sheet].sheet_state != "visible":
            raise ValueError(f"sheet '{sheet}' is hidden and not accessible — same as what an analyst opening this file would see")

    def list_sheets(self):
        return [{"name": n} for n in self.wb.sheetnames if self.wb[n].sheet_state == "visible"]

    def get_dims(self, sheet):
        self._check_visible(sheet)
        ws = self.wb[sheet]
        return {
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "note": "upper bound only; formatting artifacts often overstate real extent — confirm with map_structure",
        }

    def _type_grid(self, ws, max_row, max_col):
        def code(v):
            if v is None:
                return "."
            if isinstance(v, str):
                return "S"
            if isinstance(v, (int, float)):
                return "N"
            return "?"

        grid = []
        for r in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True):
            grid.append([code(v) for v in r])
        return grid

    def map_structure(self, sheet):
        self._check_visible(sheet)
        ws = self.wb[sheet]
        max_row = min(ws.max_row, ROW_SCAN_CAP)
        max_col = min(ws.max_column, COL_SCAN_CAP)
        grid = self._type_grid(ws, max_row, max_col)
        n_rows = len(grid)
        n_cols = len(grid[0]) if n_rows else 0
        if not n_cols:
            return {"zones": [], "note": "sheet appears empty within scan bounds"}

        blank_col = [all(grid[r][c] == "." for r in range(n_rows)) for c in range(n_cols)]

        col_bands = []
        c = 0
        while c < n_cols:
            if blank_col[c]:
                c += 1
                continue
            start = c
            while c < n_cols and not blank_col[c]:
                c += 1
            col_bands.append((start + 1, c))

        zones = []
        for (c0, c1) in col_bands:
            r = 0
            while r < n_rows:
                row_blank = all(grid[r][cc - 1] == "." for cc in range(c0, c1 + 1))
                if row_blank:
                    r += 1
                    continue
                start = r
                while r < n_rows and not all(grid[r][cc - 1] == "." for cc in range(c0, c1 + 1)):
                    r += 1
                zones.append({"rows": [start + 1, r], "cols": [c0, c1]})

        groups = {}
        for z in zones:
            groups.setdefault(tuple(z["cols"]), []).append(z)

        hints = []
        FRAGMENT_THRESHOLD = 4
        for cols, zs in groups.items():
            if len(zs) >= FRAGMENT_THRESHOLD:
                row_min = zs[0]["rows"][0]
                row_max = zs[-1]["rows"][1]
                hints.append(
                    f"{len(zs)} zones share cols {list(cols)} across rows {row_min}-{row_max} — "
                    f"if these look like one table with many line items rather than separate tables, "
                    f"try read_axis(col=<label column>) once instead of peeking each zone."
                )

        result = {"zones": zones, "scanned_rows": n_rows, "scanned_cols": n_cols}
        if hints:
            result["hints"] = hints
        if ws.max_row > ROW_SCAN_CAP or ws.max_column > COL_SCAN_CAP:
            result["note"] = f"scan bounded to {ROW_SCAN_CAP} rows x {COL_SCAN_CAP} cols; re-run peek beyond this if needed"
        return result

    def peek(self, sheet, r0, r1, c0, c1):
        self._check_visible(sheet)
        n_cells = (r1 - r0 + 1) * (c1 - c0 + 1)
        if n_cells > PEEK_MAX_CELLS:
            n_cols_req = c1 - c0 + 1
            n_rows_req = r1 - r0 + 1
            if n_cols_req <= PEEK_MAX_CELLS:
                max_rows = max(1, PEEK_MAX_CELLS // n_cols_req)
                suggested = {"r0": r0, "r1": min(r1, r0 + max_rows - 1), "c0": c0, "c1": c1}
            else:
                max_cols = max(1, PEEK_MAX_CELLS // max(1, n_rows_req))
                suggested = {"r0": r0, "r1": r0, "c0": c0, "c1": min(c1, c0 + max_cols - 1)}
            return {
                "error": f"requested {n_cells} cells exceeds the {PEEK_MAX_CELLS}-cell cap "
                         f"((r1-r0+1)*(c1-c0+1) must be <= {PEEK_MAX_CELLS}). Narrow the range.",
                "suggested_window": suggested,
            }
        ws = self.wb[sheet]
        rows = list(ws.iter_rows(min_row=r0, max_row=r1, min_col=c0, max_col=c1, values_only=True))
        result = {"rows": [list(r) for r in rows]}
        error_count = sum(1 for row in rows for v in row if is_excel_error(v))
        if error_count:
            result["warning"] = (
                f"{error_count} cell(s) in this window are formula errors (#NAME?/#DIV/0!/etc.) — "
                f"likely broken external data links, not real data"
            )
        return result

    def read_axis(self, sheet, kind, index):
        self._check_visible(sheet)
        ws = self.wb[sheet]
        if kind == "row":
            max_col = min(ws.max_column, COL_SCAN_CAP)
            vals = list(ws.iter_rows(min_row=index, max_row=index, min_col=1, max_col=max_col, values_only=True))[0]
        elif kind == "col":
            max_row = min(ws.max_row, ROW_SCAN_CAP)
            vals = [row[0] for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=index, max_col=index, values_only=True)]
        else:
            return {"error": "kind must be 'row' or 'col'"}

        values = {i + 1: v for i, v in enumerate(vals) if v is not None}
        result = {"values": values}
        warnings = []
        if kind == "col" and values:
            numeric_frac = sum(isinstance(v, (int, float)) for v in values.values()) / len(values)
            if numeric_frac > 0.5:
                warnings.append("majority of values are numeric — this may be a data column, not a label column")
        error_count = sum(1 for v in values.values() if is_excel_error(v))
        if error_count:
            warnings.append(
                f"{error_count} value(s) on this axis are formula errors (#NAME?/#DIV/0!/etc.) — "
                f"likely broken external data links, not real data"
            )
        if warnings:
            result["warnings"] = warnings
        return result

    def read_cell(self, sheet, row, col):
        self._check_visible(sheet)
        ws = self.wb[sheet]
        v = list(ws.iter_rows(min_row=row, max_row=row, min_col=col, max_col=col, values_only=True))[0][0]
        result = {"value": v}
        if is_excel_error(v):
            result["error"] = True
            result["note"] = "this is a formula error, likely a broken external data link (e.g. disconnected Bloomberg add-in), not real data"
        return result


def call_deepseek(api_key, model, messages):
    resp = requests.post(
        DEEPSEEK_BASE_URL,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={"model": model, "messages": messages, "tools": TOOLS},
        timeout=120,
    )
    resp.raise_for_status()
    return resp.json()


def run(task, model, xlsx_path, max_tool_calls=MAX_TOOL_CALLS):
    env = load_env(ENV_PATH)
    api_key = env.get("DEEPSEEK_API_KEY")
    if not api_key:
        sys.exit("DEEPSEEK_API_KEY not found in .env")

    with open(PROTOCOL_PATH) as f:
        protocol = f.read()

    workbook = Workbook(xlsx_path)
    dispatch = {
        "list_sheets": lambda **kw: workbook.list_sheets(),
        "get_dims": lambda **kw: workbook.get_dims(**kw),
        "map_structure": lambda **kw: workbook.map_structure(**kw),
        "peek": lambda **kw: workbook.peek(**kw),
        "read_axis": lambda **kw: workbook.read_axis(**kw),
        "read_cell": lambda **kw: workbook.read_cell(**kw),
    }

    system_prompt = "Follow this protocol exactly. Use only the tools provided.\n\n" + protocol
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": task},
    ]

    seen_calls = {}
    tool_call_count = 0
    usage_totals = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}
    api_turns = 0

    while True:
        data = call_deepseek(api_key, model, messages)
        api_turns += 1
        usage = data.get("usage", {})
        for k in usage_totals:
            usage_totals[k] += usage.get(k, 0)

        message = data["choices"][0]["message"]
        messages.append(message)

        tool_calls = message.get("tool_calls")
        if not tool_calls:
            print(message.get("content", ""))
            break

        budget_exhausted = False
        for tc in tool_calls:
            name = tc["function"]["name"]
            args = json.loads(tc["function"]["arguments"] or "{}")

            if tool_call_count >= max_tool_calls:
                budget_exhausted = True
                output = {"error": "tool call budget exhausted, stop and answer with what you have"}
            else:
                key = (name, json.dumps(args, sort_keys=True))
                if key in seen_calls:
                    output = {"note": "duplicate call — identical to an earlier call this session, reuse that result instead of re-reading"}
                else:
                    tool_call_count += 1
                    try:
                        output = dispatch[name](**args)
                    except Exception as e:
                        output = {"error": str(e)}
                    seen_calls[key] = True

            output_str = json.dumps(output, default=str)
            if len(output_str) > MAX_TOOL_OUTPUT_CHARS:
                output_str = output_str[:MAX_TOOL_OUTPUT_CHARS] + f"... [truncated, {len(output_str)} chars total]"

            print(f"\n[{tool_call_count}/{max_tool_calls}] {name}({args}) ->\n{output_str}")

            messages.append({"role": "tool", "tool_call_id": tc["id"], "content": output_str})

        if budget_exhausted:
            print("\n[stopped: tool call budget exhausted]")
            break

    print(
        f"\n--- usage: {api_turns} API turns, {tool_call_count} tool calls, "
        f"{usage_totals['prompt_tokens']} prompt tokens, {usage_totals['completion_tokens']} completion tokens, "
        f"{usage_totals['total_tokens']} total tokens ---"
    )


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit('usage: python3 harness.py "<task>" <xlsx_filename> [model] [max_tool_calls]')
    task_arg = sys.argv[1]
    xlsx_arg = os.path.join(HARNESS_DIR, sys.argv[2])
    model_arg = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_MODEL
    max_calls_arg = int(sys.argv[4]) if len(sys.argv) > 4 else MAX_TOOL_CALLS
    run(task_arg, model_arg, xlsx_arg, max_calls_arg)
