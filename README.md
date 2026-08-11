---
Created: 2026-08-11
Updated: 2026-08-11
Last checked: 2026-08-11
---

# PEE — Poony Excel Explorer

A toolset for LLM agents to answer point-queries against real, messy analyst-built xlsx models —
without dumping full sheets into context. The toolset is the product. `harness.py`'s DeepSeek
tool-calling loop exists only to make the tools runnable and testable; it is bare-minimum
scaffolding, not a deliverable in its own right.

## The tools

Ten tools, defined in `harness.py`'s `TOOLS` schema and `Workbook` class, callable by any agent
harness that speaks OpenAI-style function calling:

- `list_sheets` — visible sheet names, optionally scanned for metric/entity/year term matches to
  rank candidate sheets before opening any in detail.
- `get_dims` — a sheet's reported max row/col. Upper bound only.
- `map_block` — every table block on a sheet (blank-row/blank-column-delimited), with match counts
  per block. Default first structural call.
- `map_bin` — coarse whole-sheet minimap, for spatial/relative-position questions `map_block`'s
  flat list can't resolve (role-swap headers, transposed axes, adjacent column-band comparison).
- `peek_ascii` — full-resolution cell-type grid (not real values) for a bounded window.
- `peek` — real values in a small window.
- `peek_row` / `peek_col` — confirm one lineitem's or period's neighbors without a whole-axis dump.
- `all_axis` — every non-blank value along one full row or column, for when the whole axis is
  genuinely needed.
- `read_cell` — a single value.

Full tool/workflow reference: `BunNav1.md`.

## Status

Single provider — DeepSeek only, hardcoded (`DEEPSEEK_BASE_URL`, `DEEPSEEK_API_KEY`). The request
format is the OpenAI Chat Completions function-calling shape (DeepSeek's endpoint is
OpenAI-compatible), but nothing else here is provider-abstracted — no other backend is wired up.
The tools are still being actively iterated on; no multi-provider work is planned until that
settles.

## Setup

```
pip install openpyxl requests
```

Create `.env` in this directory:
```
DEEPSEEK_API_KEY=<your key>
```

Place the `.xlsx` files to query in this directory.

## Usage

```
python3 harness.py "<question>" <xlsx_filename> [model] [max_tool_calls]
```

`model` defaults to `deepseek-v4-pro`; `max_tool_calls` defaults to 40.
