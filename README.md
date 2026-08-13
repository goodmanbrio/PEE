---
Created: 2026-08-11
Updated: 2026-08-13
Last checked: 2026-08-13
---

# PEE — Poony Excel Explorer

A toolset for LLM agents to answer point-queries against real, messy analyst-built xlsx models —
without dumping full sheets into context. The toolset is the product. `harness.py`'s DeepSeek
tool-calling loop exists only to make the tools runnable and testable; it is bare-minimum
scaffolding, not a deliverable in its own right.

## The tools

Fourteen tools, defined in `harness.py`'s `TOOLS` schema and `Workbook` class, callable by any
agent harness that speaks OpenAI-style function calling. Number cell values are always shown
capped to 3 significant figures — a display cap, not the real stored precision.

- `set_terms` — fix a metric/entity/year term once for the rest of the session; the tools below
  fall back to it whenever their own term args are omitted.
- `list_sheets` — visible sheet names, optionally scanned for metric/entity/year term matches to
  rank candidate sheets before opening any in detail. `metric_term`/`entity_term` each accept a
  list of candidate wordings instead of one string, when the sheet's own wording is unknown.
- `get_dims` — a sheet's reported max row/col. Upper bound only.
- `map_block` — every table block on a sheet (blank-row/blank-column-delimited), with match counts
  per block, plus AB-run/CD-run: candidate lineitem/period axes found directly in that block, so a
  confirmed candidate can skip straight to `check_axes` instead of a blind `peek_ascii` sweep.
  Default first structural call.
- `map_bin` — coarse whole-sheet minimap, for spatial/relative-position questions `map_block`'s
  flat list can't resolve (role-swap headers, transposed axes, adjacent column-band comparison).
- `find_matches` — exact coordinates for a term across a sheet or region, when `map_block`/
  `map_bin` narrowed to "somewhere in here" but not further.
- `check_axes` — confirm a `map_block`-surfaced axis candidate is real: renders the block with the
  candidate row/column's actual values shown, everything else as a type flag.
- `peek_axes` — read a confirmed candidate axis's own values directly, no rendering, no flags.
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

`BunNav1.md` (the system prompt) plus the `TOOLS` schema — both resent in full on every turn — are
now about 42,400 characters combined, roughly 2.6x their size when this harness was first built.
Each new tool and each round of sharper tool descriptions adds to a cost paid on every single API
call, not once. Not yet addressed; trimming this down is planned for a future version.

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
