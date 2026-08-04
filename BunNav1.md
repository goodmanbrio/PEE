---
date created: 2026-08-04
date updated: 2026-08-04
date surroundings last checked: 2026-08-04
---

# BunNav1: Excel navigation protocol (tool-based)

## Purpose
Answer a point query from an xlsx file cheaply. Use only the six tools provided — `list_sheets`, `get_dims`, `map_structure`, `peek`, `read_axis`, `read_cell`. Every tool call costs money: this conversation's full history is resent to the model on every turn, so an oversized or unnecessary call gets paid for again and again, not once. Read exactly what you need, nothing "for context."

## Tools
- `list_sheets()` — sheet names + visible/hidden state. Always call first; use the exact name returned.
- `get_dims(sheet)` — `max_row`/`max_col`. Treat as an upper bound only — formatting artifacts often overstate the real extent. Confirm with `map_structure`.
- `map_structure(sheet)` — one call finds every table zone on the sheet: splits into column bands separated by fully blank columns, then splits each band into row bands separated by fully blank rows. Returns a list of `{rows:[r0,r1], cols:[c0,c1]}` bounding boxes. Always call this before `peek` — don't guess zone boundaries by eye.
- `peek(sheet, r0, r1, c0, c1)` — read real values in a small window. **Hard-capped at 50 cells total** — the call returns an error if you exceed it, no exceptions. Use it on a zone's top-left corner to identify the header row and label column.
- `read_axis(sheet, kind, index)` — `kind` is `"row"` or `"col"`. Dumps every non-blank value along that full row or column. Use once `peek` has told you which row is the period header, or which column holds row labels — not before. If a `"col"` read comes back mostly numeric, it's warning you that's probably a data column, not a label column — stop and recheck rather than trusting it.
- `read_cell(sheet, row, col)` — single value. Use for the final point-query lookup(s).

## Workflow
1. `list_sheets` — confirm the exact sheet name.
2. `get_dims` — rough upper bound only, not ground truth.
3. `map_structure` — get the real table zones in one call.
4. `peek` each candidate zone's corner (small window) to identify the header row / label column.
5. `read_axis` to confirm the full period axis and/or row-label axis, once you know which row/column they are.
6. `read_cell` for the target data point(s).
7. If the requested figure isn't a direct cell (e.g. a CAGR column exists but is blank for this row), pull the two endpoint values with `read_cell` and compute the figure yourself in your answer — no tool call needed for arithmetic.

## Judgment checklist (known failure modes)
- The header row is not always the top row of a zone — there may be a title row above it.
- A stray single label near the header row is not necessarily part of the real period axis — check whether data actually populates under it before trusting it.
- Row labels can be nested across multiple columns — check more than one column before committing to "the" label column.
- An all-text row with no numbers can be a blank spacer, not a section header — don't infer meaning from shape alone.
- A data block can share a header row with the zone above it if `map_structure` split them at a blank row — check upward if a zone has data but no header of its own.
- Derived columns (CAGR, YoY%, etc.) are sometimes populated for only some rows in a sheet — a blank cell there doesn't mean the figure doesn't exist, it may mean you compute it from two endpoints (workflow step 7).

## Constraints
- Every tool call is expensive — this session's entire history resends on every turn. Don't call a tool "just to see" or "for context" once you already have your answer.
- Never re-request a range you've already read — check your own earlier tool results before calling again. A repeated identical call will be flagged, not re-executed.
- You have a hard budget of tool calls this session. Reaching the answer efficiently is graded, not just reaching it at all.
