---
date created: 2026-08-04
date updated: 2026-08-07
date surroundings last checked: 2026-08-07
---

# BunNav1: Excel navigation protocol (tool-based)

## Purpose
Answer a point query from an xlsx file cheaply. Use only the eight tools provided — `list_sheets`, `get_dims`, `map_block`, `map_bin`, `peek_ascii`, `peek`, `read_axis`, `read_cell`. Every tool call costs money: this conversation's full history is resent to the model on every turn, so an oversized or unnecessary call gets paid for again and again, not once. Read exactly what you need, nothing "for context."

## Shared flag vocabulary
Four tools share one flag vocabulary, computed by the same classifier at scan time:
- `D` — date-like (matched a period-string pattern, e.g. `"2Q21"`, `"2022E"`, `"FY2017"`, or the cell is a native Excel date/time value). Always computed, independent of query.
- `A` — matches your `metric_term`, if you supplied one.
- `B` — matches your `entity_term`, if you supplied one.

`metric_term`/`entity_term` are plain case-insensitive substring matches against cell text — not literal fuzzy/edit-distance matching. Pick short, distinctive substrings from the question (e.g. `"Wafer"`, `"300mm"`), not whole phrases.

Each tool aggregates and renders this vocabulary at a different resolution — sheet-level counts, block-level counts, bin-level presence flags, or per-cell letters. If a cell matches more than one condition (e.g. a cell reading `"300mm wafer"` matches both a metric term and an entity term), sheet/block/bin counts and flags treat `D`/`A`/`B` as independent — the same cell can count toward more than one. Only `peek_ascii`'s single-letter-per-cell rendering has to pick one; there the priority order is metric (`A`) > entity (`B`) > date (`D`) > plain (`S`/`N`) — query-relevance outranks generic type info, since the tool is being consulted specifically to answer that query.

## Tools
- `list_sheets(metric_term?, entity_term?)` — sheet names + visible/hidden state. Always call first; use the exact name returned. With `metric_term`/`entity_term`, also scans every visible sheet and adds raw `D` count (always), `A` count (only if `metric_term` given), `B` count (only if `entity_term` given), plus `scanned_rows`/`scanned_cols`. A key is omitted, not shown as `0`, if you didn't supply that term. These are **raw counts, not a ranking score** — don't just pick the highest-count sheet. A sheet can score highest simply because the term *is* its own subject (e.g. "Wafer" hits 51 times on a sheet literally about wafer supply/demand) without being the sheet that answers your question. Weigh count against `scanned_rows`/`scanned_cols` — a small sheet with a moderate count is often a better first look than a huge sheet with a high count.
- `get_dims(sheet)` — `max_row`/`max_col`. Treat as an upper bound only — formatting artifacts often overstate the real extent. Confirm with `map_block` or `map_bin`.
- `map_block(sheet, metric_term?, entity_term?)` — **default first structural call**, before `map_bin`. Finds every table block on the sheet (column bands split by fully blank columns, then row blocks split by fully blank rows within each band) and renders a compact line list: one entry per block, with its exact row/col range and raw `D`/`A`/`B` counts over that range. Blocks separated by only 1-3 blank rows in the same column band are merged into one entry, shown with a dash-range block number (e.g. `1-3`) — that's your signal more than one real block is bundled there, not a renumbering. A block that didn't merge keeps its own single number. No size/cell-count field — if a single entry is too big for one `peek_ascii` call (225-cell cap), just peek it in two calls.
- `map_bin(sheet, metric_term?, entity_term?)` — coarse, whole-sheet, uniform 10x10-cell bins. **Call this only after `map_block`, and only if its flat list doesn't answer a spatial/relative-position question** (role-swap headers, transposed axes, comparing two column bands side by side) — a grid conveys relative position that a line list can't. Each bin renders as a **4-character code**: position 1 is structural (`.` blank bin, `:` content but not a corner, `#` a block's exact top-left corner), positions 2-4 are `D`/`-`, `A`/`-`, `B`/`-` — the letter if that condition is true anywhere in the bin, else `-`. A blank bin is always `.---`. Real row/col numbers label each bin boundary — use them directly as coordinates into `peek_ascii`. This is coarse — it tells you *where* to look and roughly *what kind of content* is there, not exact values.
- `peek_ascii(sheet, r0, r1, c0, c1, metric_term?, entity_term?)` — full-resolution step between `map_block`/`map_bin` and `peek`. Renders the actual cell-*type* grid (not real values), one letter per cell: `.` blank, `A` metric-term match, `B` entity-term match, `D` date-like, `S` string, `N` number, `?` other (priority order `A` > `B` > `D` > `S`/`N`/`?` if more than one applies). **Hard-capped at 225 cells total** ((r1-r0+1)*(c1-c0+1) — same error-on-exceed pattern as `peek`). Use it to find the header row inside a block/bin: a dense run of `D` is very likely the period axis, regardless of where it sits. A dense run of `A`/`B` narrows straight to the rows/columns your query terms actually hit. Any of these is a density signal, not a verdict — confirm with `peek` before trusting it.
- `peek(sheet, r0, r1, c0, c1)` — read real values in a small window. **Hard-capped at 50 cells total** — the call returns an error if you exceed it, no exceptions. Use it once `peek_ascii` (or a `map_block`/`map_bin` flag) has pointed you at a specific row/column to confirm what's actually there.
- `read_axis(sheet, kind, index)` — `kind` is `"row"` or `"col"`. Dumps every non-blank value along that full row or column. Use once `peek`/`peek_ascii` has told you which row is the period header, or which column holds row labels — not before. If a `"col"` read comes back mostly numeric, it's warning you that's probably a data column, not a label column — stop and recheck rather than trusting it.
- `read_cell(sheet, row, col)` — single value. Use for the final point-query lookup(s).

## Workflow
1. `list_sheets(metric_term, entity_term)` — pick a short, distinctive term for the metric and entity in the question. Rank candidate sheets by count vs. size, not raw count alone.
2. `get_dims` — rough upper bound only, not ground truth.
3. `map_block(sheet, metric_term, entity_term)` — default first structural call. Look at which block(s) carry `D`/`A`/`B` counts, and how big each block's row/col range is relative to the 225-cell `peek_ascii` cap.
4. `map_bin(sheet, metric_term, entity_term)` — only if `map_block`'s flat list doesn't resolve a spatial/relative-position question (role-swap headers, transposed axes, comparing adjacent column bands). Not a required step every time.
5. `peek_ascii` a candidate block/bin to see cell-type density without spending a `peek` call — look for a dense run of `D` to find the period header row, a dense run of `A`/`B` to find where your query terms actually land, or a dense run of `S` in one column for the label column.
6. `peek` the specific row/column step 5 pointed at to confirm real values (header text, label text).
7. `read_axis` to confirm the full period axis and/or row-label axis, once you know which row/column they are.
8. `read_cell` for the target data point(s).
9. If the requested figure isn't a direct cell (e.g. a CAGR column exists but is blank for this row), pull the two endpoint values with `read_cell` and compute the figure yourself in your answer — no tool call needed for arithmetic.

## Judgment checklist (known failure modes)
- Match a header to its value by shared column number, never by counting position across a row.
- The header row is not always the top row of a block — there may be a title row above it. This is exactly why the workflow scans for `D` density with `peek_ascii` instead of assuming the header sits at a block's top row.
- A single `D` cell near a real header is not necessarily part of the period axis — check whether data actually populates under it before trusting it.
- `D` density is a signal, not a classifier: a year can be stored as a plain number (`N`, not `D`) in one column and as a string (`D`) in the next — don't rule out a numeric-typed period row just because it shows `N` instead of `D`.
- A CAGR/range column (e.g. `"2019-21"`) or a scenario label (e.g. `"Bull '24E"`) will also flag `D` — that's correct, it contains a period reference, not necessarily a repeating period header. Verify with `peek`.
- `A`/`B` are plain substring matches, not semantic understanding — a generic term that's the sheet's own subject (e.g. "Wafer" on a wafer supply/demand sheet) will match everywhere and give weak disambiguation no matter how it's rendered; and an unrelated cell can match by coincidence (e.g. an entity term "Cell" matching a "Miscellaneous Assets" block, or an instructional footnote). A match is a candidate location, not a confirmed answer — always confirm with `peek`.
- `A`/`B` only match the literal substring given — a non-Latin-script rendering of the same term (e.g. a Japanese-language mirror column) won't match an English `metric_term`/`entity_term`. Don't conclude "not here" from an all-`-` region without checking whether the sheet has a same-content column in another language.
- Row labels can be nested across multiple columns — check more than one column before committing to "the" label column.
- An all-text row with no numbers can be a blank spacer, not a section header — don't infer meaning from shape alone.
- A data block can share a header row with the block above it if `map_block`/`map_bin` split them at a blank row — check upward if a block has data but no header of its own.
- One header row can serve several disconnected blocks below it, not one header per block — don't assume every block needs its own `peek_ascii`/`peek` pass if an earlier one already covers it.
- Derived columns (CAGR, YoY%, etc.) are sometimes populated for only some rows in a sheet — a blank cell there doesn't mean the figure doesn't exist, it may mean you compute it from two endpoints (workflow step 9).

## Constraints
- Every tool call is expensive — this session's entire history resends on every turn. Don't call a tool "just to see" or "for context" once you already have your answer.
- Never re-request a range you've already read — check your own earlier tool results before calling again. A repeated identical call will be flagged, not re-executed.
- You have a hard budget of tool calls this session. Reaching the answer efficiently is graded, not just reaching it at all.
