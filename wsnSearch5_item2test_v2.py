#!/usr/bin/env python3
"""
date created: 2026-08-09
date updated: 2026-08-09
date surroundings last checked: 2026-08-09

v2 of wsnSearch5_item2test.py, correcting the AB-run/CD-run hit conditions per user
feedback (v1 used the doc's literal text -- metric/entity-only, C-or-D-only -- and found
it essentially never fires on real single-item queries):

  AB hit-set (run-continuation): A, B, or plain string (S). Run qualifies as a candidate
  only if it's length >=3 AND contains >=1 real A/B hit (not pure S) -- blocks unrelated
  text columns/rows from qualifying on string-ness alone.

  CD hit-set: D (date-like) or plain string (S) -- C (year_term match) dropped entirely,
  since a bare "matches my exact target year" signal only ever fires on one cell and can't
  form a run on its own; D now carries the real period-detection weight. Run qualifies only
  if length >=3 AND contains >=1 real D hit (not pure S) -- blocks false positives like a
  bare "Forecast Forecast Forecast" title run (found in Wuxi Op & Non op row 1) from
  registering as a period candidate.

  is_date_like is reimplemented LOCALLY here with a sequential-year rule added (a plain
  integer counts as date-like if it's part of a run of >=3 consecutive cells, row-wise or
  col-wise, each exactly +1 from the previous, in range 1980-2035) -- NOT edited in
  dateregex.py/harness.py itself. Same precedent as wsnSearch4_ctest.py, which locally
  rewrote is_year_match to test a proposed C-signal redesign without touching the shared
  file. Orientation-specific: a cell can be part of a row-wise sequence without being part
  of a col-wise one, and vice versa -- computed as two precomputed hit-grids per block,
  not a single per-cell classification.
"""
import sys
import warnings
warnings.filterwarnings("ignore")

from openpyxl.utils.cell import column_index_from_string

from harness import Workbook, _detect_blocks, _merge_blocks, _cell_flags, _render_table
from dateregex import is_date_like as _base_is_date_like

RUN_THRESHOLD = 3
CANDIDATE_CAP = 5
SEQ_YEAR_MIN, SEQ_YEAR_MAX = 1980, 2035


def _is_plain_int_year(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)) and float(v).is_integer():
        iv = int(v)
        if SEQ_YEAR_MIN <= iv <= SEQ_YEAR_MAX:
            return iv
    return None


def _seq_hits_1d(seq):
    """seq: list of raw cell values, in order. Returns a set of indices that are part of
    a maximal run of length >=3 of consecutive integers each exactly +1 from the previous,
    within SEQ_YEAR_MIN..SEQ_YEAR_MAX."""
    hits = set()
    run_start = None
    prev_val = None
    for i, v in enumerate(seq):
        iv = _is_plain_int_year(v)
        if iv is not None and prev_val is not None and iv == prev_val + 1:
            pass  # continues current run, run_start already set
        else:
            if run_start is not None and i - run_start >= RUN_THRESHOLD:
                hits.update(range(run_start, i))
            run_start = i if iv is not None else None
        prev_val = iv
    if run_start is not None and len(seq) - run_start >= RUN_THRESHOLD:
        hits.update(range(run_start, len(seq)))
    return hits


def _is_plain_string(v):
    return isinstance(v, str) and v.strip() != ""


def _precompute_seq_grids(values, r0, r1, c0, c1):
    """seq_row[r][cc]=True if (r,cc) is part of a row-wise qualifying integer sequence;
    seq_col[r][cc]=True if part of a col-wise one. Keyed by real (row,col) tuples."""
    seq_row = {}
    for r in range(r0, r1 + 1):
        row_vals = [values[r - 1][cc - 1] for cc in range(c0, c1 + 1)]
        hit_idx = _seq_hits_1d(row_vals)
        for i in hit_idx:
            seq_row[(r, c0 + i)] = True
    seq_col = {}
    for cc in range(c0, c1 + 1):
        col_vals = [values[r - 1][cc - 1] for r in range(r0, r1 + 1)]
        hit_idx = _seq_hits_1d(col_vals)
        for i in hit_idx:
            seq_col[(r0 + i, cc)] = True
    return seq_row, seq_col


def _ab_cd_candidates_v2(values, r0, r1, c0, c1, metric_term, entity_term, year_term):
    seq_row, seq_col = _precompute_seq_grids(values, r0, r1, c0, c1)

    def ab_hit_and_anchor(v):
        _, metric, entity, _ = _cell_flags(v, metric_term, entity_term, year_term)
        anchor = metric or entity
        hit = anchor or _is_plain_string(v)
        return hit, anchor

    def cd_hit_and_anchor(v, r, cc, use_row):
        base_date = _base_is_date_like(v)
        seq = seq_row.get((r, cc), False) if use_row else seq_col.get((r, cc), False)
        anchor = base_date or seq
        hit = anchor or _is_plain_string(v)
        return hit, anchor

    ab_col_candidates = set()
    cd_col_candidates = set()
    ab_row_candidates = set()
    cd_row_candidates = set()

    # column-wise streaks
    for cc in range(c0, c1 + 1):
        ab_streak = ab_anchor_ct = 0
        cd_streak = cd_anchor_ct = 0
        for r in range(r0, r1 + 1):
            v = values[r - 1][cc - 1]
            hit, anchor = ab_hit_and_anchor(v)
            if hit:
                ab_streak += 1
                ab_anchor_ct += anchor
            else:
                if ab_streak >= RUN_THRESHOLD and ab_anchor_ct >= 1:
                    ab_col_candidates.add(cc)
                ab_streak = ab_anchor_ct = 0
            hit, anchor = cd_hit_and_anchor(v, r, cc, use_row=False)
            if hit:
                cd_streak += 1
                cd_anchor_ct += anchor
            else:
                if cd_streak >= RUN_THRESHOLD and cd_anchor_ct >= 1:
                    cd_col_candidates.add(cc)
                cd_streak = cd_anchor_ct = 0
        if ab_streak >= RUN_THRESHOLD and ab_anchor_ct >= 1:
            ab_col_candidates.add(cc)
        if cd_streak >= RUN_THRESHOLD and cd_anchor_ct >= 1:
            cd_col_candidates.add(cc)

    # row-wise streaks
    for r in range(r0, r1 + 1):
        ab_streak = ab_anchor_ct = 0
        cd_streak = cd_anchor_ct = 0
        for cc in range(c0, c1 + 1):
            v = values[r - 1][cc - 1]
            hit, anchor = ab_hit_and_anchor(v)
            if hit:
                ab_streak += 1
                ab_anchor_ct += anchor
            else:
                if ab_streak >= RUN_THRESHOLD and ab_anchor_ct >= 1:
                    ab_row_candidates.add(r)
                ab_streak = ab_anchor_ct = 0
            hit, anchor = cd_hit_and_anchor(v, r, cc, use_row=True)
            if hit:
                cd_streak += 1
                cd_anchor_ct += anchor
            else:
                if cd_streak >= RUN_THRESHOLD and cd_anchor_ct >= 1:
                    cd_row_candidates.add(r)
                cd_streak = cd_anchor_ct = 0
        if ab_streak >= RUN_THRESHOLD and ab_anchor_ct >= 1:
            ab_row_candidates.add(r)
        if cd_streak >= RUN_THRESHOLD and cd_anchor_ct >= 1:
            cd_row_candidates.add(r)

    ab_candidates = sorted([(c, "c") for c in ab_col_candidates] + [(r, "r") for r in ab_row_candidates])
    cd_candidates = sorted([(c, "c") for c in cd_col_candidates] + [(r, "r") for r in cd_row_candidates])
    return ab_candidates, cd_candidates


def _fmt_candidates(cands):
    if not cands:
        return "-", None
    shown = cands[:CANDIDATE_CAP]
    s = ",".join(f"{idx}{orient}" for idx, orient in shown)
    note = None
    if len(cands) > CANDIDATE_CAP:
        note = f"{len(cands) - CANDIDATE_CAP} more candidates not shown"
    orients = {o for _, o in shown}
    if len(orients) > 1:
        mixed = ",".join(f"{idx}{orient}" for idx, orient in shown)
        mix_note = f"candidates mix orientation ({mixed}) — check both, order is not a nesting signal"
        note = f"{note}; {mix_note}" if note else mix_note
    return s, note


def _enhanced_d_count(values, r0, r1, c0, c1):
    seq_row, seq_col = _precompute_seq_grids(values, r0, r1, c0, c1)
    n = 0
    for r in range(r0, r1 + 1):
        for cc in range(c0, c1 + 1):
            v = values[r - 1][cc - 1]
            if _base_is_date_like(v) or seq_row.get((r, cc)) or seq_col.get((r, cc)):
                n += 1
    return n


def map_block_with_runs_v2(wb, sheet, metric_term=None, entity_term=None, year_term=None):
    values, n_rows, n_cols, true_max_row, true_max_col = wb._scan(sheet)
    zones = _detect_blocks(values, n_rows, n_cols)
    entries = _merge_blocks(zones)

    rows_out = []
    per_block_notes = []
    for e in entries:
        num_str = str(e["nums"][0]) if len(e["nums"]) == 1 else f"{e['nums'][0]}-{e['nums'][-1]}"
        r0, r1 = e["rows"]
        c0, c1 = e["cols"]
        a = b = c = 0
        for r in range(r0, r1 + 1):
            for cc in range(c0, c1 + 1):
                _, metric, entity, year = _cell_flags(values[r - 1][cc - 1], metric_term, entity_term, year_term)
                a += metric
                b += entity
                c += year
        d = _enhanced_d_count(values, r0, r1, c0, c1)

        ab_cands, cd_cands = _ab_cd_candidates_v2(values, r0, r1, c0, c1, metric_term, entity_term, year_term)
        ab_str, ab_note = _fmt_candidates(ab_cands)
        cd_str, cd_note = _fmt_candidates(cd_cands)
        if ab_note:
            per_block_notes.append(f"block {num_str} AB-run {ab_note}")
        if cd_note:
            per_block_notes.append(f"block {num_str} CD-run {cd_note}")

        rows_out.append([
            num_str,
            str(r0) if r0 == r1 else f"{r0}-{r1}",
            str(c0) if c0 == c1 else f"{c0}-{c1}",
            d, a, b, c, ab_str, cd_str,
        ])

    table = _render_table(["block", "rows", "cols", "D(v2)", "A", "B", "C", "AB-run", "CD-run"], rows_out)
    return table, per_block_notes, entries


def main():
    from wsnSearch5_item2test import TEST_ROWS

    wb_cache = {}
    for t in TEST_ROWS:
        if t["file"] not in wb_cache:
            wb_cache[t["file"]] = Workbook(t["file"])
        wb = wb_cache[t["file"]]
        colstr = "".join(ch for ch in t["cell"] if ch.isalpha())
        rownum = int("".join(ch for ch in t["cell"] if ch.isdigit()))
        col = column_index_from_string(colstr)

        table, notes, entries = map_block_with_runs_v2(
            wb, t["sheet"], t["metric_term"], t["entity_term"], t["year_term"]
        )
        print("=" * 100)
        print(f"[{t['n']}] {t['file']} | {t['sheet']} | {t['cell']} (row={rownum}, col={col})")
        print(f"  terms: metric={t['metric_term']!r} entity={t['entity_term']!r} year={t['year_term']!r}")
        print(f"  true period axis: {t['true_period']}")
        print(f"  true lineitem axis: {t['true_lineitem']}")
        print(table)
        for note in notes:
            print(f"  note: {note}")


if __name__ == "__main__":
    main()
