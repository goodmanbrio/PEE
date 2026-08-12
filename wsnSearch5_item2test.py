#!/usr/bin/env python3
"""
date created: 2026-08-09
date updated: 2026-08-09
date surroundings last checked: 2026-08-09

Throwaway practicality test for wsnSearch5.md Item 2 (map_block's AB-run/CD-run columns).
Reuses harness.py's real _cell_flags/_detect_blocks/_merge_blocks/_render_table (not a
reimplementation), same convention as wsnSearch4_ctest.py/wsnSearch4_peektest.py.

Implements the run-detector exactly per wsnSearch5.md item 2 section 2:
  - one bookkeeping pass alongside map_block's existing D/A/B/C counting loop
  - col_streak[cc] per column (reset on non-hit, incremented on hit) for AB-hits
    (metric|entity) and CD-hits (date|year), tracking each column's longest run
  - row-scoped scalar streak for the same two hit-types, reset each row
  - run qualifies at length >= 3 (RUN_THRESHOLD)
  - candidates capped at 5 per axis type per block (CANDIDATE_CAP), comma-joined,
    ascending by index, "<index><r|c>"
"""
import sys
import warnings
warnings.filterwarnings("ignore")

from openpyxl.utils.cell import column_index_from_string

from harness import Workbook, _detect_blocks, _merge_blocks, _cell_flags, _render_table

RUN_THRESHOLD = 3
CANDIDATE_CAP = 5


def _ab_cd_candidates(values, r0, r1, c0, c1, metric_term, entity_term, year_term):
    """One pass over the block. Returns (ab_candidates, cd_candidates), each a list of
    (index, orientation) tuples in ascending-index order, orientation 'r' or 'c'."""
    width = c1 - c0 + 1
    col_streak_ab = [0] * width
    col_best_ab = [0] * width
    col_streak_cd = [0] * width
    col_best_cd = [0] * width
    row_best_ab = {}  # row -> best run length within that row
    row_best_cd = {}

    for r in range(r0, r1 + 1):
        row_streak_ab = 0
        row_streak_cd = 0
        rb_ab = 0
        rb_cd = 0
        for cc in range(c0, c1 + 1):
            i = cc - c0
            date, metric, entity, year = _cell_flags(values[r - 1][cc - 1], metric_term, entity_term, year_term)
            ab_hit = metric or entity
            cd_hit = date or year

            if ab_hit:
                col_streak_ab[i] += 1
            else:
                col_streak_ab[i] = 0
            col_best_ab[i] = max(col_best_ab[i], col_streak_ab[i])

            if cd_hit:
                col_streak_cd[i] += 1
            else:
                col_streak_cd[i] = 0
            col_best_cd[i] = max(col_best_cd[i], col_streak_cd[i])

            if ab_hit:
                row_streak_ab += 1
            else:
                row_streak_ab = 0
            rb_ab = max(rb_ab, row_streak_ab)

            if cd_hit:
                row_streak_cd += 1
            else:
                row_streak_cd = 0
            rb_cd = max(rb_cd, row_streak_cd)

        row_best_ab[r] = rb_ab
        row_best_cd[r] = rb_cd

    ab_candidates = []
    cd_candidates = []
    for i in range(width):
        if col_best_ab[i] >= RUN_THRESHOLD:
            ab_candidates.append((c0 + i, "c"))
        if col_best_cd[i] >= RUN_THRESHOLD:
            cd_candidates.append((c0 + i, "c"))
    for r in range(r0, r1 + 1):
        if row_best_ab[r] >= RUN_THRESHOLD:
            ab_candidates.append((r, "r"))
        if row_best_cd[r] >= RUN_THRESHOLD:
            cd_candidates.append((r, "r"))

    ab_candidates.sort(key=lambda t: t[0])
    cd_candidates.sort(key=lambda t: t[0])
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
        mix_note = f"candidates mix orientation ({mixed}) — check both, one is likely a false positive, order is not a nesting signal"
        note = f"{note}; {mix_note}" if note else mix_note
    return s, note


def map_block_with_runs(wb, sheet, metric_term=None, entity_term=None, year_term=None):
    values, n_rows, n_cols, true_max_row, true_max_col = wb._scan(sheet)
    zones = _detect_blocks(values, n_rows, n_cols)
    entries = _merge_blocks(zones)

    rows_out = []
    per_block_notes = []
    for e in entries:
        num_str = str(e["nums"][0]) if len(e["nums"]) == 1 else f"{e['nums'][0]}-{e['nums'][-1]}"
        r0, r1 = e["rows"]
        c0, c1 = e["cols"]
        d = a = b = c = 0
        for r in range(r0, r1 + 1):
            for cc in range(c0, c1 + 1):
                date, metric, entity, year = _cell_flags(values[r - 1][cc - 1], metric_term, entity_term, year_term)
                d += date
                a += metric
                b += entity
                c += year

        ab_cands, cd_cands = _ab_cd_candidates(values, r0, r1, c0, c1, metric_term, entity_term, year_term)
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

    table = _render_table(["block", "rows", "cols", "D", "A", "B", "C", "AB-run", "CD-run"], rows_out)
    return table, per_block_notes, entries


# ---- test rows: file | sheet | answer cell | metric_term | entity_term | year_term ----
# metric_term/entity_term chosen from real label text confirmed by direct sheet inspection
# (not guessed blind) -- see wsnSearch5_item2test_notes for the recon this was built from.
TEST_ROWS = [
    dict(n=1, file="300450_Wuxi Lead_20210422_client.xlsx", sheet="Op & Non op", cell="L22",
         metric_term="Marketing", entity_term=None, year_term=2022,
         true_period="row 2 (Rmb mn | 2012..2023, numeric-typed) -- SEPARATE block [1,16], not merged with data block [18,38] (BLOCK_MERGE_MAX_ROWSUM=20 exceeded)",
         true_lineitem="col 1 (label col), row 22 'Marketing Expenses' -- single row, term doesn't repeat"),
    dict(n=2, file="300450_Wuxi Lead_20210422_client.xlsx", sheet="Consolidate", cell="M46",
         metric_term="Growth", entity_term="Prime Operating", year_term=2022,
         true_period="row 2 (2012..2022E, mixed numeric/string) -- SEPARATE block [rows 2,2]x[17,27], own 1-row band",
         true_lineitem="col 2, row 46 'Prime Operating Business' -- single row"),
    dict(n=3, file="300450_Wuxi Lead_20210422_client.xlsx", sheet="PL", cell="AA6",
         metric_term="Gross Profit", entity_term=None, year_term=2020,
         true_period="row 3 (1H15..2H20, string half-year labels) -- SAME block as data [1,31]x[17,30]",
         true_lineitem="col 1/3 (bilingual labels, different col band) row 6 'Gross Profit'"),
    dict(n=4, file="2024 1H Hit ratio.xlsx", sheet="Model Port Updates", cell="AC603",
         metric_term="PnL", entity_term="CGNX", year_term=2024,
         true_period="none -- flat trade log, no period axis at all",
         true_lineitem="col 3/4 Ticker/Name, but query needs 3-field co-match (date+ticker+name), not a single axis"),
    dict(n=5, file="1070.HK Model Dec 2020.xlsx", sheet="Sheet1", cell="L6",
         metric_term="Sales volume", entity_term="TCL", year_term=2020,
         true_period="row 2 (1Q20,2Q20,3Q20,9M20, string) -- SAME block [1,18]x[10,13]",
         true_lineitem="col 1, rows 3-10 (TCL TV / Sales volume / ASP block) -- from testruns3 q05 log"),
    dict(n=6, file="20220404_Parade_BOE.xlsx", sheet="Shipment", cell="G10",
         metric_term="MacBook", entity_term=None, year_term=2018,
         true_period="row 2 (1Q17..3Q20, string) -- SAME block [2,26]x[2,34]",
         true_lineitem="col 2, row ~10 'MacBook' -- single row"),
    dict(n=7, file="20220404_Parade_BOE.xlsx", sheet="Parade BOE", cell="G202",
         metric_term="DisplayPort", entity_term=None, year_term=2020,
         true_period="header row for this section not in first 6 rows of block -- need wider probe",
         true_lineitem="col 2, row 202 'DisplayPort Solutions (DP)' -- single row"),
    dict(n=8, file="20220404_Parade_BOE.xlsx", sheet="Parade BOE", cell="J168",
         metric_term="Source Driver", entity_term=None, year_term=2023,
         true_period="header row for this section not in first 6 rows of block -- need wider probe",
         true_lineitem="col 2, row 168 'Source Driver (SD)' -- single row"),
    dict(n=9, file="[Quarter added]Advantest 6857 JP BOE.xlsx", sheet="Advantest BOE", cell="AI86",
         metric_term="Revenue", entity_term=None, year_term=2022,
         true_period="role-swap: row3='CAGR' (this col only), row4='2022-2025E' (this col only) -- expected MISS (out of scope per testqs)",
         true_lineitem="n/a -- expected miss"),
    dict(n=10, file="[Quarter added]Advantest 6857 JP BOE.xlsx", sheet="Advantest BOE", cell="AC86",
         metric_term="Revenue", entity_term="Advantest", year_term=2022,
         true_period="row 3 (FY3/22 form, string) -- SAME block [72,88]x[24,33] header not shown in first 6 rows shown, need wider probe",
         true_lineitem="col 1/2, row 86 'Total Revenue' -- from testruns3 q10 log"),
    dict(n=11, file="[Quarter added]Advantest 6857 JP BOE.xlsx", sheet="Advantest BOE", cell="D8",
         metric_term="SoC Test Systems", entity_term=None, year_term=2020,
         true_period="row 3 (4QFY20.., string) + row 4 (1Q20.., string) -- SAME block [3,19]x[4,20], nested 2-level",
         true_lineitem="col 2, row 8 'SoC Test Systems' -- single row"),
    dict(n=12, file="20220204_WaferSupplyDemandModel.xlsx", sheet="Wafer_SemiSummary", cell="F44",
         metric_term="300mm", entity_term=None, year_term=2018,
         true_period="col 3 (years 1990..2018+, numeric-typed, TRANSPOSED axis) -- SAME block [12,141]x[2,48]",
         true_lineitem="row 15 ('Total','YoY','300mm','200mm',...) -- row-form lineitem header"),
    dict(n=13, file="20220204_WaferSupplyDemandModel.xlsx", sheet="WaferSupplyDemand", cell="BJ6",
         metric_term="Shin-Etsu", entity_term=None, year_term=2011,
         true_period="row 4 (years, sparse merged-cell) + row 5 (quarters, dense) -- nested, SAME block [3,51]x[45,130]",
         true_lineitem="col 46/47, row 6 'Shin-Etsu Handotai (JPN)' -- single row"),
    dict(n=14, file="20220204_WaferSupplyDemandModel.xlsx", sheet="PriceAssumption", cell="F37",
         metric_term="300mm", entity_term=None, year_term=2018,
         true_period="row 36 (YYYY年末 CJK form) -- SAME block [35,48]x[2,11]",
         true_lineitem="col 3, rows 37/40 '300mm wafer' -- appears 2x, not adjacent (separated by 200mm row)"),
    dict(n=15, file="99-00 Study Raw.xlsx", sheet="Existing", cell="G5",
         metric_term="indexed", entity_term="INFOR", year_term=1998,
         true_period="row 4 (native datetime dtype) -- SAME block [3,3754]x[2,41]",
         true_lineitem="col 3, row 5 'INFOR GLOBAL SOL' -- single row"),
    dict(n=16, file="99-00 Study Raw.xlsx", sheet="Existing", cell="H603",
         metric_term="indexed", entity_term="PSYCHIATRIC", year_term=1998,
         true_period="row 4 (native datetime dtype) -- SAME block [3,3754]x[2,41]",
         true_lineitem="col 3, row 603 'PSYCHIATRIC SOLU' -- single row, beyond ROW_SCAN_BUDGET? 603<10000 ok"),
    dict(n=17, file="99-00 Study Raw.xlsx", sheet="Existing", cell="H1000",
         metric_term="indexed", entity_term="BOB EVANS", year_term=1998,
         true_period="row 4 (native datetime dtype) -- SAME block [3,3754]x[2,41]",
         true_lineitem="col 3, row 1000 'BOB EVANS FARMS' -- single row"),
    dict(n=18, file="2.元戎启行资金消耗统计 2019.2-2021.4.xlsx", sheet="元戎启行资金消耗统计", cell="C4",
         metric_term="经营活动净流出", entity_term=None, year_term=2019,
         true_period="row 3 (CJK full-year-range + monthly, string) -- SAME block [2,6]x[2,21]",
         true_lineitem="col 2, row 4 '经营活动净流出' -- single row"),
    dict(n=19, file="2.元戎启行资金消耗统计 2019.2-2021.4.xlsx", sheet="元戎启行资金消耗统计", cell="E5",
         metric_term="资本性支出", entity_term=None, year_term=2020,
         true_period="row 3 (CJK monthly) -- SAME block [2,6]x[2,21]",
         true_lineitem="col 2, row 5 '资本性支出' -- single row"),
    dict(n=20, file="2.元戎启行资金消耗统计 2019.2-2021.4.xlsx", sheet="元戎启行资金消耗统计", cell="S6",
         metric_term="合计", entity_term=None, year_term=2021,
         true_period="row 3 (CJK monthly) -- SAME block [2,6]x[2,21]",
         true_lineitem="col 2, row 6 '合计' -- single row"),
]


def main():
    wb_cache = {}
    for t in TEST_ROWS:
        if t["file"] not in wb_cache:
            wb_cache[t["file"]] = Workbook(t["file"])
        wb = wb_cache[t["file"]]
        colstr = "".join(ch for ch in t["cell"] if ch.isalpha())
        rownum = int("".join(ch for ch in t["cell"] if ch.isdigit()))
        col = column_index_from_string(colstr)

        table, notes, entries = map_block_with_runs(
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
