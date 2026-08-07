#!/usr/bin/env python3
"""
date created: 2026-08-05
date updated: 2026-08-05
date surroundings last checked: 2026-08-05
"""
import datetime
import re

# Each entry: (id, category, mode, compiled_regex).
# mode "full"   -> re.fullmatch against the normalized string (strip + casefold).
# mode "search" -> re.search anywhere in the normalized string (for period tokens
#                  embedded in a longer label, e.g. "H1 Total Move").
# Source: periodHeaderSurvey.md section 1, one entry per catalog row.
_PATTERNS = [
    # -- Quarter --
    ("q_yy", "quarter", "full", r"[1-4]q\d{2}"),
    ("q_yy_e", "quarter", "full", r"[1-4]q\d{2}e"),
    ("qfy_yy", "quarter", "full", r"[1-4]qfy\d{2}e?"),
    ("fy3_yy", "quarter", "full", r"fy3/\d{2}e?"),
    ("q_bare", "quarter", "full", r"[1-4]q"),
    ("q_bare_e", "quarter", "full", r"[1-4]q\s?e"),
    # -- Half-year --
    ("h_yy", "half_year", "full", r"[1-2]h\d{2}"),
    ("h_embedded", "half_year", "search", r"\b(?:[1-2]h|h[1-2])\b"),
    # -- Cumulative / YTD --
    ("m_yy", "ytd", "full", r"\d{1,2}m\d{2}"),
    # -- Year --
    ("yyyy", "year", "full", r"\d{4}"),
    ("yyyy_e", "year", "full", r"\d{4}'?e"),
    ("fy_yyyy", "year", "full", r"fy\d{4}"),
    ("end_yyyy", "year", "full", r"end-\d{4}e?"),
    ("from_yyyy_q", "year", "full", r"from\s\d{4}\s[1-4]q"),
    # -- Ranges (two points, not one period; still date-like) --
    ("range_yyyy_yy", "range", "full", r"\d{4}-\d{2}"),
    ("range_yyyy_yy_e", "range", "full", r"\d{4}-\d{2}\s?e"),
    ("range_yyyy_yyyy", "range", "full", r"\d{4}-\d{4}e?"),
    # -- Scenario-tagged --
    ("scenario_yy_e", "scenario", "full", r"(?:bull|bear)\s'\d{2}e"),
    # -- Full dates (plain text; native datetime/time cells are caught by dtype check) --
    ("yyyymmdd", "full_date", "full", r"\d{8}"),
    ("mm_dd_yyyy", "full_date", "full", r"\d{2}/\d{2}/\d{4}"),
    # -- CJK --
    ("cjk_year_month", "cjk", "full", r"\d{4}年\d{1,2}月"),
    ("cjk_year_month_range", "cjk", "full", r"\d{4}年\d{1,2}-\d{1,2}月"),
    ("cjk_year_end", "cjk", "full", r"\d{4}年末"),
    ("cjk_year_q_from", "cjk", "full", r"\d{4}年[1-4]qより"),
]

PATTERNS = [(pid, cat, mode, re.compile(pat)) for pid, cat, mode, pat in _PATTERNS]


def is_date_like(v):
    if isinstance(v, (datetime.datetime, datetime.time)):
        return True
    if isinstance(v, str):
        s = v.strip().casefold()
        if not s:
            return False
        for _id, _cat, mode, regex in PATTERNS:
            if mode == "full":
                if regex.fullmatch(s):
                    return True
            else:
                if regex.search(s):
                    return True
        return False
    return False
