#!/usr/bin/env python3
"""
date created: 2026-08-05
date updated: 2026-08-08
date surroundings last checked: 2026-08-08
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


# ---- C signal: is_year_match, wsnSearch4.md item 2 section 3 ----
# Separate pattern tables from PATTERNS above: these need capture groups to compare
# against a specific target_year, which the D-signal patterns (existence-only) don't carry.
# 4-digit patterns: existing ids from PATTERNS re-expressed with capture groups, plus
# three new ones (ye_yyyy, cy_yyyy, yyyy_a).
_C_4DIGIT = [
    ("yyyy", re.compile(r"(\d{4})")),
    ("yyyy_e", re.compile(r"(\d{4})'?e")),
    ("fy_yyyy", re.compile(r"fy(\d{4})")),
    ("end_yyyy", re.compile(r"end-(\d{4})e?")),
    ("cjk_year_end", re.compile(r"(\d{4})年末")),
    ("ye_yyyy", re.compile(r"ye\s?(\d{4})")),
    ("cy_yyyy", re.compile(r"cy(\d{4})")),
    ("yyyy_a", re.compile(r"(\d{4})a")),
]
# 2-digit fiscal-embedded patterns, compared against target_year % 100 — no
# century-guessing, since target_year is already known absolutely.
_C_2DIGIT = [
    ("fy3_yy", re.compile(r"fy3/(\d{2})e?")),
    ("q_yy", re.compile(r"[1-4]q(\d{2})")),
    ("q_yy_e", re.compile(r"[1-4]q(\d{2})e")),
    ("qfy_yy", re.compile(r"[1-4]qfy(\d{2})e?")),
    ("h_yy", re.compile(r"[1-2]h(\d{2})")),
    ("m_yy", re.compile(r"\d{1,2}m(\d{2})")),
    ("range_yyyy_yy", re.compile(r"\d{4}-(\d{2})")),
    ("range_yyyy_yy_e", re.compile(r"\d{4}-(\d{2})\s?e")),
    ("scenario_yy_e", re.compile(r"(?:bull|bear)\s'(\d{2})e")),
]
_CJK_YEAR_MONTH_RANGE = re.compile(r"(\d{4})年(\d{1,2})-(\d{1,2})月")
_CJK_YEAR_MONTH = re.compile(r"(\d{4})年(\d{1,2})月")
_FROM_YYYY_Q = re.compile(r"from\s(\d{4})\s[1-4]q")

_C_YEAR_MIN, _C_YEAR_MAX = 1980, 2035


def is_year_match(v, target_year):
    """Query-specific year triangulation, an anchor to check not a settled answer —
    same hedge as is_date_like. See wsnSearch4.md item 2 section 3 for the full spec,
    including the month-granularity limitation (cjk_year_month / native datetime match
    on year alone, not month) and the explicit range-membership / quarter-specific
    exclusions."""
    if isinstance(v, datetime.datetime):
        return v.year == target_year
    if isinstance(v, datetime.time):
        return False  # bare time-of-day carries no year
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if float(v).is_integer():
            iv = int(v)
            if _C_YEAR_MIN <= iv <= _C_YEAR_MAX:
                return iv == target_year
        return False
    if isinstance(v, str):
        s = v.strip().casefold()
        if not s:
            return False
        m = _CJK_YEAR_MONTH_RANGE.fullmatch(s)
        if m:
            y, m1, m2 = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return y == target_year and m1 == 1 and m2 == 12
        if _FROM_YYYY_Q.fullmatch(s):
            return False
        for _pid, regex in _C_4DIGIT:
            m = regex.fullmatch(s)
            if m:
                return int(m.group(1)) == target_year
        for _pid, regex in _C_2DIGIT:
            m = regex.fullmatch(s)
            if m:
                return int(m.group(1)) == target_year % 100
        m = _CJK_YEAR_MONTH.fullmatch(s)
        if m:
            return int(m.group(1)) == target_year
        return False
    return False
