"""
Generic parsing helpers used by the CSV and OCR layers.

These are pure functions with no I/O.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Optional

import pandas as pd


def parse_number(raw: Optional[str]) -> Optional[float]:
    """Parse a number string, handling EU (1.234,56) and Anglo (1,234.56) formats."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).replace("€", "").replace("$", "").replace("£", "").strip()
    if not s:
        return None

    last_dot = s.rfind(".")
    last_comma = s.rfind(",")
    if last_comma > last_dot:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")

    try:
        return float(s)
    except ValueError:
        return None


def parse_date(raw: Optional[str]) -> Optional[str]:
    """Parse a date into ISO yyyy-mm-dd, trying common European formats."""
    if raw is None:
        return None
    s = str(raw).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y",
                "%d.%m.%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


# Month names in German and English (lower-case, no umlaut variations both listed)
MONTH_NAMES = {
    # German
    "januar": 1, "februar": 2, "märz": 3, "marz": 3, "april": 4,
    "mai": 5, "juni": 6, "juli": 7, "august": 8,
    "september": 9, "oktober": 10, "november": 11, "dezember": 12,
    # English
    "january": 1, "march": 3, "may": 5, "june": 6, "july": 7, "october": 10,
    # Common English abbreviations
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}


def parse_date_with_month_name(text: str) -> Optional[str]:
    """Extracts the first 'day month-name year' date from text, e.g. '19. Februar 2026'."""
    match = re.search(r"\b(\d{1,2})\.?\s+([A-Za-zäöüÄÖÜ]+)\s+(\d{4})\b", text)
    if not match:
        return None
    day, month_name, year = match.groups()
    month = MONTH_NAMES.get(month_name.lower())
    if not month:
        return None
    try:
        return f"{int(year):04d}-{month:02d}-{int(day):02d}"
    except ValueError:
        return None


def find_column_index(headers_lower: list[str], aliases: list[str]) -> int:
    """Returns the first index in `headers_lower` matching one of `aliases`, or -1."""
    for alias in aliases:
        if alias in headers_lower:
            return headers_lower.index(alias)
    return -1
