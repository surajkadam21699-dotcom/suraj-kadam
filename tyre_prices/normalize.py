"""Parse the loosely formatted numbers, dates, units and currencies on trade-data pages."""

from __future__ import annotations

import re
from datetime import date, datetime

# Kilograms in one reported unit, for the units that measure weight.
_KG_PER_UNIT = {
    "G": 0.001, "GM": 0.001, "GMS": 0.001, "GRM": 0.001,
    "KG": 1.0, "KGS": 1.0, "KGM": 1.0, "KILOGRAM": 1.0, "KILOGRAMS": 1.0,
    "LB": 0.45359237, "LBS": 0.45359237,
    "QTL": 100.0, "QUINTAL": 100.0,
    "T": 1000.0, "MT": 1000.0, "MTS": 1000.0, "MTON": 1000.0, "TNE": 1000.0,
    "TON": 1000.0, "TONS": 1000.0, "TONNE": 1000.0, "TONNES": 1000.0,
}

# Western (123,456) and Indian (1,23,456) digit grouping, optional decimals.
_NUMBER = re.compile(r"\d+(?:,\d{2,3})*(?:\.\d+)?|\.\d+")
# Figures the sites partially hide from non-subscribers, e.g. "2,3**" or "12XX".
_MASKED = re.compile(r"\*|\d[xX]|[xX]\d")

# Indian sites write day before month, so 04/10/2026 is 4 October.
_DATE_FORMATS = (
    "%d-%b-%Y", "%d %b %Y", "%d-%b-%y", "%d-%B-%Y", "%d %B %Y",
    "%b %d, %Y", "%b %d %Y", "%B %d, %Y", "%B %d %Y",
    "%Y-%m-%d", "%Y/%m/%d",
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y",
)

_CURRENCIES = (
    ("INR", re.compile(r"\binr\b|₹|\brs\b|rupee", re.IGNORECASE)),
    ("USD", re.compile(r"\busd\b|\$", re.IGNORECASE)),
    ("EUR", re.compile(r"\beur\b|€", re.IGNORECASE)),
    ("GBP", re.compile(r"\bgbp\b|£", re.IGNORECASE)),
)


def parse_number(text: str | float | None) -> float | None:
    """First number in ``text`` ("₹ 1,23,456.50" -> 123456.5), or None if absent or masked."""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    if _MASKED.search(text):
        return None
    match = _NUMBER.search(text)
    return float(match.group().replace(",", "")) if match else None


def parse_date(text: str | None) -> str | None:
    """ISO date (YYYY-MM-DD) for the date formats trade-data sites use, or None."""
    if not text:
        return None
    cleaned = " ".join(text.split())
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt).date().isoformat()
        except ValueError:
            pass
    try:
        return date.fromisoformat(cleaned[:10]).isoformat()  # e.g. "2026-04-10T00:00:00"
    except ValueError:
        return None


def normalize_unit(text: str | None) -> str | None:
    """Upper-case unit code without punctuation: "Kgs." -> "KGS", "M.T." -> "MT"."""
    code = re.sub(r"[^A-Za-z]", "", text or "").upper()
    return code or None


def kg_per_unit(unit: str | None) -> float | None:
    """Kilograms in one ``unit``, or None when it is not a weight (NOS, PCS, ...)."""
    return _KG_PER_UNIT.get(normalize_unit(unit) or "")


def detect_currency(text: str | None) -> str | None:
    """Currency named in a column header such as "Value (INR)" or "Unit Price (US $)"."""
    for code, pattern in _CURRENCIES:
        if text and pattern.search(text):
            return code
    return None
