"""Recognise scrap tyre shipments and sort them into the grades traders quote prices for."""

from __future__ import annotations

import re

SCRAP_RUBBER_HS = "4004"  # waste, parings and scrap of rubber: where Indian customs files scrap tyres

_TYRE = re.compile(r"\b(?:tyre|tire)", re.IGNORECASE)
_SCRAP = re.compile(
    r"scrap|waste|shred|\bcut|piece|\bpcs?\b|bale|crumb|granul|powder|buffing"
    r"|used|worn|\bold\b|chip|\btdf\b|sidewall",
    re.IGNORECASE,
)

# Checked in order; the first grade whose pattern matches the description wins.
_GRADES = (
    ("crumb", re.compile(r"crumb|granul|powder|mesh|buffing", re.IGNORECASE)),
    ("shredded", re.compile(r"shred|chip|\btdf\b|derived fuel", re.IGNORECASE)),
    ("cut", re.compile(r"\bcut|piece|\bpcs?\b|sidewall|halves", re.IGNORECASE)),
    ("baled", re.compile(r"\bbale", re.IGNORECASE)),
    ("whole", re.compile(r"whole|used|worn|\bold\b|casing", re.IGNORECASE)),
)


def is_scrap_tyre(description: str | None, hs_code: str | None = None) -> bool:
    """True when a shipment reads like scrap or used tyres rather than other tyre-related goods."""
    if not description or not _TYRE.search(description):
        return False
    hs_code = hs_code or ""
    if hs_code and not hs_code.startswith("40"):
        return False  # outside the rubber chapter: carbon black, tyre cord fabric, moulds, ...
    return hs_code.startswith(SCRAP_RUBBER_HS) or bool(_SCRAP.search(description))


def classify_grade(description: str | None) -> str:
    """Grade bucket for a scrap tyre description: crumb, shredded, cut, baled, whole or unspecified."""
    for grade, pattern in _GRADES:
        if description and pattern.search(description):
            return grade
    return "unspecified"
