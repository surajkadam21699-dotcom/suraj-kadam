import pytest

from tyre_prices.grades import classify_grade, is_scrap_tyre


@pytest.mark.parametrize("description, hs_code, expected", [
    ("RUBBER TYRE SCRAP CUT IN THREE PIECES", "40040000", True),
    ("USED TYRES IN BALES", None, True),
    ("SCRAP TIRES WHOLE, PASSENGER", None, True),
    ("PNEUMATIC TYRES", "40040000", True),  # filed under scrap rubber, so scrap by definition
    ("NEW RADIAL TYRES 295/80R22.5", "40112010", False),
    ("RUBBER SCRAP (EPDM PROFILES)", "40040000", False),
    ("CARBON BLACK USED IN TYRE MANUFACTURING", "28030010", False),
    ("", None, False),
    (None, None, False),
])
def test_is_scrap_tyre(description, hs_code, expected):
    assert is_scrap_tyre(description, hs_code) is expected


@pytest.mark.parametrize("description, expected", [
    ("RUBBER TYRE SCRAP CUT IN THREE PIECES", "cut"),
    ("WASTE TYRES - 2 CUT (PCR)", "cut"),
    ("TYRE SCRAP 3 PCS", "cut"),
    ("SHREDDED TYRE SCRAP (TDF)", "shredded"),
    ("TYRE CHIPS 50MM", "shredded"),
    ("TYRE BUFFING / CRUMB RUBBER POWDER 30 MESH", "crumb"),
    ("USED TYRES IN BALES", "baled"),
    ("SCRAP TIRES WHOLE, PASSENGER", "whole"),
    ("USED TRUCK TYRES", "whole"),
    ("RUBBER TYRE SCRAP", "unspecified"),
    (None, "unspecified"),
])
def test_classify_grade(description, expected):
    assert classify_grade(description) == expected
