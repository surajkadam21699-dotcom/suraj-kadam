import pytest

from tyre_prices.normalize import detect_currency, kg_per_unit, normalize_unit, parse_date, parse_number


@pytest.mark.parametrize("text, expected", [
    ("1,23,456.50", 123456.5),
    ("₹ 21,60,000", 2160000.0),
    ("$11,812.50", 11812.5),
    ("123,456", 123456.0),
    ("0.32", 0.32),
    ("20", 20.0),
    ("12,345 KGS", 12345.0),
    ("2,3**", None),
    ("12XX", None),
    ("XXXX", None),
    ("-", None),
    ("", None),
    (None, None),
    (42, 42.0),
])
def test_parse_number(text, expected):
    assert parse_number(text) == expected


@pytest.mark.parametrize("text, expected", [
    ("05-Jan-2026", "2026-01-05"),
    ("05-JAN-2026", "2026-01-05"),
    ("5 Jan 2026", "2026-01-05"),
    ("05-Jan-26", "2026-01-05"),
    ("Apr 10, 2026", "2026-04-10"),
    ("April 10, 2026", "2026-04-10"),
    ("2026-04-10", "2026-04-10"),
    ("2026-04-10T00:00:00", "2026-04-10"),
    ("12/04/2026", "2026-04-12"),  # day before month, as Indian sites write it
    ("15-04-2026", "2026-04-15"),
    ("15.04.2026", "2026-04-15"),
    ("not a date", None),
    ("", None),
    (None, None),
])
def test_parse_date(text, expected):
    assert parse_date(text) == expected


@pytest.mark.parametrize("text, expected", [("Kgs.", "KGS"), ("M.T.", "MT"), ("kg", "KG"), ("", None), (None, None)])
def test_normalize_unit(text, expected):
    assert normalize_unit(text) == expected


@pytest.mark.parametrize("unit, expected", [
    ("KGS", 1.0), ("Kgs.", 1.0), ("kg", 1.0), ("MTS", 1000.0), ("M.T.", 1000.0), ("TON", 1000.0),
    ("LBS", 0.45359237), ("NOS", None), ("PCS", None), ("", None), (None, None),
])
def test_kg_per_unit(unit, expected):
    assert kg_per_unit(unit) == expected


@pytest.mark.parametrize("header, expected", [
    ("Value (INR)", "INR"),
    ("Per Unit (Rs.)", "INR"),
    ("Value in ₹", "INR"),
    ("Total Value (USD)", "USD"),
    ("Unit Price (US $)", "USD"),
    ("Value (EUR)", "EUR"),
    ("Value (GBP)", "GBP"),
    ("Value", None),
    ("Hours", None),
    (None, None),
])
def test_detect_currency(header, expected):
    assert detect_currency(header) == expected
