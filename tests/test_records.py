import csv

import pytest

from tyre_prices.records import FIELDNAMES, build_record, write_csv


def test_price_per_tonne_from_value_and_weight():
    record = build_record(
        source="t", period="2026-01-05", quantity=108000, unit="KGS", value=2160000, currency="INR", inr_per_usd=90,
    )
    assert record.tonnes == 108
    assert record.unit_price == 20
    assert record.price_per_tonne == 20000
    assert record.price_per_tonne_usd == pytest.approx(222.22, abs=0.01)


def test_price_per_tonne_from_unit_price_alone():
    record = build_record(source="t", period=None, unit="KGS", unit_price=21.5, currency="INR")
    assert record.price_per_tonne == 21500
    assert record.value is None and record.tonnes is None
    assert record.price_per_tonne_usd is None  # no exchange rate given


def test_value_derived_from_unit_price_and_quantity():
    record = build_record(source="t", period=None, quantity=25, unit="MTS", unit_price=21000, currency="INR")
    assert record.value == 525000
    assert record.price_per_tonne == 21000


def test_count_units_have_no_price_per_tonne():
    record = build_record(source="t", period=None, quantity=500, unit="NOS", value=750000, currency="INR")
    assert record.unit_price == 1500
    assert record.tonnes is None and record.price_per_tonne is None


def test_reported_weight_wins_over_quantity():
    record = build_record(
        source="t", period=None, quantity=1000, unit="u", weight_kg=8000, value=4000, currency="USD",
    )
    assert record.tonnes == 8
    assert record.unit_price == 4
    assert record.price_per_tonne == record.price_per_tonne_usd == 500


def test_zero_quantity_is_not_divided_by():
    record = build_record(source="t", period=None, quantity=0, unit="KGS", value=100, currency="USD")
    assert record.unit_price is None and record.price_per_tonne is None


def test_write_csv(tmp_path):
    record = build_record(source="t", period="2026-01-05", quantity=108000, unit="KGS", value=2160000, currency="INR")
    path = tmp_path / "records.csv"
    write_csv([record], path)
    with open(path, newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    assert list(rows[0]) == FIELDNAMES
    assert rows[0]["price_per_tonne"] == "20000"
    assert rows[0]["tonnes"] == "108"
    assert rows[0]["port"] == ""
