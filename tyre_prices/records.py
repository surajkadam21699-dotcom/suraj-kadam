"""The normalised price record every source produces, and CSV output."""

from __future__ import annotations

import csv
from dataclasses import asdict, dataclass, fields
from pathlib import Path

from .normalize import kg_per_unit, normalize_unit


@dataclass
class PriceRecord:
    source: str
    period: str | None  # YYYY-MM-DD for a shipment, YYYY-MM or YYYY for monthly/annual statistics
    origin: str | None
    port: str | None
    hs_code: str | None
    grade: str | None
    description: str | None
    quantity: float | None
    unit: str | None
    tonnes: float | None
    value: float | None
    currency: str | None
    unit_price: float | None  # value per reported unit, in `currency`
    price_per_tonne: float | None  # in `currency`
    price_per_tonne_usd: float | None

    @property
    def month(self) -> str | None:
        return self.period[:7] if self.period else None


def build_record(
    *,
    source: str,
    period: str | None,
    origin: str | None = None,
    port: str | None = None,
    hs_code: str | None = None,
    grade: str | None = None,
    description: str | None = None,
    quantity: float | None = None,
    unit: str | None = None,
    value: float | None = None,
    currency: str | None = None,
    unit_price: float | None = None,
    weight_kg: float | None = None,
    inr_per_usd: float | None = None,
) -> PriceRecord:
    """Build a record, deriving tonnes, unit price and price per tonne from what the source reported."""
    unit = normalize_unit(unit)
    kg_each = kg_per_unit(unit)
    if not weight_kg and quantity and kg_each:
        weight_kg = quantity * kg_each
    tonnes = weight_kg / 1000 if weight_kg else None

    if unit_price is None and value and quantity:
        unit_price = value / quantity
    if value is None and unit_price and quantity:
        value = unit_price * quantity

    price_per_tonne = None
    if value and tonnes:
        price_per_tonne = value / tonnes
    elif unit_price and kg_each:
        price_per_tonne = unit_price * 1000 / kg_each

    return PriceRecord(
        source=source,
        period=period,
        origin=origin,
        port=port,
        hs_code=hs_code,
        grade=grade,
        description=description,
        quantity=quantity,
        unit=unit,
        tonnes=tonnes,
        value=value,
        currency=currency,
        unit_price=unit_price,
        price_per_tonne=price_per_tonne,
        price_per_tonne_usd=to_usd(price_per_tonne, currency, inr_per_usd),
    )


def to_usd(amount: float | None, currency: str | None, inr_per_usd: float | None) -> float | None:
    if amount is None:
        return None
    if currency == "USD":
        return amount
    if currency == "INR" and inr_per_usd:
        return amount / inr_per_usd
    return None


FIELDNAMES = [field.name for field in fields(PriceRecord)]


def write_csv(records: list[PriceRecord], path: Path) -> None:
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
        writer.writeheader()
        for record in records:
            writer.writerow({name: _cell(value) for name, value in asdict(record).items()})


def _cell(value):
    if isinstance(value, float):
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return "" if value is None else value
