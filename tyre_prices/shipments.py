"""Shipment-level import prices from trade-data web pages (Zauba, Seair and similar sites).

These sites publish sample Indian customs records as an HTML table with columns
such as Date, HS Code, Description, Origin Country, Port of Discharge, Unit,
Quantity, Value (INR) and Per Unit (INR). The column names differ from site to
site, so shipment tables are recognised by their header text rather than by
site-specific markup, and a page saved from a browser parses the same way.
"""

from __future__ import annotations

import logging
import re
import time
from collections.abc import Iterable, Iterator
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

from .grades import classify_grade, is_scrap_tyre
from .normalize import detect_currency, parse_date, parse_number
from .records import PriceRecord, build_record

log = logging.getLogger(__name__)

# Public pages listing sample Indian import shipments of scrap tyres (HS 4004).
DEFAULT_URLS = (
    "https://www.zauba.com/import-scrap-tyre/hs-code-40040000-hs-code.html",
    "https://www.zauba.com/import-tyre+scrap-hs-code.html",
    "https://www.seair.co.in/scrap-tyre-import-data.aspx",
    "https://www.seair.co.in/tyre-scrap-import-data/hs-code-4004000.aspx",
    "https://www.seair.co.in/shredded-tyre-import-data/hs-code-40040000.aspx",
)

# Header names seen on the sites, after _header_key() normalisation.
_KNOWN_HEADERS = {
    "date": "date", "shipment date": "date", "arrival date": "date", "be date": "date",
    "bill of entry date": "date",
    "hs code": "hs_code", "hsn code": "hs_code", "hs": "hs_code", "hsn": "hs_code", "hscode": "hs_code",
    "ritc code": "hs_code",
    "description": "description", "product description": "description", "item description": "description",
    "goods description": "description", "product": "description",
    "origin country": "origin", "country of origin": "origin", "origin": "origin", "country": "origin",
    "supplier country": "origin",
    "port of discharge": "port", "indian port": "port", "port": "port", "destination port": "port",
    "port of unloading": "port",
    "unit": "unit", "uqc": "unit", "std unit": "unit", "qty unit": "unit", "quantity unit": "unit",
    "unit of measure": "unit", "uom": "unit",
    "quantity": "quantity", "qty": "quantity", "std quantity": "quantity", "std qty": "quantity",
    "value": "value", "total value": "value", "assessable value": "value", "cif value": "value",
    "amount": "value",
    "per unit": "unit_price", "unit price": "unit_price", "std unit price": "unit_price",
    "unit rate": "unit_price", "rate": "unit_price", "price": "unit_price", "price per unit": "unit_price",
    "unit value": "unit_price",
}
# Columns to leave alone even though they contain a known word (trading parties, loading port, row numbers).
_IGNORED_HEADER = re.compile(
    r"importer|exporter|buyer|shipper|consignee|supplier name|address|loading|foreign port|serial|^(?:s|sr) ?no$|^#$"
)
# Fallback for unfamiliar headers, checked in order; the first match wins.
_HEADER_RULES = (
    ("unit_price", re.compile(r"per unit|unit (?:price|rate|value)|price per|\brate\b|\bprice\b")),
    ("value", re.compile(r"value|amount")),
    ("unit", re.compile(r"\bunits?\b|\buqc\b|\buom\b")),
    ("quantity", re.compile(r"quantity|\bqty\b")),
    ("date", re.compile(r"date")),
    ("hs_code", re.compile(r"\bhsn?\b|\britc\b|\bcth\b|tariff")),
    ("description", re.compile(r"descr|product|item|goods|commodity")),
    ("port", re.compile(r"\bport\b|discharge|\bicd\b")),
    ("origin", re.compile(r"origin|country")),
)


def _header_key(text: str) -> str:
    """Normalise a header: "Per Unit (INR)" -> "per unit", "Std. Quantity" -> "std quantity"."""
    text = re.sub(r"\(.*?\)|\[.*?\]", " ", text.lower())
    return " ".join(re.sub(r"[^a-z0-9#]+", " ", text).split())


def _classify_header(text: str) -> str | None:
    key = _header_key(text)
    if key in _KNOWN_HEADERS:
        return _KNOWN_HEADERS[key]
    if not key or _IGNORED_HEADER.search(key):
        return None
    for field, pattern in _HEADER_RULES:
        if pattern.search(key):
            return field
    return None


def _map_columns(header: list[str]) -> dict[str, tuple[int, str | None]]:
    """Field -> (column index, currency named in the header); the first column for a field wins."""
    columns: dict[str, tuple[int, str | None]] = {}
    for index, text in enumerate(header):
        field = _classify_header(text)
        if field and field not in columns:
            columns[field] = (index, detect_currency(text))
    return columns


def _is_shipment_header(columns: dict) -> bool:
    return (
        len(columns) >= 3
        and ("value" in columns or "unit_price" in columns)
        and any(field in columns for field in ("description", "hs_code", "quantity"))
    )


def _table_rows(table) -> Iterator[list[str]]:
    """Cell texts of each row belonging to ``table`` itself (not to nested tables), colspans expanded."""
    for tr in table.find_all("tr"):
        if tr.find_parent("table") is not table:
            continue
        cells = []
        for cell in tr.find_all(["td", "th"], recursive=False):
            text = " ".join(cell.get_text(" ", strip=True).split())
            span = str(cell.get("colspan") or "1")
            cells.extend([text] * (min(int(span), 50) if span.isdecimal() and int(span) > 0 else 1))
        if cells:
            yield cells


def parse_page(
    html: str | bytes,
    *,
    source: str,
    default_currency: str = "INR",
    inr_per_usd: float | None = None,
    scrap_only: bool = True,
) -> list[PriceRecord]:
    """Shipment records from every shipment table on a page."""
    soup = BeautifulSoup(html, "html.parser")
    records, tables = [], 0
    for table in soup.find_all("table"):
        rows = list(_table_rows(table))
        for position, cells in enumerate(rows):
            columns = _map_columns(cells)
            if _is_shipment_header(columns):
                tables += 1
                records += _parse_rows(
                    rows[position + 1:], columns, source=source, default_currency=default_currency,
                    inr_per_usd=inr_per_usd, scrap_only=scrap_only,
                )
                break
    if not tables:
        log.warning("%s: no shipment table found (the page may need JavaScript, a login, or be a bot check)", source)
    return records


def _parse_rows(rows, columns, *, source, default_currency, inr_per_usd, scrap_only) -> list[PriceRecord]:
    value_currency = columns.get("value", (None, None))[1]
    price_currency = columns.get("unit_price", (None, None))[1]
    currency = value_currency or price_currency or default_currency
    # A per-unit column quoted in another currency than the value column is ignored; value / quantity replaces it.
    use_unit_price = "unit_price" in columns and ("value" not in columns or price_currency in (None, currency))

    records = []
    for cells in rows:
        description = _cell(cells, columns, "description")
        hs_code = re.sub(r"\D", "", _cell(cells, columns, "hs_code") or "") or None
        if scrap_only and not is_scrap_tyre(description, hs_code):
            continue
        value = parse_number(_cell(cells, columns, "value"))
        unit_price = parse_number(_cell(cells, columns, "unit_price")) if use_unit_price else None
        if value is None and unit_price is None:
            continue  # masked figures, "view more" rows, repeated headers
        quantity = _cell(cells, columns, "quantity")
        # Some sites have no Unit column and write the unit after the number: "1,08,000 KGS".
        unit = _cell(cells, columns, "unit") or _trailing_unit(quantity)
        records.append(build_record(
            source=source,
            period=parse_date(_cell(cells, columns, "date")),
            origin=_cell(cells, columns, "origin"),
            port=_cell(cells, columns, "port"),
            hs_code=hs_code,
            grade=classify_grade(description),
            description=description,
            quantity=parse_number(quantity),
            unit=unit,
            value=value,
            currency=currency,
            unit_price=unit_price,
            inr_per_usd=inr_per_usd,
        ))
    return records


def _cell(cells: list[str], columns: dict, field: str) -> str | None:
    index = columns.get(field, (None,))[0]
    if index is None or index >= len(cells):
        return None
    return cells[index] or None


def _trailing_unit(quantity: str | None) -> str | None:
    match = re.search(r"\d\s*([A-Za-z][A-Za-z.]*)\s*$", quantity or "")
    return match.group(1) if match else None


def expand_urls(urls: Iterable[str], pages: int) -> Iterator[str]:
    """Substitute page numbers 1..pages into URLs containing "{page}"."""
    for url in urls:
        if "{page}" in url:
            for page in range(1, pages + 1):
                yield url.replace("{page}", str(page))
        else:
            yield url


def _shipment_key(record: PriceRecord) -> tuple:
    return (record.period, record.description, record.origin, record.port, record.quantity, record.value)


def scrape(
    urls: Iterable[str] = (),
    files: Iterable[str | Path] = (),
    *,
    session: requests.Session,
    pages: int = 1,
    pause: float = 1.5,
    **parse_options,
) -> tuple[list[PriceRecord], list[str]]:
    """Parse every URL and saved page; returns (records, URLs that could not be fetched).

    Shipments already seen on an earlier page are dropped, since sites list the same
    records under several product and HS code pages; repeats within a page are kept.
    """
    pages_html: list[tuple[str, bytes]] = []
    failures = []
    for index, url in enumerate(expand_urls(urls, pages)):
        if index and pause:
            time.sleep(pause)
        try:
            response = session.get(url, timeout=30)
            response.raise_for_status()
        except requests.RequestException as error:
            log.warning("%s: %s", url, error)
            failures.append(url)
            continue
        pages_html.append((urlparse(url).netloc.removeprefix("www."), response.content))
    for path in files:
        pages_html.append((Path(path).name, Path(path).read_bytes()))

    records, seen = [], set()
    for source, html in pages_html:
        found = parse_page(html, source=source, **parse_options)
        fresh = [record for record in found if _shipment_key(record) not in seen]
        seen.update(_shipment_key(record) for record in found)
        log.info("%s: %d scrap tyre shipments (%d new)", source, len(found), len(fresh))
        records += fresh
    return records, failures
