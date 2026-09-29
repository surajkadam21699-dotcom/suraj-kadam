"""India's official import statistics from the UN Comtrade API.

Comtrade publishes the customs value (CIF, USD) and net weight of imports per
partner country, HS code and month, so value / weight gives the average import
price per tonne. Indian customs files scrap tyres under HS 400400 ("waste,
parings and scrap of rubber"). The public preview endpoint needs no key but
returns at most 500 rows per call; set COMTRADE_API_KEY (free registration at
https://comtradedeveloper.un.org) to use the full endpoint instead.
"""

from __future__ import annotations

import logging
import time

import requests

from .records import PriceRecord, build_record

log = logging.getLogger(__name__)

PREVIEW_URL = "https://comtradeapi.un.org/public/v1/preview/C/{freq}/HS"
DATA_URL = "https://comtradeapi.un.org/data/v1/get/C/{freq}/HS"
PREVIEW_ROW_LIMIT = 500
KEYED_ROW_LIMIT = 100_000
INDIA = 699
SCRAP_TYRE_HS = ("400400",)


class ComtradeError(RuntimeError):
    """The API answered, but with an error message instead of data."""


def month_periods(start: tuple[int, int], end: tuple[int, int]) -> list[str]:
    """Comtrade period codes from ``start`` to ``end`` inclusive: (2025, 12), (2026, 1) -> ["202512", "202601"]."""
    (year, month), periods = start, []
    while (year, month) <= end:
        periods.append(f"{year}{month:02d}")
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return periods


def fetch(
    periods: list[str],
    *,
    session: requests.Session,
    hs_codes: tuple[str, ...] | list[str] = SCRAP_TYRE_HS,
    reporter: int = INDIA,
    freq: str = "M",
    api_key: str | None = None,
    pause: float = 1.0,
) -> list[dict]:
    """Raw import rows for each period, one request per period as the official client does.

    A period that fails is logged and skipped; the error is raised only if every period fails.
    """
    url = (DATA_URL if api_key else PREVIEW_URL).format(freq=freq)
    row_limit = KEYED_ROW_LIMIT if api_key else PREVIEW_ROW_LIMIT
    rows, errors = [], []
    for index, period in enumerate(periods):
        if index and pause:
            time.sleep(pause)
        params = {
            "reportercode": reporter,
            "period": period,
            "cmdCode": ",".join(hs_codes),
            "flowCode": "M",
            "includeDesc": "true",
            "maxRecords": row_limit,
        }
        if api_key:
            params["subscription-key"] = api_key
        try:
            response = session.get(url, params=params, timeout=60)
            response.raise_for_status()
            payload = response.json()
            if payload.get("error"):
                raise ComtradeError(payload["error"])
        except (requests.RequestException, ValueError, ComtradeError) as error:
            log.warning("Comtrade period %s failed: %s", period, error)
            errors.append(error)
            continue
        data = payload.get("data") or []
        if len(data) >= row_limit:
            log.warning("Comtrade period %s hit the %d-row limit; results may be incomplete", period, row_limit)
        log.info("Comtrade period %s: %d rows", period, len(data))
        rows.extend(data)
    if periods and len(errors) == len(periods):
        raise errors[-1]
    return rows


def parse(rows: list[dict]) -> list[PriceRecord]:
    """Price records from Comtrade rows, keeping only the headline figure per period, partner and HS code."""
    records, seen = [], set()
    for row in rows:
        if not _is_headline(row):
            continue  # transport-mode / customs-procedure breakdowns repeat the totals
        key = (str(row.get("period")), str(row.get("partnerCode")), str(row.get("cmdCode")))
        if key in seen:
            continue
        seen.add(key)
        records.append(build_record(
            source="comtrade",
            period=_period_label(row),
            origin=_partner_name(row),
            hs_code=str(row.get("cmdCode") or "") or None,
            description=row.get("cmdDesc"),
            quantity=row.get("qty"),
            unit=row.get("qtyUnitAbbr"),
            weight_kg=row.get("netWgt"),
            value=row.get("primaryValue"),
            currency="USD",
        ))
    return records


def _is_headline(row: dict) -> bool:
    """Imports in total: no second partner, transport mode or customs procedure breakdown."""
    return (
        str(row.get("flowCode") or "M") == "M"
        and str(row.get("partner2Code") or 0) == "0"
        and str(row.get("motCode") or 0) == "0"
        and str(row.get("customsCode") or "C00") == "C00"
    )


def _period_label(row: dict) -> str | None:
    period = str(row.get("period") or "")
    return f"{period[:4]}-{period[4:6]}" if len(period) == 6 else period or None


def _partner_name(row: dict) -> str:
    if str(row.get("partnerCode")) == "0":
        return "World"
    return row.get("partnerDesc") or row.get("partnerISO") or str(row.get("partnerCode"))
