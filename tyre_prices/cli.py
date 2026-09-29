"""Command line: python -m tyre_prices {comtrade,shipments} [options]."""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
from dataclasses import asdict
from datetime import date, datetime, timezone
from pathlib import Path

import requests

from . import comtrade, shipments
from .net import make_session
from .records import PriceRecord, write_csv
from .summary import GroupStats, render_markdown, render_text, summarize

log = logging.getLogger("tyre_prices")

Table = tuple[str, str, list[GroupStats]]  # title, group column label, rows


def main(argv: list[str] | None = None, session: requests.Session | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s: %(message)s")
    return args.handler(args, session or make_session())


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m tyre_prices",
        description="Scrape scrap tyre import prices (India, HS 4004) and summarise them per tonne.",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="show debug logging")
    output = argparse.ArgumentParser(add_help=False)
    output.add_argument("--out", default="output", help="directory for the CSV and JSON results (default: output)")
    output.add_argument(
        "--summary-md", metavar="PATH",
        help="also append the summary tables as Markdown to PATH, e.g. $GITHUB_STEP_SUMMARY",
    )
    commands = parser.add_subparsers(dest="command", required=True)

    stats = commands.add_parser(
        "comtrade", parents=[output],
        help="official import statistics from UN Comtrade: USD per tonne by month and origin country",
    )
    stats.add_argument("--months", type=int, default=24, help="months to fetch, ending at --end (default: 24)")
    stats.add_argument("--start", type=_period_arg, help="first period: YYYY-MM, or YYYY with --annual")
    stats.add_argument("--end", type=_period_arg, help="last period (default: last month, or last year with --annual)")
    stats.add_argument("--annual", action="store_true", help="fetch yearly instead of monthly figures")
    stats.add_argument(
        "--hs", action="append", metavar="CODE",
        help="HS code to fetch, repeatable (default: 400400, scrap rubber including scrap tyres)",
    )
    stats.add_argument("--reporter", type=int, default=comtrade.INDIA, help="Comtrade reporter (default: 699, India)")
    stats.add_argument("--pause", type=float, default=1.0, help="seconds between API calls (default: 1)")
    stats.set_defaults(handler=run_comtrade)

    pages = commands.add_parser("shipments", parents=[output], help="shipment-level prices from trade-data web pages")
    pages.add_argument(
        "--url", action="append",
        help="page to scrape, repeatable; may contain {page} (default: built-in scrap tyre pages)",
    )
    pages.add_argument("--file", action="append", help="saved HTML page to parse instead of fetching one, repeatable")
    pages.add_argument("--pages", type=int, default=1, help="page numbers to fill into {page} in URLs (default: 1)")
    pages.add_argument("--currency", default="INR", help="currency of value columns naming none (default: INR)")
    pages.add_argument("--inr-per-usd", type=float, help="exchange rate to add USD prices to INR records, e.g. 88.5")
    pages.add_argument("--all-rows", action="store_true", help="keep rows that do not look like scrap tyres")
    pages.add_argument("--pause", type=float, default=1.5, help="seconds between page requests (default: 1.5)")
    pages.set_defaults(handler=run_shipments)
    return parser


def _period_arg(text: str) -> str:
    if not re.fullmatch(r"\d{4}(?:-(?:0[1-9]|1[0-2]))?", text):
        raise argparse.ArgumentTypeError(f"expected YYYY-MM or YYYY, got {text!r}")
    return text


def comtrade_periods(start: str | None, end: str | None, months: int, annual: bool, today: date) -> list[str]:
    """Comtrade period codes for the requested range; by default the months up to last month."""
    if annual:
        if any(period and len(period) != 4 for period in (start, end)):
            raise ValueError("--annual takes years (YYYY) for --start and --end")
        last = int(end) if end else today.year - 1
        first = int(start) if start else last - 2
        periods = [str(year) for year in range(first, last + 1)]
    else:
        if any(period and len(period) != 7 for period in (start, end)):
            raise ValueError("monthly periods are YYYY-MM (add --annual for years)")
        if months < 1:
            raise ValueError("--months must be at least 1")
        last = _month(end) if end else _add_months((today.year, today.month), -1)
        first = _month(start) if start else _add_months(last, 1 - months)
        periods = comtrade.month_periods(first, last)
    if not periods:
        raise ValueError("--start is after --end")
    return periods


def _month(text: str) -> tuple[int, int]:
    return int(text[:4]), int(text[5:7])


def _add_months(month: tuple[int, int], delta: int) -> tuple[int, int]:
    year, index = divmod(month[0] * 12 + month[1] - 1 + delta, 12)
    return year, index + 1


def run_comtrade(args: argparse.Namespace, session: requests.Session) -> int:
    try:
        periods = comtrade_periods(args.start, args.end, args.months, args.annual, date.today())
    except ValueError as error:
        log.error("%s", error)
        return 2
    hs_codes = args.hs or list(comtrade.SCRAP_TYRE_HS)
    api_key = os.environ.get("COMTRADE_API_KEY") or None
    log.info(
        "UN Comtrade%s: reporter %s, HS %s, %s to %s",
        "" if api_key else " (public preview API)", args.reporter, ",".join(hs_codes), periods[0], periods[-1],
    )
    try:
        rows = comtrade.fetch(
            periods, session=session, hs_codes=hs_codes, reporter=args.reporter,
            freq="A" if args.annual else "M", api_key=api_key, pause=args.pause,
        )
    except (requests.RequestException, ValueError, comtrade.ComtradeError) as error:
        log.error("UN Comtrade request failed: %s", error)
        return 1
    records = comtrade.parse(rows)
    if not records:
        log.error("UN Comtrade has no import figures for these periods yet (monthly data appears a few months late)")
        return 1

    several = len(hs_codes) > 1
    reported = sorted(record.period for record in records if record.period and record.price_per_tonne)
    reporter = "India" if args.reporter == comtrade.INDIA else f"Reporter {args.reporter}"
    partners = [record for record in records if record.origin != "World"]

    def by_origin(record: PriceRecord) -> str | None:
        return f"{record.origin} (HS {record.hs_code})" if several else record.origin

    tables: list[Table] = [
        (
            f"{reporter} imports of HS {', '.join(hs_codes)}: CIF price per tonne, all origins",
            "Period",
            summarize(
                [record for record in records if record.origin == "World"],
                lambda record: f"{record.period} HS {record.hs_code}" if several else record.period,
            ),
        ),
    ]
    if reported:
        tables += [
            (
                f"By origin country, {reported[-1]} (latest)",
                "Origin",
                summarize([r for r in partners if r.period == reported[-1]], by_origin, sort_by="tonnes"),
            ),
            (
                f"By origin country, {reported[0]} to {reported[-1]}",
                "Origin",
                summarize(partners, by_origin, sort_by="tonnes"),
            ),
        ]
    _emit(args, "comtrade", records, tables)
    return 0


def run_shipments(args: argparse.Namespace, session: requests.Session) -> int:
    files = args.file or []
    urls = args.url or ([] if files else list(shipments.DEFAULT_URLS))
    records, failures = shipments.scrape(
        urls, files, session=session, pages=args.pages, pause=args.pause,
        default_currency=args.currency.upper(), inr_per_usd=args.inr_per_usd, scrap_only=not args.all_rows,
    )
    if not records:
        log.error(
            "no scrap tyre shipments found%s. If a site blocks automated requests, save the page "
            "from your browser and parse it with --file",
            f" ({len(failures)} page(s) could not be fetched)" if failures else "",
        )
        return 1
    if args.inr_per_usd is None and any(record.currency == "INR" for record in records):
        log.info("prices are in INR; add --inr-per-usd RATE to convert them to USD")

    subject = "Shipments" if args.all_rows else "Scrap tyre shipments"
    tables: list[Table] = [
        (f"{subject}: price per tonne by grade", "Grade", summarize(records, lambda r: r.grade, sort_by="tonnes")),
        ("By origin country", "Origin", summarize(records, lambda r: r.origin, sort_by="tonnes")),
        ("By month", "Month", summarize(records, lambda r: r.month)),
    ]
    _emit(args, "shipments", records, tables)
    return 0


def _emit(args: argparse.Namespace, name: str, records: list[PriceRecord], tables: list[Table]) -> None:
    """Write <name>_records.csv and <name>_summary.json, print the tables, append Markdown if asked."""
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    csv_path, json_path = out / f"{name}_records.csv", out / f"{name}_summary.json"
    write_csv(records, csv_path)
    priced = sum(1 for record in records if record.price_per_tonne is not None)
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "records": len(records),
        "records_with_price_per_tonne": priced,
        "tables": [
            {"title": title, "group": label, "rows": [_rounded(asdict(stat)) for stat in stats]}
            for title, label, stats in tables
        ],
    }
    json_path.write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")

    print(f"{len(records)} records, {priced} with a price per tonne\n")
    for title, label, stats in tables:
        print(render_text(title, stats, label) + "\n")
    if args.summary_md:
        with open(args.summary_md, "a", encoding="utf-8") as handle:
            handle.write("\n\n".join(render_markdown(title, stats, label) for title, label, stats in tables) + "\n\n")
    log.info("wrote %s and %s", csv_path, json_path)


def _rounded(row: dict) -> dict:
    return {key: round(value, 4) if isinstance(value, float) else value for key, value in row.items()}
