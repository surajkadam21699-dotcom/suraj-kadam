"""Aggregate price records into per-group price-per-tonne statistics and render them as tables."""

from __future__ import annotations

import statistics
from collections import defaultdict
from collections.abc import Callable, Iterable
from dataclasses import dataclass

from .records import PriceRecord


@dataclass
class GroupStats:
    group: str
    currency: str
    records: int
    tonnes: float
    weighted_avg: float | None  # total value / total tonnes
    median: float
    low: float
    high: float


def record_price(record: PriceRecord, prefer_usd: bool = True) -> tuple[float | None, str | None]:
    """The per-tonne price summaries use, and its currency (USD whenever it is known)."""
    if prefer_usd and record.price_per_tonne_usd is not None:
        return record.price_per_tonne_usd, "USD"
    if record.price_per_tonne is not None:
        return record.price_per_tonne, record.currency or "?"
    return None, None


def summarize(
    records: Iterable[PriceRecord],
    key: Callable[[PriceRecord], str | None],
    sort_by: str = "group",
    prefer_usd: bool = True,
) -> list[GroupStats]:
    """Price statistics per ``key(record)`` and currency; records without a per-tonne price are skipped."""
    groups: dict[tuple[str, str], list[tuple[float, float]]] = defaultdict(list)
    for record in records:
        price, currency = record_price(record, prefer_usd)
        if price is not None:
            groups[(key(record) or "unknown", currency)].append((price, record.tonnes or 0.0))

    stats = []
    for (group, currency), items in groups.items():
        prices = [price for price, _ in items]
        tonnes = sum(weight for _, weight in items)
        stats.append(GroupStats(
            group=group,
            currency=currency,
            records=len(items),
            tonnes=tonnes,
            weighted_avg=sum(price * weight for price, weight in items) / tonnes if tonnes else None,
            median=statistics.median(prices),
            low=min(prices),
            high=max(prices),
        ))
    if sort_by == "tonnes":
        return sorted(stats, key=lambda s: (-s.tonnes, s.group))
    return sorted(stats, key=lambda s: (s.group, s.currency))


def _header(label: str) -> tuple[str, ...]:
    return (label, "Records", "Tonnes", "Wtd avg/t", "Median/t", "Min/t", "Max/t", "Currency")


def _cells(stat: GroupStats) -> tuple[str, ...]:
    def money(amount):
        return "-" if amount is None else f"{amount:,.0f}"

    return (
        stat.group, f"{stat.records:,}", f"{stat.tonnes:,.1f}",
        money(stat.weighted_avg), money(stat.median), money(stat.low), money(stat.high), stat.currency,
    )


def render_text(title: str, stats: list[GroupStats], label: str = "Group") -> str:
    if not stats:
        return f"{title}\n  (no records with a per-tonne price)"
    rows = [_header(label)] + [_cells(stat) for stat in stats]
    widths = [max(len(row[i]) for row in rows) for i in range(len(rows[0]))]
    lines = [title]
    for row in rows:
        # Group and currency read left to right; the numbers line up on the right.
        cells = [
            cell.ljust(width) if i in (0, len(row) - 1) else cell.rjust(width)
            for i, (cell, width) in enumerate(zip(row, widths, strict=True))
        ]
        lines.append("  ".join(cells).rstrip())
    return "\n".join(lines)


def render_markdown(title: str, stats: list[GroupStats], label: str = "Group") -> str:
    lines = [f"### {title}", ""]
    if not stats:
        return "\n".join(lines + ["_No records with a per-tonne price._"])
    header = _header(label)
    lines.append("| " + " | ".join(header) + " |")
    lines.append("|" + "|".join([" --- "] + [" ---: "] * (len(header) - 2) + [" --- "]) + "|")
    for stat in stats:
        lines.append("| " + " | ".join(cell.replace("|", "\\|") for cell in _cells(stat)) + " |")
    return "\n".join(lines)
