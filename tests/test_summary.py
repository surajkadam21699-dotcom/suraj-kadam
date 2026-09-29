from tyre_prices.records import PriceRecord
from tyre_prices.summary import GroupStats, render_markdown, render_text, summarize


def priced(price, tonnes, origin, currency="USD"):
    return PriceRecord(
        source="t", period="2026-01-01", origin=origin, port=None, hs_code=None, grade=None, description=None,
        quantity=None, unit=None, tonnes=tonnes, value=None, currency=currency, unit_price=None,
        price_per_tonne=price, price_per_tonne_usd=price if currency == "USD" else None,
    )


def test_weighted_average_median_and_order_by_tonnes():
    stats = summarize(
        [priced(200, 10, "UK"), priced(300, 30, "UK"), priced(250, 5, "AU")],
        key=lambda record: record.origin,
        sort_by="tonnes",
    )
    uk, au = stats
    assert (uk.group, uk.records, uk.tonnes, uk.currency) == ("UK", 2, 40, "USD")
    assert uk.weighted_avg == 275  # (200 * 10 + 300 * 30) / 40
    assert (uk.median, uk.low, uk.high) == (250, 200, 300)
    assert au.group == "AU"


def test_currencies_are_kept_apart():
    stats = summarize([priced(21000, 10, "UK", currency="INR"), priced(230, 10, "UK")], key=lambda r: r.origin)
    assert [(stat.currency, stat.median) for stat in stats] == [("INR", 21000), ("USD", 230)]


def test_unpriced_records_are_skipped_and_missing_groups_labelled():
    stats = summarize([priced(None, 10, "UK"), priced(100, None, None)], key=lambda r: r.origin)
    assert len(stats) == 1
    assert stats[0].group == "unknown"
    assert stats[0].weighted_avg is None  # no weight to average over


def test_records_under_the_minimum_weight_are_left_out():
    records = [priced(95, 20, "Bahrain"), priced(92752, 0.001, "Bahrain")]
    [stat] = summarize(records, key=lambda r: r.origin)
    assert (stat.records, stat.high) == (1, 95)
    assert summarize(records, key=lambda r: r.origin, min_tonnes=0)[0].high == 92752


STAT = GroupStats("United Kingdom", "USD", 2, 40.0, 275.0, 250.0, 200.0, 300.0)


def test_render_text():
    lines = render_text("Title", [STAT], "Origin").splitlines()
    assert lines[0] == "Title"
    assert lines[1].split() == ["Origin", "Records", "Tonnes", "Wtd", "avg/t", "Median/t", "Min/t", "Max/t", "Currency"]
    assert lines[2].split() == ["United", "Kingdom", "2", "40.0", "275", "250", "200", "300", "USD"]
    assert len(lines[1]) >= len(lines[2])


def test_render_markdown():
    markdown = render_markdown("Title", [STAT], "Origin")
    assert markdown.splitlines()[0] == "### Title"
    assert "| United Kingdom | 2 | 40.0 | 275 | 250 | 200 | 300 | USD |" in markdown


def test_render_empty_tables():
    assert "no records" in render_text("Title", [])
    assert "No records" in render_markdown("Title", [])
