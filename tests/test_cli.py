import csv
import json
from datetime import date

import pytest

from tyre_prices import shipments
from tyre_prices.cli import comtrade_periods, main

TODAY = date(2026, 9, 28)


def unreachable(url, params):
    raise AssertionError(f"unexpected request to {url}")


def test_shipments_from_saved_pages(tmp_path, capsys, fixtures_dir, fake_session):
    session = fake_session(unreachable)
    summary_md = tmp_path / "summary.md"
    code = main([
        "shipments",
        "--file", str(fixtures_dir / "shipments_zauba_style.html"),
        "--file", str(fixtures_dir / "shipments_usd_layout.html"),
        "--inr-per-usd", "90",
        "--out", str(tmp_path / "out"),
        "--summary-md", str(summary_md),
    ], session=session)

    assert code == 0
    assert session.calls == []
    with open(tmp_path / "out" / "shipments_records.csv", newline="", encoding="utf-8") as handle:
        assert len(list(csv.DictReader(handle))) == 7
    summary = json.loads((tmp_path / "out" / "shipments_summary.json").read_text(encoding="utf-8"))
    assert (summary["records"], summary["records_with_price_per_tonne"]) == (7, 6)
    assert [table["group"] for table in summary["tables"]] == ["Grade", "Origin", "Month"]
    assert {row["currency"] for table in summary["tables"] for row in table["rows"]} == {"USD"}
    assert "Scrap tyre shipments: price per tonne by grade" in capsys.readouterr().out
    assert "### By origin country" in summary_md.read_text(encoding="utf-8")


def test_shipments_fails_when_every_default_page_is_blocked(tmp_path, fake_session, fake_response):
    session = fake_session(lambda url, params: fake_response(status=403))
    assert main(["shipments", "--pause", "0", "--out", str(tmp_path)], session=session) == 1
    assert [url for url, _ in session.calls] == list(shipments.DEFAULT_URLS)


def test_comtrade_command(tmp_path, capsys, monkeypatch, fixtures_dir, fake_session, fake_response):
    monkeypatch.delenv("COMTRADE_API_KEY", raising=False)
    payload = json.loads((fixtures_dir / "comtrade_india_400400.json").read_text(encoding="utf-8"))
    empty = {"count": 0, "data": [], "error": ""}
    session = fake_session(lambda url, params: fake_response(payload if params["period"] == "202601" else empty))

    code = main(["comtrade", "--start", "2025-12", "--end", "2026-02", "--pause", "0", "--out", str(tmp_path)],
                session=session)

    assert code == 0
    assert [params["period"] for _, params in session.calls] == ["202512", "202601", "202602"]
    tables = json.loads((tmp_path / "comtrade_summary.json").read_text(encoding="utf-8"))["tables"]
    by_period, latest, by_origin = tables
    assert [(row["group"], row["weighted_avg"]) for row in by_period["rows"]] == [("2026-01", 210)]
    assert latest["title"] == "By origin country, 2026-01 (latest)"
    assert [row["group"] for row in latest["rows"]] == ["United Kingdom", "Australia", "New Zealand"]
    assert [row["group"] for row in by_origin["rows"]] == ["United Kingdom", "Australia", "New Zealand"]
    assert "India imports of HS 400400: CIF price per tonne, all origins" in capsys.readouterr().out


def test_comtrade_latest_table_uses_the_last_month_with_prices(tmp_path, monkeypatch, fixtures_dir, fake_session,
                                                               fake_response):
    monkeypatch.delenv("COMTRADE_API_KEY", raising=False)
    january = json.loads((fixtures_dir / "comtrade_india_400400.json").read_text(encoding="utf-8"))
    february = json.loads(json.dumps(january).replace('"202601"', '"202602"'))
    february["data"] = [row for row in february["data"] if row["partnerDesc"] in ("World", "Australia")]
    session = fake_session(lambda url, params: fake_response(january if params["period"] == "202601" else february))

    assert main(["comtrade", "--start", "2026-01", "--end", "2026-02", "--pause", "0", "--out", str(tmp_path)],
                session=session) == 0
    by_period, latest, by_origin = json.loads((tmp_path / "comtrade_summary.json").read_text())["tables"]
    assert [row["group"] for row in by_period["rows"]] == ["2026-01", "2026-02"]
    assert [row["group"] for row in latest["rows"]] == ["Australia"]
    assert by_origin["title"] == "By origin country, 2026-01 to 2026-02"
    assert [row["records"] for row in by_origin["rows"]] == [2, 1, 1]  # Australia, United Kingdom, New Zealand


def test_comtrade_command_without_published_data(tmp_path, monkeypatch, fake_session, fake_response):
    monkeypatch.delenv("COMTRADE_API_KEY", raising=False)
    session = fake_session(lambda url, params: fake_response({"count": 0, "data": [], "error": ""}))
    assert main(["comtrade", "--months", "2", "--pause", "0", "--out", str(tmp_path)], session=session) == 1


def test_default_period_is_the_24_months_to_last_month():
    periods = comtrade_periods(None, None, 24, False, TODAY)
    assert (len(periods), periods[0], periods[-1]) == (24, "202409", "202608")


def test_explicit_month_ranges():
    assert comtrade_periods("2025-11", "2026-02", 24, False, TODAY) == ["202511", "202512", "202601", "202602"]
    assert comtrade_periods(None, "2026-03", 3, False, TODAY) == ["202601", "202602", "202603"]


def test_annual_periods():
    assert comtrade_periods(None, None, 24, True, TODAY) == ["2023", "2024", "2025"]
    assert comtrade_periods("2020", "2021", 24, True, TODAY) == ["2020", "2021"]


@pytest.mark.parametrize("start, end, annual", [
    ("2026", None, False),
    ("2026-01", None, True),
    ("2026-05", "2026-01", False),
    ("2026", "2025", True),
])
def test_invalid_period_ranges(start, end, annual):
    with pytest.raises(ValueError):
        comtrade_periods(start, end, 24, annual, TODAY)


def test_malformed_period_is_rejected_by_the_parser():
    with pytest.raises(SystemExit):
        main(["comtrade", "--start", "2026-13"])
