import json

import pytest

from tyre_prices import comtrade


@pytest.fixture
def payload(fixtures_dir):
    return json.loads((fixtures_dir / "comtrade_india_400400.json").read_text(encoding="utf-8"))


def test_parse_keeps_headline_rows_and_prices_them_per_tonne(payload):
    records = {record.origin: record for record in comtrade.parse(payload["data"])}
    assert set(records) == {"World", "United Kingdom", "Australia", "New Zealand", "Areas, nes"}
    world = records["World"]
    assert (world.source, world.period, world.hs_code, world.currency) == ("comtrade", "2026-01", "400400", "USD")
    assert world.tonnes == 100000
    assert world.price_per_tonne == world.price_per_tonne_usd == 210
    assert records["Australia"].price_per_tonne == 220
    # The sea-transport breakdown repeats the United Kingdom total and must not replace or double it.
    assert records["United Kingdom"].tonnes == 40000
    assert records["United Kingdom"].price_per_tonne == 210


def test_parse_uses_quantity_when_net_weight_is_missing(payload):
    new_zealand = next(r for r in comtrade.parse(payload["data"]) if r.origin == "New Zealand")
    assert new_zealand.tonnes == 25000
    assert new_zealand.price_per_tonne == 196


def test_parse_leaves_price_empty_without_any_weight(payload):
    unknown = next(r for r in comtrade.parse(payload["data"]) if r.origin == "Areas, nes")
    assert unknown.value == 1000
    assert unknown.tonnes is None and unknown.price_per_tonne is None


def test_parse_drops_repeated_rows(payload):
    assert len(comtrade.parse(payload["data"] * 2)) == 5


def test_month_periods_cross_the_year():
    assert comtrade.month_periods((2025, 11), (2026, 2)) == ["202511", "202512", "202601", "202602"]
    assert comtrade.month_periods((2026, 3), (2026, 2)) == []


def test_fetch_uses_the_preview_api_without_a_key(fake_session, fake_response):
    session = fake_session(lambda url, params: fake_response({"data": [{"period": params["period"]}], "error": ""}))
    rows = comtrade.fetch(["202601", "202602"], session=session, pause=0)
    assert rows == [{"period": "202601"}, {"period": "202602"}]
    url, params = session.calls[0]
    assert url == "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
    assert params["reportercode"] == 699
    assert params["cmdCode"] == "400400"
    assert params["flowCode"] == "M"
    assert params["includeDesc"] == "true"
    assert params["maxRecords"] == 500
    assert "subscription-key" not in params


def test_fetch_uses_the_data_api_with_a_key(fake_session, fake_response):
    session = fake_session(lambda url, params: fake_response({"data": [], "error": ""}))
    comtrade.fetch(["2025"], session=session, freq="A", api_key="secret", hs_codes=["400400", "401220"], pause=0)
    url, params = session.calls[0]
    assert url == "https://comtradeapi.un.org/data/v1/get/C/A/HS"
    assert params["subscription-key"] == "secret"
    assert params["cmdCode"] == "400400,401220"


def test_fetch_skips_a_failed_period(fake_session, fake_response):
    def respond(url, params):
        if params["period"] == "202601":
            return fake_response(status=500)
        return fake_response({"data": [{"period": params["period"]}], "error": ""})

    rows = comtrade.fetch(["202601", "202602"], session=fake_session(respond), pause=0)
    assert rows == [{"period": "202602"}]


@pytest.mark.parametrize("response", [
    {"payload": {"data": [], "error": "Invalid reporterCode"}},
    {"content": b"<html>maintenance</html>"},
])
def test_fetch_raises_when_every_period_fails(fake_session, fake_response, response):
    session = fake_session(lambda url, params: fake_response(**response))
    with pytest.raises((comtrade.ComtradeError, ValueError)):
        comtrade.fetch(["202601", "202602"], session=session, pause=0)
