import logging

import pytest

from tyre_prices import shipments


def parse_fixture(fixtures_dir, name, **options):
    return shipments.parse_page((fixtures_dir / name).read_bytes(), source=name, **options)


@pytest.mark.parametrize("header, field", [
    ("Date", "date"),
    ("HS Code", "hs_code"),
    ("Product Name", "description"),
    ("Country of Origin", "origin"),
    ("Supplier Country", "origin"),
    ("Port of Discharge", "port"),
    ("Indian Port", "port"),
    ("Port of Loading", None),
    ("Importer Name", None),
    ("Exporter", None),
    ("S.No.", None),
    ("Unit", "unit"),
    ("Qty Unit", "unit"),
    ("Quantity (Units)", "quantity"),
    ("Total Quantity", "quantity"),
    ("Value (INR)", "value"),
    ("Value In ₹", "value"),
    ("Total Amount (USD)", "value"),
    ("Per Unit (INR)", "unit_price"),
    ("Unit Price (USD)", "unit_price"),
])
def test_header_names_map_to_fields(header, field):
    assert shipments._classify_header(header) == field


def test_zauba_style_page(fixtures_dir):
    records = parse_fixture(fixtures_dir, "shipments_zauba_style.html", inr_per_usd=90)
    assert [record.description for record in records] == [
        "RUBBER TYRE SCRAP CUT IN THREE PIECES",
        "SHREDDED TYRE SCRAP (TDF)",
        "USED TYRES IN BALES",
        "USED TRUCK TYRES",
    ]
    cut = records[0]
    assert (cut.period, cut.hs_code, cut.origin, cut.port, cut.unit, cut.grade) == (
        "2026-01-05", "40040000", "United Kingdom", "Mundra", "KGS", "cut",
    )
    assert (cut.quantity, cut.value, cut.currency) == (108000, 2160000, "INR")
    assert cut.price_per_tonne == 20000
    assert cut.price_per_tonne_usd == pytest.approx(222.22, abs=0.01)
    assert (records[1].grade, records[1].price_per_tonne) == ("shredded", 21000)
    assert (records[2].grade, records[2].price_per_tonne) == ("baled", 21000)
    trucks = records[3]
    assert (trucks.unit, trucks.unit_price, trucks.price_per_tonne) == ("NOS", 1500, None)


def test_all_rows_keeps_other_goods_but_not_masked_rows(fixtures_dir):
    descriptions = [r.description for r in parse_fixture(fixtures_dir, "shipments_zauba_style.html", scrap_only=False)]
    assert "RUBBER SCRAP (EPDM PROFILES)" in descriptions
    assert "CARBON BLACK USED IN TYRE MANUFACTURING" in descriptions
    assert "TYRE SCRAP" not in descriptions  # value and price hidden from non-subscribers
    assert len(descriptions) == 6


def test_usd_page_with_header_in_td_cells(fixtures_dir):
    records = parse_fixture(fixtures_dir, "shipments_usd_layout.html")
    assert [record.period for record in records] == ["2026-04-10", "2026-04-12", "2026-04-15"]
    assert {record.currency for record in records} == {"USD"}
    cut, crumb, whole = records
    assert (cut.hs_code, cut.grade, cut.tonnes) == ("40040000", "cut", 52.5)
    assert cut.price_per_tonne == cut.price_per_tonne_usd == 225
    assert crumb.grade == "crumb"
    assert crumb.price_per_tonne == pytest.approx(320)
    assert (whole.grade, whole.unit) == ("whole", "LBS")
    assert whole.tonnes == pytest.approx(45.359237)
    assert whole.unit_price == pytest.approx(0.0907185)
    assert whole.price_per_tonne == pytest.approx(200, abs=0.01)


def test_unit_price_in_another_currency_is_ignored():
    html = """<table>
      <tr><th>Description</th><th>Unit</th><th>Quantity</th><th>Value (USD)</th><th>Per Unit (INR)</th></tr>
      <tr><td>TYRE SCRAP</td><td>KGS</td><td>1000</td><td>250</td><td>21</td></tr>
    </table>"""
    [record] = shipments.parse_page(html, source="t")
    assert record.currency == "USD"
    assert record.unit_price == 0.25
    assert record.price_per_tonne == 250


def test_unit_written_inside_the_quantity_cell():
    html = """<table>
      <tr><th>Date</th><th>Product Description</th><th>Quantity</th><th>Value (INR)</th></tr>
      <tr><td>05-Jan-2026</td><td>TYRE SCRAP 3 CUT</td><td>1,08,000 KGS</td><td>21,60,000</td></tr>
    </table>"""
    [record] = shipments.parse_page(html, source="t")
    assert (record.quantity, record.unit, record.tonnes) == (108000, "KGS", 108)
    assert record.price_per_tonne == 20000


def test_default_currency_applies_when_headers_name_none():
    html = """<table>
      <tr><th>Description</th><th>Quantity</th><th>Unit</th><th>Value</th></tr>
      <tr><td>TYRE SCRAP</td><td>2</td><td>MT</td><td>500</td></tr>
    </table>"""
    assert shipments.parse_page(html, source="t")[0].currency == "INR"
    assert shipments.parse_page(html, source="t", default_currency="USD")[0].price_per_tonne_usd == 250


def test_page_without_a_shipment_table_is_reported(caplog):
    with caplog.at_level(logging.WARNING):
        assert shipments.parse_page("<html><body><p>Checking your browser</p></body></html>", source="blocked") == []
    assert "no shipment table" in caplog.text


def test_expand_urls():
    urls = shipments.expand_urls(["https://a.example/p-{page}.html", "https://b.example/"], 3)
    assert list(urls) == [
        "https://a.example/p-1.html", "https://a.example/p-2.html", "https://a.example/p-3.html", "https://b.example/",
    ]


def test_scrape_drops_shipments_repeated_on_later_pages(fixtures_dir, fake_session, fake_response):
    html = (fixtures_dir / "shipments_zauba_style.html").read_bytes()

    def respond(url, params):
        return fake_response(status=403) if "blocked" in url else fake_response(content=html)

    session = fake_session(respond)
    records, failures = shipments.scrape(
        ["https://www.zauba.com/a", "https://www.zauba.com/b", "https://blocked.example/c"], session=session, pause=0,
    )
    assert failures == ["https://blocked.example/c"]
    assert len(records) == 4
    assert {record.source for record in records} == {"zauba.com"}


def test_scrape_keeps_identical_shipments_within_a_page(tmp_path, fake_session):
    row = "<tr><td>TYRE SCRAP</td><td>KGS</td><td>1000</td><td>250</td></tr>"
    page = tmp_path / "saved.html"
    page.write_text(f"<table><tr><th>Description</th><th>Unit</th><th>Quantity</th><th>Value</th></tr>{row}{row}</table>")
    records, failures = shipments.scrape(files=[page], session=fake_session(None), pause=0)
    assert failures == []
    assert [record.source for record in records] == ["saved.html", "saved.html"]
