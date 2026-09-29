# Scrap tyre import prices

Scrapes what India pays for imported scrap tyres (HS 4004, "waste, parings and scrap of rubber") and summarises
it as a price per tonne: by month, by origin country and by grade.

| Command | Source | What you get |
| --- | --- | --- |
| `comtrade` | [UN Comtrade API](https://comtradeplus.un.org) (official customs statistics) | Monthly CIF value and net weight of India's imports, as USD per tonne: overall by month, and by origin country for the latest month and for the whole period. Complete and official, but published a few months late and not split by grade. |
| `shipments` | Trade-data pages such as [Zauba](https://www.zauba.com) and [Seair](https://www.seair.co.in) | Individual shipments (date, description, origin, port, quantity, value), as price per tonne by grade: cut, baled, shredded, crumb, whole. Recent, but limited to the sample rows the sites show publicly. |

## Run it on GitHub

Open **Actions → Scrap tyre import prices → Run workflow**. The price tables appear on the run's summary page, and
the CSV/JSON files are attached as the `tyre-import-prices` artifact. The workflow also runs by itself on the 5th of
every month.

Without a key, UN Comtrade's public preview API is used (at most 500 rows per request, plenty for one HS code). To use
the full API, get a free key at [comtradedeveloper.un.org](https://comtradedeveloper.un.org) and add it as the
repository secret `COMTRADE_API_KEY`.

## Run it locally

```bash
pip install -r requirements.txt

python -m tyre_prices comtrade                                # last 24 months, HS 400400
python -m tyre_prices comtrade --start 2025-01 --end 2025-12
python -m tyre_prices comtrade --annual --start 2021 --end 2025
python -m tyre_prices comtrade --hs 400400 --hs 401220        # add used pneumatic tyres

python -m tyre_prices shipments --inr-per-usd 88.5            # built-in scrap tyre pages, prices also in USD
python -m tyre_prices shipments --url "https://example.com/scrap-tyre/p-{page}.html" --pages 5
python -m tyre_prices shipments --file saved-page.html        # a page saved from your browser
```

Each command prints its summary tables and writes `output/<command>_records.csv` (one row per shipment or monthly
figure) and `output/<command>_summary.json`. `--out DIR` changes the directory; `--summary-md PATH` also appends the
tables as Markdown.

CSV columns: `source, period, origin, port, hs_code, grade, description, quantity, unit, tonnes, value, currency,
unit_price, price_per_tonne, price_per_tonne_usd`.

## Things to know

- **Comtrade prices are averages** (total CIF value ÷ net weight). HS 400400 is mostly tyre scrap for India but also
  covers other rubber scrap and crumb, so treat it as the market average rather than a quote for one grade.
- **Trade-data sites** show only a sample of shipments to non-subscribers, hide some figures (those rows are skipped)
  and may block automated requests. If a site blocks the scraper, open the page in your browser, save it (Ctrl+S) and
  pass it with `--file`. Check a site's terms of use before scraping it.
- Shipment tables are recognised by their column headers (Date, HS Code, Description, Origin Country, Port of
  Discharge, Unit, Quantity, Value (INR), Per Unit (INR) and common variants), so pages from other sites with a similar
  table work too. Tables that a page builds with JavaScript are not supported.
- Summary tables leave out figures for less than 1 tonne (a few kilograms of samples give absurd per-tonne prices);
  the CSV keeps them.
- Values whose column header names no currency are taken as INR (`--currency USD` changes that). INR records get USD
  prices only when you pass `--inr-per-usd`.
- A shipment is kept when its description mentions tyres and reads like scrap (scrap, waste, cut, bales, shredded,
  crumb, used, ...) or it is filed under HS 4004. `--all-rows` keeps every row.

## Development

```bash
pip install -r requirements-dev.txt
pytest
```

The tests run offline against synthetic pages and API responses in `tests/fixtures`.
