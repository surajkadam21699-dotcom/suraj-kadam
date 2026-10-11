# Control Transformer Manufacturing — Business Course

A practical, end-to-end course for starting and running a small-to-medium
business that designs, manufactures, tests and sells **control transformers**
(typically 25 VA – 5 kVA, single-phase, ≤ 1000 V) used in electrical control
panels, machine tools, MCCs, HVAC, lifts, pumps and automation systems.

The course combines three tracks:

| Track | What you learn |
|---|---|
| **Technical** | How control transformers work, how to size and design them, materials, manufacturing process, testing and standards |
| **Operations** | Factory layout, machinery, raw-material sourcing, quality system, production planning |
| **Business** | Market, product range, costing & pricing, registrations, sales channels, finance, scaling |

## Who this is for

- Electrical / electronics diploma or degree holders planning a manufacturing unit
- Existing winding / rewinding workshops moving into branded products
- Panel builders who want to make transformers in-house
- Investors evaluating a small electrical-components factory

No prior manufacturing experience is assumed. Basic AC circuit knowledge
(voltage, current, power, power factor) helps.

## Syllabus

| # | Module | Outcome |
|---|---|---|
| 01 | [Fundamentals of Control Transformers](modules/01-fundamentals.md) | Explain what a control transformer is and why panels need one |
| 02 | [Sizing & Selection](modules/02-sizing-and-selection.md) | Size a transformer from a load list (sealed + inrush VA) |
| 03 | [Electrical Design](modules/03-electrical-design.md) | Calculate core, turns, wire gauge and check window fill |
| 04 | [Materials & Bill of Materials](modules/04-materials-and-bom.md) | Specify laminations, copper, insulation, varnish, hardware |
| 05 | [Manufacturing Process](modules/05-manufacturing-process.md) | Run the shop-floor flow from winding to packing |
| 06 | [Testing & Quality Control](modules/06-testing-and-quality.md) | Perform routine and type tests; build a QC system |
| 07 | [Standards, Certification & Compliance](modules/07-standards-and-compliance.md) | Know which standards and registrations apply |
| 08 | [Factory Setup: Machinery, Layout & Investment](modules/08-factory-setup.md) | Plan a unit at three investment levels |
| 09 | [Costing & Pricing](modules/09-costing-and-pricing.md) | Build a unit cost sheet and set profitable prices |
| 10 | [Market, Sales & Marketing](modules/10-market-sales-marketing.md) | Find customers, build channels, win orders |
| 11 | [Finance, Legal & Business Plan](modules/11-finance-and-business-plan.md) | Prepare a bankable project report |
| 12 | [Operations, Scaling & Growth](modules/12-operations-and-scaling.md) | Run day-to-day and grow product lines |
| — | [Launch Roadmap (90 days)](modules/13-launch-roadmap.md) | Step-by-step action plan |

## Tools & templates

| File | Purpose |
|---|---|
| [`tools/ct_calc.py`](tools/ct_calc.py) | Command-line calculator: sizing from a load list, core/turns/wire design, unit costing |
| [`templates/load-list.csv`](templates/load-list.csv) | Sample control-circuit load list for sizing |
| [`templates/cost-sheet.csv`](templates/cost-sheet.csv) | Unit cost sheet template |
| [`templates/test-report.md`](templates/test-report.md) | Routine test report format |
| [`templates/business-plan-outline.md`](templates/business-plan-outline.md) | Project report / business plan skeleton |

Quick start:

```bash
cd control-transformer-course
python tools/ct_calc.py size templates/load-list.csv
python tools/ct_calc.py design --va 500 --vp 415 --vs 110
python tools/ct_calc.py cost templates/cost-sheet.csv --margin 25
python -m pytest -q tests
```

## How to use this course

1. Read modules 01–03 first — every business decision later depends on
   understanding the product.
2. Work every **Exercise** at the end of each module; answers to numeric
   exercises can be checked with `ct_calc.py`.
3. Treat all money figures as **illustrative**. Material prices, machine
   prices, tax rates and regulations change — always re-quote locally before
   committing money.
4. Use module 13 as your execution checklist.

> **Safety note:** Transformers involve mains voltages and high-voltage
> testing (2–4 kV). Only trained people should operate HV test equipment, and
> test areas must be interlocked and fenced.
