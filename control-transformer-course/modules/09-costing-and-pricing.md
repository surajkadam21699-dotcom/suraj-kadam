# Module 09 — Costing & Pricing

## 9.1 Unit cost structure

```
Unit cost = Direct materials (+ scrap allowance)
          + Direct labour
          + Variable overheads (power, consumables)
          + Fixed overhead share (rent, salaries, depreciation, interest)
          + Warranty provision
```

Typical shares for small control transformers (very indicative):

| Head | Share of ex-works cost |
|---|---|
| Copper | 25–40 % |
| Laminations | 15–25 % |
| Other materials (bobbin, insulation, hardware, packing) | 10–20 % |
| Labour | 8–15 % |
| Overheads | 10–20 % |

## 9.2 Building the cost sheet

Use [`templates/cost-sheet.csv`](../templates/cost-sheet.csv). The rates in
the template are **illustrative placeholders for a 100 VA unit** — replace
them with your own supplier quotations and time-studied labour rates.

```bash
python tools/ct_calc.py cost templates/cost-sheet.csv --scrap 3 --margin 25
```

Steps:

1. **Material quantities** from the design (copper and core mass from
   `ct_calc.py design`, other items from BOM).
2. **Material rates** from current purchase prices (landed: incl. freight,
   non-recoverable taxes).
3. **Scrap allowance** — 2–5 % on materials.
4. **Labour** — time-study each operation (hours × fully-loaded wage rate).
5. **Overheads** — monthly fixed costs ÷ monthly production (in
   standard-hours or units).
6. **Warranty** — % of sales based on failure history (start with ~1 %).

## 9.3 Pricing

### Cost-plus

```
Price = Cost ÷ (1 − target gross margin)
```

A 25 % gross margin on a cost of 100 gives price 133.3 (a 33 % mark-up).
Do not confuse **margin** and **mark-up**.

### Market-based

- Collect competitors' price lists and the **actual net prices** after
  discounts (ask panel builders and distributors).
- Position: *economy* (price leader), *value* (good quality, fair price —
  recommended for new entrants), or *premium* (certified, branded, fast
  delivery).

### Price list and discount structure

- Publish a **list price** per model.
- Distributor discount and OEM/contract pricing off the list.
- Quantity breaks (e.g. 1–9, 10–49, 50+).
- **Copper price-variation clause** for long contracts — e.g. price adjusts
  when the copper reference price moves beyond ±5 % from the base.

## 9.4 Break-even

```
Contribution per unit = Price − Variable cost per unit
Break-even units/month = Monthly fixed costs ÷ Contribution per unit
```

Example: price 1,475; variable cost 1,000; fixed costs 2,50,000 / month →
contribution 475 → break-even ≈ 527 units / month (≈ 21 per working day).

## 9.5 Profit levers (in order of impact)

1. **Copper & steel optimisation** — right-sized designs; don't over-design.
2. **Purchasing** — volume contracts, two suppliers, timing of copper buys.
3. **First-pass yield** — every reworked unit costs labour twice.
4. **Product mix** — custom & multi-tap units carry higher margins.
5. **Overhead absorption** — higher volume spreads fixed costs.
6. **Payment terms** — faster collection reduces interest cost.

## Exercises

1. Complete a cost sheet for 50, 100, 250 and 500 VA using real quotations.
2. Calculate break-even for your Level A plan.
3. Build your first price list with distributor and OEM discount levels.
