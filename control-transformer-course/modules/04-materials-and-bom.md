# Module 04 — Materials & Bill of Materials

Materials are typically **60–75 % of the ex-factory cost** of a control
transformer, and copper alone is often the single biggest item. Smart
purchasing is therefore one of the biggest profit levers you have.

## 4.1 Core steel (laminations)

| Material | Notes | Typical use |
|---|---|---|
| **CRNO** (cold-rolled non-oriented), 0.50 / 0.65 mm | Isotropic, cheaper, good for EI stamping | Most small control transformers |
| **CRGO** (cold-rolled grain-oriented), 0.23–0.35 mm | Lower loss in rolling direction, higher B | Better efficiency / smaller size; must be genuine prime grade |
| Hot-rolled silicon steel | Obsolete, high loss | Avoid |

Buying options:

1. **Ready-stamped EI laminations** from a lamination maker — lowest
   investment; recommended to start.
2. **Coil / sheet + own press tool** — needs a power press and dies; worth it
   only at volume.

Quality checks on receipt: grade certificate (mill test certificate), thickness,
insulation coating present on both sides, burr height (low burr → lower loss),
dimensional fit of E and I, specific loss test on sample (Epstein or
single-sheet tester at your supplier).

> Beware of "secondary" or "defective" CRGO sold as prime — it raises no-load
> loss and noise. Buy from traceable sources with test certificates.

## 4.2 Winding wire

| Item | Specification |
|---|---|
| Material | Electrolytic copper (aluminium only for cost-sensitive larger units) |
| Insulation | Polyester-imide / dual-coat enamel, thermal class 155–200 °C |
| Grade | Grade 1 (thin) / Grade 2 (medium) — grade 2 preferred for HV windings |
| Standard | IEC 60317 series / national equivalent (e.g. IS 13730 in India) |
| Larger currents | Paper- or Nomex-covered rectangular strip, or parallel round wires |

Receiving checks: diameter (micrometer), enamel continuity / pinholes,
elongation, solderability, spool label vs order.

## 4.3 Insulation materials

| Material | Thermal class | Use |
|---|---|---|
| Kraft / presspahn paper | A (105 °C) | Economy layer insulation |
| Polyester film (Mylar-type) | B (130 °C) | Layer and inter-winding insulation |
| DMD / NMN laminates | F / H (155 / 180 °C) | Inter-winding and slot liners |
| Aramid paper (Nomex-type) | H / C | High-temperature designs |
| Polyester / glass tape | B–H | Outer wrap and margin tapes |
| Fibreglass / silicone sleeving | F–H | Lead insulation |

**Insulation system rule:** every material in the winding must be rated for
the class you declare on the nameplate. A class F transformer with class A
paper inside is really class A.

## 4.4 Impregnating varnish

- Polyester or alkyd **air-drying / baking varnish** (class F/H)
- Solvent-less resins for VPI (vacuum-pressure impregnation) in larger units
- Purpose: bonds turns (reduces hum), improves heat transfer, keeps out
  moisture and dust.
- Check: viscosity (flow cup), solid content, shelf life; store cool.

## 4.5 Mechanical parts

- Bobbin / coil former (PBT, nylon-GF, phenolic) or fabricated former
- Mounting clamps / brackets — powder-coated or zinc-plated steel
- Through-bolts or welded / riveted core clamping
- Terminal block (barrier type or DIN-style, finger-safe) with tap markings
- Earth terminal
- Fuse holders (optional)
- Rating plate / label (durable, legible: rating, voltages, frequency, class,
  ambient, serial number, standard, manufacturer)
- Packing: corrugated box, foam inserts, polybag, instruction leaflet

## 4.6 Sample BOM — 100 VA, 415 / 110 V

| # | Item | Qty | Unit |
|---|---|---|---|
| 1 | EI lamination (tongue ≈ 30 mm), CRNO 0.5 mm | 1.6 | kg |
| 2 | Enamelled Cu wire 0.355 mm (primary) | ~0.18 | kg |
| 3 | Enamelled Cu wire 0.63 mm (secondary) | ~0.20 | kg |
| 4 | Bobbin to suit lamination and stack | 1 | pc |
| 5 | Polyester film 0.05 mm | 0.05 | m² |
| 6 | Polyester tape 19 mm | 1.5 | m |
| 7 | Sleeving | 0.5 | m |
| 8 | Varnish | 0.08 | kg |
| 9 | Mounting bracket set | 1 | set |
| 10 | Terminal block 6-way + earth | 1 | pc |
| 11 | Rating label | 1 | pc |
| 12 | Carton + inserts | 1 | set |

Use `tools/ct_calc.py design` to estimate copper and core mass per model and
build BOMs for your whole range.

## 4.7 Purchasing strategy

- Identify **two approved suppliers** for every critical item (laminations,
  wire, bobbins).
- Link copper purchases to a published copper reference price so you
  understand price movements; consider **price-variation clauses** in large
  customer contracts.
- Buy wire in the few diameters that cover most of your range — standardise
  designs to reduce inventory.
- Keep 4–6 weeks of stock of A-class items; order C-class items quarterly.
- Maintain an **Approved Vendor List** and incoming-inspection records
  (required for ISO 9001).

## Exercises

1. Prepare BOMs for 50, 250 and 1000 VA models.
2. Call three lamination suppliers and three wire suppliers; record price,
   minimum order, lead time and certificates offered.
3. Explain why class F insulation can let you build a smaller transformer for
   the same VA — and what the trade-off is.
