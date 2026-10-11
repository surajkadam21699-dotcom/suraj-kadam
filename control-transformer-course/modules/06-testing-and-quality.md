# Module 06 — Testing & Quality Control

Your reputation is built on **zero field failures**. Every transformer must
pass routine tests before it leaves the factory, and every new design must
pass type tests before you sell it.

## 6.1 Routine tests (100 % of production)

| Test | Purpose | Typical method / instrument |
|---|---|---|
| Visual & dimensional | Workmanship, labels, terminals | Checklist |
| Winding resistance | Wire size / joints correct | Micro-ohmmeter or bridge |
| Voltage ratio & polarity | Correct turns at every tap | Ratio meter, or apply rated primary and measure secondary |
| No-load current & loss | Core quality and stacking | Wattmeter / power analyser at rated voltage |
| Dielectric (HV) withstand | Insulation integrity: pri-sec, pri-core, sec-core | HV tester, AC test voltage per standard for 60 s (production may allow shorter time at higher voltage — check standard) |
| Insulation resistance | Dry, clean insulation | Megger, 500 V / 1000 V DC |
| Induced over-voltage (sample or 100 %) | Inter-turn insulation | Apply higher voltage at higher frequency, or surge / impulse winding tester |
| Earth continuity | Frame bonded to earth terminal | Low-resistance ohmmeter |

Record results against the serial number (see
[`templates/test-report.md`](../templates/test-report.md)).

## 6.2 Type tests (each new design, and after design changes)

| Test | What it proves |
|---|---|
| Temperature rise at rated load | Winding temperature within insulation-class limits (resistance method) |
| Regulation & efficiency | Published performance figures |
| **Inrush capability (40 % PF, 90 % voltage)** | Published max inrush VA (Module 02) |
| Short-circuit / overload behaviour | Safe under fault with recommended fuses |
| Dielectric at full type-test level | Insulation design margins |
| Creepage & clearance measurement | Compliance with standard |
| Mechanical / terminal tests | Terminal strength, marking durability |
| Noise (optional) | Hum level |

For certification (CE, UL, etc.) type tests are normally witnessed or
performed by an accredited laboratory.

## 6.3 Test bench (minimum)

- Variac (auto-transformer) for each power range
- True-RMS multimeters, clamp meter
- Power analyser / wattmeter (V, A, W, PF)
- HV tester (AC, up to 5 kV, with trip current setting) — **interlocked**
- Insulation resistance tester
- Winding resistance meter
- Turns-ratio meter or calibrated reference transformer
- Load bank: resistive + inductive (to set 40 % PF for inrush tests)
- Temperature measurement: thermocouples / data logger
- Surge (impulse) winding tester — highly recommended for inter-turn faults

Calibrate instruments yearly at an accredited lab and keep certificates.

## 6.4 Quality management system

Build a simple, **ISO 9001-ready** system from day one:

1. **Quality policy & objectives** — e.g. "field failure rate < 0.1 %".
2. **Control plan** for each product: characteristic → spec → method →
   frequency → reaction plan.
3. **Incoming inspection** procedure and records.
4. **In-process checks** (turns, ratio, stacking weight).
5. **Final test** 100 % with traceable records.
6. **Non-conformance** handling: tag, segregate, root cause, corrective
   action (8D / 5-Why).
7. **Calibration** register.
8. **Customer complaints** log with analysis.
9. **Internal audits** and management review.

## 6.5 Traceability

Serial number on every unit → links to: date, operator, wire lot, lamination
lot, varnish batch, test results. When a complaint arrives you can find the
cause quickly and contain affected batches.

## 6.6 Warranty and failure analysis

- Typical warranty in this industry: 12–24 months.
- Collect every failed unit; dismantle and classify the failure (insulation,
  overload, wrong connection, mechanical). Customer misuse and genuine defects
  both teach you something.

## Exercises

1. Write a control plan for the 250 VA model.
2. Design your test-report format and serial-number scheme.
3. Plan the layout of an interlocked HV test booth.
