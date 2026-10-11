# Module 03 — Electrical Design

This module gives you a **first-cut design method** for single-phase,
shell-type (EI lamination) control transformers from 25 VA to ~5 kVA. Every
design must then be proven with prototypes and temperature-rise tests
(Module 06).

## 3.1 Design inputs

| Input | Symbol | Typical value |
|---|---|---|
| Rating | VA | from Module 02 |
| Primary / secondary voltage | Vp, Vs | 415 / 110 V |
| Frequency | f | 50 Hz |
| Peak flux density | Bmax | CRNO 1.1–1.3 T; CRGO 1.4–1.6 T |
| Current density | J | 2.5–3.5 A/mm² (lower for larger units / class B) |
| Core constant | k | 1.0–1.2 |
| Stacking factor | Sf | 0.92–0.96 |
| Efficiency (assumed) | η | 0.85 (25 VA) … 0.97 (5 kVA) |

## 3.2 Step-by-step method

### Step 1 — Core area

```
A_core (cm²) ≈ k × √VA
```

### Step 2 — Lamination and stack

Standard **scrapless EI laminations** have centre-limb (tongue) width `T`,
window width `T/2` and window height `1.5 T`, so:

```
Window area = 0.75 × T²
Stack height S = A_core / (Sf × T)      (aim S ≈ 1.0 – 1.5 × T)
```

Choose the nearest lamination size your supplier stocks (e.g. tongue 19, 25,
32, 38, 44, 50, 64 mm — sizes vary by supplier).

### Step 3 — Turns per volt

```
Turns/volt = 1 / (4.44 × f × Bmax × A_net)      (A_net in m²)
```

### Step 4 — Turns

```
Np = Vp × Turns/volt
Ns = Vs × Turns/volt × (1 + regulation compensation)
```

Regulation compensation adds extra secondary turns so full-load voltage is
close to nameplate. Typical: 10–15 % at 25–50 VA, 5–8 % at 100–500 VA,
2–4 % above 1 kVA. Calibrate from your prototypes.

For **multi-tap primaries** (e.g. 0-380-415-440 V) compute turns at each tap
from the same turns/volt.

### Step 5 — Currents and wire

```
Is = VA / Vs
Ip = VA / (η × Vp)
Wire copper area = I / J      →  pick next standard diameter
```

### Step 6 — Window fill check

```
Fill = (Np × ap + Ns × as) / Window area      ≤ 0.35 – 0.40
```

If fill is too high: choose a bigger lamination, increase stack (fewer turns),
or (carefully) increase J. If fill is very low you are over-sized — try a
smaller lamination.

### Step 7 — Losses, temperature and regulation (verify)

- **Iron loss** = core mass × specific loss (W/kg at your B, from the steel
  datasheet).
- **Copper loss** = Ip²Rp + Is²Rs, R from wire length (turns × mean length
  of turn) and resistivity 0.0172 Ω·mm²/m at 20 °C (≈ 0.021 at 75 °C).
- **Regulation (resistive)** ≈ copper loss / VA.
- **Temperature rise** — confirm by test; for class B insulation keep average
  winding rise ≤ ~80 K at 40 °C ambient (check your standard's limit).

## 3.3 Worked example — 500 VA, 415 V / 110 V, 50 Hz

| Step | Result |
|---|---|
| Core area | 1.1 × √500 ≈ 24.6 cm² |
| Lamination | Tongue 46 mm, stack 57 mm → net area 24.9 cm² |
| Turns/volt | 1 / (4.44 × 50 × 1.3 × 0.00249) ≈ 1.39 |
| Primary turns | 415 × 1.39 ≈ 577 |
| Secondary turns | 110 × 1.39 × 1.05 ≈ 161 |
| Currents | Ip ≈ 1.27 A, Is ≈ 4.55 A |
| Wire (J = 3) | Primary 0.80 mm, secondary 1.40 mm |
| Window fill | ≈ 0.34 ✓ |

Reproduce: `python tools/ct_calc.py design --va 500 --vp 415 --vs 110`

## 3.4 Insulation design

| Location | Typical insulation |
|---|---|
| Core to winding | Moulded bobbin (nylon / PBT / phenolic) or presspahn / Nomex former |
| Between layers | Kraft / presspahn / polyester film, 0.05–0.13 mm |
| Primary to secondary | 2–3 layers of insulation + margin tapes (creepage per standard) |
| Outer wrap | Polyester tape / presspahn |
| Leads | Sleeving (fibreglass / PVC rated to insulation class) |

Creepage and clearance distances and the primary-to-secondary test voltage
come from your chosen standard (Module 07). Design the bobbin with **margins**
at both ends to achieve creepage, or use a **split (two-section) bobbin**
which gives excellent primary–secondary separation and is popular for
control transformers.

## 3.5 Optional features

- **Electrostatic (Faraday) screen** between windings, brought out to earth
  terminal — reduces noise for PLC / electronic loads.
- **Thermal cut-out** embedded in winding.
- **Fuse holders** mounted on the terminal block.
- **Finger-safe (IP20) terminals** and transparent covers.
- **Dual-voltage secondary** (e.g. 0-24 V + 0-110 V) for mixed panels.

## 3.6 Design documentation (per model)

Keep a **design sheet** for every model containing: lamination size, stack,
turns per tap, wire sizes, insulation schedule, winding sequence, bobbin
drawing, expected no-load current / losses, measured test results, and
revision history. This becomes your most valuable intellectual property.

## Exercises

1. Design a 250 VA, 230 V / 24 V transformer. What happens to secondary wire
   size and why? *(Is ≈ 10.4 A — consider parallel strands or rectangular
   conductor.)*
2. Recalculate the 500 VA example with Bmax = 1.5 T (CRGO). What changes and
   what is the risk? *(Fewer turns, smaller window, but higher no-load
   current and noise; risk of saturation at +10 % supply.)*
3. Write a design sheet for your first three models.
