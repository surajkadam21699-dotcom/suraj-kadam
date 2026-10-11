# Module 02 — Sizing & Selection

Customers will ask you "which transformer do I need?". Answering this well
wins orders and avoids field failures. Every salesperson and application
engineer in your company should master this module.

## 2.1 Data you need from the customer

1. Supply (primary) voltage and frequency
2. Control (secondary) voltage
3. A **load list**: every coil, lamp, timer and device on the control circuit
   with its **sealed VA** and **inrush VA** (from the device datasheet)
4. Which devices can be energised **at the same time**
5. Ambient temperature and enclosure conditions
6. Any certification needed (CE, UL, etc.)

## 2.2 The four-step sizing method

This is the method used by most major control-transformer manufacturers.

### Step 1 — Total sealed VA

Add the sealed (holding) VA of **all** devices that can be on at the same
time, plus lamps, timers, relays and any other continuous load.

```
Sealed_total = Σ sealed VA of all simultaneously-energised devices
```

The transformer's **continuous VA rating must be ≥ Sealed_total**.

### Step 2 — Total inrush VA

Add the inrush VA of the coils that can be **energised at the same instant**
(usually the worst-case start sequence), plus the sealed VA of everything that
is already on at that moment.

```
Inrush_total = Σ inrush VA (coils closing together) + Σ sealed VA (devices already on)
```

### Step 3 — Required maximum inrush rating

Combine continuous and inrush loading:

```
Required inrush VA = √( Sealed_total² + Inrush_total² )
```

Choose a transformer whose **maximum inrush VA** (published at 40 % power
factor and **90 % secondary voltage**, i.e. ≤ 10 % voltage drop) is ≥ this
value. Most coils drop out or chatter if the voltage dips below ~85 %, so the
90 % criterion gives margin.

### Step 4 — Select the rating

Pick the smallest standard size that satisfies **both**:

- Continuous VA ≥ Sealed_total
- Max inrush VA (at 90 % voltage) ≥ Required inrush VA

## 2.3 Worked example

| Device | Qty | Sealed VA each | Inrush VA each |
|---|---|---|---|
| Main contactor (40 A) | 1 | 12 | 110 |
| Star contactor (25 A) | 1 | 8 | 70 |
| Delta contactor (40 A) | 1 | 12 | 110 |
| Timer | 1 | 3 | 3 |
| Pilot lamps | 3 | 2 | 2 |
| Auxiliary relays | 2 | 3 | 20 |

- Sealed_total = 12 + 8 + 12 + 3 + 6 + 6 = **47 VA**
- Worst instant: main + star close together; everything else is counted at
  its sealed VA (conservatively including the delta contactor) →
  Inrush_total = 110 + 70 + 12 + 3 + 6 + 6 = **207 VA**
- Required inrush = √(47² + 207²) ≈ **212 VA**

If your 50 VA model is published at 300 VA max inrush (at 90 % voltage), a
**50 VA** unit is technically enough; most panel builders would still choose
**75–100 VA** for spare capacity. If the 50 VA unit is only rated 180 VA
inrush, it is not suitable.

> Run this example: `python tools/ct_calc.py size templates/load-list.csv`

## 2.4 Your inrush rating table — a key product document

As a manufacturer you must **test and publish** each model's inrush
capability. Build a table like:

| Rating (VA) | Max inrush VA @ 40 % PF, 90 % Vs | Regulation at rated load |
|---|---|---|
| 50 | — measure — | — measure — |
| 100 | — measure — | — measure — |
| 250 | — measure — | — measure — |

How to measure: load the secondary with a resistive–inductive bank set to
40 % PF, increase load until secondary voltage falls to 90 % of rated, and
record VA. Repeat on three samples and publish the lowest value. Typical
ratios of inrush-VA to continuous-VA for small open-frame designs lie roughly
in the 3× to 10× range — larger units and units with better regulation are
at the high end.

## 2.5 Other selection checks

- **Primary taps**: offer taps so customers can match actual supply (e.g. 380,
  400, 415, 440 V) and keep secondary voltage correct.
- **Fusing**: recommend primary and secondary fuse ratings for each model
  (overcurrent protection is required by most installation standards).
- **Ambient derating**: if ambient > rated ambient, derate (rule of thumb:
  ~1 % per °C above rating — confirm by your own temperature-rise tests).
- **Frequency**: a 50 Hz design works on 60 Hz (lower flux), but a 60 Hz
  design on 50 Hz will over-flux. Design dual-rated products at 50 Hz.
- **Earthing of secondary**: advise which secondary terminal to earth.

## 2.6 Selling tip

Offer a free "**transformer selection service**": customers e-mail their load
list, you return the recommended model within hours. This builds trust with
panel builders and captures their repeat orders.

## Exercises

1. A control circuit has sealed load 120 VA and worst-case inrush of 650 VA.
   Calculate the required inrush VA. *(≈ 661 VA)*
2. Explain why selecting purely on sealed VA leads to contactor chatter.
3. Draft a one-page "How to select our control transformer" guide for your
   catalogue.
