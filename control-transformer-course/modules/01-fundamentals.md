# Module 01 — Fundamentals of Control Transformers

## 1.1 What is a control transformer?

A **control transformer** (also called a *machine tool transformer* or
*industrial control transformer*) is a small, isolating, step-down
transformer that supplies the **control circuit** of electrical equipment —
contactor and relay coils, solenoid valves, pilot lamps, timers, PLC power
supplies, push-button stations and similar devices.

Typical example: a motor starter panel fed at 415 V three-phase uses a
415 V → 110 V control transformer so that operators touch only 110 V
push-buttons, and so that the control circuit is isolated from the power
circuit.

## 1.2 Why panels use control transformers

| Reason | Explanation |
|---|---|
| **Safety** | Lower voltage (24 V / 48 V / 110 V) at operator controls reduces shock risk |
| **Isolation** | Separate windings isolate the control circuit from the supply; one side can be earthed deliberately so an earth fault trips a fuse instead of energising a coil |
| **Standard coil voltage** | Panels sold to many countries can use one coil voltage (e.g. 110 V or 24 V) whatever the supply |
| **Inrush capability** | Designed to hold voltage up when several contactor coils pull in together |
| **Code compliance** | Machinery standards (e.g. IEC 60204-1) require a transformer for control circuits above a certain size |

## 1.3 How it differs from other small transformers

| Feature | Control transformer | General-purpose / power transformer | SMPS / electronic supply |
|---|---|---|---|
| Key rating | **Sealed VA + inrush VA** | Continuous kVA | Continuous W |
| Regulation | Designed for good voltage at **inrush (≈ 40 % PF)** | Designed for rated load | Electronically regulated |
| Construction | Open-frame, EI laminations, terminal block on top | Often enclosed | PCB |
| Primary | Often multi-tap (e.g. 0-380-415-440 V or 220-240 V) | Fixed | Wide range |
| Secondary | 24 / 48 / 110 / 230 V, often split or multi-tap | Fixed | DC |
| Overload behaviour | Robust, tolerant of short inrush overloads | Thermal limited | Current-limited / foldback |

## 1.4 Working principle (in one page)

1. AC voltage on the primary winding drives an alternating flux in the iron
   core.
2. The flux links the secondary winding and induces a voltage in it.
3. The voltage ratio equals the turns ratio: `Vp / Vs ≈ Np / Ns`.
4. The induced voltage is set by the **EMF equation**:

   ```
   E = 4.44 × f × N × Bmax × A
   ```

   - `E` = RMS voltage (V), `f` = frequency (Hz), `N` = turns,
     `Bmax` = peak flux density (tesla), `A` = net core cross-section (m²)

5. Losses: **iron (core) loss** — constant whenever energised; **copper
   loss** — rises with the square of load current.
6. **Regulation** — secondary voltage falls under load because of winding
   resistance and leakage reactance. For control transformers regulation at
   high inrush current is the critical figure.

## 1.5 Typical product specification

| Parameter | Common values |
|---|---|
| Rating | 25, 50, 75, 100, 150, 250, 350, 500, 750, 1000, 1500, 2000, 3000, 5000 VA |
| Primary | 230 V; 415 V; multi-tap 0-220-240 V; 0-380-415-440 V; 0-230-415 V |
| Secondary | 24 V, 48 V, 110 V, 230 V; centre-tapped 55-0-55 V; dual 0-24 + 0-110 V |
| Frequency | 50 Hz (India/Europe/Asia), 50/60 Hz dual-rated for export |
| Insulation class | B (130 °C), F (155 °C), H (180 °C) |
| Construction | Open-frame, EI laminations, varnish-impregnated, mounting feet, terminal block |
| Protection | Optional primary/secondary fuse holders, electrostatic screen, IP-rated cover |
| Ambient | 40 °C (standard) or 50 °C (Indian/Middle-East market preference) |

## 1.6 Where they are used (your customers)

- Motor control centres (MCC) and power control centres (PCC)
- Machine tools — lathes, CNC, presses, injection moulding machines
- HVAC, chillers, air handling units
- Lifts / elevators and escalators
- Pump panels, water treatment plants, irrigation
- Cranes and hoists
- Textile, packaging and food-processing machinery
- Railway signalling, solar and EV-charger auxiliary panels

## Key terms

- **VA** — volt-amperes, apparent power rating.
- **Sealed VA** — continuous VA drawn by a coil once the armature has closed.
- **Inrush VA** — momentary VA drawn when a coil is first energised (typically
  5–10 × sealed VA, at low power factor).
- **Regulation** — % drop in secondary voltage from no-load to load.
- **Isolation transformer** — separate, insulated primary and secondary
  windings (unlike an autotransformer).

## Exercises

1. List five devices in a typical star-delta starter panel that would be fed
   from the control transformer.
2. Why is a contactor coil's inrush more important than its sealed VA when
   sizing a control transformer?
3. A transformer has 830 primary turns and 220 secondary turns. With 415 V on
   the primary, what is the no-load secondary voltage? *(≈ 110 V)*
