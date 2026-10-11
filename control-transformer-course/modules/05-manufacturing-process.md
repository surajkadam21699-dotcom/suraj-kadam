# Module 05 — Manufacturing Process

## 5.1 Process flow

```
Incoming inspection
      │
      ▼
Bobbin / former preparation ──► Coil winding (primary → insulation → secondary)
                                         │
                                         ▼
                                 Lead termination & sleeving
                                         │
                                         ▼
                                 In-process test (resistance, ratio)
                                         │
                                         ▼
                Core stacking (interleaved E & I) ──► Clamping / bracket fitting
                                         │
                                         ▼
                           Pre-heat ► Varnish dip / VPI ► Drain ► Bake
                                         │
                                         ▼
                      Terminal block fitting, wiring, tap marking
                                         │
                                         ▼
                           Final routine tests (100 %)
                                         │
                                         ▼
                      Labelling ► Packing ► Finished-goods store
```

## 5.2 Operation details

### Winding
- Machine: semi-automatic **linear coil-winding machine** with turns counter,
  wire tensioner and traverse (layer) guide; CNC multi-spindle machines for
  volume.
- Wind **primary first** (usually higher voltage, smaller wire, closer to
  core) — or follow the design sheet.
- Layer-wind neatly; insert layer insulation as specified; keep margins for
  creepage.
- Check **turns count** at each tap; mark tap leads.
- Wire tension: enough to lay flat, not enough to stretch (stretching thins the
  wire and damages enamel).

### Lead termination
- Strip enamel (mechanical stripper, solder pot or thermal stripper).
- Crimp or solder to flexible leads / pins; sleeve every joint.
- Anchor leads with tape so they cannot pull on fine wire.

### Core stacking
- **Interleave** E and I laminations (alternate sides each 1–2 sheets) to
  minimise air gap and no-load current.
- Stack to the specified height/weight — weigh the stack for consistency.
- Press tight; fit clamps / brackets; burr-free, square assembly.

### Impregnation
- Pre-heat coil assembly (e.g. 100–120 °C) to drive out moisture.
- Dip in varnish (or VPI) until bubbling stops; drain.
- Bake per varnish datasheet (typical 2–6 h at 120–150 °C).
- Result: rigid, quiet, moisture-resistant unit.

### Final assembly
- Fit terminal block, connect leads per wiring diagram, torque screws.
- Fit earth terminal and bond to core/frame.
- Apply rating label with serial number.

## 5.3 Work instructions

Write a one-page **work instruction (WI)** with photos for each operation
and each model family: tools, settings, steps, checkpoints, common defects.
Display them at the workstation. This lets you train new operators quickly and
is required for ISO 9001.

## 5.4 Common defects and prevention

| Defect | Cause | Prevention |
|---|---|---|
| Wrong ratio | Turns miscount, wrong tap | Counter verification, in-process ratio test |
| Inter-turn short | Enamel damage, tension too high | Wire guides, tension control, surge test |
| HV breakdown | Insufficient margin, damaged insulation | WI with margin dimensions, HV test 100 % |
| High no-load current / hum | Poor interleaving, air gap, wrong steel | Stacking WI, stack-weight check, steel certificate |
| Loose laminations | Insufficient clamping / varnish | Clamp torque, proper impregnation |
| Overheating | Under-sized wire, poor varnish | Design review, temperature-rise sample tests |
| Terminal loosening | Under-torqued screws | Torque screwdriver, audit |

## 5.5 Productivity benchmarks (to measure yourself against)

Track: units per operator per shift, first-pass yield at final test (target
≥ 98 %), rework %, scrap %, on-time delivery. Time-study each operation and use
the figures in your cost sheet (Module 09).

## 5.6 Health, safety and environment

- Varnish area: ventilation / fume extraction, fire extinguishers, no naked
  flames, solvent storage rules.
- Ovens: temperature controllers with over-temperature cut-off.
- HV test area: interlocked enclosure, dead-man switch, warning lamp, rubber
  mat, trained operators only.
- PPE: gloves, safety glasses, safety shoes; first-aid kit.
- Copper and lamination scrap: segregate and sell to recyclers — it is
  valuable.

## Exercises

1. Draw the process flow for your planned product range and mark inspection
   points.
2. Write a work instruction for winding the 100 VA model.
3. List five KPIs you will display on the shop-floor board.
