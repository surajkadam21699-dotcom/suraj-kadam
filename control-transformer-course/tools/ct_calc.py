#!/usr/bin/env python3
"""Control transformer calculator for the course.

Sub-commands:
    size    Size a transformer from a load-list CSV (sealed + inrush VA method)
    design  Rough electrical design of a single-phase EI-core transformer
    cost    Unit cost and selling price from a cost-sheet CSV

Results are first-cut engineering estimates for learning and quoting.
Always verify a design by building and testing prototypes.
"""

import argparse
import csv
import math
import sys
from dataclasses import dataclass

# Standard control-transformer ratings (VA)
STANDARD_RATINGS_VA = [
    25, 50, 75, 100, 150, 200, 250, 350, 500, 750,
    1000, 1500, 2000, 3000, 5000,
]

# Nominal enamelled copper wire diameters, mm (IEC R20 series)
STANDARD_WIRE_MM = [
    0.100, 0.112, 0.125, 0.140, 0.160, 0.180, 0.200, 0.224, 0.250, 0.280,
    0.315, 0.355, 0.400, 0.450, 0.500, 0.560, 0.630, 0.710, 0.800, 0.900,
    1.000, 1.120, 1.250, 1.400, 1.600, 1.800, 2.000, 2.240, 2.500, 2.800,
    3.150,
]

COPPER_DENSITY_G_PER_CM3 = 8.89
STEEL_DENSITY_G_PER_CM3 = 7.65


# --------------------------------------------------------------------------
# Sizing
# --------------------------------------------------------------------------

@dataclass
class SizingResult:
    sealed_total_va: float
    inrush_total_va: float
    required_inrush_va: float
    min_continuous_rating_va: int | None


def size_from_loads(loads):
    """Size from a list of dicts with keys: qty, sealed_va, inrush_va, starts_together.

    Devices with starts_together true contribute inrush VA; all others
    contribute their sealed VA to the inrush moment (they are already on).
    """
    sealed_total = 0.0
    inrush_total = 0.0
    for load in loads:
        qty = float(load["qty"])
        sealed = float(load["sealed_va"]) * qty
        inrush = float(load["inrush_va"]) * qty
        sealed_total += sealed
        inrush_total += inrush if load["starts_together"] else sealed
    required = math.hypot(sealed_total, inrush_total)
    rating = next((r for r in STANDARD_RATINGS_VA if r >= sealed_total), None)
    return SizingResult(sealed_total, inrush_total, required, rating)


def read_load_list(path):
    loads = []
    with open(path, newline="", encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            loads.append({
                "device": row["device"],
                "qty": row["qty"],
                "sealed_va": row["sealed_va"],
                "inrush_va": row["inrush_va"],
                "starts_together": row["starts_together"].strip().lower()
                in ("1", "y", "yes", "true"),
            })
    return loads


# --------------------------------------------------------------------------
# Electrical design
# --------------------------------------------------------------------------

@dataclass
class Winding:
    volts: float
    amps: float
    turns: int
    wire_mm: float
    copper_area_mm2: float


@dataclass
class DesignResult:
    va: float
    core_area_cm2: float
    tongue_mm: float
    stack_mm: float
    turns_per_volt: float
    primary: Winding
    secondary: Winding
    window_area_mm2: float
    window_fill: float
    core_mass_kg: float
    copper_mass_kg: float


def pick_wire(required_area_mm2):
    """Smallest standard wire whose copper area >= required area."""
    for d in STANDARD_WIRE_MM:
        if math.pi * d * d / 4 >= required_area_mm2:
            return d
    raise ValueError(
        f"needs {required_area_mm2:.2f} mm2 copper; use strip or parallel wires"
    )


def design(va, vp, vs, freq=50.0, b_tesla=1.3, j_a_mm2=3.0, k_core=1.1,
           efficiency=0.95, reg_comp=0.05, stacking=0.95, tongue_mm=None):
    """First-cut design for a shell-type EI transformer with scrapless laminations.

    Scrapless EI geometry: window width = T/2, window height = 1.5 T,
    so window area = 0.75 T^2 (T = centre-limb / tongue width).
    """
    if min(va, vp, vs, freq, b_tesla, j_a_mm2) <= 0:
        raise ValueError("all inputs must be positive")

    core_area_cm2 = k_core * math.sqrt(va)
    if tongue_mm is None:
        # aim for stack ~= 1.25 x tongue, rounded to 2 mm steps
        tongue_mm = math.sqrt(core_area_cm2 * 100 / (stacking * 1.25))
        tongue_mm = max(10.0, round(tongue_mm / 2) * 2)
    stack_mm = core_area_cm2 * 100 / (stacking * tongue_mm)
    stack_mm = math.ceil(stack_mm)
    net_area_m2 = stacking * tongue_mm * stack_mm * 1e-6

    tpv = 1 / (4.44 * freq * b_tesla * net_area_m2)

    ip = va / (efficiency * vp)
    is_ = va / vs
    np_ = round(vp * tpv)
    ns = round(vs * tpv * (1 + reg_comp))
    dp = pick_wire(ip / j_a_mm2)
    ds = pick_wire(is_ / j_a_mm2)
    ap = math.pi * dp * dp / 4
    as_ = math.pi * ds * ds / 4

    window_area = 0.75 * tongue_mm ** 2
    fill = (np_ * ap + ns * as_) / window_area

    # Core mass: scrapless EI outline is 3T x 2.5T, minus two windows
    lam_area_mm2 = (3 * tongue_mm) * (2.5 * tongue_mm) - 2 * window_area
    core_mass = lam_area_mm2 * stack_mm * stacking * 1e-3 * STEEL_DENSITY_G_PER_CM3 / 1000

    # Copper mass from mean turn length around the bobbin (approximate)
    build = tongue_mm / 2
    mlt_mm = 2 * (tongue_mm + stack_mm) + math.pi * build
    cu_vol_cm3 = (np_ * ap + ns * as_) * mlt_mm / 1000
    copper_mass = cu_vol_cm3 * COPPER_DENSITY_G_PER_CM3 / 1000

    return DesignResult(
        va=va,
        core_area_cm2=stacking * tongue_mm * stack_mm / 100,
        tongue_mm=tongue_mm,
        stack_mm=stack_mm,
        turns_per_volt=tpv,
        primary=Winding(vp, ip, np_, dp, ap),
        secondary=Winding(vs, is_, ns, ds, as_),
        window_area_mm2=window_area,
        window_fill=fill,
        core_mass_kg=core_mass,
        copper_mass_kg=copper_mass,
    )


# --------------------------------------------------------------------------
# Costing
# --------------------------------------------------------------------------

@dataclass
class CostResult:
    by_category: dict
    scrap_allowance: float
    total_cost: float
    selling_price: float


def cost_from_rows(rows, scrap_pct=3.0, margin_pct=25.0):
    """Rows: dicts with category, quantity, rate.

    Scrap allowance is applied to the 'material' category.
    margin_pct is gross margin on selling price: price = cost / (1 - m).
    """
    if not 0 <= margin_pct < 100:
        raise ValueError("margin must be in [0, 100)")
    by_cat = {}
    for row in rows:
        cat = row["category"].strip().lower()
        by_cat[cat] = by_cat.get(cat, 0.0) + float(row["quantity"]) * float(row["rate"])
    scrap = by_cat.get("material", 0.0) * scrap_pct / 100
    total = sum(by_cat.values()) + scrap
    price = total / (1 - margin_pct / 100)
    return CostResult(by_cat, scrap, total, price)


def read_cost_sheet(path):
    with open(path, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _cmd_size(args):
    r = size_from_loads(read_load_list(args.csv))
    print(f"Sealed VA total          : {r.sealed_total_va:8.1f} VA")
    print(f"Inrush VA total          : {r.inrush_total_va:8.1f} VA")
    print(f"Required max inrush VA   : {r.required_inrush_va:8.1f} VA "
          "(at 40% PF, 90% secondary voltage)")
    if r.min_continuous_rating_va:
        print(f"Minimum continuous rating: {r.min_continuous_rating_va} VA "
              "(also check its published inrush rating)")
    else:
        print("Sealed load exceeds largest standard rating")


def _cmd_design(args):
    d = design(args.va, args.vp, args.vs, freq=args.freq, b_tesla=args.b,
               j_a_mm2=args.j, k_core=args.k, tongue_mm=args.tongue)
    print(f"Rating            : {d.va:.0f} VA, {d.primary.volts:.0f} V / {d.secondary.volts:.0f} V")
    print(f"Core              : tongue {d.tongue_mm:.0f} mm x stack {d.stack_mm:.0f} mm, "
          f"net area {d.core_area_cm2:.1f} cm2")
    print(f"Turns per volt    : {d.turns_per_volt:.3f}")
    for name, w in (("Primary", d.primary), ("Secondary", d.secondary)):
        print(f"{name:<18}: {w.turns} turns, {w.amps:.2f} A, wire {w.wire_mm:.3f} mm")
    print(f"Window fill       : {d.window_fill:.2f} (aim <= 0.40)")
    print(f"Core mass (approx): {d.core_mass_kg:.2f} kg")
    print(f"Copper mass (appr): {d.copper_mass_kg:.2f} kg")
    if d.window_fill > 0.40:
        print("WARNING: winding may not fit - increase tongue size or lower current density")


def _cmd_cost(args):
    c = cost_from_rows(read_cost_sheet(args.csv), args.scrap, args.margin)
    for cat, amount in c.by_category.items():
        print(f"{cat.capitalize():<16}: {amount:10.2f}")
    print(f"{'Scrap allowance':<16}: {c.scrap_allowance:10.2f}")
    print(f"{'Total cost':<16}: {c.total_cost:10.2f}")
    print(f"{'Selling price':<16}: {c.selling_price:10.2f}  ({args.margin:.0f}% gross margin)")


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("size", help="size from load-list CSV")
    s.add_argument("csv")
    s.set_defaults(func=_cmd_size)

    d = sub.add_parser("design", help="first-cut EI transformer design")
    d.add_argument("--va", type=float, required=True)
    d.add_argument("--vp", type=float, required=True, help="primary volts")
    d.add_argument("--vs", type=float, required=True, help="secondary volts")
    d.add_argument("--freq", type=float, default=50.0)
    d.add_argument("--b", type=float, default=1.3, help="peak flux density, T")
    d.add_argument("--j", type=float, default=3.0, help="current density, A/mm2")
    d.add_argument("--k", type=float, default=1.1, help="core-area constant")
    d.add_argument("--tongue", type=float, default=None, help="tongue width, mm")
    d.set_defaults(func=_cmd_design)

    c = sub.add_parser("cost", help="unit cost and price from cost-sheet CSV")
    c.add_argument("csv")
    c.add_argument("--scrap", type=float, default=3.0, help="material scrap %%")
    c.add_argument("--margin", type=float, default=25.0, help="gross margin %%")
    c.set_defaults(func=_cmd_cost)

    args = p.parse_args(argv)
    try:
        args.func(args)
    except (ValueError, KeyError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
