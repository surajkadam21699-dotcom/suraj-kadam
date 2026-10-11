import math
import os
import sys

import pytest

HERE = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(HERE, "..", "tools"))

import ct_calc  # noqa: E402

TEMPLATES = os.path.join(HERE, "..", "templates")


def test_size_worked_example_from_module_02():
    loads = ct_calc.read_load_list(os.path.join(TEMPLATES, "load-list.csv"))
    r = ct_calc.size_from_loads(loads)
    assert r.sealed_total_va == pytest.approx(47)
    assert r.inrush_total_va == pytest.approx(207)
    assert r.required_inrush_va == pytest.approx(math.hypot(47, 207))
    assert r.min_continuous_rating_va == 50


def test_size_exercise_from_module_02():
    r = ct_calc.size_from_loads([
        {"qty": 1, "sealed_va": 120, "inrush_va": 650, "starts_together": True},
    ])
    # inrush total here is the inrush figure itself
    assert r.required_inrush_va == pytest.approx(661, abs=1)


def test_design_obeys_emf_equation_and_fits_window():
    d = ct_calc.design(500, 415, 110)
    net_area_m2 = d.core_area_cm2 * 1e-4
    emf = 4.44 * 50 * d.primary.turns * 1.3 * net_area_m2
    assert emf == pytest.approx(415, rel=0.01)
    assert d.secondary.turns > 110 * d.turns_per_volt  # regulation compensation
    assert d.window_fill <= 0.40
    assert d.primary.copper_area_mm2 * 3.0 >= d.primary.amps
    assert d.secondary.copper_area_mm2 * 3.0 >= d.secondary.amps


def test_pick_wire_rounds_up_to_standard_size():
    assert ct_calc.pick_wire(0.42) == 0.800
    with pytest.raises(ValueError):
        ct_calc.pick_wire(100)


def test_cost_scrap_and_margin():
    rows = [
        {"category": "material", "quantity": 2, "rate": 50},
        {"category": "labour", "quantity": 1, "rate": 20},
    ]
    c = ct_calc.cost_from_rows(rows, scrap_pct=10, margin_pct=20)
    assert c.scrap_allowance == pytest.approx(10)
    assert c.total_cost == pytest.approx(130)
    assert c.selling_price == pytest.approx(162.5)


def test_cli_runs_on_templates(capsys):
    assert ct_calc.main(["size", os.path.join(TEMPLATES, "load-list.csv")]) == 0
    assert ct_calc.main(["cost", os.path.join(TEMPLATES, "cost-sheet.csv")]) == 0
    assert ct_calc.main(["design", "--va", "100", "--vp", "230", "--vs", "24"]) == 0
    assert "Window fill" in capsys.readouterr().out
