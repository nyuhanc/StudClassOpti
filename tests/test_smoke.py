"""Phase 1 smoke tests.

These don't pin exact objective values (CP-SAT + random shuffles are not
deterministic enough across machines for that). They verify the core wiring:
config round-trips, validation catches bad data, the model is feasible, and the
hard constraints actually hold in the returned solution.

Run with:  pytest   (install pytest first: pip install pytest)
"""

from __future__ import annotations

import os

import pytest

from opti.core import (
    SolverConfig,
    load_excel,
    preprocess,
    validate,
    best_nat_sci_pair,
    solve,
)
from opti.core.constraints.registry import REGISTRY, enabled_constraints, validate_constraints

# The Slovene spreadsheet is the real input; translate its headers to the English
# names core uses via the same bridge the GUI applies on load.
from gui.i18n import excel_to_internal

EXCEL = os.path.join(os.path.dirname(__file__), "..", "dijaki_2025.xlsx")


@pytest.fixture(scope="module")
def data():
    config = SolverConfig()
    return preprocess(excel_to_internal(load_excel(EXCEL)), config)


def test_config_json_roundtrip(tmp_path):
    config = SolverConfig()
    config.max_class_size = 25
    config.constraints.best_pair = ("Biology", "Physics")
    config.objective.stratification = 3

    path = tmp_path / "config.json"
    config.to_json(path)
    loaded = SolverConfig.from_json(path)

    assert loaded.max_class_size == 25
    assert loaded.constraints.best_pair == ("Biology", "Physics")
    assert loaded.objective.stratification == 3


def test_validation_passes_on_real_data(data):
    assert validate(data, SolverConfig()) == []


def test_validation_flags_bad_permutation(data):
    config = SolverConfig()
    bad = data.copy()
    # Break one student's language priorities (duplicate a value).
    bad.loc[bad.index[0], "French"] = bad.loc[bad.index[0], "Italian"]
    errors = validate(bad, config)
    assert any("language priorities" in str(e) for e in errors)


def test_male_split_gender_matching_is_case_insensitive(data):
    """Token 'm' matches 'M'/'Ž' data; only a genuine mismatch is flagged."""
    config = SolverConfig()  # male split on, token 'm'
    upper = data.copy()
    upper["Gender"] = upper["Gender"].str.upper()  # 'M'/'Ž'
    # Case difference alone must NOT error and must NOT silently ignore males.
    assert not any("gender" in str(e) for e in validate_constraints(upper, config))
    males = upper[upper["Gender"].str.lower() == "m"]["Student"].tolist()
    assert males, "test data should contain males"

    # A token matching nobody (even case-insensitively) is still flagged.
    config.constraints.male_gender_value = "x"
    assert any("no student has gender 'x'" in str(e) for e in validate_constraints(upper, config))


def test_best_pair_is_a_valid_pair(data):
    config = SolverConfig()
    pair, count, counts = best_nat_sci_pair(data, config)
    assert pair[0] in config.nat_sci_classes
    assert pair[1] in config.nat_sci_classes
    assert count >= 0


def test_registry_toggle_controls_enabled_set():
    """enabled_constraints() honours the per-constraint flags."""
    config = SolverConfig()

    # The redundant science-capacity constraint (c03) ships off by default.
    enabled_ids = {c.id for c in enabled_constraints(config)}
    assert "c03" not in enabled_ids
    # ...and the rest of the default-on constraints are present.
    assert {"c01", "c02", "c04", "c13"} <= enabled_ids

    # Flipping a flag adds/removes exactly that constraint.
    config.constraints.nat_sci_capacity_enabled = True
    assert "c03" in {c.id for c in enabled_constraints(config)}

    config.constraints.subject_pairing_enabled = False
    assert "c13" not in {c.id for c in enabled_constraints(config)}

    # Every registered constraint maps to a real flag on the config.
    for c in REGISTRY:
        assert hasattr(config.constraints, c.enabled_field)


def test_solve_quick_feasible_and_constraints_hold(data):
    """A fast 1-shuffle, short-time-limit run that still respects constraints."""
    config = SolverConfig()
    config.shuffles = 1
    config.time_limit_per_shuffle = 20
    config.num_workers = 4
    config.random_seed = 0

    result = solve(data, config)
    assert result.found, "expected a feasible solution"

    out = result.best_data

    # Constraint 1: class sizes within limit.
    assert (out["Class"].value_counts() <= config.max_class_size).all()

    # Constraint 2: at most (multiplier * max_class_size) students per language.
    lang_cap = config.constraints.lang_capacity_multiplier * config.max_class_size
    assert (out["Language"].value_counts() <= lang_cap).all()

    # Constraint 4: the two sciences differ for every student.
    assert (out["NatSci1"] != out["NatSci2"]).all()

    takes_physics = (out["NatSci1"] == "Physics") | (out["NatSci2"] == "Physics")

    # Constraint 10: at most max_class_size students take Physics.
    assert takes_physics.sum() <= config.max_class_size

    # Constraint 11: students ranking Physics 1st or 2nd must receive it.
    wants_physics = out["Physics"].isin([1, 2])
    assert takes_physics[wants_physics].all()

    # Constraint 12: students ranking Physics last (3rd) must not receive it.
    rejects_physics = out["Physics"] == 3
    assert (~takes_physics[rejects_physics]).all()

    # Constraint 13: anyone with Physics has Biology in the other slot.
    phys = out[takes_physics]
    assert ((phys["NatSci1"] == "Biology") | (phys["NatSci2"] == "Biology")).all()

    # Constraint 9: students whose top language is Russian get Russian.
    top_russian = out["Russian"] == 1
    assert (out.loc[top_russian, "Language"] == "Russian").all()

    # Constraint 14: students whose top language is Spanish never get Italian.
    top_spanish = out["Spanish"] == 1
    assert (out.loc[top_spanish, "Language"] != "Italian").all()


def test_join_language_class_gathers_one_class(data):
    """Constraint 15 (off by default): all French students land in one class."""
    config = SolverConfig()
    config.shuffles = 1
    config.time_limit_per_shuffle = 25
    config.num_workers = 4
    config.random_seed = 0
    config.constraints.join_language_class_enabled = True
    config.constraints.join_language = "French"

    result = solve(data, config)
    assert result.found, "expected a feasible solution"

    french = result.best_data[result.best_data["Language"] == "French"]
    # All French students share a single class...
    assert french["Class"].nunique() <= 1
    # ...and therefore fit within one class's capacity.
    assert len(french) <= config.max_class_size


def test_join_subject_class_gathers_one_class(data):
    """Constraint 16 (off by default): all Physics takers land in one class."""
    config = SolverConfig()
    config.shuffles = 1
    config.time_limit_per_shuffle = 25
    config.num_workers = 4
    config.random_seed = 0
    config.constraints.join_subject_class_enabled = True
    config.constraints.join_subject = "Physics"

    result = solve(data, config)
    assert result.found, "expected a feasible solution"

    out = result.best_data
    takes_physics = (out["NatSci1"] == "Physics") | (out["NatSci2"] == "Physics")
    physics = out[takes_physics]
    # Everyone taking Physics (in either slot) shares a single class...
    assert physics["Class"].nunique() <= 1
    # ...and therefore fits within one class's capacity.
    assert takes_physics.sum() <= config.max_class_size


def test_special_needs_split_distributes_pp(data):
    """Constraint 17 (off by default): PP students spread across N classes."""
    config = SolverConfig()
    config.shuffles = 1
    config.time_limit_per_shuffle = 25
    config.num_workers = 4
    config.random_seed = 0
    config.constraints.pp_split_enabled = True
    config.constraints.pp_classes = 2
    config.constraints.pp_min_per_class = 1
    config.constraints.pp_max_per_class = 4

    result = solve(data, config)
    assert result.found, "expected a feasible solution"

    out = result.best_data
    per_class = out[out["PP"] == 1]["Class"].value_counts()
    # PP students occupy exactly pp_classes classes...
    assert len(per_class) == config.constraints.pp_classes
    # ...each within the configured min..max bounds.
    assert per_class.min() >= config.constraints.pp_min_per_class
    assert per_class.max() <= config.constraints.pp_max_per_class
