"""Configuration objects for the solver.

Everything the original ``main_CP_v2.py`` hard-coded as module-level globals,
inline magic numbers, or interactive prompts now lives here as serializable
data. A ``SolverConfig`` fully determines a run, can be saved to / loaded from
JSON (so a configuration can be reused year to year), and is the single object
the future GUI will bind its widgets to.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional


# Canonical option orderings. The 1-based index of each item is its internal
# id used throughout the model (e.g. French -> language 1, Biology -> science 1).
DEFAULT_LANGUAGES = ["French", "Italian", "German", "Russian", "Spanish"]
DEFAULT_NAT_SCI = ["Biology", "Physics", "Chemistry"]


@dataclass
class ObjectiveWeights:
    """Weights of the maximization objective.

    These are exactly the hand-tuned knobs from the original script. Higher
    ``stratification`` makes top-ranked preferences disproportionately more
    valuable than lower-ranked ones.
    """

    lang_importance: int = 1
    lang_penalty: int = 10
    nat_sci_1_importance: int = 100
    nat_sci_2_importance: int = 1
    nat_sci_penalty: int = 100
    stratification: int = 2


@dataclass
class ConstraintsConfig:
    """Which hard constraints are active and their parameters.

    Each ``*_enabled`` flag toggles a constraint; the accompanying fields are
    that constraint's tunable parameters. Numbering follows the original script
    (5 and 7 are intentionally absent). In Phase 2 each of these becomes a
    self-describing Constraint plugin; for now they are plain fields so Phase 1
    can reproduce the existing model exactly.
    """

    # 1. Max students per class.
    max_class_size_enabled: bool = True

    # 2. Max students sharing a language (multiplier * max_class_size).
    lang_capacity_enabled: bool = True
    lang_capacity_multiplier: int = 1

    # 3. Max student-slots per science subject (multiplier * max_class_size).
    #    Off by default: with constraint 4 (distinct sciences) each science fills
    #    at most one slot per student, so slots-per-science <= num students <=
    #    num_of_classes * max_class_size, which equals the cap at multiplier 3 --
    #    redundant for the default dimensions. Kept as a toggle for setups
    #    (smaller classes / more students) where it can actually bind.
    nat_sci_capacity_enabled: bool = False
    nat_sci_capacity_multiplier: int = 3

    # 4. A student's two science subjects must differ.
    distinct_sciences_enabled: bool = True

    # 5. Keep schoolmates together: a student naming a partner in the
    #    `schoolmate_column` is placed in that partner's class. Off by default
    #    (preserves original behaviour); needs validate() for bad references.
    schoolmate_enabled: bool = False
    schoolmate_column: str = "Schoolmate"

    # 6. Concentrate the most common science pair into one class.
    #    best_pair = None  -> auto-pick the most popular pair from the data.
    #    best_pair = ("Biology", "Physics") -> force that pair.
    best_pair_class_enabled: bool = True
    best_pair: Optional[tuple[str, str]] = None

    # 8. Males split into exactly N classes, each holding min..max males,
    #    the remaining class(es) holding none.
    #    NOTE: the input data encodes gender as 'm' (male) / 'ž' (female);
    #    male_gender_value is the token treated as "male".
    male_split_enabled: bool = True
    male_gender_value: str = "m"
    male_classes: int = 2
    male_min_per_class: int = 4
    male_max_per_class: int = 7

    # 9. Students whose top language is `forced_top_language` must get it.
    force_top_language_enabled: bool = True
    forced_top_language: str = "Russian"

    # 10. Cap students assigned to `capacity_subject` (max_class_size each).
    subject_capacity_enabled: bool = True
    capacity_subject: str = "Physics"

    # 11. Students with `gated_subject` as 1st/2nd priority must receive it.
    require_subject_for_top_priority_enabled: bool = True
    # 12. Students with `gated_subject` as 3rd priority must NOT receive it.
    exclude_subject_for_low_priority_enabled: bool = True
    gated_subject: str = "Physics"

    # 13. Anyone assigned `paired_from` must take `paired_to` in the other slot.
    subject_pairing_enabled: bool = True
    paired_from: str = "Physics"
    paired_to: str = "Biology"

    # 14. Students with `exclusion_top_language` as top priority must NOT be
    #     assigned `excluded_language`.
    language_exclusion_enabled: bool = True
    exclusion_top_language: str = "Spanish"
    excluded_language: str = "Italian"

    # 15. Everyone assigned `join_language` is gathered into one class (which
    #     caps that language at max_class_size students). Off by default: it is
    #     a strong rule that can make the problem infeasible.
    join_language_class_enabled: bool = False
    join_language: str = "French"

    # 16. Everyone taking `join_subject` (in either science slot) is gathered
    #     into one class (which caps that subject at max_class_size students).
    #     Off by default: a strong rule that can make the problem infeasible.
    join_subject_class_enabled: bool = False
    join_subject: str = "Biology"

    # 17. Special-needs students (spreadsheet column PP == 1) split into exactly
    #     N classes, each holding min..max of them; the rest hold none. Mirrors
    #     the male split (8). Off by default; no-ops if the PP column is absent.
    pp_split_enabled: bool = False
    pp_classes: int = 2
    pp_min_per_class: int = 1
    pp_max_per_class: int = 3


@dataclass
class SolverConfig:
    """Full specification of an optimization run."""

    # Problem dimensions.
    num_of_classes: int = 3
    max_class_size: int = 28
    # Expected number of students. All students in the sheet are always
    # processed (no truncation).
    #   None -> accept however many rows the spreadsheet has (auto-extracted).
    #   int  -> validate() flags an error unless the sheet has exactly this
    #           many students (guards against a wrong / mis-edited file).
    expected_students: Optional[int] = None

    # Option universes (1-based index of each item is its internal id).
    languages: list[str] = field(default_factory=lambda: list(DEFAULT_LANGUAGES))
    nat_sci_classes: list[str] = field(default_factory=lambda: list(DEFAULT_NAT_SCI))

    # Search/solver tuning.
    shuffles: int = 20
    time_limit_per_shuffle: int = 120  # seconds
    num_workers: int = 8
    random_seed: Optional[int] = None  # set for reproducible shuffles/tests

    objective: ObjectiveWeights = field(default_factory=ObjectiveWeights)
    constraints: ConstraintsConfig = field(default_factory=ConstraintsConfig)

    # ---- index helpers (1-based, matching the model's encoding) ----
    def language_index(self, name: str) -> int:
        return self.languages.index(name) + 1

    def nat_sci_index(self, name: str) -> int:
        return self.nat_sci_classes.index(name) + 1

    # ---- JSON persistence ----
    def to_dict(self) -> dict:
        # tuples survive asdict as lists; that's fine for JSON round-tripping.
        return asdict(self)

    def to_json(self, path: str | Path) -> None:
        Path(path).write_text(json.dumps(self.to_dict(), indent=2), encoding="utf-8")

    @classmethod
    def from_dict(cls, d: dict) -> "SolverConfig":
        d = dict(d)
        obj = d.pop("objective", {})
        cons = dict(d.pop("constraints", {}))
        config = cls(**d)
        config.objective = ObjectiveWeights(**obj)
        bp = cons.get("best_pair")
        if bp is not None:
            cons["best_pair"] = tuple(bp)
        config.constraints = ConstraintsConfig(**cons)
        return config

    @classmethod
    def from_json(cls, path: str | Path) -> "SolverConfig":
        return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))
