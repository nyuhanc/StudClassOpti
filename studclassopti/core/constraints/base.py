"""Building blocks for self-describing constraints.

Each hard constraint of the model is a :class:`Constraint`: it knows its own
toggle flag, label, description, and parameter schema, and applies itself to a
shared :class:`ModelContext`. The GUI (Phase 3) walks the registry and renders a
checkbox + a widget per :class:`Parameter` with no constraint-specific code.

``ModelContext`` owns the decision variables and the reified indicator bools, so
constraints (and the objective) share the same ``in_class`` / ``has_lang`` /
``ns*_is`` helpers instead of each rebuilding identical bool vars.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
from ortools.sat.python import cp_model

from ..config import SolverConfig
from ..data import DataError


@dataclass
class Parameter:
    """GUI/schema description of one tunable field on ``ConstraintsConfig``.

    ``key`` is the attribute name on ``ConstraintsConfig`` (the typed config
    stays the source of truth; the GUI reads/writes it via getattr/setattr).
    ``choices_from`` names a list on ``SolverConfig`` ("languages" /
    "nat_sci_classes") whose items are the valid options for a "choice" field.
    """

    key: str
    label: str
    kind: str  # "int" | "str" | "choice"
    minimum: Optional[int] = None
    maximum: Optional[int] = None
    choices_from: Optional[str] = None


class ModelContext:
    """Decision variables + memoized indicator bools for one model build."""

    def __init__(self, data: pd.DataFrame, config: SolverConfig, best_pair: tuple[str, str]):
        self.data = data
        self.config = config
        self.best_pair = best_pair
        self.model = cp_model.CpModel()
        self.students: list[int] = data["Student"].tolist()
        self.num_langs = len(config.languages)
        self.num_ns = len(config.nat_sci_classes)

        # ---- decision variables ----
        self.class_var: dict[int, cp_model.IntVar] = {}
        self.lang_var: dict[int, cp_model.IntVar] = {}
        self.ns1_var: dict[int, cp_model.IntVar] = {}
        self.ns2_var: dict[int, cp_model.IntVar] = {}
        for s in self.students:
            self.class_var[s] = self.model.NewIntVar(1, config.num_of_classes, f"{s}_class")
            self.lang_var[s] = self.model.NewIntVar(1, self.num_langs, f"{s}_lang")
            self.ns1_var[s] = self.model.NewIntVar(1, self.num_ns, f"{s}_nat_sci_1")
            self.ns2_var[s] = self.model.NewIntVar(1, self.num_ns, f"{s}_nat_sci_2")

        # ---- preference priority lists (plain ints, known at build time) ----
        # lang_priorities[s][k] = 1-based id of the language ranked at position k+1.
        self.lang_priorities: dict[int, list[int]] = {}
        self.ns_priorities: dict[int, list[int]] = {}
        for s in self.students:
            row = data[data["Student"] == s].iloc[0]
            lang_ratings = [row[lang] for lang in config.languages]
            self.lang_priorities[s] = [
                lang_ratings.index(rank) + 1 for rank in range(1, self.num_langs + 1)
            ]
            ns_ratings = [row[ns] for ns in config.nat_sci_classes]
            self.ns_priorities[s] = [
                ns_ratings.index(rank) + 1 for rank in range(1, self.num_ns + 1)
            ]

        self._bool_cache: dict[tuple, cp_model.IntVar] = {}

    # ---- memoized reified indicator bools (var == value) ----
    def _eq_bool(self, kind: str, var: cp_model.IntVar, value: int, s: int) -> cp_model.IntVar:
        key = (kind, s, value)
        b = self._bool_cache.get(key)
        if b is None:
            b = self.model.NewBoolVar(f"{s}_{kind}_{value}")
            self.model.Add(var == value).OnlyEnforceIf(b)
            self.model.Add(var != value).OnlyEnforceIf(b.Not())
            self._bool_cache[key] = b
        return b

    def in_class(self, s: int, i: int) -> cp_model.IntVar:
        return self._eq_bool("in_class", self.class_var[s], i, s)

    def has_lang(self, s: int, i: int) -> cp_model.IntVar:
        return self._eq_bool("has_lang", self.lang_var[s], i, s)

    def ns1_is(self, s: int, i: int) -> cp_model.IntVar:
        return self._eq_bool("ns1_is", self.ns1_var[s], i, s)

    def ns2_is(self, s: int, i: int) -> cp_model.IntVar:
        return self._eq_bool("ns2_is", self.ns2_var[s], i, s)


class Constraint(ABC):
    """A toggleable, self-describing hard constraint."""

    id: str
    enabled_field: str  # name of the ``*_enabled`` flag on ConstraintsConfig
    label: str
    description: str
    parameters: list[Parameter] = []

    @abstractmethod
    def apply(self, ctx: ModelContext) -> None:
        """Add this constraint to ``ctx.model`` using ``ctx.config``."""

    def validate(self, data: pd.DataFrame, config: SolverConfig) -> list[DataError]:
        """Return data problems that would make this constraint unusable.

        Default: nothing to check. Override for constraints with data-dependent
        prerequisites (e.g. schoolmate references).
        """
        return []
