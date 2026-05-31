"""5. Students who request a schoolmate must share that schoolmate's class.

Off by default. The request comes from an optional ``Schoolmate`` column whose
value is the partner student's id (blank / 0 / NaN means "no request"). The
constraint is symmetric in effect: ``class[s] == class[mate]``.

``validate`` only catches the obvious data errors (unknown partner id,
self-reference). It does not try to prove that the resulting togetherness groups
fit within class capacity -- that is left to the solver.
"""

from __future__ import annotations

import pandas as pd

from ..config import SolverConfig
from ..data import DataError
from .base import Constraint, ModelContext, Parameter


def _requests(data: pd.DataFrame, column: str) -> list[tuple[int, int]]:
    """Yield (student, partner) pairs from a Schoolmate column, skipping blanks."""
    if column not in data.columns:
        return []
    pairs: list[tuple[int, int]] = []
    for _, row in data.iterrows():
        raw = row[column]
        if pd.isna(raw):
            continue
        try:
            mate = int(raw)
        except (ValueError, TypeError):
            continue
        if mate == 0:
            continue
        pairs.append((int(row["Student"]), mate))
    return pairs


class Schoolmate(Constraint):
    id = "c05"
    enabled_field = "schoolmate_enabled"
    label = "Keep schoolmates together"
    description = (
        "A student who names a schoolmate (in the Schoolmate column) is placed in "
        "the same class as that schoolmate."
    )
    parameters = [
        Parameter("schoolmate_column", "Schoolmate column name", "str"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        column = ctx.config.constraints.schoolmate_column
        valid = set(ctx.students)
        for student, mate in _requests(ctx.data, column):
            if mate in valid and mate != student:
                ctx.model.Add(ctx.class_var[student] == ctx.class_var[mate])

    def validate(self, data: pd.DataFrame, config: SolverConfig) -> list[DataError]:
        column = config.constraints.schoolmate_column
        if column not in data.columns:
            return [DataError(None, f"schoolmate constraint enabled but column '{column}' is missing")]
        valid = set(int(s) for s in data["Student"].tolist())
        errors: list[DataError] = []
        for student, mate in _requests(data, column):
            if mate == student:
                errors.append(DataError(student, "names itself as a schoolmate"))
            elif mate not in valid:
                errors.append(DataError(student, f"names unknown schoolmate {mate}"))
        return errors
