"""4. A student's two science subjects must differ."""

from __future__ import annotations

from .base import Constraint, ModelContext


class DistinctSciences(Constraint):
    id = "c04"
    enabled_field = "distinct_sciences_enabled"
    label = "Distinct sciences"
    description = "A student's two science subjects must be different."
    parameters = []

    def apply(self, ctx: ModelContext) -> None:
        for s in ctx.students:
            ctx.model.Add(ctx.ns1_var[s] != ctx.ns2_var[s])
