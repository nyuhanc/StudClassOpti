"""1. At most ``max_class_size`` students per class."""

from __future__ import annotations

from .base import Constraint, ModelContext


class ClassSize(Constraint):
    id = "c01"
    enabled_field = "max_class_size_enabled"
    label = "Max class size"
    description = "At most max_class_size students per class."
    parameters = []  # uses the top-level SolverConfig.max_class_size

    def apply(self, ctx: ModelContext) -> None:
        for i in range(1, ctx.config.num_of_classes + 1):
            members = [ctx.in_class(s, i) for s in ctx.students]
            ctx.model.Add(sum(members) <= ctx.config.max_class_size)
