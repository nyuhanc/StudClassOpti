"""2. At most (multiplier * max_class_size) students per language."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class LangCapacity(Constraint):
    id = "c02"
    enabled_field = "lang_capacity_enabled"
    label = "Language capacity"
    description = "At most (multiplier * max_class_size) students may share a language."
    parameters = [
        Parameter("lang_capacity_multiplier", "Capacity multiplier", "int", minimum=1, maximum=10),
    ]

    def apply(self, ctx: ModelContext) -> None:
        cap = ctx.config.constraints.lang_capacity_multiplier * ctx.config.max_class_size
        for i in range(1, ctx.num_langs + 1):
            holders = [ctx.has_lang(s, i) for s in ctx.students]
            ctx.model.Add(sum(holders) <= cap)
