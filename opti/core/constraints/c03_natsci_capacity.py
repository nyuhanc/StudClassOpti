"""3. At most (multiplier * max_class_size) student-slots per science subject."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class NatSciCapacity(Constraint):
    id = "c03"
    enabled_field = "nat_sci_capacity_enabled"
    label = "Science capacity"
    description = (
        "At most (multiplier * max_class_size) student-slots per science subject "
        "(each student fills two slots)."
    )
    parameters = [
        Parameter("nat_sci_capacity_multiplier", "Capacity multiplier", "int", minimum=1, maximum=10),
    ]

    def apply(self, ctx: ModelContext) -> None:
        cap = ctx.config.constraints.nat_sci_capacity_multiplier * ctx.config.max_class_size
        for i in range(1, ctx.num_ns + 1):
            slots = []
            for s in ctx.students:
                slots.append(ctx.ns1_is(s, i))
                slots.append(ctx.ns2_is(s, i))
            ctx.model.Add(sum(slots) <= cap)
