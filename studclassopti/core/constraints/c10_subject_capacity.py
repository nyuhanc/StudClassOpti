"""10. Cap students assigned to the capacity subject (max_class_size slots)."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class SubjectCapacity(Constraint):
    id = "c10"
    enabled_field = "subject_capacity_enabled"
    label = "Subject slot cap"
    description = "At most max_class_size student-slots may be assigned to `capacity_subject`."
    parameters = [
        Parameter("capacity_subject", "Subject", "choice", choices_from="nat_sci_classes"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        subj = ctx.config.nat_sci_index(ctx.config.constraints.capacity_subject)
        slots = []
        for s in ctx.students:
            slots.append(ctx.ns1_is(s, subj))
            slots.append(ctx.ns2_is(s, subj))
        ctx.model.Add(sum(slots) <= ctx.config.max_class_size)
