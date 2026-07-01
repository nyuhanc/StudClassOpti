"""11. Gated subject as 1st/2nd priority -> the student must be assigned it."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class RequireSubjectForTopPriority(Constraint):
    id = "c11"
    enabled_field = "require_subject_for_top_priority_enabled"
    label = "Require subject for high priority"
    description = (
        "Students who rank `gated_subject` as their 1st or 2nd science priority "
        "are guaranteed it in one of their two slots."
    )
    parameters = [
        Parameter("gated_subject", "Subject", "choice", choices_from="nat_sci_classes"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        gated = ctx.config.nat_sci_index(ctx.config.constraints.gated_subject)
        for s in ctx.students:
            if gated in (ctx.ns_priorities[s][0], ctx.ns_priorities[s][1]):
                ctx.model.AddBoolOr([ctx.ns1_is(s, gated), ctx.ns2_is(s, gated)])
