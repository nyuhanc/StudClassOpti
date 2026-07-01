"""12. Gated subject as 3rd priority -> the student must NOT be assigned it."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class ExcludeSubjectForLowPriority(Constraint):
    id = "c12"
    enabled_field = "exclude_subject_for_low_priority_enabled"
    label = "Exclude subject for low priority"
    description = (
        "Students who rank `gated_subject` as their lowest science priority are "
        "never assigned it. (Shares `gated_subject` with the require-subject rule.)"
    )
    parameters = [
        Parameter("gated_subject", "Subject", "choice", choices_from="nat_sci_classes"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        gated = ctx.config.nat_sci_index(ctx.config.constraints.gated_subject)
        for s in ctx.students:
            if ctx.ns_priorities[s][2] == gated:
                ctx.model.Add(ctx.ns1_var[s] != gated)
                ctx.model.Add(ctx.ns2_var[s] != gated)
