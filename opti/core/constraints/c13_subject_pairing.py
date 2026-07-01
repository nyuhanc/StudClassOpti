"""13. Anyone assigned `paired_from` must take `paired_to` in the other slot."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class SubjectPairing(Constraint):
    id = "c13"
    enabled_field = "subject_pairing_enabled"
    label = "Pair science subjects"
    description = "A student assigned `paired_from` must take `paired_to` as their other science."
    parameters = [
        Parameter("paired_from", "From subject", "choice", choices_from="nat_sci_classes"),
        Parameter("paired_to", "To subject", "choice", choices_from="nat_sci_classes"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        pfrom = ctx.config.nat_sci_index(ctx.config.constraints.paired_from)
        pto = ctx.config.nat_sci_index(ctx.config.constraints.paired_to)
        for s in ctx.students:
            ns1_is = ctx.ns1_is(s, pfrom)
            ns2_is = ctx.ns2_is(s, pfrom)
            ctx.model.Add(ctx.ns2_var[s] == pto).OnlyEnforceIf(ns1_is)
            ctx.model.Add(ctx.ns1_var[s] == pto).OnlyEnforceIf(ns2_is)
