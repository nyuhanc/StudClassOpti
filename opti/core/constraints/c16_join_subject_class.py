"""16. All students taking ``join_subject`` (either slot) share one class."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class JoinSubjectClass(Constraint):
    id = "c16"
    enabled_field = "join_subject_class_enabled"
    label = "Join science subject into one class"
    description = (
        "Every student taking `join_subject` in either science slot is placed in a single "
        "common class. This caps that subject at one class -- at most max_class_size students "
        "may take it."
    )
    parameters = [
        Parameter("join_subject", "Subject", "choice", choices_from="nat_sci_classes"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        subj = ctx.config.nat_sci_index(ctx.config.constraints.join_subject)

        # Per-student "takes this subject in either slot" indicator.
        takes = []
        for s in ctx.students:
            t = ctx.model.NewBoolVar(f"{s}_takes_ns_{subj}")
            got = [ctx.ns1_is(s, subj), ctx.ns2_is(s, subj)]
            ctx.model.AddBoolOr(got).OnlyEnforceIf(t)
            ctx.model.AddBoolAnd([g.Not() for g in got]).OnlyEnforceIf(t.Not())
            takes.append(t)

        # A single solver-chosen class holds everyone who takes this subject.
        join_class = ctx.model.NewIntVar(1, ctx.config.num_of_classes, f"join_ns_{subj}_class")
        for s, t in zip(ctx.students, takes):
            ctx.model.Add(ctx.class_var[s] == join_class).OnlyEnforceIf(t)

        # Implied: they all fit in that one class.
        ctx.model.Add(sum(takes) <= ctx.config.max_class_size)
