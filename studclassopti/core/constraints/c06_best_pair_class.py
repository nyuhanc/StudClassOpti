"""6. Concentrate the most popular science pair into a single class."""

from __future__ import annotations

from .base import Constraint, ModelContext


class BestPairClass(Constraint):
    id = "c06"
    enabled_field = "best_pair_class_enabled"
    label = "Concentrate best science pair"
    description = (
        "Students taking the most popular science pair (auto-detected, or set "
        "explicitly) are gathered into one class."
    )
    parameters = []  # best_pair is auto-computed / set on the run, not a GUI scalar

    def apply(self, ctx: ModelContext) -> None:
        a_idx = ctx.config.nat_sci_index(ctx.best_pair[0])
        b_idx = ctx.config.nat_sci_index(ctx.best_pair[1])
        best_class = ctx.model.NewIntVar(1, ctx.config.num_of_classes, "best_matching_class")
        for s in ctx.students:
            ab = ctx.model.NewBoolVar(f"{s}_pair_ab")
            ba = ctx.model.NewBoolVar(f"{s}_pair_ba")
            ctx.model.Add(ctx.ns1_var[s] == a_idx).OnlyEnforceIf(ab)
            ctx.model.Add(ctx.ns2_var[s] == b_idx).OnlyEnforceIf(ab)
            ctx.model.Add(ctx.ns1_var[s] == b_idx).OnlyEnforceIf(ba)
            ctx.model.Add(ctx.ns2_var[s] == a_idx).OnlyEnforceIf(ba)
            # in_best uses the IntVar best_class (not a constant) -> not memoizable.
            in_best = ctx.model.NewBoolVar(f"{s}_in_best_class")
            ctx.model.Add(ctx.class_var[s] == best_class).OnlyEnforceIf(in_best)
            ctx.model.Add(ctx.class_var[s] != best_class).OnlyEnforceIf(in_best.Not())
            ctx.model.Add(ab + ba == in_best)
