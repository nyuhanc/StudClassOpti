"""17. Special-needs (PP == 1) students split into exactly N classes (min..max each)."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class SpecialNeedsSplit(Constraint):
    id = "c17"
    enabled_field = "pp_split_enabled"
    label = "Distribute special-needs students"
    description = (
        "Students flagged special-needs (PP column == 1) are concentrated into exactly "
        "`pp_classes` classes, each holding between `pp_min_per_class` and `pp_max_per_class` "
        "of them; the remaining classes hold none."
    )
    parameters = [
        Parameter("pp_classes", "Number of classes with PP students", "int", minimum=1),
        Parameter("pp_min_per_class", "Min PP students per such class", "int", minimum=0),
        Parameter("pp_max_per_class", "Max PP students per such class", "int", minimum=0),
    ]

    def apply(self, ctx: ModelContext) -> None:
        c = ctx.config.constraints
        if "PP" not in ctx.data.columns:
            return
        pp = ctx.data[ctx.data["PP"] == 1]["Student"].tolist()
        if not pp:
            return

        counts = []
        for i in range(1, ctx.config.num_of_classes + 1):
            in_i = [ctx.in_class(s, i) for s in pp]
            cnt = ctx.model.NewIntVar(0, len(pp), f"pp_count_in_class_{i}")
            ctx.model.Add(cnt == sum(in_i))
            counts.append(cnt)

        has_pp_flags = []
        for i, cnt in enumerate(counts, start=1):
            has = ctx.model.NewBoolVar(f"class_{i}_has_pp")
            ctx.model.Add(cnt >= 1).OnlyEnforceIf(has)
            ctx.model.Add(cnt == 0).OnlyEnforceIf(has.Not())
            has_pp_flags.append(has)
        ctx.model.Add(sum(has_pp_flags) == c.pp_classes)

        for cnt in counts:
            nz = ctx.model.NewBoolVar(f"{cnt.Name()}_nonzero")
            ctx.model.Add(cnt >= 1).OnlyEnforceIf(nz)
            ctx.model.Add(cnt == 0).OnlyEnforceIf(nz.Not())
            ctx.model.Add(cnt >= c.pp_min_per_class).OnlyEnforceIf(nz)
            ctx.model.Add(cnt <= c.pp_max_per_class).OnlyEnforceIf(nz)
