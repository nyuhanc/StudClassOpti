"""8. Males split into exactly N classes (min..max each); the rest hold none."""

from __future__ import annotations

import pandas as pd

from ..config import SolverConfig
from ..data import DataError
from .base import Constraint, ModelContext, Parameter


def _male_mask(data: pd.DataFrame, token: str) -> pd.Series:
    """Case-insensitive match of the Gender column against the male token.

    So a 'm' token matches sheets that encode gender as 'M' (and 'ž'/'Ž')."""
    return data["Gender"].astype(str).str.lower() == str(token).lower()


class MaleSplit(Constraint):
    id = "c08"
    enabled_field = "male_split_enabled"
    label = "Distribute male students"
    description = (
        "Male students are concentrated into exactly `male_classes` classes, each "
        "holding between `male_min_per_class` and `male_max_per_class` of them."
    )
    parameters = [
        Parameter("male_gender_value", "Male gender token", "str"),
        Parameter("male_classes", "Number of classes with males", "int", minimum=1),
        Parameter("male_min_per_class", "Min males per such class", "int", minimum=0),
        Parameter("male_max_per_class", "Max males per such class", "int", minimum=0),
    ]

    def apply(self, ctx: ModelContext) -> None:
        c = ctx.config.constraints
        if "Gender" not in ctx.data.columns:
            return
        males = ctx.data[_male_mask(ctx.data, c.male_gender_value)]["Student"].tolist()
        if not males:
            return

        counts = []
        for i in range(1, ctx.config.num_of_classes + 1):
            in_i = [ctx.in_class(s, i) for s in males]
            cnt = ctx.model.NewIntVar(0, len(males), f"male_count_in_class_{i}")
            ctx.model.Add(cnt == sum(in_i))
            counts.append(cnt)

        has_males_flags = []
        for i, cnt in enumerate(counts, start=1):
            has = ctx.model.NewBoolVar(f"class_{i}_has_males")
            ctx.model.Add(cnt >= 1).OnlyEnforceIf(has)
            ctx.model.Add(cnt == 0).OnlyEnforceIf(has.Not())
            has_males_flags.append(has)
        ctx.model.Add(sum(has_males_flags) == c.male_classes)

        for cnt in counts:
            nz = ctx.model.NewBoolVar(f"{cnt.Name()}_nonzero")
            ctx.model.Add(cnt >= 1).OnlyEnforceIf(nz)
            ctx.model.Add(cnt == 0).OnlyEnforceIf(nz.Not())
            ctx.model.Add(cnt >= c.male_min_per_class).OnlyEnforceIf(nz)
            ctx.model.Add(cnt <= c.male_max_per_class).OnlyEnforceIf(nz)

    def validate(self, data: pd.DataFrame, config: SolverConfig) -> list[DataError]:
        """Flag the silent no-op: enabled but the gender token matches nobody.

        Matching is case-insensitive (see ``_male_mask``), so this only fires on a
        genuine mismatch (wrong token or missing column), not a mere case diff.
        """
        c = config.constraints
        if "Gender" not in data.columns:
            return [DataError(None, "male split is enabled but the Gender column is missing")]
        if not _male_mask(data, c.male_gender_value).any():
            found = ", ".join(sorted(str(v) for v in data["Gender"].dropna().unique()))
            return [DataError(
                None,
                f"male split is enabled but no student has gender '{c.male_gender_value}' "
                f"(Gender column contains: {found}) -- set the male gender token to match the data",
            )]
        return []
