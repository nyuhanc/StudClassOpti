"""14. Top-priority language == exclusion language -> must NOT get the excluded one."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class LanguageExclusion(Constraint):
    id = "c14"
    enabled_field = "language_exclusion_enabled"
    label = "Language exclusion"
    description = (
        "Students whose first-choice language is `exclusion_top_language` are never "
        "assigned `excluded_language`."
    )
    parameters = [
        Parameter("exclusion_top_language", "Top-choice language", "choice", choices_from="languages"),
        Parameter("excluded_language", "Excluded language", "choice", choices_from="languages"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        c = ctx.config.constraints
        excl_top = ctx.config.language_index(c.exclusion_top_language)
        excluded = ctx.config.language_index(c.excluded_language)
        for s in ctx.students:
            if ctx.lang_priorities[s][0] == excl_top:
                ctx.model.Add(ctx.lang_var[s] != excluded)
