"""9. Students whose top language is the forced one must receive it."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class ForceTopLanguage(Constraint):
    id = "c09"
    enabled_field = "force_top_language_enabled"
    label = "Force top language"
    description = "Students whose first-choice language is `forced_top_language` are guaranteed it."
    parameters = [
        Parameter("forced_top_language", "Language", "choice", choices_from="languages"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        forced = ctx.config.language_index(ctx.config.constraints.forced_top_language)
        for s in ctx.students:
            if ctx.lang_priorities[s][0] == forced:
                ctx.model.Add(ctx.lang_var[s] == forced)
