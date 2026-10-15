"""15. All students assigned ``join_language`` share one class."""

from __future__ import annotations

from .base import Constraint, ModelContext, Parameter


class JoinLanguageClass(Constraint):
    id = "c15"
    enabled_field = "join_language_class_enabled"
    label = "Join language into one class"
    description = (
        "Every student assigned `join_language` is placed in a single common class. "
        "This caps that language at one class -- at most max_class_size students may take it."
    )
    parameters = [
        Parameter("join_language", "Language", "choice", choices_from="languages"),
    ]

    def apply(self, ctx: ModelContext) -> None:
        lang = ctx.config.language_index(ctx.config.constraints.join_language)
        got_lang = [ctx.has_lang(s, lang) for s in ctx.students]

        # A single solver-chosen class holds everyone who gets this language.
        join_class = ctx.model.NewIntVar(1, ctx.config.num_of_classes, f"join_lang_{lang}_class")
        for s, got in zip(ctx.students, got_lang):
            ctx.model.Add(ctx.class_var[s] == join_class).OnlyEnforceIf(got)

        # Implied: they all fit in that one class.
        ctx.model.Add(sum(got_lang) <= ctx.config.max_class_size)
