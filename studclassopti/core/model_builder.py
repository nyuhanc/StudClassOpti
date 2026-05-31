"""Assemble a CP-SAT model from data + a SolverConfig.

The hard constraints now live as self-describing plugins in ``core.constraints``;
``build_model`` builds a shared :class:`ModelContext`, applies every enabled
constraint from the registry, then adds the objective. Behaviour with the default
config matches the original ``main_CP_v2.py`` model.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
from ortools.sat.python import cp_model

from .config import SolverConfig
from .constraints.base import ModelContext
from .constraints.registry import enabled_constraints


@dataclass
class BuiltModel:
    """A model ready to solve plus the handles needed to read a solution."""

    model: cp_model.CpModel
    class_var: dict[int, cp_model.IntVar]
    lang_var: dict[int, cp_model.IntVar]
    ns1_var: dict[int, cp_model.IntVar]
    ns2_var: dict[int, cp_model.IntVar]


def build_objective(ctx: ModelContext) -> None:
    """Add the maximization objective (preference scoring + penalties).

    Kept separate from the constraint registry: the objective isn't a toggle, it
    is a set of weights the GUI exposes in an advanced panel.
    """
    w = ctx.config.objective
    terms = []
    for s in ctx.students:
        # Language scoring.
        for rank in range(1, ctx.num_langs + 1):
            lang_at_rank = ctx.lang_priorities[s][rank - 1]
            got = ctx.has_lang(s, lang_at_rank)
            terms.append(w.lang_importance * got * (ctx.num_langs - rank) ** w.stratification)
            if rank == 1:
                terms.append(-w.lang_penalty * got.Not())

        # Science scoring.
        ns1_top = None
        ns2_top = None
        for rank in range(1, ctx.num_ns + 1):
            ns_at_rank = ctx.ns_priorities[s][rank - 1]
            g1 = ctx.ns1_is(s, ns_at_rank)
            terms.append(w.nat_sci_1_importance * g1 * (ctx.num_ns - rank) ** w.stratification)
            g2 = ctx.ns2_is(s, ns_at_rank)
            terms.append(w.nat_sci_2_importance * g2 * (ctx.num_ns - rank) ** w.stratification)
            if rank == 1:
                ns1_top = g1
                ns2_top = g2

        # Penalize once if neither slot holds the top science choice.
        top_in_any = ctx.model.NewBoolVar(f"{s}_top_ns_in_any_slot")
        ctx.model.AddBoolOr([ns1_top, ns2_top]).OnlyEnforceIf(top_in_any)
        ctx.model.AddBoolAnd([ns1_top.Not(), ns2_top.Not()]).OnlyEnforceIf(top_in_any.Not())
        terms.append(-w.nat_sci_penalty * top_in_any.Not())

    ctx.model.Maximize(sum(terms))


def build_model(
    data: pd.DataFrame,
    config: SolverConfig,
    best_pair: tuple[str, str],
) -> BuiltModel:
    """Build and return the CP-SAT model for one (already shuffled) frame."""
    ctx = ModelContext(data, config, best_pair)
    for constraint in enabled_constraints(config):
        constraint.apply(ctx)
    build_objective(ctx)
    return BuiltModel(ctx.model, ctx.class_var, ctx.lang_var, ctx.ns1_var, ctx.ns2_var)
