"""Run the multi-shuffle optimization.

The shuffle loop, progress reporting, and best-solution bookkeeping that lived
in the body of ``main_CP_v2.py`` move here behind a clean function signature:

    result = solve(data, config, progress_cb=..., should_cancel=...)

* ``progress_cb`` is called once per shuffle with a ``SolveProgress`` so a CLI
  can print it and a GUI can drive a progress bar / live best-score readout.
* ``should_cancel`` is polled between shuffles so a GUI Cancel button (or Ctrl-C
  wrapper) can stop the run early and still return the best solution so far.

No prints, no prompts, no file I/O.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Optional

import pandas as pd
from ortools.sat.python import cp_model

from .config import SolverConfig
from .data import best_nat_sci_pair
from .model_builder import build_model


@dataclass
class SolveProgress:
    """Status snapshot emitted after each shuffle."""

    shuffle: int            # 1-based index of the shuffle just finished
    total_shuffles: int
    status: str             # CP-SAT status name, e.g. "OPTIMAL", "INFEASIBLE"
    score: Optional[float]  # this shuffle's objective, or None if no solution
    best_score: float       # best objective seen so far


@dataclass
class SolverResult:
    """Outcome of a full run."""

    best_score: float
    best_data: Optional[pd.DataFrame]  # input frame + Class/Language/NatSci1/NatSci2
    best_pair: tuple[str, str]
    shuffles_run: int
    cancelled: bool = False
    progress: list[SolveProgress] = field(default_factory=list)

    @property
    def found(self) -> bool:
        return self.best_data is not None


ProgressCb = Callable[[SolveProgress], None]
CancelCb = Callable[[], bool]


def _assign_solution(
    data: pd.DataFrame, built, solver: cp_model.CpSolver, config: SolverConfig
) -> pd.DataFrame:
    """Attach the solver's assignment as result columns to a copy of ``data``."""
    out = data.copy()
    out["Class"] = None
    out["Language"] = None
    out["NatSci1"] = None
    out["NatSci2"] = None
    for s in data["Student"].tolist():
        out.loc[out["Student"] == s, "Class"] = solver.Value(built.class_var[s])
        out.loc[out["Student"] == s, "Language"] = config.languages[solver.Value(built.lang_var[s]) - 1]
        out.loc[out["Student"] == s, "NatSci1"] = config.nat_sci_classes[solver.Value(built.ns1_var[s]) - 1]
        out.loc[out["Student"] == s, "NatSci2"] = config.nat_sci_classes[solver.Value(built.ns2_var[s]) - 1]
    return out


def solve(
    data: pd.DataFrame,
    config: SolverConfig,
    progress_cb: Optional[ProgressCb] = None,
    should_cancel: Optional[CancelCb] = None,
) -> SolverResult:
    """Run the configured number of shuffles and return the best solution.

    ``data`` should already be preprocessed (see ``core.data.preprocess``) and
    validated. The best science pair for constraint 6 is taken from the config
    if set, otherwise auto-computed from the data.
    """
    if config.constraints.best_pair is not None:
        best_pair = tuple(config.constraints.best_pair)
    else:
        best_pair, _, _ = best_nat_sci_pair(data, config)

    best_score = float("-inf")
    best_data: Optional[pd.DataFrame] = None
    progress: list[SolveProgress] = []
    shuffles_run = 0
    cancelled = False

    # Reproducible shuffling when a seed is configured (useful for tests).
    rng_state = config.random_seed

    for idx in range(config.shuffles):
        if should_cancel is not None and should_cancel():
            cancelled = True
            break

        if config.random_seed is not None:
            shuffled = data.sample(frac=1, random_state=rng_state).reset_index(drop=True)
            rng_state += 1  # vary per shuffle but deterministically
        else:
            shuffled = data.sample(frac=1).reset_index(drop=True)

        built = build_model(shuffled, config, best_pair)

        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = config.num_workers
        solver.parameters.max_time_in_seconds = config.time_limit_per_shuffle

        status = solver.Solve(built.model)
        shuffles_run = idx + 1

        score: Optional[float] = None
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            score = solver.ObjectiveValue()
            if score > best_score:
                best_score = score
                best_data = _assign_solution(shuffled, built, solver, config)

        snap = SolveProgress(
            shuffle=idx + 1,
            total_shuffles=config.shuffles,
            status=solver.StatusName(status),
            score=score,
            best_score=best_score if best_data is not None else 0.0,
        )
        progress.append(snap)
        if progress_cb is not None:
            progress_cb(snap)

    return SolverResult(
        best_score=best_score if best_data is not None else 0.0,
        best_data=best_data,
        best_pair=tuple(best_pair),
        shuffles_run=shuffles_run,
        cancelled=cancelled,
        progress=progress,
    )
