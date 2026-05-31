"""Data loading, validation, and light preprocessing.

The original script validated the spreadsheet by printing errors and then
blocking on ``input(...)``. Here validation returns the problems as data, so a
CLI can print them and a GUI can render them in a panel and disable the Run
button until the file is clean. Nothing in this module prompts or exits.
"""

from __future__ import annotations

import itertools
from dataclasses import dataclass
from typing import Optional

import pandas as pd

from .config import SolverConfig


@dataclass
class DataError:
    """A single problem found in the input spreadsheet."""

    student: Optional[int]  # None for file-level/structural errors
    message: str

    def __str__(self) -> str:
        if self.student is None:
            return self.message
        return f"Student {self.student}: {self.message}"


def load_excel(path: str) -> pd.DataFrame:
    """Read the spreadsheet. Raises on a missing/unreadable file."""
    return pd.read_excel(path, engine="openpyxl")


def preprocess(data: pd.DataFrame, config: SolverConfig) -> pd.DataFrame:
    """Coerce id/int columns. Processes *all* students (no truncation).

    Returns a new frame (does not mutate the caller's).
    """
    data = data.copy()
    if "Chemistry" in data.columns:
        data["Chemistry"] = data["Chemistry"].astype(int)
    if "Student" in data.columns:
        data["Student"] = data["Student"].astype(int)
    return data


def validate(data: pd.DataFrame, config: SolverConfig) -> list[DataError]:
    """Return all problems found in ``data``. Empty list == clean.

    Checks performed:
      * required columns are present;
      * if a student count is configured, it matches the sheet;
      * each student's language priorities form a 1..N permutation;
      * each student's science priorities form a 1..M permutation.
    """
    errors: list[DataError] = []

    required = ["Student", *config.languages, *config.nat_sci_classes]
    missing = [c for c in required if c not in data.columns]
    if missing:
        errors.append(DataError(None, f"Missing required column(s): {', '.join(missing)}"))
        # Without these columns the per-student checks can't run meaningfully.
        return errors

    # If an expected student count was specified, it must match the file exactly.
    if config.expected_students is not None and len(data) != config.expected_students:
        errors.append(
            DataError(
                None,
                f"expected {config.expected_students} students but the "
                f"spreadsheet contains {len(data)}",
            )
        )

    expected_lang = set(range(1, len(config.languages) + 1))
    expected_ns = set(range(1, len(config.nat_sci_classes) + 1))

    for _, row in data.iterrows():
        student = int(row["Student"])

        lang_pris = [row[lang] for lang in config.languages]
        if set(lang_pris) != expected_lang:
            errors.append(
                DataError(
                    student,
                    f"language priorities must be a permutation of {sorted(expected_lang)}, "
                    f"got {lang_pris}",
                )
            )

        ns_pris = [row[ns] for ns in config.nat_sci_classes]
        if set(ns_pris) != expected_ns:
            errors.append(
                DataError(
                    student,
                    f"science priorities must be a permutation of {sorted(expected_ns)}, "
                    f"got {ns_pris}",
                )
            )

    return errors


def best_nat_sci_pair(
    data: pd.DataFrame, config: SolverConfig
) -> tuple[tuple[str, str], int, dict]:
    """Find the science pair most students rank as their top two.

    Replaces the interactive "please adjust constraint 6" step: the most common
    {rank-1, rank-2} science pair is computed automatically and can be fed
    straight into the model.

    Returns ``(best_pair, best_count, all_counts)`` where ``all_counts`` maps
    each ordered (a, b) pair (a before b in config order) to its match count.
    """
    counts: dict[tuple[str, str], int] = {}
    best_pair: tuple[str, str] = (config.nat_sci_classes[0], config.nat_sci_classes[1])
    best_count = 0

    pairs = [
        (a, b)
        for a, b in itertools.product(config.nat_sci_classes, repeat=2)
        if config.nat_sci_classes.index(a) < config.nat_sci_classes.index(b)
    ]
    for ns1, ns2 in pairs:
        match = 0
        for _, row in data.iterrows():
            if (row[ns1] == 1 and row[ns2] == 2) or (row[ns2] == 1 and row[ns1] == 2):
                match += 1
        counts[(ns1, ns2)] = match
        if match > best_count:
            best_count = match
            best_pair = (ns1, ns2)

    return best_pair, best_count, counts
