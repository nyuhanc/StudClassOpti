"""Command-line front end that reproduces the behaviour of main_CP_v2.py
using the refactored ``studclassopti.core`` package.

This exists to de-risk Phase 1: it lets us confirm the extracted core produces
the same kind of result as the original script before any GUI is built. All the
logic now lives in ``core``; this file only handles console I/O and file output.

Usage:
    python run_cli.py [path/to/students.xlsx] [--config config.json]
"""

from __future__ import annotations

import argparse
import os

from studclassopti.core import (
    SolverConfig,
    load_excel,
    preprocess,
    validate,
    best_nat_sci_pair,
    solve,
)
from studclassopti.core.solver import SolveProgress


def _print_progress(p: SolveProgress) -> None:
    score_txt = f"{p.score:.0f}" if p.score is not None else "no solution"
    print(
        f"  Shuffle {p.shuffle}/{p.total_shuffles}: {p.status}, "
        f"score={score_txt}, best so far={p.best_score:.0f}"
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Student class optimizer (CLI).")
    parser.add_argument("excel", nargs="?", default="students_list_2025.xlsx")
    parser.add_argument("--config", help="Path to a SolverConfig JSON file.")
    args = parser.parse_args()

    config = SolverConfig.from_json(args.config) if args.config else SolverConfig()

    data = load_excel(args.excel)
    data = preprocess(data, config)

    # ---- validation (was a blocking input() prompt) ----
    errors = validate(data, config)
    if errors:
        print("Errors in the data:")
        for e in errors:
            print(f"  - {e}")
        print("\nFix the errors and run again.")
        return 1
    print("Data validation passed: no errors found.")

    # ---- best science pair (was a blocking input() prompt) ----
    best_pair, best_count, counts = best_nat_sci_pair(data, config)
    print("\nScience pair popularity (students ranking both as their top two):")
    for (a, b), n in counts.items():
        print(f"  {a} + {b}: {n}")
    if config.constraints.best_pair is not None:
        used_pair = tuple(config.constraints.best_pair)
        print(f"Using configured pair for constraint 6: {used_pair[0]} + {used_pair[1]}")
    else:
        used_pair = best_pair
        print(
            f"Auto-selecting most popular pair for constraint 6: "
            f"{best_pair[0]} + {best_pair[1]} ({best_count} students)"
        )

    # ---- solve ----
    print(
        f"\nRunning {config.shuffles} shuffles "
        f"({config.time_limit_per_shuffle}s each, {config.num_workers} workers)..."
    )
    result = solve(data, config, progress_cb=_print_progress)

    if not result.found:
        print("\nNo solution found in any shuffle.")
        return 1

    best = result.best_data.sort_values(by=["Class", "Language", "NatSci1", "NatSci2"])
    print(f"\nBest total score: {result.best_score:.0f}")
    print("\nClass sizes:")
    print(best["Class"].value_counts().sort_index().to_string())

    # ---- save (unchanged output format) ----
    results_name = input("\nSave into dir name (inside results dir): ").strip()
    out_dir = os.path.join("results", results_name)
    os.makedirs(out_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(args.excel))[0]
    best.to_excel(os.path.join(out_dir, f"{base}_{results_name}.xlsx"), index=False)

    w = config.objective
    with open(os.path.join(out_dir, f"{base}_{results_name}_model_parameters.txt"), "w") as f:
        f.write("Objective function parameters:\n")
        f.write(f"lang_importance = {w.lang_importance}\n")
        f.write(f"lang_penalty = {w.lang_penalty}\n")
        f.write(f"nat_sci_1_importance = {w.nat_sci_1_importance}\n")
        f.write(f"nat_sci_2_importance = {w.nat_sci_2_importance}\n")
        f.write(f"nat_sci_penalty = {w.nat_sci_penalty}\n")
        f.write(f"stratification = {w.stratification}\n\n")
        f.write("Other information:\n")
        f.write(f"Max class size: {config.max_class_size}\n")
        f.write(f"Number of classes: {config.num_of_classes}\n")
        f.write(f"Number of shuffles: {config.shuffles}\n")
        f.write(f"Time limit per shuffle: {config.time_limit_per_shuffle}s\n")
        f.write(f"Search workers: {config.num_workers}\n")
        f.write(f"Best score: {result.best_score:.0f}\n")
        f.write(f"Best science pair: {used_pair[0]} and {used_pair[1]}\n")

    print(f"\nSaved results to {out_dir}/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
