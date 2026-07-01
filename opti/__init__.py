"""StudClassOpti - student-to-class assignment optimizer.

Phase 1 package: the `core` subpackage holds all optimization logic with no
UI dependencies and no blocking prompts. It is driven entirely by a
``SolverConfig`` and communicates progress/results through return values and
callbacks, so it can be reused by a CLI, a GUI, or tests interchangeably.
"""
