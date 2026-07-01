"""Run the solver off the GUI thread.

``core.solver.solve`` is a long, CPU-bound call. Running it on a ``QThread``
keeps the window responsive; its ``progress_cb`` and ``should_cancel`` hooks map
onto a Qt signal and a thread-safe flag.

Cancellation is cooperative and only polled between shuffles, so a click on
Cancel takes effect after the current shuffle's time limit at the latest.
"""

from __future__ import annotations

import pandas as pd
from PySide6.QtCore import QThread, Signal

from opti.core import SolverConfig, solve
from opti.core.solver import SolveProgress, SolverResult


class SolverWorker(QThread):
    progress = Signal(object)  # SolveProgress
    finished_ok = Signal(object)  # SolverResult
    failed = Signal(str)

    def __init__(self, data: pd.DataFrame, config: SolverConfig):
        super().__init__()
        self._data = data
        self._config = config
        self._cancel = False

    def cancel(self) -> None:
        """Request a stop; honoured between shuffles."""
        self._cancel = True

    def run(self) -> None:
        try:
            result: SolverResult = solve(
                self._data,
                self._config,
                progress_cb=self.progress.emit,
                should_cancel=lambda: self._cancel,
            )
        except Exception as exc:  # surface any solver/data error to the UI
            self.failed.emit(str(exc))
            return
        self.finished_ok.emit(result)
