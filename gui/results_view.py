"""Show a solved assignment and export it.

Displays the best assignment in a table plus a class-size summary, and writes
the same two files the old CLI produced: the ``.xlsx`` and a
``_model_parameters.txt`` sidecar describing the run.
"""

from __future__ import annotations

import os

import pandas as pd
from PySide6.QtWidgets import (
    QFileDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from studclassopti.core import SolverConfig
from studclassopti.core.solver import SolverResult

SORT_COLS = ["Class", "Language", "NatSci1", "NatSci2"]


class ResultsView(QWidget):
    def __init__(self):
        super().__init__()
        self._result: SolverResult | None = None
        self._config: SolverConfig | None = None
        self._base_name = "students"

        layout = QVBoxLayout(self)
        self.summary = QLabel("No results yet. Load a file and run.")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table, 1)

        self.export_btn = QPushButton("Export to Excel…")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self._export)
        layout.addWidget(self.export_btn)

    def show_result(self, result: SolverResult, config: SolverConfig, base_name: str) -> None:
        self._result = result
        self._config = config
        self._base_name = base_name

        if not result.found:
            self.summary.setText("No feasible solution found in any shuffle.")
            self.table.clear()
            self.table.setRowCount(0)
            self.export_btn.setEnabled(False)
            return

        df = result.best_data.sort_values(by=SORT_COLS).reset_index(drop=True)
        sizes = df["Class"].value_counts().sort_index()
        sizes_txt = ", ".join(f"class {c}: {n}" for c, n in sizes.items())
        cancelled = " (cancelled early)" if result.cancelled else ""
        self.summary.setText(
            f"Best score: {result.best_score:.0f}{cancelled}  |  "
            f"pair: {result.best_pair[0]} + {result.best_pair[1]}  |  {sizes_txt}"
        )
        self._fill_table(df)
        self.export_btn.setEnabled(True)

    def _fill_table(self, df: pd.DataFrame) -> None:
        self.table.clear()
        self.table.setColumnCount(len(df.columns))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        for r in range(len(df)):
            for col in range(len(df.columns)):
                self.table.setItem(r, col, QTableWidgetItem(str(df.iat[r, col])))
        self.table.resizeColumnsToContents()

    def _export(self) -> None:
        if self._result is None or not self._result.found:
            return
        default = f"{self._base_name}_results.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Export results", default, "Excel (*.xlsx)")
        if not path:
            return
        df = self._result.best_data.sort_values(by=SORT_COLS).reset_index(drop=True)
        try:
            df.to_excel(path, index=False)
            self._write_parameters(path)
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))
            return
        QMessageBox.information(self, "Exported", f"Saved:\n{path}")

    def _write_parameters(self, xlsx_path: str) -> None:
        config = self._config
        w = config.objective
        txt_path = os.path.splitext(xlsx_path)[0] + "_model_parameters.txt"
        with open(txt_path, "w") as f:
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
            f.write(f"Best score: {self._result.best_score:.0f}\n")
            f.write(f"Best science pair: {self._result.best_pair[0]} and {self._result.best_pair[1]}\n")
