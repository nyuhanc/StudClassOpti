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

from opti.core import SolverConfig
from opti.core.constraints.base import Parameter
from opti.core.constraints.registry import REGISTRY
from opti.core.solver import SolverResult

from . import i18n

SORT_COLS = ["Class", "Language", "NatSci1", "NatSci2"]


class ResultsView(QWidget):
    def __init__(self):
        super().__init__()
        self._result: SolverResult | None = None
        self._config: SolverConfig | None = None
        self._base_name = "students"

        layout = QVBoxLayout(self)
        self.summary = QLabel(i18n.NO_RESULTS)
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table, 1)

        self.export_btn = QPushButton(i18n.EXPORT_BTN)
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self._export)
        layout.addWidget(self.export_btn)

    def show_result(self, result: SolverResult, config: SolverConfig, base_name: str) -> None:
        self._result = result
        self._config = config
        self._base_name = base_name

        if not result.found:
            self.summary.setText(i18n.NO_FEASIBLE)
            self.table.clear()
            self.table.setRowCount(0)
            self.export_btn.setEnabled(False)
            return

        df = result.best_data.sort_values(by=SORT_COLS).reset_index(drop=True)
        sizes = df["Class"].value_counts().sort_index()
        sizes_txt = ", ".join(f"razred {c}: {n}" for c, n in sizes.items())
        cancelled = i18n.CANCELLED_EARLY if result.cancelled else ""
        self.summary.setText(
            i18n.summary_text(result.best_score, cancelled, result.best_pair, sizes_txt)
        )
        self._fill_table(i18n.localize_result_df(df))
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
        default = f"{self._base_name}_rezultati.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, i18n.EXPORT_TITLE, default, "Excel (*.xlsx)")
        if not path:
            return
        df = self._result.best_data.sort_values(by=SORT_COLS).reset_index(drop=True)
        try:
            i18n.localize_result_df(df).to_excel(path, index=False)
            self._write_parameters(path)
        except Exception as exc:
            QMessageBox.critical(self, i18n.EXPORT_FAILED, str(exc))
            return
        QMessageBox.information(self, i18n.EXPORTED, i18n.exported_text(path))

    def _write_parameters(self, xlsx_path: str) -> None:
        config = self._config
        w = config.objective
        pair = self._result.best_pair
        txt_path = os.path.splitext(xlsx_path)[0] + "_parametri_modela.txt"
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(i18n.PARAM_OBJECTIVE_HEADER + "\n")
            for label, field in i18n.PARAM_OBJECTIVE.items():
                f.write(f"{label} = {getattr(w, field)}\n")
            f.write("\n")
            f.write(i18n.PARAM_CONSTRAINTS_HEADER + "\n")
            for c in REGISTRY:
                if c.id in i18n.HIDDEN_CONSTRAINTS:
                    continue
                enabled = getattr(config.constraints, c.enabled_field)
                f.write(f"{i18n.constraint_label(c.id)}: {'da' if enabled else 'ne'}\n")
                for p in c.parameters:
                    label = i18n.param_label(c.id, p.key, p.label)
                    value = self._param_value_sl(p, getattr(config.constraints, p.key))
                    f.write(f"  {label} = {value}\n")
                if c.id == "c06":
                    bp = config.constraints.best_pair
                    pair_txt = (
                        i18n.AUTO_PAIR if bp is None
                        else f"{i18n.subject_to_sl(bp[0])} + {i18n.subject_to_sl(bp[1])}"
                    )
                    f.write(f"  {i18n.SCIENCE_PAIR} = {pair_txt}\n")
            f.write("\n")
            f.write(i18n.PARAM_OTHER_HEADER + "\n")
            f.write(f"Največja velikost razreda: {config.max_class_size}\n")
            f.write(f"Število razredov: {config.num_of_classes}\n")
            f.write(f"Število poskusov: {config.shuffles}\n")
            f.write(f"Časovna omejitev na poskus: {config.time_limit_per_shuffle}s\n")
            f.write(f"Število niti za iskanje: {config.num_workers}\n")
            f.write(f"Najboljši rezultat: {self._result.best_score:.0f}\n")
            f.write(
                f"Najboljši naravoslovni par: "
                f"{i18n.subject_to_sl(pair[0])} in {i18n.subject_to_sl(pair[1])}\n"
            )

    @staticmethod
    def _param_value_sl(p: Parameter, value):
        """A constraint parameter's value in Slovene for the report."""
        if p.kind == "choice":  # a language / science name
            return i18n.subject_to_sl(value)
        if p.key == "schoolmate_column":  # a spreadsheet column name
            return i18n.col_to_sl(value)
        return value  # int / plain str
