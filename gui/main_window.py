"""The application window: load a file, configure, run, watch, export.

Left pane is the configuration form (general settings, the registry-driven
constraint panel, and an advanced objective-weights section); right pane shows
progress and results. The solver runs on a ``SolverWorker`` thread so the UI
stays live and can cancel mid-run. Both old ``input()`` pauses are gone: data
errors disable Run, and the best science pair is auto-selected (overridable in
the constraint panel).
"""

from __future__ import annotations

import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from studclassopti.core import SolverConfig, load_excel, preprocess, validate
from studclassopti.core.constraints.registry import validate_constraints

from .constraint_panel import ConstraintPanel
from .results_view import ResultsView
from .solver_worker import SolverWorker


def _spin(minimum: int, maximum: int, value: int) -> QSpinBox:
    w = QSpinBox()
    w.setRange(minimum, maximum)
    w.setValue(value)
    return w


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("StudClassOpti")
        self.resize(1100, 760)

        self._data = None  # preprocessed DataFrame, or None until a file loads
        self._base_name = "students"
        self._worker: SolverWorker | None = None

        defaults = SolverConfig()
        self._build_general(defaults)
        self._build_objective(defaults)
        self.constraint_panel = ConstraintPanel(defaults)
        self.constraint_panel.load_from(defaults)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_left())
        splitter.addWidget(self._build_right())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

        self._refresh_run_enabled()

    # ---- left pane (configuration) ----
    def _build_general(self, c: SolverConfig) -> None:
        self.num_classes = _spin(1, 20, c.num_of_classes)
        self.max_class_size = _spin(1, 60, c.max_class_size)
        self.shuffles = _spin(1, 1000, c.shuffles)
        self.time_limit = _spin(1, 3600, c.time_limit_per_shuffle)
        self.num_workers = _spin(1, 64, c.num_workers)

        self.general_box = QGroupBox("General")
        form = QFormLayout(self.general_box)
        form.addRow("Number of classes", self.num_classes)
        form.addRow("Max class size", self.max_class_size)
        form.addRow("Shuffles", self.shuffles)
        form.addRow("Time limit per shuffle (s)", self.time_limit)
        form.addRow("Search workers", self.num_workers)

    def _build_objective(self, c: SolverConfig) -> None:
        w = c.objective
        self.lang_importance = _spin(0, 1000, w.lang_importance)
        self.lang_penalty = _spin(0, 1000, w.lang_penalty)
        self.ns1_importance = _spin(0, 1000, w.nat_sci_1_importance)
        self.ns2_importance = _spin(0, 1000, w.nat_sci_2_importance)
        self.ns_penalty = _spin(0, 1000, w.nat_sci_penalty)
        self.stratification = _spin(0, 10, w.stratification)

        # Collapsible: the group's checkbox shows/hides the (advanced) contents.
        self.objective_box = QGroupBox("Objective weights (advanced)")
        self.objective_box.setCheckable(True)
        self.objective_box.setChecked(False)
        content = QWidget()
        form = QFormLayout(content)
        form.addRow("Language importance", self.lang_importance)
        form.addRow("Language penalty", self.lang_penalty)
        form.addRow("Science 1 importance", self.ns1_importance)
        form.addRow("Science 2 importance", self.ns2_importance)
        form.addRow("Science penalty", self.ns_penalty)
        form.addRow("Stratification", self.stratification)
        outer = QVBoxLayout(self.objective_box)
        outer.addWidget(content)
        content.setVisible(False)
        self.objective_box.toggled.connect(content.setVisible)

    def _build_left(self) -> QWidget:
        # File row.
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText("No spreadsheet loaded")
        browse = QPushButton("Browse…")
        browse.clicked.connect(self._browse)
        file_row = QHBoxLayout()
        file_row.addWidget(self.file_edit, 1)
        file_row.addWidget(browse)

        # Config save/load row.
        save_cfg = QPushButton("Save config…")
        save_cfg.clicked.connect(self._save_config)
        load_cfg = QPushButton("Load config…")
        load_cfg.clicked.connect(self._load_config)
        cfg_row = QHBoxLayout()
        cfg_row.addWidget(save_cfg)
        cfg_row.addWidget(load_cfg)

        inner = QWidget()
        v = QVBoxLayout(inner)
        v.addLayout(file_row)
        v.addLayout(cfg_row)
        v.addWidget(self.general_box)
        v.addWidget(QLabel("Constraints"))
        v.addWidget(self.constraint_panel)
        v.addWidget(self.objective_box)
        v.addStretch(1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)
        scroll.setMinimumWidth(420)
        return scroll

    # ---- right pane (run + results) ----
    def _build_right(self) -> QWidget:
        self.run_btn = QPushButton("Run")
        self.run_btn.clicked.connect(self._run)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel)
        btn_row = QHBoxLayout()
        btn_row.addWidget(self.run_btn)
        btn_row.addWidget(self.cancel_btn)

        self.progress = QProgressBar()
        self.status = QLabel("")
        self.status.setWordWrap(True)
        self.results = ResultsView()

        right = QWidget()
        v = QVBoxLayout(right)
        v.addLayout(btn_row)
        v.addWidget(self.progress)
        v.addWidget(self.status)
        v.addWidget(self.results, 1)
        return right

    # ---- config <-> widgets ----
    def _build_config(self) -> SolverConfig:
        c = SolverConfig()
        c.num_of_classes = self.num_classes.value()
        c.max_class_size = self.max_class_size.value()
        c.shuffles = self.shuffles.value()
        c.time_limit_per_shuffle = self.time_limit.value()
        c.num_workers = self.num_workers.value()
        c.objective.lang_importance = self.lang_importance.value()
        c.objective.lang_penalty = self.lang_penalty.value()
        c.objective.nat_sci_1_importance = self.ns1_importance.value()
        c.objective.nat_sci_2_importance = self.ns2_importance.value()
        c.objective.nat_sci_penalty = self.ns_penalty.value()
        c.objective.stratification = self.stratification.value()
        self.constraint_panel.apply_to(c)
        return c

    def _apply_config_to_widgets(self, c: SolverConfig) -> None:
        self.num_classes.setValue(c.num_of_classes)
        self.max_class_size.setValue(c.max_class_size)
        self.shuffles.setValue(c.shuffles)
        self.time_limit.setValue(c.time_limit_per_shuffle)
        self.num_workers.setValue(c.num_workers)
        self.lang_importance.setValue(c.objective.lang_importance)
        self.lang_penalty.setValue(c.objective.lang_penalty)
        self.ns1_importance.setValue(c.objective.nat_sci_1_importance)
        self.ns2_importance.setValue(c.objective.nat_sci_2_importance)
        self.ns_penalty.setValue(c.objective.nat_sci_penalty)
        self.stratification.setValue(c.objective.stratification)
        self.constraint_panel.load_from(c)

    def _save_config(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Save config", "config.json", "JSON (*.json)")
        if path:
            self._build_config().to_json(path)

    def _load_config(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Load config", "", "JSON (*.json)")
        if path:
            self._apply_config_to_widgets(SolverConfig.from_json(path))
            self._revalidate()

    # ---- file load + validation ----
    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Open spreadsheet", "", "Excel (*.xlsx)")
        if not path:
            return
        try:
            self._data = preprocess(load_excel(path), self._build_config())
        except Exception as exc:
            self._data = None
            self.status.setText(f"Could not read file: {exc}")
            self._refresh_run_enabled()
            return
        self.file_edit.setText(path)
        self._base_name = os.path.splitext(os.path.basename(path))[0]
        self._revalidate()

    def _revalidate(self) -> list:
        """Validate the loaded data against current config; update status."""
        if self._data is None:
            self.status.setText("")
            self._refresh_run_enabled()
            return []
        config = self._build_config()
        errors = validate(self._data, config) + validate_constraints(self._data, config)
        if errors:
            shown = "\n".join(f"  • {e}" for e in errors[:10])
            more = f"\n  …and {len(errors) - 10} more" if len(errors) > 10 else ""
            self.status.setText(f"{len(errors)} data error(s) — fix before running:\n{shown}{more}")
        else:
            self.status.setText("Data OK. Ready to run.")
        self._refresh_run_enabled(errors)
        return errors

    def _refresh_run_enabled(self, errors: list | None = None) -> None:
        running = self._worker is not None
        ok = self._data is not None and not errors
        self.run_btn.setEnabled(ok and not running)

    # ---- run / cancel ----
    def _run(self) -> None:
        if self._worker is not None:
            return
        if self._revalidate():  # re-check; errors block the run
            return
        config = self._build_config()
        self.progress.setRange(0, config.shuffles)
        self.progress.setValue(0)
        self.status.setText("Running…")

        self._worker = SolverWorker(self._data, config)
        self._run_config = config
        self._worker.progress.connect(self._on_progress)
        self._worker.finished_ok.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self.run_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self._worker.start()

    def _cancel(self) -> None:
        if self._worker is not None:
            self._worker.cancel()
            self.cancel_btn.setEnabled(False)
            self.status.setText("Cancelling after the current shuffle…")

    def _on_progress(self, p) -> None:
        self.progress.setValue(p.shuffle)
        score = f"{p.score:.0f}" if p.score is not None else "no solution"
        self.status.setText(
            f"Shuffle {p.shuffle}/{p.total_shuffles}: {p.status}, "
            f"score={score}, best={p.best_score:.0f}"
        )

    def _on_finished(self, result) -> None:
        self.results.show_result(result, self._run_config, self._base_name)
        if result.found:
            self.status.setText(f"Done. Best score {result.best_score:.0f}.")
        else:
            self.status.setText("Done. No feasible solution found.")
        self._teardown_worker()

    def _on_failed(self, message: str) -> None:
        self.status.setText(f"Solver error: {message}")
        self._teardown_worker()

    def _teardown_worker(self) -> None:
        self._worker.wait()  # thread's run() has returned; finishes near-instantly
        self._worker = None
        self.cancel_btn.setEnabled(False)
        self._refresh_run_enabled(None if self._data is not None else [1])
