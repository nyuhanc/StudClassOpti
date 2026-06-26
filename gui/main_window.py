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

from . import i18n
from .constraint_panel import ConstraintPanel
from .results_view import ResultsView
from .solver_worker import SolverWorker
from .widgets import info_toggle


def _spin(minimum: int, maximum: int, value: int) -> QSpinBox:
    w = QSpinBox()
    w.setRange(minimum, maximum)
    w.setValue(value)
    return w


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(i18n.WINDOW_TITLE)
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

        self.general_box = QGroupBox(i18n.GENERAL)
        form = QFormLayout(self.general_box)
        self._add_general_row(form, i18n.NUM_CLASSES, i18n.NUM_CLASSES_DESC, self.num_classes)
        self._add_general_row(form, i18n.MAX_CLASS_SIZE, i18n.MAX_CLASS_SIZE_DESC, self.max_class_size)
        self._add_general_row(form, i18n.SHUFFLES, i18n.SHUFFLES_DESC, self.shuffles)
        self._add_general_row(form, i18n.TIME_LIMIT, i18n.TIME_LIMIT_DESC, self.time_limit)
        self._add_general_row(form, i18n.SEARCH_WORKERS, i18n.SEARCH_WORKERS_DESC, self.num_workers)

    @staticmethod
    def _add_general_row(form: QFormLayout, label: str, description: str, field: QWidget) -> None:
        """A form row whose label carries an ⓘ that expands a description below."""
        info, desc = info_toggle(description)
        cell = QWidget()
        h = QHBoxLayout(cell)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(QLabel(label))
        h.addWidget(info)
        h.addStretch(1)
        form.addRow(cell, field)
        form.addRow(desc)  # spans both columns, just under the field

    def _build_objective(self, c: SolverConfig) -> None:
        w = c.objective
        self.lang_importance = _spin(0, 1000, w.lang_importance)
        self.lang_penalty = _spin(0, 1000, w.lang_penalty)
        self.ns1_importance = _spin(0, 1000, w.nat_sci_1_importance)
        self.ns2_importance = _spin(0, 1000, w.nat_sci_2_importance)
        self.ns_penalty = _spin(0, 1000, w.nat_sci_penalty)
        self.stratification = _spin(0, 10, w.stratification)

        # Collapsible: the group's checkbox shows/hides the (advanced) contents.
        self.objective_box = QGroupBox(i18n.OBJECTIVE_BOX)
        self.objective_box.setCheckable(True)
        self.objective_box.setChecked(False)
        content = QWidget()
        form = QFormLayout(content)
        self._add_general_row(form, i18n.LANG_IMPORTANCE, i18n.LANG_IMPORTANCE_DESC, self.lang_importance)
        self._add_general_row(form, i18n.LANG_PENALTY, i18n.LANG_PENALTY_DESC, self.lang_penalty)
        self._add_general_row(form, i18n.NS1_IMPORTANCE, i18n.NS1_IMPORTANCE_DESC, self.ns1_importance)
        self._add_general_row(form, i18n.NS2_IMPORTANCE, i18n.NS2_IMPORTANCE_DESC, self.ns2_importance)
        self._add_general_row(form, i18n.NS_PENALTY, i18n.NS_PENALTY_DESC, self.ns_penalty)
        self._add_general_row(form, i18n.STRATIFICATION, i18n.STRATIFICATION_DESC, self.stratification)
        outer = QVBoxLayout(self.objective_box)
        outer.addWidget(content)
        content.setVisible(False)
        self.objective_box.toggled.connect(content.setVisible)

    def _build_left(self) -> QWidget:
        # File row.
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText(i18n.NO_FILE)
        browse = QPushButton(i18n.BROWSE)
        browse.clicked.connect(self._browse)
        file_row = QHBoxLayout()
        file_row.addWidget(self.file_edit, 1)
        file_row.addWidget(browse)

        # Config save/load row.
        save_cfg = QPushButton(i18n.SAVE_CONFIG_BTN)
        save_cfg.clicked.connect(self._save_config)
        load_cfg = QPushButton(i18n.LOAD_CONFIG_BTN)
        load_cfg.clicked.connect(self._load_config)
        cfg_row = QHBoxLayout()
        cfg_row.addWidget(save_cfg)
        cfg_row.addWidget(load_cfg)

        inner = QWidget()
        v = QVBoxLayout(inner)
        v.addLayout(file_row)
        v.addLayout(cfg_row)
        v.addWidget(self.general_box)
        v.addWidget(QLabel(i18n.CONSTRAINTS))
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
        self.run_btn = QPushButton(i18n.RUN)
        self.run_btn.clicked.connect(self._run)
        self.cancel_btn = QPushButton(i18n.CANCEL)
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
        path, _ = QFileDialog.getSaveFileName(self, i18n.SAVE_CONFIG_TITLE, "config.json", "JSON (*.json)")
        if path:
            self._build_config().to_json(path)

    def _load_config(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, i18n.LOAD_CONFIG_TITLE, "", "JSON (*.json)")
        if path:
            self._apply_config_to_widgets(SolverConfig.from_json(path))
            self._revalidate()

    # ---- file load + validation ----
    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, i18n.OPEN_SPREADSHEET_TITLE, "", "Excel (*.xlsx)")
        if not path:
            return
        try:
            raw = i18n.excel_to_internal(load_excel(path))  # Slovene headers -> internal names
            self._data = preprocess(raw, self._build_config())
        except Exception as exc:
            self._data = None
            self.status.setText(i18n.could_not_read(exc))
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
            shown = "\n".join(f"  • {i18n.localize_errors(str(e))}" for e in errors[:10])
            more = i18n.and_more(len(errors) - 10) if len(errors) > 10 else ""
            self.status.setText(f"{i18n.data_errors_header(len(errors))}\n{shown}{more}")
        else:
            self.status.setText(i18n.DATA_OK)
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
        self.status.setText(i18n.RUNNING)

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
            self.status.setText(i18n.CANCELLING)

    def _on_progress(self, p) -> None:
        self.progress.setValue(p.shuffle)
        self.status.setText(i18n.progress_text(p))

    def _on_finished(self, result) -> None:
        self.results.show_result(result, self._run_config, self._base_name)
        if result.found:
            self.status.setText(i18n.done_best(result.best_score))
        else:
            self.status.setText(i18n.DONE_INFEASIBLE)
        self._teardown_worker()

    def _on_failed(self, message: str) -> None:
        self.status.setText(i18n.solver_error(message))
        self._teardown_worker()

    def _teardown_worker(self) -> None:
        self._worker.wait()  # thread's run() has returned; finishes near-instantly
        self._worker = None
        self.cancel_btn.setEnabled(False)
        self._refresh_run_enabled(None if self._data is not None else [1])
