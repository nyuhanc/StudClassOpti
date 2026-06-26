"""Auto-generated constraint controls.

Walks ``constraints.registry.REGISTRY`` and renders one checkable group box per
constraint (the checkbox is its ``enabled_field``) plus one widget per
``Parameter``. Adding a constraint to the registry makes it appear here with no
GUI code -- that is the whole point of the Phase 2 plugin design.

The only hand-written control is the best-pair override for constraint 6, whose
value (``best_pair``) isn't a scalar ``Parameter`` but an optional subject pair.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from studclassopti.core import SolverConfig
from studclassopti.core.constraints.base import Constraint, Parameter
from studclassopti.core.constraints.registry import REGISTRY

from . import i18n
from .widgets import info_toggle


def _science_pairs(config: SolverConfig) -> list[tuple[str, str]]:
    ns = config.nat_sci_classes
    return [(a, b) for i, a in enumerate(ns) for b in ns[i + 1 :]]


class ConstraintPanel(QWidget):
    def __init__(self, config: SolverConfig):
        super().__init__()
        self._config = config  # source of option lists (languages / sciences)
        self._checks: list[tuple[Constraint, QCheckBox]] = []
        self._params: dict[tuple[str, str], QWidget] = {}
        self._best_pair: QComboBox | None = None

        layout = QVBoxLayout(self)
        starred = False
        for constraint in REGISTRY:
            if constraint.id in i18n.HIDDEN_CONSTRAINTS:
                continue
            layout.addWidget(self._build_group(constraint))
            starred = starred or constraint.id in i18n.SINGLE_CHOICE_CONSTRAINTS

        if starred:  # explain the red "*" once, at the bottom
            footnote = QLabel(i18n.SINGLE_CHOICE_FOOTNOTE)
            footnote.setWordWrap(True)
            footnote.setStyleSheet("color: #666;")
            layout.addWidget(footnote)
        layout.addStretch(1)

    def _build_group(self, c: Constraint) -> QGroupBox:
        box = QGroupBox()
        outer = QVBoxLayout(box)

        # Header: enable checkbox (the label) + an info button, both always
        # enabled so the description can be read before turning the rule on.
        check = QCheckBox(i18n.constraint_label(c.id))
        check.setStyleSheet("font-weight: bold;")
        description = i18n.constraint_description(c.id)
        if c.id in i18n.SINGLE_CHOICE_CONSTRAINTS:
            description += i18n.SINGLE_CHOICE_MARK
        info, desc = info_toggle(description)
        header = QHBoxLayout()
        header.addWidget(check)
        header.addStretch(1)
        header.addWidget(info)
        outer.addLayout(header)

        # Collapsible description (hidden until the info button is toggled).
        outer.addWidget(desc)

        # Parameters live in their own widget so they grey out when the rule is
        # off, without disabling the header (checkbox + info).
        params = QWidget()
        form = QFormLayout(params)
        form.setContentsMargins(0, 0, 0, 0)
        for p in c.parameters:
            widget = self._build_param_widget(p)
            self._params[(c.id, p.key)] = widget
            form.addRow(i18n.param_label(c.id, p.key, p.label), widget)

        if c.id == "c06":
            self._best_pair = QComboBox()
            self._best_pair.addItem(i18n.AUTO_PAIR, None)
            for a, b in _science_pairs(self._config):
                self._best_pair.addItem(f"{i18n.subject_to_sl(a)} + {i18n.subject_to_sl(b)}", (a, b))
            form.addRow(i18n.SCIENCE_PAIR, self._best_pair)

        outer.addWidget(params)
        check.toggled.connect(params.setEnabled)
        params.setEnabled(check.isChecked())

        self._checks.append((c, check))
        return box

    def _build_param_widget(self, p: Parameter) -> QWidget:
        if p.kind == "int":
            w = QSpinBox()
            w.setMinimum(p.minimum if p.minimum is not None else 0)
            w.setMaximum(p.maximum if p.maximum is not None else 9999)
            return w
        if p.kind == "choice":
            w = QComboBox()
            for opt in getattr(self._config, p.choices_from):
                w.addItem(i18n.subject_to_sl(opt), opt)  # show Slovene, store English
            return w
        return QLineEdit()  # "str"

    # ---- config <-> widgets ----
    def load_from(self, config: SolverConfig) -> None:
        cons = config.constraints
        for c, check in self._checks:
            check.setChecked(getattr(cons, c.enabled_field))
            for p in c.parameters:
                value = getattr(cons, p.key)
                if (c.id, p.key) == ("c05", "schoolmate_column"):
                    value = i18n.col_to_sl(value)  # show the Slovene column name
                self._set_widget(self._params[(c.id, p.key)], value)

        if self._best_pair is not None:
            if cons.best_pair is None:
                self._best_pair.setCurrentIndex(0)
            else:
                idx = self._best_pair.findData(tuple(cons.best_pair))
                self._best_pair.setCurrentIndex(idx if idx >= 0 else 0)

    def apply_to(self, config: SolverConfig) -> None:
        cons = config.constraints
        for c, check in self._checks:
            setattr(cons, c.enabled_field, check.isChecked())
            for p in c.parameters:
                value = self._read_widget(self._params[(c.id, p.key)])
                if (c.id, p.key) == ("c05", "schoolmate_column"):
                    value = i18n.col_to_en(value)  # store the internal English name
                setattr(cons, p.key, value)

        if self._best_pair is not None:
            cons.best_pair = self._best_pair.currentData()  # None or (a, b)

    @staticmethod
    def _set_widget(w: QWidget, value) -> None:
        if isinstance(w, QSpinBox):
            w.setValue(int(value))
        elif isinstance(w, QComboBox):
            idx = w.findData(value)  # match the English value stored as item data
            w.setCurrentIndex(idx if idx >= 0 else 0)
        else:
            w.setText(str(value))

    @staticmethod
    def _read_widget(w: QWidget):
        if isinstance(w, QSpinBox):
            return w.value()
        if isinstance(w, QComboBox):
            return w.currentData()
        return w.text()
