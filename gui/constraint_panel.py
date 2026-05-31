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
    QComboBox,
    QFormLayout,
    QGroupBox,
    QLineEdit,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from studclassopti.core import SolverConfig
from studclassopti.core.constraints.base import Constraint, Parameter
from studclassopti.core.constraints.registry import REGISTRY

AUTO_PAIR = "Auto (most popular)"


def _science_pairs(config: SolverConfig) -> list[tuple[str, str]]:
    ns = config.nat_sci_classes
    return [(a, b) for i, a in enumerate(ns) for b in ns[i + 1 :]]


class ConstraintPanel(QWidget):
    def __init__(self, config: SolverConfig):
        super().__init__()
        self._config = config  # source of option lists (languages / sciences)
        self._boxes: list[tuple[Constraint, QGroupBox]] = []
        self._params: dict[tuple[str, str], QWidget] = {}
        self._best_pair: QComboBox | None = None

        layout = QVBoxLayout(self)
        for constraint in REGISTRY:
            layout.addWidget(self._build_group(constraint))
        layout.addStretch(1)

    def _build_group(self, c: Constraint) -> QGroupBox:
        box = QGroupBox(c.label)
        box.setCheckable(True)  # the checkbox is the constraint's enable flag
        box.setToolTip(c.description)
        form = QFormLayout(box)

        for p in c.parameters:
            widget = self._build_param_widget(p)
            self._params[(c.id, p.key)] = widget
            form.addRow(p.label, widget)

        if c.id == "c06":
            self._best_pair = QComboBox()
            self._best_pair.addItem(AUTO_PAIR)
            for a, b in _science_pairs(self._config):
                self._best_pair.addItem(f"{a} + {b}")
            form.addRow("Science pair", self._best_pair)

        self._boxes.append((c, box))
        return box

    def _build_param_widget(self, p: Parameter) -> QWidget:
        if p.kind == "int":
            w = QSpinBox()
            w.setMinimum(p.minimum if p.minimum is not None else 0)
            w.setMaximum(p.maximum if p.maximum is not None else 9999)
            return w
        if p.kind == "choice":
            w = QComboBox()
            w.addItems(getattr(self._config, p.choices_from))
            return w
        return QLineEdit()  # "str"

    # ---- config <-> widgets ----
    def load_from(self, config: SolverConfig) -> None:
        cons = config.constraints
        for c, box in self._boxes:
            box.setChecked(getattr(cons, c.enabled_field))
            for p in c.parameters:
                self._set_widget(self._params[(c.id, p.key)], getattr(cons, p.key))

        if self._best_pair is not None:
            if cons.best_pair is None:
                self._best_pair.setCurrentIndex(0)
            else:
                self._best_pair.setCurrentText(f"{cons.best_pair[0]} + {cons.best_pair[1]}")

    def apply_to(self, config: SolverConfig) -> None:
        cons = config.constraints
        for c, box in self._boxes:
            setattr(cons, c.enabled_field, box.isChecked())
            for p in c.parameters:
                setattr(cons, p.key, self._read_widget(self._params[(c.id, p.key)]))

        if self._best_pair is not None:
            if self._best_pair.currentIndex() == 0:
                cons.best_pair = None
            else:
                a, b = self._best_pair.currentText().split(" + ")
                cons.best_pair = (a, b)

    @staticmethod
    def _set_widget(w: QWidget, value) -> None:
        if isinstance(w, QSpinBox):
            w.setValue(int(value))
        elif isinstance(w, QComboBox):
            w.setCurrentText(str(value))
        else:
            w.setText(str(value))

    @staticmethod
    def _read_widget(w: QWidget):
        if isinstance(w, QSpinBox):
            return w.value()
        if isinstance(w, QComboBox):
            return w.currentText()
        return w.text()
