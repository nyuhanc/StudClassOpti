"""Small shared GUI helpers."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QToolButton


def info_toggle(description: str) -> tuple[QToolButton, QLabel]:
    """An ⓘ button paired with a collapsible description label, wired together.

    The label starts hidden; clicking the button shows/hides it. Both are
    returned so the caller can place them in its own layout.
    """
    button = QToolButton()
    button.setText("ⓘ")
    button.setAutoRaise(True)
    button.setCheckable(True)
    button.setCursor(Qt.PointingHandCursor)
    button.setToolTip("Pokaži/skrij opis")

    label = QLabel(description)
    label.setWordWrap(True)
    label.setStyleSheet("color: #666;")
    label.setVisible(False)

    button.toggled.connect(label.setVisible)
    return button, label
