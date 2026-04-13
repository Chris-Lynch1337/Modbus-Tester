"""
AboutDialog and ColorButton widget.
"""
from __future__ import annotations

from datetime import datetime

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QColorDialog, QDialog, QLabel, QPushButton, QVBoxLayout,
)

from ..constants import (
    ACCENT_BLUE, APP_COMPANY, APP_CONTACT, APP_NAME, APP_VERSION,
    STYLESHEET, TEXT_DIM, TEXT_PRIMARY,
)


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"About {APP_NAME}")
        self.setFixedSize(400, 260)
        self.setStyleSheet(STYLESHEET)
        v = QVBoxLayout(self)
        v.setSpacing(12)
        v.setContentsMargins(24, 24, 24, 24)

        title = QLabel(f"▣  {APP_NAME}")
        title.setFont(QFont("Consolas", 14, QFont.Bold))
        title.setStyleSheet(f"color: {ACCENT_BLUE}; letter-spacing: 2px;")
        v.addWidget(title)

        for text, dim in [
            (f"Version {APP_VERSION}", False),
            (f"© {datetime.now().year} {APP_COMPANY}", True),
            ("", True),
            ("Designed for AutomationDirect Productivity Suite", True),
            ("Modbus TCP holding register testing utility", True),
            ("", True),
            (f"Support: {APP_CONTACT}", True),
        ]:
            lbl = QLabel(text)
            lbl.setStyleSheet(f"color: {TEXT_DIM if dim else TEXT_PRIMARY}; font-size: 11px;")
            v.addWidget(lbl)

        v.addStretch()
        btn = QPushButton("Close")
        btn.clicked.connect(self.accept)
        v.addWidget(btn)



# ─── Color picker button ──────────────────────────────────────────────────────
class ColorButton(QPushButton):
    color_changed = pyqtSignal(str)

    def __init__(self, color: str = "#ffffff", parent=None):
        super().__init__(parent)
        self._color = color
        self.setFixedSize(48, 28)
        self._update_style()
        self.clicked.connect(self._pick)

    def _update_style(self):
        self.setStyleSheet(
            f"background-color: {self._color}; border: 2px solid #363a4a; border-radius: 3px;"
        )

    def _pick(self):
        c = QColorDialog.getColor(QColor(self._color), self, "Pick Color")
        if c.isValid():
            self._color = c.name()
            self._update_style()
            self.color_changed.emit(self._color)

    def color(self) -> str:
        return self._color

    def set_color(self, c: str):
        self._color = c
        self._update_style()

