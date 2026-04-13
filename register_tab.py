"""Register tab UI builder."""
from __future__ import annotations
from datetime import datetime
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QCheckBox, QComboBox, QFrame, QGridLayout, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QProgressBar,
    QPushButton, QScrollArea, QSpinBox, QVBoxLayout, QWidget,
)
from ..constants import (
    ACCENT_AMBER, ACCENT_BLUE, ACCENT_GREEN, ACCENT_RED,
    BG_DARK, BG_FIELD, BG_MID, BG_PANEL, BORDER,
    DTYPE_OPTIONS, DTYPE_INFO, DTYPE_TOOLTIPS,
    MAX_SWEEP_TAGS, TEXT_DIM, TEXT_PRIMARY,
)
from ..datatypes import (
    OperationRequest, ValidationError, pack_value, preview_pack,
)


def build_register_tab(win) -> QWidget:
    w = QWidget(); v = QVBoxLayout(w); v.setSpacing(10)
    win.reg_status_label = QLabel("")
    win.reg_status_label.setStyleSheet("color: #f0a500; font-size: 11px; font-weight: bold;")
    win.reg_status_label.setVisible(False)
    v.addWidget(win.reg_status_label)

    sg = QGroupBox("Single Register Write  (FC06 for 16-bit / FC16 for 32-bit)")
    sv = QVBoxLayout(sg); sv.setSpacing(8)

    r1 = QHBoxLayout()
    r1.addWidget(QLabel("Address (6x):"))
    win.reg_addr = QSpinBox(); win.reg_addr.setRange(400001, 499999); win.reg_addr.setValue(400001); win.reg_addr.setFixedWidth(100)
    win.reg_addr.setToolTip("Productivity Suite 6x register address (400001+)")
    r1.addWidget(win.reg_addr); r1.addSpacing(20)
    r1.addWidget(QLabel("Data Type:"))
    win.reg_dtype = QComboBox(); win.reg_dtype.setFixedWidth(280); win.reg_dtype.addItems(DTYPE_OPTIONS)
    win.reg_dtype.setToolTip("Select the data type matching your PLC tag definition")
    r1.addWidget(win.reg_dtype); r1.addStretch(); sv.addLayout(r1)

    r2 = QHBoxLayout()
    r2.addWidget(QLabel("Value:"))
    win.reg_value_edit = QLineEdit("0"); win.reg_value_edit.setFixedWidth(160)
    win.reg_value_edit.setPlaceholderText("integer or float")
    win.reg_value_edit.setToolTip("Enter value to write. Supports decimal or hex (0x prefix).")
    r2.addWidget(win.reg_value_edit)
    win.reg_preview = QLabel("")
    win.reg_preview.setStyleSheet(f"color: {ACCENT_AMBER}; font-size: 11px; font-family: Consolas;")
    win.reg_preview.setMinimumWidth(280); r2.addWidget(win.reg_preview); r2.addStretch()
    write_btn = QPushButton("WRITE  [Enter]"); write_btn.setObjectName("write_btn")
    write_btn.clicked.connect(win.write_single_register); r2.addWidget(write_btn); sv.addLayout(r2)
    win.reg_write_btn = write_btn

    win.reg_type_info = QLabel("")
    win.reg_type_info.setStyleSheet(f"color: {TEXT_DIM}; font-size: 10px;"); sv.addWidget(win.reg_type_info)

    r4 = QHBoxLayout()
    r4.addWidget(QLabel("Read Count:"))
    win.reg_read_count = QSpinBox(); win.reg_read_count.setRange(1, 125); win.reg_read_count.setValue(2); win.reg_read_count.setFixedWidth(70)
    win.reg_read_count.setToolTip("Number of raw 16-bit registers to read")
    r4.addWidget(win.reg_read_count)
    read_btn = QPushButton("READ"); read_btn.setObjectName("read_btn"); read_btn.clicked.connect(win.read_registers); r4.addWidget(read_btn)
    win.reg_read_btn = read_btn
    r4.addSpacing(10); r4.addWidget(QLabel("Decode As:"))
    win.reg_read_dtype = QComboBox(); win.reg_read_dtype.setFixedWidth(280); win.reg_read_dtype.addItems(DTYPE_OPTIONS)
    win.reg_read_dtype.setCurrentText("INT32   –  Lo word first  (CD AB)")
    win.reg_read_dtype.setToolTip("How to interpret the raw register bytes when displaying the result")
    r4.addWidget(win.reg_read_dtype); r4.addSpacing(10); r4.addWidget(QLabel("Value:"))
    win.reg_read_display = QLineEdit("")
    win.reg_read_display.setReadOnly(True); win.reg_read_display.setFixedWidth(180)
    win.reg_read_display.setPlaceholderText("decoded value")
    win.reg_read_display.setStyleSheet("background-color: #111318; color: #f0a500; font-size: 13px; font-weight: bold; border: 1px solid #f0a500; border-radius: 3px; padding: 4px 8px;")
    r4.addWidget(win.reg_read_display); r4.addStretch(); sv.addLayout(r4)
    v.addWidget(sg)

    mg = QGroupBox("Register Quick Set  (8 × UINT16 registers, configurable base)")
    mgv = QVBoxLayout(mg)
    ba_row = QHBoxLayout(); ba_row.addWidget(QLabel("Base Address (6x):"))
    win.reg_base_addr = QSpinBox(); win.reg_base_addr.setRange(400001, 499992); win.reg_base_addr.setValue(400001); win.reg_base_addr.setFixedWidth(100)
    ba_row.addWidget(win.reg_base_addr); ba_row.addStretch()
    write_all_btn = QPushButton("WRITE ALL"); write_all_btn.setObjectName("write_btn")
    write_all_btn.clicked.connect(win.write_all_quick_registers); ba_row.addWidget(write_all_btn); mgv.addLayout(ba_row)
    win.reg_write_all_btn = write_all_btn
    grid = QGridLayout(); grid.setSpacing(8); win.reg_quick_fields: List[QSpinBox] = []
    for i in range(8):
        lbl = QLabel(f"+{i}"); lbl.setStyleSheet(f"color: {TEXT_DIM}; font-size: 10px;"); lbl.setAlignment(Qt.AlignCenter)
        sp = QSpinBox(); sp.setRange(0, 65535); sp.setFixedWidth(100); sp.setValue(0)
        grid.addWidget(lbl, 0, i); grid.addWidget(sp, 1, i); win.reg_quick_fields.append(sp)
    mgv.addLayout(grid); v.addWidget(mg); v.addStretch()

    win.reg_value_edit.textChanged.connect(win._update_reg_preview)
    win.reg_dtype.currentIndexChanged.connect(win._update_reg_preview)
    win.reg_dtype.currentIndexChanged.connect(win._update_type_info)
    win.reg_dtype.currentIndexChanged.connect(win._update_dtype_tooltip)
    win._update_type_info(); win._update_reg_preview(); win._update_dtype_tooltip()
    return w
