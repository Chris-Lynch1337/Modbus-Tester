"""Batch / Ramp tab UI builder."""
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


def build_batch_tab(win) -> QWidget:
    w = QWidget(); v = QVBoxLayout(w); v.setSpacing(10)
    win.batch_status_label = QLabel("")
    win.batch_status_label.setStyleSheet("color: #f0a500; font-size: 11px; font-weight: bold;")
    win.batch_status_label.setVisible(False)
    v.addWidget(win.batch_status_label)

    bg = QGroupBox("Batch Write  —  comma-separated UINT16 values"); bv = QVBoxLayout(bg)
    rl = QHBoxLayout(); rl.addWidget(QLabel("Reg Addr (6x):"))
    win.batch_reg_addr = QSpinBox(); win.batch_reg_addr.setRange(400001, 499999); win.batch_reg_addr.setValue(400001); win.batch_reg_addr.setFixedWidth(100)
    rl.addWidget(win.batch_reg_addr); rl.addWidget(QLabel("Values (0–65535):"))
    win.batch_reg_vals = QLineEdit("100,200,300,400"); rl.addWidget(win.batch_reg_vals)
    r_write = QPushButton("WRITE REGS"); r_write.setObjectName("write_btn"); r_write.clicked.connect(win.batch_write_registers); rl.addWidget(r_write)
    win.batch_write_btn = r_write
    bv.addLayout(rl); v.addWidget(bg)

    rg = QGroupBox("Ramp Generator"); rv = QVBoxLayout(rg)
    r1 = QHBoxLayout()
    r1.addWidget(QLabel("Mode:"))
    win.ramp_mode = QComboBox(); win.ramp_mode.setFixedWidth(140)
    win.ramp_mode.addItems(["Linear (+step)", "Powers of 2 (×2)"])
    win.ramp_mode.setToolTip("Linear: adds step value each tick.\nPowers of 2: doubles each tick (1,2,4,8…). Use for bit testing.")
    win.ramp_mode.currentIndexChanged.connect(win._on_ramp_mode_change); r1.addWidget(win.ramp_mode); r1.addSpacing(12)
    r1.addWidget(QLabel("Register (6x):"))
    win.ramp_addr = QSpinBox(); win.ramp_addr.setRange(400001, 499999); win.ramp_addr.setValue(400001); win.ramp_addr.setFixedWidth(100); r1.addWidget(win.ramp_addr)
    r1.addWidget(QLabel("Start:"))
    win.ramp_start = QComboBox(); win.ramp_start.setFixedWidth(160)
    for b in range(1, 33):
        win.ramp_start.addItem(f"Bit {b} = {2**(b-1)}", 2**(b-1))
    win.ramp_start.setCurrentIndex(0); r1.addWidget(win.ramp_start)
    r1.addWidget(QLabel("End:"))
    win.ramp_end = QComboBox(); win.ramp_end.setFixedWidth(160)
    for b in range(1, 33):
        win.ramp_end.addItem(f"Bit {b} = {2**(b-1)}", 2**(b-1))
    win.ramp_end.setCurrentIndex(5); r1.addWidget(win.ramp_end)
    r1.addWidget(QLabel("Delay (ms):"))
    win.ramp_delay = QSpinBox(); win.ramp_delay.setRange(50, 5000); win.ramp_delay.setValue(200); win.ramp_delay.setFixedWidth(80); r1.addWidget(win.ramp_delay)
    r1.addSpacing(12); r1.addWidget(QLabel("Word Order:"))
    win.ramp_dtype = QComboBox(); win.ramp_dtype.setFixedWidth(160)
    win.ramp_dtype.addItems(["UINT16 (single reg)", "UINT32 Lo/Hi (CD AB)", "UINT32 Hi/Lo (AB CD)"])
    win.ramp_dtype.setCurrentIndex(1)
    win.ramp_dtype.setToolTip("UINT32 Lo/Hi (CD AB) matches Productivity Suite INT32/DINT tag format")
    r1.addWidget(win.ramp_dtype); r1.addStretch(); rv.addLayout(r1)

    r2 = QHBoxLayout()
    win.ramp_step_lbl = QLabel("Step:")
    win.ramp_step = QSpinBox(); win.ramp_step.setRange(1, 65535); win.ramp_step.setValue(1); win.ramp_step.setFixedWidth(80)
    r2.addWidget(win.ramp_step_lbl); r2.addWidget(win.ramp_step); r2.addSpacing(12)
    win.ramp_btn = QPushButton("▶  START RAMP"); win.ramp_btn.setObjectName("write_btn"); win.ramp_btn.setMinimumWidth(160)
    win.ramp_btn.clicked.connect(win.toggle_ramp); r2.addWidget(win.ramp_btn); r2.addSpacing(12)
    win.ramp_preview_lbl = QLabel("")
    win.ramp_preview_lbl.setStyleSheet(f"color: {ACCENT_AMBER}; font-size: 10px; font-family: Consolas;")
    r2.addWidget(win.ramp_preview_lbl); r2.addStretch(); rv.addLayout(r2)
    v.addWidget(rg); v.addStretch()

    for sig in [win.ramp_start.currentIndexChanged, win.ramp_step.valueChanged,
                win.ramp_end.currentIndexChanged,   win.ramp_mode.currentIndexChanged,
                win.ramp_dtype.currentIndexChanged]:
        sig.connect(win._update_ramp_preview)
    win.ramp_timer = QTimer(win); win.ramp_timer.timeout.connect(win._ramp_step)
    return w
