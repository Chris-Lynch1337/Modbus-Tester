"""Tag Sweep tab UI builder."""
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


def build_sweep_tab(win) -> QWidget:
    w = QWidget(); v = QVBoxLayout(w); v.setSpacing(14)
    v.setContentsMargins(8, 8, 8, 8)
    win.sweep_status_label = QLabel("")
    win.sweep_status_label.setStyleSheet("color: #f0a500; font-size: 11px; font-weight: bold;")
    win.sweep_status_label.setVisible(False)
    v.addWidget(win.sweep_status_label)

    cfg = QGroupBox("Tag Sweep Configuration"); cv = QVBoxLayout(cfg); cv.setSpacing(16)
    cv.setContentsMargins(16, 20, 16, 16)

    # ── Row 1: Address settings ──────────────────────────────────────────
    r1 = QHBoxLayout(); r1.setSpacing(8)
    r1.addWidget(QLabel("Start Address (6x):"))
    win.sweep_start_addr = QSpinBox()
    win.sweep_start_addr.setRange(400001, 499999); win.sweep_start_addr.setValue(400001); win.sweep_start_addr.setFixedWidth(100)
    r1.addWidget(win.sweep_start_addr)

    r1.addSpacing(24); r1.addWidget(QLabel("Address Step:"))
    win.sweep_addr_step = QSpinBox()
    win.sweep_addr_step.setRange(1, 100); win.sweep_addr_step.setValue(2); win.sweep_addr_step.setFixedWidth(70)
    win.sweep_addr_step.setToolTip("How many registers between tags.\nFor INT32 tags: 2 (each tag takes 2 registers).")
    r1.addWidget(win.sweep_addr_step)

    r1.addSpacing(24); r1.addWidget(QLabel("Tag Count:"))
    win.sweep_tag_count = QSpinBox()
    win.sweep_tag_count.setRange(1, MAX_SWEEP_TAGS); win.sweep_tag_count.setValue(200); win.sweep_tag_count.setFixedWidth(70)
    r1.addWidget(win.sweep_tag_count)

    r1.addSpacing(24); r1.addWidget(QLabel("Word Order:"))
    win.sweep_dtype = QComboBox(); win.sweep_dtype.setFixedWidth(190)
    win.sweep_dtype.addItems(["UINT32 Lo/Hi (CD AB)", "UINT32 Hi/Lo (AB CD)", "UINT16"])
    win.sweep_dtype.setToolTip("Match this to your PLC tag data type.\nUINT32 Lo/Hi (CD AB) = Productivity Suite INT32/DINT")
    r1.addWidget(win.sweep_dtype); r1.addStretch()
    cv.addLayout(r1)

    # Divider
    div = QFrame(); div.setFrameShape(QFrame.HLine); div.setObjectName("separator")
    cv.addWidget(div)

    # ── Row 2: Timing + value mode ───────────────────────────────────────
    r2 = QHBoxLayout(); r2.setSpacing(8)
    r2.addWidget(QLabel("Value Delay (ms):"))
    win.sweep_val_delay = QSpinBox()
    win.sweep_val_delay.setRange(50, 10000); win.sweep_val_delay.setValue(600); win.sweep_val_delay.setFixedWidth(80)
    win.sweep_val_delay.setToolTip("Time between each value write (1, 2, 3 ... to max)")
    r2.addWidget(win.sweep_val_delay)

    r2.addSpacing(24); r2.addWidget(QLabel("Tag Switch Delay (ms):"))
    win.sweep_tag_delay = QSpinBox()
    win.sweep_tag_delay.setRange(50, 10000); win.sweep_tag_delay.setValue(1500); win.sweep_tag_delay.setFixedWidth(80)
    win.sweep_tag_delay.setToolTip("Pause after finishing one tag before starting the next")
    r2.addWidget(win.sweep_tag_delay)

    r2.addSpacing(24); r2.addWidget(QLabel("Value Mode:"))
    win.sweep_val_mode = QComboBox(); win.sweep_val_mode.setFixedWidth(145)
    win.sweep_val_mode.addItems(["Linear (1 to N)", "Powers of 2"])
    win.sweep_val_mode.setToolTip("Linear: writes 1,2,3...N. Powers of 2: writes 1,2,4,8...2^N for 32-bit bit testing.")
    win.sweep_val_mode.currentIndexChanged.connect(win._on_sweep_val_mode_change)
    r2.addWidget(win.sweep_val_mode)

    r2.addSpacing(16)
    win.sweep_max_val_lbl = QLabel("Send 1 to:")
    r2.addWidget(win.sweep_max_val_lbl)
    win.sweep_max_val = QSpinBox()
    win.sweep_max_val.setRange(1, 65535); win.sweep_max_val.setValue(64); win.sweep_max_val.setFixedWidth(80)
    r2.addWidget(win.sweep_max_val)

    win.sweep_max_bit_lbl = QLabel("Up to Bit:")
    win.sweep_max_bit_lbl.setVisible(False); r2.addWidget(win.sweep_max_bit_lbl)
    win.sweep_max_bit = QComboBox(); win.sweep_max_bit.setFixedWidth(175); win.sweep_max_bit.setVisible(False)
    for b in range(1, 33):
        win.sweep_max_bit.addItem(f"Bit {b}  =  {2**(b-1)}", 2**(b-1))
    win.sweep_max_bit.setCurrentIndex(5)
    r2.addWidget(win.sweep_max_bit)

    r2.addSpacing(24); r2.addWidget(QLabel("Tag Prefix:"))
    win.sweep_tag_prefix = QLineEdit("aiw"); win.sweep_tag_prefix.setFixedWidth(70)
    win.sweep_tag_prefix.setToolTip("Tag name prefix in the log (e.g. aiw -> aiw001, aiw002...)")
    r2.addWidget(win.sweep_tag_prefix); r2.addStretch()
    cv.addLayout(r2)
    v.addWidget(cfg)

    # ── Progress ─────────────────────────────────────────────────────────
    pg = QGroupBox("Progress"); pv = QVBoxLayout(pg); pv.setSpacing(12)
    pv.setContentsMargins(16, 20, 16, 16)
    row1 = QHBoxLayout(); row1.setSpacing(8)
    row1.addWidget(QLabel("Tag:"))
    win.sweep_tag_lbl = QLabel("—")
    win.sweep_tag_lbl.setStyleSheet("color: #3a9bd5; font-size: 14px; font-weight: bold; letter-spacing: 1px;")
    win.sweep_tag_lbl.setToolTip("Click to copy tag name")
    win.sweep_tag_lbl.mousePressEvent = lambda e: QApplication.clipboard().setText(win.sweep_tag_lbl.text())
    win.sweep_tag_lbl.setCursor(Qt.PointingHandCursor)
    row1.addWidget(win.sweep_tag_lbl)
    row1.addSpacing(24); row1.addWidget(QLabel("Address:"))
    win.sweep_addr_lbl = QLabel("—")
    win.sweep_addr_lbl.setStyleSheet("color: #f0a500; font-size: 14px; font-weight: bold;")
    win.sweep_addr_lbl.setToolTip("Click to copy address")
    win.sweep_addr_lbl.mousePressEvent = lambda e: QApplication.clipboard().setText(win.sweep_addr_lbl.text())
    win.sweep_addr_lbl.setCursor(Qt.PointingHandCursor)
    row1.addWidget(win.sweep_addr_lbl)
    row1.addSpacing(24); row1.addWidget(QLabel("Sending Value:"))
    win.sweep_val_lbl = QLabel("—")
    win.sweep_val_lbl.setStyleSheet("color: #4caf7d; font-size: 14px; font-weight: bold;")
    win.sweep_val_lbl.setToolTip("Click to copy value")
    win.sweep_val_lbl.mousePressEvent = lambda e: QApplication.clipboard().setText(win.sweep_val_lbl.text())
    win.sweep_val_lbl.setCursor(Qt.PointingHandCursor)
    row1.addWidget(win.sweep_val_lbl)
    row1.addStretch()
    win.sweep_pause_btn = QPushButton("⏸  PAUSE")
    win.sweep_pause_btn.setObjectName("read_btn")
    win.sweep_pause_btn.clicked.connect(win.pause_sweep)
    win.sweep_pause_btn.setEnabled(False)
    win.sweep_pause_btn.setFixedWidth(110)
    row1.addWidget(win.sweep_pause_btn)
    row1.addSpacing(6)
    win.sweep_btn = QPushButton("▶  START SWEEP"); win.sweep_btn.setObjectName("write_btn"); win.sweep_btn.setMinimumWidth(160)
    win.sweep_btn.clicked.connect(win.toggle_sweep); row1.addWidget(win.sweep_btn)
    pv.addLayout(row1)
    # Start-from row
    row2 = QHBoxLayout(); row2.setSpacing(8)
    row2.addWidget(QLabel("Start from tag #:"))
    win.sweep_start_from = QSpinBox()
    win.sweep_start_from.setRange(1, MAX_SWEEP_TAGS)
    win.sweep_start_from.setValue(1)
    win.sweep_start_from.setFixedWidth(80)
    win.sweep_start_from.setToolTip("Begin sweep from this tag number instead of tag 1")
    row2.addWidget(win.sweep_start_from)
    row2.addSpacing(16)
    win.sweep_test_one_btn = QPushButton("▶  TEST CURRENT TAG")
    win.sweep_test_one_btn.setObjectName("read_btn")
    win.sweep_test_one_btn.setToolTip("Run the full value sequence on the current tag only, then stop")
    win.sweep_test_one_btn.clicked.connect(win._test_single_tag)
    row2.addWidget(win.sweep_test_one_btn)
    row2.addStretch()
    pv.addLayout(row2)
    win.sweep_progress = QProgressBar(); win.sweep_progress.setRange(0, 100); win.sweep_progress.setValue(0)
    win.sweep_progress.setFixedHeight(22)
    pv.addWidget(win.sweep_progress)
    export_row = QHBoxLayout(); export_row.addStretch()
    win.sweep_export_btn = QPushButton("⬇  EXPORT REPORT")
    win.sweep_export_btn.setObjectName("export_btn")
    win.sweep_export_btn.setEnabled(False)
    win.sweep_export_btn.setToolTip("Export sweep summary as CSV")
    win.sweep_export_btn.clicked.connect(win._export_sweep_report)
    export_row.addWidget(win.sweep_export_btn)
    pv.addLayout(export_row)
    v.addWidget(pg); v.addStretch()

    win.sweep_val_timer = QTimer(win); win.sweep_val_timer.timeout.connect(win._sweep_val_step)
    win.sweep_tag_timer = QTimer(win); win.sweep_tag_timer.setSingleShot(True)
    win.sweep_tag_timer.timeout.connect(win._sweep_next_tag)
    return w

