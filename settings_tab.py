"""Settings tab UI builder."""
from __future__ import annotations
import shutil
from datetime import datetime
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox, QColorDialog, QComboBox, QFileDialog, QFrame, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton,
    QScrollArea, QSpinBox, QVBoxLayout, QWidget,
)
from ..constants import (
    ACCENT_BLUE, ACCENT_GREEN, ACCENT_RED, ACCENT_AMBER,
    APP_COMPANY, APP_CONTACT, APP_NAME, APP_VERSION,
    BG_DARK, BG_FIELD, BG_PANEL, BORDER,
    DTYPE_OPTIONS, SETTINGS_FILE, TEXT_DIM, TEXT_PRIMARY,
    APPEARANCE_THEMES,
)
from ..ui.dialogs import ColorButton


# ── Settings tab ──────────────────────────────────────────────────────────
def build_settings_tab(win) -> QWidget:
    outer = QWidget()
    scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setWidget(outer)
    scroll.setStyleSheet("QScrollArea { border: none; }")
    v = QVBoxLayout(outer); v.setSpacing(14); v.setContentsMargins(16, 16, 16, 16)

    # ── Appearance ────────────────────────────────────────────────────────
    ap = QGroupBox("Appearance"); al = QVBoxLayout(ap); al.setSpacing(12)

    def color_row(label, attr, default):
        row = QHBoxLayout()
        lbl = QLabel(label); lbl.setMinimumWidth(160)
        btn = ColorButton(default)
        btn.color_changed.connect(win._apply_appearance)
        setattr(win, attr, btn)
        row.addWidget(lbl); row.addWidget(btn); row.addStretch()
        return row

    al.addLayout(color_row("Background (dark):",  "st_bg_dark",  BG_DARK))
    al.addLayout(color_row("Background (panels):", "st_bg_panel", BG_PANEL))
    al.addLayout(color_row("Input field background:", "st_bg_field", BG_FIELD))
    al.addLayout(color_row("Accent / highlight color:", "st_accent", ACCENT_BLUE))
    al.addLayout(color_row("Primary text color:", "st_text",    TEXT_PRIMARY))
    al.addLayout(color_row("Log success color:", "st_log_ok",  ACCENT_GREEN))
    al.addLayout(color_row("Log error color:",   "st_log_err", ACCENT_RED))

    font_row = QHBoxLayout()
    font_row.addWidget(QLabel("Font family:")); 
    win.st_font_family = QComboBox(); win.st_font_family.setFixedWidth(200)
    win.st_font_family.addItems([
        "Consolas", "Courier New", "Lucida Console", "Cascadia Code", "Fira Code", "Monaco",
        "Arial", "Helvetica", "Comic Sans MS",
    ])
    win.st_font_family.currentTextChanged.connect(win._apply_appearance)
    font_row.addWidget(win.st_font_family)
    font_row.addSpacing(24); font_row.addWidget(QLabel("Font size:"))
    win.st_font_size = QSpinBox(); win.st_font_size.setRange(8, 30); win.st_font_size.setValue(12); win.st_font_size.setSuffix(" pt")
    win.st_font_size.setFixedWidth(75); win.st_font_size.valueChanged.connect(win._apply_appearance)
    font_row.addWidget(win.st_font_size); font_row.addStretch()
    al.addLayout(font_row)

    preview_row = QHBoxLayout()
    win.st_preview = QLabel("The quick brown fox — 0x1A2B3C4D — Register 400001")
    win.st_preview.setStyleSheet(f"color: {TEXT_PRIMARY}; background: {BG_FIELD}; padding: 6px 10px; border: 1px solid {BORDER}; border-radius: 3px;")
    preview_row.addWidget(win.st_preview); preview_row.addStretch()
    al.addLayout(preview_row)

    theme_row = QHBoxLayout()
    theme_row.addWidget(QLabel("Preset theme:"))
    win.st_theme_combo = QComboBox(); win.st_theme_combo.setFixedWidth(200)
    win.st_theme_combo.addItems(list(APPEARANCE_THEMES.keys()))
    win.st_theme_custom_label = "Custom (modified)"
    win.st_theme_combo.addItem(win.st_theme_custom_label)
    theme_row.addWidget(win.st_theme_combo)
    apply_theme_btn = QPushButton("Apply Theme")
    apply_theme_btn.setObjectName("write_btn")
    apply_theme_btn.clicked.connect(win._apply_selected_theme)
    theme_row.addWidget(apply_theme_btn)
    reset_theme_btn = QPushButton("Reset Appearance")
    reset_theme_btn.setObjectName("read_btn")
    reset_theme_btn.clicked.connect(win._reset_appearance_defaults)
    theme_row.addWidget(reset_theme_btn)
    theme_row.addStretch()
    al.addLayout(theme_row)
    v.addWidget(ap)

    # ── Behavior ──────────────────────────────────────────────────────────
    bh = QGroupBox("Behavior"); bl = QVBoxLayout(bh); bl.setSpacing(10)

    def spin_row(label, attr, lo, hi, default, suffix=""):
        row = QHBoxLayout(); row.addWidget(QLabel(label))
        sp = QSpinBox(); sp.setRange(lo, hi); sp.setValue(default); sp.setFixedWidth(90)
        if suffix: sp.setSuffix(suffix)
        setattr(win, attr, sp); row.addWidget(sp); row.addStretch()
        return row

    def combo_row(label, attr, items, default_text=""):
        row = QHBoxLayout(); row.addWidget(QLabel(label))
        cb = QComboBox(); cb.setFixedWidth(300); cb.addItems(items)
        if default_text: cb.setCurrentText(default_text)
        setattr(win, attr, cb); row.addWidget(cb); row.addStretch()
        return row

    def check_row(label, attr, default=True):
        row = QHBoxLayout()
        cb = QCheckBox(label); cb.setChecked(default)
        setattr(win, attr, cb); row.addWidget(cb); row.addStretch()
        return row

    bl.addLayout(spin_row("Default Unit ID:", "st_default_unit", 1, 247, 1))
    bl.addLayout(spin_row("Default Timeout (s):", "st_default_timeout", 1, 30, 3))
    bl.addLayout(combo_row("Default Word Order:", "st_default_word_order", DTYPE_OPTIONS, DTYPE_OPTIONS[3]))
    bl.addLayout(spin_row("Default Value Delay (ms):", "st_default_val_delay", 50, 10000, 600, " ms"))
    bl.addLayout(spin_row("Default Tag Switch Delay (ms):", "st_default_tag_delay", 50, 10000, 1500, " ms"))
    bl.addLayout(spin_row("Auto-reconnect Interval (s):", "st_reconnect_interval", 1, 60, 3, " s"))

    ts_row = QHBoxLayout(); ts_row.addWidget(QLabel("Timestamp format:"))
    win.st_ts_format = QComboBox(); win.st_ts_format.setFixedWidth(200)
    win.st_ts_format.addItems(["HH:MM:SS.ms  (default)", "HH:MM:SS"])
    ts_row.addWidget(win.st_ts_format); ts_row.addStretch()
    bl.addLayout(ts_row)

    bl.addLayout(check_row("Show queue depth in header", "st_show_queue", True))
    bl.addLayout(check_row("Show command queue panel", "st_show_queue_panel", True))
    bl.addLayout(check_row("Confirm before closing if ramp/sweep active", "st_confirm_exit", True))
    v.addWidget(bh)

    # ── Sweep defaults ────────────────────────────────────────────────────
    sd = QGroupBox("Sweep Defaults"); sl = QVBoxLayout(sd); sl.setSpacing(10)
    sl.addLayout(spin_row("Default tag count:", "st_def_tag_count", 1, 200, 200))
    sl.addLayout(spin_row("Default address step:", "st_def_addr_step", 1, 100, 2))

    pfx_row = QHBoxLayout(); pfx_row.addWidget(QLabel("Default tag prefix:"))
    win.st_def_prefix = QLineEdit("aiw"); win.st_def_prefix.setFixedWidth(100)
    pfx_row.addWidget(win.st_def_prefix); pfx_row.addStretch()
    sl.addLayout(pfx_row)
    v.addWidget(sd)

    # ── Import / Export / Reset ───────────────────────────────────────────
    io = QGroupBox("Settings File"); il = QHBoxLayout(io); il.setSpacing(12)
    export_btn = QPushButton("Export Settings…"); export_btn.setObjectName("write_btn")
    export_btn.clicked.connect(win._export_settings)
    import_btn = QPushButton("Import Settings…"); import_btn.setObjectName("read_btn")
    import_btn.clicked.connect(win._import_settings)
    reset_btn = QPushButton("Reset to Defaults"); reset_btn.setObjectName("disconnect_btn")
    reset_btn.clicked.connect(win._reset_settings)
    il.addWidget(export_btn); il.addWidget(import_btn); il.addWidget(reset_btn); il.addStretch()
    v.addWidget(io)

    # ── About / Info ──────────────────────────────────────────────────────
    ab = QGroupBox("About"); av = QVBoxLayout(ab); av.setSpacing(6)
    for text, bold in [
        (f"{APP_NAME}  v{APP_VERSION}", True),
        (f"© {datetime.now().year} {APP_COMPANY}", False),
        ("", False),
        ("Designed for AutomationDirect Productivity Suite", False),
        ("Modbus TCP holding register testing and commissioning utility", False),
        ("", False),
        (f"Support: {APP_CONTACT}", False),
        (f"Settings file: {SETTINGS_FILE}", False),
    ]:
        lbl = QLabel(text)
        style = f"color: {ACCENT_BLUE}; font-size: 13px; font-weight: bold;" if bold else f"color: {TEXT_DIM}; font-size: 11px;"
        lbl.setStyleSheet(style)
        av.addWidget(lbl)
    v.addWidget(ab)
    v.addStretch()

    # Apply defaults from sweep defaults to sweep tab when changed
    win.st_def_tag_count.valueChanged.connect(lambda v: win.sweep_tag_count.setValue(v))
    win.st_def_addr_step.valueChanged.connect(lambda v: win.sweep_addr_step.setValue(v))
    win.st_def_prefix.textChanged.connect(lambda t: win.sweep_tag_prefix.setText(t))
    win.st_default_unit.valueChanged.connect(lambda v: win.unit_spin.setValue(v))
    win.st_default_timeout.valueChanged.connect(lambda v: win.timeout_spin.setValue(v))
    win.st_default_val_delay.valueChanged.connect(lambda v: win.sweep_val_delay.setValue(v))
    win.st_default_tag_delay.valueChanged.connect(lambda v: win.sweep_tag_delay.setValue(v))
    win.st_reconnect_interval.valueChanged.connect(win._apply_behavior)
    win.st_show_queue.stateChanged.connect(win._apply_behavior)
    win.st_show_queue_panel.stateChanged.connect(win._apply_behavior)

    # Wrap in a widget for the tab
    wrapper = QWidget(); wl = QVBoxLayout(wrapper); wl.setContentsMargins(0, 0, 0, 0)
    wl.addWidget(scroll)
    return wrapper
