"""
Main application window — ModbusTester QMainWindow.

Tab content is built by the ui/ tab modules.
Business logic (connect, ramp, sweep, settings apply) lives here.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import QByteArray, Qt, QThread, QTimer, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QKeySequence
from PyQt5.QtWidgets import (
    QApplication, QCheckBox, QColorDialog, QComboBox, QFileDialog,
    QFrame, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMainWindow,
    QListWidget, QListWidgetItem, QMenu, QMessageBox, QProgressBar,
    QPushButton, QScrollArea, QShortcut, QSpinBox, QStatusBar,
    QTabWidget, QTextEdit, QVBoxLayout, QWidget, QInputDialog,
)

from .constants import (
    ACCENT_AMBER, ACCENT_BLUE, ACCENT_GREEN, ACCENT_RED,
    APP_NAME, APP_VERSION, APP_COMPANY, APP_CONTACT,
    BG_DARK, BG_FIELD, BG_MID, BG_PANEL, BORDER,
    DTYPE_OPTIONS, DTYPE_INFO, DTYPE_TOOLTIPS,
    MAX_LOG_LINES, MAX_SWEEP_TAGS, RECONNECT_DELAY_MS,
    SETTINGS_FILE, STALL_TIMEOUT_MS, STYLESHEET,
    TEXT_DIM, TEXT_PRIMARY,
    APPEARANCE_THEMES,
)
from .datatypes import (
    DecodeResult, OperationRequest, OperationResult, ValidationError,
    pack_value, preview_pack, validate_host,
)
from .demo import DemoServer
from .workers import CommandProcessor, ConnectionWorker, PYMODBUS_AVAILABLE
from .ramp import RampController
from .sweep import SweepController
from .ui.dialogs import ColorButton
from .ui.register_tab  import build_register_tab
from .ui.batch_tab     import build_batch_tab
from .ui.sweep_tab     import build_sweep_tab
from .ui.settings_tab  import build_settings_tab


# ─── Main window ──────────────────────────────────────────────────────────────
class ModbusTester(QMainWindow):
    def __init__(self):
        super().__init__()
        self.client              = None
        self.connected           = False
        self.connection_worker: Optional[ConnectionWorker] = None
        self.command_thread:    Optional[QThread]          = None
        self.command_processor: Optional[CommandProcessor] = None
        self.pending_shutdown    = False
        self.log_line_count      = 0
        self.write_count         = 0
        self.error_count         = 0
        self.last_write_ts       = ""
        self._auto_reconnect     = False
        self._reconnect_timer    = QTimer(self)
        self._reconnect_timer.setSingleShot(True)
        self._reconnect_timer.timeout.connect(self._do_auto_reconnect)

        self.ramp_running  = False
        self.ramp_current  = 0
        self.sweep_running     = False
        self.sweep_paused      = False
        self.sweep_tag_idx     = 0
        self.sweep_val         = 1
        self.sweep_start_time  = None
        self.sweep_error_count = 0
        self.sweep_write_count = 0
        self.sweep_tags_with_errors: list = []
        self._single_tag_mode = False
        self._last_sweep_summary: dict = {}
        self._connection_profiles: List[dict] = []
        self.demo_server = DemoServer()
        self.demo_mode_enabled = False
        self._demo_prev_host: Optional[str] = None
        self._demo_prev_port: Optional[int] = None
        self._pending_requests: List[str] = []
        self.ramp_controller = RampController(self)
        self.sweep_controller = SweepController(self)

        # Settings state (loaded before UI build so defaults apply)
        self._app_settings: dict = {}
        self._load_app_settings()

        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumSize(1200, 780)
        self.setStyleSheet(STYLESHEET)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        root.addWidget(self._build_header())
        root.addWidget(self._build_connection_bar())
        sep = QFrame(); sep.setObjectName("separator"); sep.setFrameShape(QFrame.HLine)
        root.addWidget(sep)
        self.tabs = QTabWidget()
        self.tabs.addTab(build_register_tab(self),  "  ⚙  Registers  ")
        self.tabs.addTab(build_batch_tab(self),     "  📈  Ramp / Batch  ")
        self.tabs.addTab(build_sweep_tab(self), "  🔄  Tag Sweep  ")
        self.settings_tab_index = self.tabs.addTab(build_settings_tab(self),  "  ☰  Settings  ")
        self.tabs.currentChanged.connect(self._on_tab_changed)
        root.addWidget(self.tabs, stretch=1)
        self.queue_panel = self._build_queue_panel()
        root.addWidget(self.queue_panel)
        root.addWidget(self._build_log())
        self._tab_status_labels = {
            "register": getattr(self, "reg_status_label", None),
            "batch":    getattr(self, "batch_status_label", None),
            "sweep":    getattr(self, "sweep_status_label", None),
        }
        self.ramp_controller.initialize()
        self.sweep_controller.initialize()

        self._build_status_bar()
        self._build_shortcuts()
        self._set_controls_enabled(False)
        self._restore_settings()
        self._refresh_profile_combo()
        self._update_panel_visibility()

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self) -> QWidget:
        w = QWidget(); h = QHBoxLayout(w); h.setContentsMargins(0, 0, 0, 4)
        title = QLabel(f"▣  {APP_NAME.upper()}")
        title.setFont(QFont("Consolas", 14, QFont.Bold))
        title.setStyleSheet(f"color: {ACCENT_BLUE}; letter-spacing: 3px;")

        self.queue_label = QLabel("QUEUE: 0")
        self.queue_label.setStyleSheet(f"color: {TEXT_DIM}; font-size: 10px;")

        self.status_label = QLabel("● DISCONNECTED")
        self.status_label.setObjectName("conn_banner_disconnected")
        self.status_label.setFont(QFont("Consolas", 10, QFont.Bold))
        self.activity_label = QLabel("")
        self.activity_label.setStyleSheet("color: #f0a500; font-size: 10px; font-family: Consolas; letter-spacing: 1px; font-weight: bold;")

        h.addWidget(title); h.addStretch()
        h.addWidget(self.activity_label); h.addSpacing(16)
        h.addWidget(self.queue_label); h.addSpacing(16)
        h.addWidget(self.status_label)
        return w

    # ── Connection bar ────────────────────────────────────────────────────────
    def _build_connection_bar(self) -> QGroupBox:
        grp = QGroupBox("Connection"); h = QHBoxLayout(grp); h.setSpacing(8)
        h.addWidget(QLabel("Host:"))
        self.host_edit = QLineEdit("192.168.1.1")
        self.host_edit.setFixedWidth(150)
        self.host_edit.setToolTip("IP address or hostname of the Modbus TCP device")
        h.addWidget(self.host_edit)

        h.addWidget(QLabel("Port:"))
        self.port_spin = QSpinBox(); self.port_spin.setRange(1, 65535); self.port_spin.setValue(502); self.port_spin.setFixedWidth(70)
        self.port_spin.setToolTip("Modbus TCP port (default: 502)")
        h.addWidget(self.port_spin)

        h.addWidget(QLabel("Unit ID:"))
        self.unit_spin = QSpinBox(); self.unit_spin.setRange(1, 247); self.unit_spin.setValue(1); self.unit_spin.setFixedWidth(60)
        self.unit_spin.setToolTip("Modbus slave/unit ID (usually 1)")
        h.addWidget(self.unit_spin)

        h.addWidget(QLabel("Timeout (s):"))
        self.timeout_spin = QSpinBox(); self.timeout_spin.setRange(1, 30); self.timeout_spin.setValue(3); self.timeout_spin.setFixedWidth(60)
        h.addWidget(self.timeout_spin)

        from PyQt5.QtWidgets import QCheckBox
        self.auto_reconnect_chk = QCheckBox("Auto-reconnect")
        self.auto_reconnect_chk.setToolTip("Automatically attempt to reconnect if the connection is lost")
        self.auto_reconnect_chk.setChecked(True)
        h.addWidget(self.auto_reconnect_chk)

        self.demo_mode_chk = QCheckBox("Demo simulator")
        self.demo_mode_chk.setToolTip("Launch a built-in Modbus TCP simulator on localhost for quick testing")
        self.demo_mode_chk.setEnabled(self.demo_server.available)
        self.demo_mode_chk.stateChanged.connect(self._on_demo_mode_toggled)
        h.addWidget(self.demo_mode_chk)

        h.addSpacing(12)
        self.profile_combo = QComboBox()
        self.profile_combo.setFixedWidth(160)
        self.profile_combo.addItem("Saved Profiles")
        self.profile_combo.currentIndexChanged.connect(self._on_profile_selected)
        h.addWidget(self.profile_combo)

        self.profile_save_btn = QPushButton("Save Profile")
        self.profile_save_btn.setObjectName("write_btn")
        self.profile_save_btn.clicked.connect(self._save_current_profile)
        h.addWidget(self.profile_save_btn)

        self.profile_delete_btn = QPushButton("Delete")
        self.profile_delete_btn.setObjectName("disconnect_btn")
        self.profile_delete_btn.setEnabled(False)
        self.profile_delete_btn.clicked.connect(self._delete_selected_profile)
        h.addWidget(self.profile_delete_btn)

        h.addStretch()
        self.connect_btn = QPushButton("CONNECT  [F5]")
        self.connect_btn.setObjectName("connect_btn")
        self.connect_btn.clicked.connect(self.do_connect)
        h.addWidget(self.connect_btn)

        self.disconnect_btn = QPushButton("DISCONNECT")
        self.disconnect_btn.setObjectName("disconnect_btn")
        self.disconnect_btn.clicked.connect(self.do_disconnect)
        self.disconnect_btn.setEnabled(False)
        h.addWidget(self.disconnect_btn)
        return grp

    def _on_demo_mode_toggled(self, state: int) -> None:
        if not hasattr(self, "demo_mode_chk"):
            return
        requested = state == Qt.Checked
        if requested == self.demo_mode_enabled:
            return
        if not self.demo_server.available:
            self._set_demo_checkbox(False)
            self.log_msg("Demo simulator unavailable \u2014 install pymodbus to enable it.", error=True)
            return
        if self.connected:
            self._set_demo_checkbox(self.demo_mode_enabled)
            self.log_msg("Disconnect before toggling the demo simulator.", error=True)
            return
        if requested:
            started = self.demo_server.start()
            if not started:
                err = self.demo_server.last_error
                detail = f": {err}" if err else ""
                self.log_msg(f"Demo simulator failed to start{detail}", error=True)
                self._set_demo_checkbox(False)
                return
            self._demo_prev_host = self.host_edit.text()
            self._demo_prev_port = self.port_spin.value()
            self.host_edit.setText(self.demo_server.host)
            self.port_spin.setValue(self.demo_server.port)
            self.host_edit.setEnabled(False)
            self.port_spin.setEnabled(False)
            self.demo_mode_enabled = True
            self.log_msg(f"Demo simulator running on {self.demo_server.host}:{self.demo_server.port}")
        else:
            self._stop_demo_mode()
        self._set_demo_checkbox(self.demo_mode_enabled)

    def _set_demo_checkbox(self, checked: bool) -> None:
        if not hasattr(self, "demo_mode_chk"):
            return
        self.demo_mode_chk.blockSignals(True)
        self.demo_mode_chk.setChecked(checked)
        self.demo_mode_chk.blockSignals(False)

    def _stop_demo_mode(self) -> None:
        if not self.demo_mode_enabled:
            return
        self.demo_server.stop()
        self.demo_mode_enabled = False
        self.host_edit.setEnabled(True)
        self.port_spin.setEnabled(True)
        if self._demo_prev_host is not None:
            self.host_edit.setText(self._demo_prev_host)
        if self._demo_prev_port is not None:
            self.port_spin.setValue(self._demo_prev_port)
        self._demo_prev_host = None
        self._demo_prev_port = None
        self._set_demo_checkbox(False)
        self.log_msg("Demo simulator stopped")

    def _build_queue_panel(self) -> QGroupBox:
        grp = QGroupBox("Command Queue")
        v = QVBoxLayout(grp); v.setContentsMargins(8, 6, 8, 8); v.setSpacing(6)
        self.queue_list = QListWidget()
        self.queue_list.setAlternatingRowColors(True)
        self.queue_list.setSelectionMode(QListWidget.NoSelection)
        self.queue_list.setFixedHeight(120)
        v.addWidget(self.queue_list)
        self._refresh_queue_list()
        return grp

    def _refresh_queue_list(self) -> None:
        if not hasattr(self, "queue_list"):
            return
        self.queue_list.clear()
        if not self._pending_requests:
            item = QListWidgetItem("Queue empty")
            item.setFlags(Qt.NoItemFlags)
            self.queue_list.addItem(item)
            return
        for entry in self._pending_requests[:20]:
            self.queue_list.addItem(entry)

    def _clear_queue_entries(self) -> None:
        self._pending_requests.clear()
        self._refresh_queue_list()

    def _add_queue_entry(self, request: OperationRequest) -> None:
        self._pending_requests.append(self._format_request_summary(request))
        self._refresh_queue_list()

    def _consume_queue_entry(self) -> None:
        if self._pending_requests:
            self._pending_requests.pop(0)
            self._refresh_queue_list()

    def _format_request_summary(self, request: OperationRequest) -> str:
        addr_6x = request.address + 400001
        if request.op == "read_registers":
            detail = f"Read {request.count} reg(s)"
        elif request.op == "write_registers":
            total = len(request.values or [])
            detail = f"Write {max(1, total)} reg(s)"
        else:
            detail = "Write single"
        if request.user_text:
            detail += f" [{request.user_text}]"
        return f"{detail} @ {addr_6x}  (unit {request.unit})"


    def _apply_appearance(self) -> None:
        """Rebuild stylesheet from current color/font settings and apply."""
        bd  = self.st_bg_dark.color()
        bp  = self.st_bg_panel.color()
        bf  = self.st_bg_field.color()
        acc = self.st_accent.color()
        txt = self.st_text.color()
        ff  = self.st_font_family.currentText()
        fs  = self.st_font_size.value()
        # Derive border and dim from base colors (slightly lighter/darker)
        brd = BORDER; dim = TEXT_DIM
        bm  = BG_MID
        new_ss = f"""
QMainWindow, QWidget {{
    background-color: {bd}; color: {txt};
    font-family: '{ff}', 'Courier New', monospace; font-size: {fs}px;
}}
QGroupBox {{
    background-color: {bp}; border: 1px solid {brd}; border-radius: 4px;
    margin-top: 20px; padding: 14px 12px; font-size: 11px; font-weight: bold; color: {dim}; letter-spacing: 1px;
}}
QGroupBox::title {{ subcontrol-origin: margin; subcontrol-position: top left; padding: 2px 8px; background-color: {bp}; color: {acc}; font-size: {max(10, fs-2)}px; }}
QLineEdit, QSpinBox, QComboBox {{ background-color: {bf}; border: 1px solid {brd}; border-radius: 3px; padding: 5px 8px; color: {txt}; font-family: '{ff}', monospace; font-size: {fs}px; selection-background-color: {acc}; }}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{ border: 1px solid {acc}; }}
QPushButton {{ background-color: {bp}; border: 1px solid {brd}; border-radius: 3px; padding: 6px 16px; color: {txt}; font-family: '{ff}', monospace; font-size: 11px; font-weight: bold; letter-spacing: 1px; }}
QPushButton:hover {{ background-color: #32364a; border-color: {acc}; color: {acc}; }}
QPushButton:pressed {{ background-color: #1e2230; }}
QPushButton:disabled {{ color: {dim}; border-color: #2a2d38; }}
QPushButton#connect_btn    {{ background-color: #1e3a28; border-color: {ACCENT_GREEN}; color: {ACCENT_GREEN}; min-width: 100px; }}
QPushButton#connect_btn:hover {{ background-color: #264830; }}
QPushButton#disconnect_btn {{ background-color: #3a1e1e; border-color: {ACCENT_RED}; color: {ACCENT_RED}; min-width: 100px; }}
QPushButton#disconnect_btn:hover {{ background-color: #4a2424; }}
QPushButton#write_btn {{ background-color: #1e2e40; border-color: {acc}; color: {acc}; }}
QPushButton#write_btn:hover {{ background-color: #253848; }}
QPushButton#read_btn  {{ background-color: #2d2a18; border-color: {ACCENT_AMBER}; color: {ACCENT_AMBER}; }}
QPushButton#read_btn:hover  {{ background-color: #3a3520; }}
QPushButton#export_btn {{ background-color: #1e2e40; border-color: {dim}; color: {dim}; }}
QPushButton#export_btn:hover {{ border-color: {acc}; color: {acc}; }}
QTextEdit {{ background-color: {bf}; border: 1px solid {brd}; border-radius: 3px; color: {self.st_log_ok.color()}; font-family: '{ff}', monospace; font-size: {fs}px; padding: 6px; }}
QTabWidget::pane {{ border: 1px solid {brd}; background-color: {bm}; }}
QTabBar::tab {{ background-color: {bd}; color: {dim}; border: 1px solid {brd}; border-bottom: none; padding: 10px 28px; font-family: '{ff}', monospace; font-size: 12px; letter-spacing: 2px; min-width: 140px; }}
QTabBar::tab:selected {{ background-color: {bm}; color: {acc}; border-bottom: 2px solid {acc}; }}
QTabBar::tab:hover:!selected {{ background-color: {bp}; color: {txt}; }}
QLabel#status_connected    {{ color: {ACCENT_GREEN}; font-weight: bold; }}
QLabel#status_disconnected {{ color: {ACCENT_RED};   font-weight: bold; }}
QLabel#status_connecting   {{ color: {ACCENT_AMBER}; font-weight: bold; }}
QFrame#separator {{ background-color: {brd}; max-height: 1px; }}
QStatusBar {{ background-color: {bp}; color: {dim}; font-size: 10px; border-top: 1px solid {brd}; }}
QProgressBar {{ background-color: {bf}; border: 1px solid {brd}; border-radius: 3px; height: 18px; text-align: center; color: {txt}; font-family: Consolas; font-size: 10px; }}
QProgressBar::chunk {{ background-color: {acc}; border-radius: 2px; }}
QToolTip {{ background-color: {bp}; color: {txt}; border: 1px solid {acc}; padding: 4px; font-size: 11px; }}
QScrollBar:vertical {{ background: {bd}; width: 8px; border: none; }}
QScrollBar::handle:vertical {{ background: {brd}; border-radius: 4px; min-height: 20px; }}
QScrollBar::handle:vertical:hover {{ background: {acc}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """
        self.setStyleSheet(new_ss)
        self.st_preview.setStyleSheet(f"color: {txt}; background: {bf}; padding: 6px 10px; border: 1px solid {brd}; border-radius: 3px; font-family: '{ff}'; font-size: {fs}px;")
        self._update_theme_combo_selection()
        self._save_settings()

    def _set_theme_values(self, theme: dict) -> None:
        if not theme:
            return
        color_map = [
            ("st_bg_dark", "bg_dark"),
            ("st_bg_panel", "bg_panel"),
            ("st_bg_field", "bg_field"),
            ("st_accent", "accent"),
            ("st_text", "text"),
            ("st_log_ok", "log_ok"),
            ("st_log_err", "log_err"),
        ]
        for attr, key in color_map:
            btn = getattr(self, attr, None)
            if btn is not None and key in theme:
                btn.set_color(theme[key])
        if hasattr(self, "st_font_family"):
            self.st_font_family.blockSignals(True)
            self.st_font_family.setCurrentText(theme.get("font_family", self.st_font_family.currentText()))
            self.st_font_family.blockSignals(False)
        if hasattr(self, "st_font_size"):
            self.st_font_size.blockSignals(True)
            self.st_font_size.setValue(theme.get("font_size", self.st_font_size.value()))
            self.st_font_size.blockSignals(False)
        self._apply_appearance()

    def _apply_selected_theme(self) -> None:
        if not hasattr(self, "st_theme_combo"):
            return
        name = self.st_theme_combo.currentText()
        theme = APPEARANCE_THEMES.get(name)
        if not theme:
            return
        self._set_theme_values(theme)
        self.log_msg(f"Applied theme: {name}")

    def _reset_appearance_defaults(self, log: bool = True) -> None:
        default_name = "Classic Dark" if "Classic Dark" in APPEARANCE_THEMES else next(iter(APPEARANCE_THEMES.keys()), "")
        theme = APPEARANCE_THEMES.get(default_name, {})
        self._set_theme_values(theme)
        if hasattr(self, "st_theme_combo") and default_name:
            self.st_theme_combo.blockSignals(True)
            self.st_theme_combo.setCurrentText(default_name)
            self.st_theme_combo.blockSignals(False)
        if log:
            self.log_msg("Appearance reset to defaults")
        self._update_theme_combo_selection()

    def _current_theme_snapshot(self) -> dict:
        if not hasattr(self, "st_bg_dark"):
            return {}
        return {
            "bg_dark":   self.st_bg_dark.color(),
            "bg_panel":  self.st_bg_panel.color(),
            "bg_field":  self.st_bg_field.color(),
            "accent":    self.st_accent.color(),
            "text":      self.st_text.color(),
            "log_ok":    self.st_log_ok.color(),
            "log_err":   self.st_log_err.color(),
            "font_family": self.st_font_family.currentText(),
            "font_size": self.st_font_size.value(),
        }

    def _find_matching_theme(self) -> Optional[str]:
        snapshot = self._current_theme_snapshot()
        if not snapshot:
            return None
        for name, theme in APPEARANCE_THEMES.items():
            if self._theme_snapshot_equals(snapshot, theme):
                return name
        return None

    def _theme_snapshot_equals(self, current: dict, theme: dict) -> bool:
        def norm(s): return (s or "").strip().lower()
        color_keys = ["bg_dark", "bg_panel", "bg_field", "accent", "text", "log_ok", "log_err"]
        for key in color_keys:
            if norm(current.get(key)) != norm(theme.get(key)):
                return False
        if norm(current.get("font_family")) != norm(theme.get("font_family")):
            return False
        return current.get("font_size") == theme.get("font_size")

    def _update_theme_combo_selection(self) -> None:
        if not hasattr(self, "st_theme_combo"):
            return
        match = self._find_matching_theme()
        label = getattr(self, "st_theme_custom_label", "Custom (modified)")
        target = match or label
        self.st_theme_combo.blockSignals(True)
        idx = self.st_theme_combo.findText(target)
        if idx == -1:
            self.st_theme_combo.addItem(target)
            idx = self.st_theme_combo.findText(target)
        if idx >= 0:
            self.st_theme_combo.setCurrentIndex(idx)
        self.st_theme_combo.blockSignals(False)

    def _apply_behavior(self) -> None:
        show_q = self.st_show_queue.isChecked()
        self.queue_label.setVisible(show_q)
        self._update_panel_visibility()
        self._save_settings()
        return

    def _update_panel_visibility(self) -> None:
        show_panel_pref = True
        if hasattr(self, "st_show_queue_panel"):
            show_panel_pref = self.st_show_queue_panel.isChecked()
        on_settings_tab = hasattr(self, "settings_tab_index") and self.tabs.currentIndex() == self.settings_tab_index
        show_queue_panel = show_panel_pref and not on_settings_tab
        if hasattr(self, "queue_panel"):
            self.queue_panel.setVisible(show_queue_panel)
        show_log = not on_settings_tab
        if hasattr(self, "log_container"):
            self.log_container.setVisible(show_log)
    def _export_settings(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Export Settings", "modbus_tester_settings.json", "JSON Files (*.json)")
        if not path: return
        dest = Path(path)
        if dest.exists():
            reply = QMessageBox.question(
                self,
                "Overwrite File?",
                f"{dest.name} already exists. Overwrite it?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return
        try:
            self._save_settings()
            shutil.copy(SETTINGS_FILE, dest)
            self.log_msg(f"Settings exported to {dest}")
        except Exception as exc:
            self.log_msg(f"Export failed: {exc}", error=True)

    def _import_settings(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Import Settings", "", "JSON Files (*.json)")
        if not path: return
        reply = QMessageBox.question(
            self,
            "Import Settings",
            "Importing will replace your current settings. A backup of the existing file will be created first.\nContinue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.No:
            return
        try:
            backup = SETTINGS_FILE.with_suffix(".bak")
            if SETTINGS_FILE.exists():
                shutil.copy(SETTINGS_FILE, backup)
                self.log_msg(f"Existing settings backed up to {backup}")
            shutil.copy(path, SETTINGS_FILE)
            self._load_app_settings()
            self._restore_settings()
            self.log_msg(f"Settings imported from {path}")
        except Exception as exc:
            self.log_msg(f"Import failed: {exc}", error=True)

    def _reset_settings(self) -> None:
        reply = QMessageBox.question(self, "Reset Settings",
            "Reset all settings to defaults? This cannot be undone.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No: return
        try:
            SETTINGS_FILE.unlink(missing_ok=True)
        except Exception: pass
        self._app_settings = {}
        self._reset_appearance_defaults(log=False)
        self.st_default_unit.setValue(1); self.st_default_timeout.setValue(3)
        self.st_default_word_order.setCurrentText(DTYPE_OPTIONS[3])
        self.st_default_val_delay.setValue(600); self.st_default_tag_delay.setValue(1500)
        self.st_reconnect_interval.setValue(3)
        self.st_confirm_exit.setChecked(True); self.st_show_queue.setChecked(True)
        self.st_ts_format.setCurrentIndex(0)
        self.log_msg("Settings reset to defaults")

    # ── Log ───────────────────────────────────────────────────────────────────
    def _build_log(self) -> QGroupBox:
        grp = QGroupBox("Activity Log"); v = QVBoxLayout(grp); v.setContentsMargins(6, 6, 6, 6)
        self.log = QTextEdit(); self.log.setReadOnly(True)
        self.log.setMinimumHeight(80)
        self.log.setSizePolicy(self.log.sizePolicy().Expanding, self.log.sizePolicy().Expanding)
        self.log.setContextMenuPolicy(Qt.CustomContextMenu)
        self.log.customContextMenuRequested.connect(self._log_context_menu)
        v.addWidget(self.log)
        h = QHBoxLayout()
        h.addWidget(QLabel("Log Height:"))
        self.log_height_spin = QSpinBox()
        self.log_height_spin.setRange(80, 600)
        self.log_height_spin.setValue(140)
        self.log_height_spin.setSuffix(" px")
        self.log_height_spin.setFixedWidth(80)
        self.log_height_spin.setToolTip("Drag to resize the activity log")
        self.log_height_spin.valueChanged.connect(lambda h: self.log.setFixedHeight(h))
        self.log.setFixedHeight(140)
        h.addWidget(self.log_height_spin)
        h.addStretch()
        self.verbose_log_combo = QComboBox(); self.verbose_log_combo.addItems(["Normal log", "Quiet log"])
        self.verbose_log_combo.setFixedWidth(120)
        self.verbose_log_combo.setToolTip("Quiet log: suppresses individual ramp/sweep write confirmations")
        h.addWidget(self.verbose_log_combo)
        export_btn = QPushButton("EXPORT LOG"); export_btn.setObjectName("export_btn")
        export_btn.clicked.connect(self._export_log)
        export_btn.setToolTip("Save the activity log to a .txt file")
        h.addWidget(export_btn)
        clr = QPushButton("CLEAR LOG"); clr.setFixedWidth(100); clr.clicked.connect(self._clear_log); h.addWidget(clr)
        v.addLayout(h)
        self.log_container = grp
        return grp

    # ── Status bar ────────────────────────────────────────────────────────────
    def _build_status_bar(self) -> None:
        sb = QStatusBar(); self.setStatusBar(sb)
        self.sb_writes = QLabel("✓  Writes: 0")
        self.sb_writes.setStyleSheet("color: #4caf7d; margin-right: 20px; font-family: Consolas; font-size: 11px;")
        self.sb_errors = QLabel("✗  Errors: 0")
        self.sb_errors.setStyleSheet("color: #e05c5c; margin-right: 20px; font-family: Consolas; font-size: 11px;")
        self.sb_last = QLabel("⏱  Last write: —")
        self.sb_last.setStyleSheet("color: #7a8099; font-family: Consolas; font-size: 11px; margin-right: 20px;")
        self.sb_mode = QLabel("")
        self.sb_mode.setStyleSheet("color: #f0a500; font-family: Consolas; font-size: 11px; font-weight: bold;")
        sb.addWidget(self.sb_writes); sb.addWidget(self.sb_errors)
        sb.addWidget(self.sb_last); sb.addWidget(self.sb_mode)
        self._pulse_state = False
        self._pulse_timer = QTimer(self); self._pulse_timer.setInterval(600)
        self._pulse_timer.timeout.connect(self._pulse_activity)

    def _pulse_activity(self) -> None:
        self._pulse_state = not self._pulse_state
        if self.ramp_running:
            self.sb_mode.setText("⬤  RAMP RUNNING" if self._pulse_state else "○  RAMP RUNNING")
            self.activity_label.setText("⬤  RAMP" if self._pulse_state else "")
        elif self.sweep_running:
            self.sb_mode.setText("⬤  SWEEP RUNNING" if self._pulse_state else "○  SWEEP RUNNING")
            self.activity_label.setText("⬤  SWEEP" if self._pulse_state else "")

    def _start_activity_pulse(self, mode: str) -> None:
        self._pulse_state = False
        self._pulse_timer.start()
        color = "#4caf7d" if mode == "ramp" else "#3a9bd5"
        self.sb_mode.setStyleSheet(f"color: {color}; font-family: Consolas; font-size: 11px; font-weight: bold;")

    def _stop_activity_pulse(self) -> None:
        self._pulse_timer.stop()
        self.sb_mode.setText("")
        self.activity_label.setText("")

    # ── Keyboard shortcuts ────────────────────────────────────────────────────
    def _build_shortcuts(self) -> None:
        QShortcut(QKeySequence("F5"),     self).activated.connect(self.do_connect)
        QShortcut(QKeySequence("Escape"), self).activated.connect(self._shortcut_stop)
        QShortcut(QKeySequence("Return"), self).activated.connect(self._shortcut_write)

    def _shortcut_stop(self) -> None:
        if self.ramp_running:   self.toggle_ramp()
        if self.sweep_running:  self.toggle_sweep()

    def _shortcut_write(self) -> None:
        if self.tabs.currentIndex() == 0:
            self.write_single_register()

    def _on_tab_changed(self, index: int) -> None:
        self._update_panel_visibility()

    # ── Settings ──────────────────────────────────────────────────────────────
    def _load_app_settings(self) -> None:
        """Load raw JSON — called before UI is built."""
        if not SETTINGS_FILE.exists(): return
        try:
            self._app_settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            self._connection_profiles = list(self._app_settings.get("connections", []))
        except Exception:
            self._app_settings = {}
            self._connection_profiles = []

    def _restore_settings(self) -> None:
        """Apply loaded settings to UI widgets — called after UI is built."""
        data = self._app_settings
        if not data: return
        self.host_edit.setText(str(data.get("host", self.host_edit.text())))
        self.port_spin.setValue(int(data.get("port", self.port_spin.value())))
        self.unit_spin.setValue(int(data.get("unit", self.unit_spin.value())))
        self.timeout_spin.setValue(int(data.get("timeout", self.timeout_spin.value())))
        self.reg_addr.setValue(int(data.get("reg_addr", self.reg_addr.value())))
        self.reg_dtype.setCurrentText(str(data.get("reg_dtype", self.reg_dtype.currentText())))
        self.reg_read_dtype.setCurrentText(str(data.get("reg_read_dtype", self.reg_read_dtype.currentText())))
        self.reg_base_addr.setValue(int(data.get("reg_base_addr", self.reg_base_addr.value())))
        self.batch_reg_addr.setValue(int(data.get("batch_reg_addr", self.batch_reg_addr.value())))
        self.ramp_addr.setValue(int(data.get("ramp_addr", self.ramp_addr.value())))
        self.ramp_dtype.setCurrentIndex(int(data.get("ramp_dtype_idx", self.ramp_dtype.currentIndex())))
        self.ramp_start.setCurrentIndex(int(data.get("ramp_start_idx", self.ramp_start.currentIndex())))
        self.ramp_end.setCurrentIndex(int(data.get("ramp_end_idx", self.ramp_end.currentIndex())))
        self.sweep_start_addr.setValue(int(data.get("sweep_start_addr", self.sweep_start_addr.value())))
        self.sweep_addr_step.setValue(int(data.get("sweep_addr_step", self.sweep_addr_step.value())))
        self.sweep_tag_count.setValue(int(data.get("sweep_tag_count", self.sweep_tag_count.value())))
        self.sweep_max_val.setValue(int(data.get("sweep_max_val", self.sweep_max_val.value())))
        self.sweep_val_mode.setCurrentIndex(int(data.get("sweep_val_mode", 0)))
        self.sweep_max_bit.setCurrentIndex(int(data.get("sweep_max_bit_idx", 5)))
        self._on_sweep_val_mode_change()
        self.sweep_tag_prefix.setText(str(data.get("sweep_tag_prefix", self.sweep_tag_prefix.text())))
        self.verbose_log_combo.setCurrentIndex(int(data.get("verbose_log_mode", 0)))
        self.auto_reconnect_chk.setChecked(bool(data.get("auto_reconnect", True)))
        log_h = int(data.get("log_height", 140))
        self.log_height_spin.setValue(log_h)
        self.log.setFixedHeight(log_h)
        if "window_geometry" in data:
            try:
                from PyQt5.QtCore import QByteArray
                self.restoreGeometry(QByteArray.fromBase64(data["window_geometry"].encode()))
            except Exception: pass
        # Apply appearance settings
        if "appearance" in data:
            ap = data["appearance"]
            self.st_bg_dark.set_color(ap.get("bg_dark", BG_DARK))
            self.st_bg_panel.set_color(ap.get("bg_panel", BG_PANEL))
            self.st_bg_field.set_color(ap.get("bg_field", BG_FIELD))
            self.st_accent.set_color(ap.get("accent", ACCENT_BLUE))
            self.st_text.set_color(ap.get("text", TEXT_PRIMARY))
            self.st_log_ok.set_color(ap.get("log_ok", ACCENT_GREEN))
            self.st_log_err.set_color(ap.get("log_err", ACCENT_RED))
            self.st_font_size.setValue(int(ap.get("font_size", 12)))
            self.st_font_family.setCurrentText(str(ap.get("font_family", "Consolas")))
            self._apply_appearance()
        # Apply behavior settings
        if "behavior" in data:
            bh = data["behavior"]
            self.st_default_unit.setValue(int(bh.get("default_unit", 1)))
            self.st_default_timeout.setValue(int(bh.get("default_timeout", 3)))
            self.st_default_word_order.setCurrentText(str(bh.get("default_word_order", DTYPE_OPTIONS[3])))
            self.st_default_val_delay.setValue(int(bh.get("default_val_delay", 600)))
            self.st_default_tag_delay.setValue(int(bh.get("default_tag_delay", 1500)))
            self.st_reconnect_interval.setValue(int(bh.get("reconnect_interval", 3)))
            self.st_confirm_exit.setChecked(bool(bh.get("confirm_exit", True)))
            self.st_show_queue.setChecked(bool(bh.get("show_queue", True)))
            if hasattr(self, "st_show_queue_panel"):
                self.st_show_queue_panel.setChecked(bool(bh.get("show_queue_panel", True)))
            self.st_ts_format.setCurrentIndex(int(bh.get("ts_format", 0)))
            self._apply_behavior()
        self._refresh_profile_combo()

    def _save_settings(self) -> None:
        data = {
            "host": self.host_edit.text().strip(),
            "port": self.port_spin.value(), "unit": self.unit_spin.value(),
            "timeout": self.timeout_spin.value(),
            "reg_addr": self.reg_addr.value(), "reg_dtype": self.reg_dtype.currentText(),
            "reg_read_dtype": self.reg_read_dtype.currentText(),
            "reg_base_addr": self.reg_base_addr.value(),
            "batch_reg_addr": self.batch_reg_addr.value(),
            "ramp_addr": self.ramp_addr.value(),
            "ramp_dtype_idx": self.ramp_dtype.currentIndex(),
            "ramp_start_idx": self.ramp_start.currentIndex(),
            "ramp_end_idx": self.ramp_end.currentIndex(),
            "sweep_start_addr": self.sweep_start_addr.value(),
            "sweep_addr_step": self.sweep_addr_step.value(),
            "sweep_tag_count": self.sweep_tag_count.value(),
            "sweep_max_val": self.sweep_max_val.value(),
            "sweep_val_mode": self.sweep_val_mode.currentIndex(),
            "sweep_max_bit_idx": self.sweep_max_bit.currentIndex(),
            "sweep_tag_prefix": self.sweep_tag_prefix.text(),
            "verbose_log_mode": self.verbose_log_combo.currentIndex(),
            "auto_reconnect": self.auto_reconnect_chk.isChecked(),
            "log_height": self.log_height_spin.value(),
        }
        # Guard: color/behavior widgets may not exist on very first launch
        if hasattr(self, "st_bg_dark"):
            data["appearance"] = {
                "bg_dark":    self.st_bg_dark.color(),
                "bg_panel":   self.st_bg_panel.color(),
                "bg_field":   self.st_bg_field.color(),
                "accent":     self.st_accent.color(),
                "text":       self.st_text.color(),
                "log_ok":     self.st_log_ok.color(),
                "log_err":    self.st_log_err.color(),
                "font_size":  self.st_font_size.value(),
                "font_family": self.st_font_family.currentText(),
            }
        if hasattr(self, "st_default_unit"):
            data["behavior"] = {
                "default_unit":       self.st_default_unit.value(),
                "default_timeout":    self.st_default_timeout.value(),
                "default_word_order": self.st_default_word_order.currentText(),
                "default_val_delay":  self.st_default_val_delay.value(),
                "default_tag_delay":  self.st_default_tag_delay.value(),
                "reconnect_interval": self.st_reconnect_interval.value(),
                "confirm_exit":       self.st_confirm_exit.isChecked(),
                "show_queue":         self.st_show_queue.isChecked(),
                "show_queue_panel":   self.st_show_queue_panel.isChecked(),
                "ts_format":          self.st_ts_format.currentIndex(),
            }
        data["connections"] = self._connection_profiles
        try:
            geo = self.saveGeometry().toBase64().data().decode()
            data["window_geometry"] = geo
            SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except Exception: pass

    # ── Connection profiles ─────────────────────────────────────────────
    def _refresh_profile_combo(self, selected_name: Optional[str] = None) -> None:
        if not hasattr(self, "profile_combo"):
            return
        current_name = selected_name
        if current_name is None and self.profile_combo.currentIndex() > 0:
            current_name = self.profile_combo.currentText()
        self._connection_profiles.sort(key=lambda p: p.get("name", "").lower())
        self.profile_combo.blockSignals(True)
        self.profile_combo.clear()
        self.profile_combo.addItem("Saved Profiles")
        selected_index = 0
        for idx, profile in enumerate(self._connection_profiles, start=1):
            name = profile.get("name", f"Profile {idx}")
            self.profile_combo.addItem(name)
            if current_name and name == current_name:
                selected_index = idx
        self.profile_combo.setCurrentIndex(selected_index)
        self.profile_combo.blockSignals(False)
        self.profile_delete_btn.setEnabled(selected_index > 0)

    def _on_profile_selected(self, index: int) -> None:
        if index <= 0:
            if hasattr(self, "profile_delete_btn"):
                self.profile_delete_btn.setEnabled(False)
            return
        if index - 1 >= len(self._connection_profiles):
            return
        profile = self._connection_profiles[index - 1]
        self.host_edit.setText(profile.get("host", self.host_edit.text()))
        self.port_spin.setValue(int(profile.get("port", self.port_spin.value())))
        self.unit_spin.setValue(int(profile.get("unit", self.unit_spin.value())))
        self.timeout_spin.setValue(int(profile.get("timeout", self.timeout_spin.value())))
        self.profile_delete_btn.setEnabled(True)

    def _save_current_profile(self) -> None:
        default_name = self.host_edit.text().strip() or "New Profile"
        name, ok = QInputDialog.getText(self, "Save Connection Profile", "Profile name:", text=default_name)
        if not ok:
            return
        name = name.strip()
        if not name:
            QMessageBox.warning(self, "Invalid Name", "Profile name cannot be empty.")
            return
        profile = {
            "name": name,
            "host": self.host_edit.text().strip(),
            "port": self.port_spin.value(),
            "unit": self.unit_spin.value(),
            "timeout": self.timeout_spin.value(),
        }
        replaced = False
        for existing in self._connection_profiles:
            if existing.get("name", "").lower() == name.lower():
                existing.update(profile)
                existing["name"] = name
                replaced = True
                break
        if not replaced:
            self._connection_profiles.append(profile)
        self._refresh_profile_combo(selected_name=name)
        self._save_settings()
        self.log_msg(f"Saved connection profile '{name}'.")

    def _delete_selected_profile(self) -> None:
        if not hasattr(self, "profile_combo"):
            return
        idx = self.profile_combo.currentIndex()
        if idx <= 0:
            return
        name = self.profile_combo.currentText()
        reply = QMessageBox.question(
            self,
            "Delete Profile",
            f"Delete connection profile '{name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.No:
            return
        del self._connection_profiles[idx - 1]
        self._refresh_profile_combo()
        self._save_settings()
        self.log_msg(f"Deleted connection profile '{name}'.")

    # ── Connect / Disconnect ──────────────────────────────────────────────────
    def do_connect(self) -> None:
        if not PYMODBUS_AVAILABLE:
            self.log_msg("pymodbus not installed. Run: pip install pymodbus", error=True); return
        if self.connected or self.connection_worker is not None: return
        host = self.host_edit.text().strip()
        if not validate_host(host):
            self.log_msg(f"Invalid host: '{host}'. Enter a valid IP address or hostname.", error=True); return
        self._set_status("● CONNECTING…", "status_connecting")
        self.connect_btn.setEnabled(False); self.disconnect_btn.setEnabled(False)
        self.connection_worker = ConnectionWorker(host, self.port_spin.value(), self.timeout_spin.value())
        self.connection_worker.finished_signal.connect(self._on_connection_finished)
        self.connection_worker.finished.connect(self._cleanup_connection_worker)
        self.connection_worker.start()

    def _cleanup_connection_worker(self) -> None:
        if self.connection_worker:
            self.connection_worker.deleteLater(); self.connection_worker = None

    def _on_connection_finished(self, ok: bool, message: str, client) -> None:
        if ok:
            self.client = client; self.connected = True
            self._auto_reconnect = False
            self._set_status("● CONNECTED", "status_connected")
            self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}  —  {self.host_edit.text().strip()}:{self.port_spin.value()}")
            self.connect_btn.setEnabled(False); self.disconnect_btn.setEnabled(True)
            self._set_controls_enabled(True)
            self._start_command_processor()
            self.log_msg(message); self._save_settings()
        else:
            self.client = None; self.connected = False
            self._set_status("● DISCONNECTED", "status_disconnected")
            self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
            self.connect_btn.setEnabled(True); self.disconnect_btn.setEnabled(False)
            self._set_controls_enabled(False)
            self.log_msg(message, error=True)
            if self._auto_reconnect and self.auto_reconnect_chk.isChecked():
                _rdelay = self.st_reconnect_interval.value() * 1000 if hasattr(self, "st_reconnect_interval") else RECONNECT_DELAY_MS
                self.log_msg(f"Reconnecting in {_rdelay//1000}s…")
                self._reconnect_timer.start(_rdelay)

    def _do_auto_reconnect(self) -> None:
        if not self.connected:
            self.log_msg("Auto-reconnect attempt…"); self.do_connect()

    def do_disconnect(self) -> None:
        self._reconnect_timer.stop(); self._auto_reconnect = False
        self._stop_motion_features(); self._stop_command_processor()
        if self.client:
            try: self.client.close()
            except Exception: pass
            self.client = None
        self.connected = False
        self._set_status("● DISCONNECTED", "status_disconnected")
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.connect_btn.setEnabled(True); self.disconnect_btn.setEnabled(False)
        self._set_controls_enabled(False)
        self._stop_activity_pulse()
        self.log_msg("Disconnected")

    def _start_command_processor(self) -> None:
        self._stop_command_processor()
        self._clear_queue_entries()
        self.command_thread    = QThread(self)
        self.command_processor = CommandProcessor()
        self.command_processor.moveToThread(self.command_thread)
        self.command_thread.started.connect(self.command_processor.start)
        self.command_processor.result_ready.connect(self._on_operation_result)
        self.command_processor.queue_depth_changed.connect(self._set_queue_depth)
        self.command_processor.fatal_error.connect(self._on_worker_fatal_error)
        self.command_processor.stall_detected.connect(self._on_stall_detected)
        self.command_processor.set_client(self.client)
        self.command_thread.start()

    def _stop_command_processor(self) -> None:
        if self.command_processor is not None:
            # Invoke stop() on the processor's own thread so timers are stopped correctly
            from PyQt5.QtCore import QMetaObject, Qt as _Qt
            QMetaObject.invokeMethod(self.command_processor, "stop", _Qt.BlockingQueuedConnection)
            self.command_processor.deleteLater()
            self.command_processor = None
        if self.command_thread is not None:
            self.command_thread.quit()
            self.command_thread.wait(2000)
            self.command_thread.deleteLater()
            self.command_thread = None
        self._set_queue_depth(0)
        self._clear_queue_entries()

    def _stop_motion_features(self) -> None:
        if self.ramp_running:
            self.ramp_running = False; self.ramp_timer.stop()
            self.ramp_btn.setText("▶  START RAMP")
        if self.sweep_running or self.sweep_paused:
            self.sweep_running = False; self.sweep_paused = False
            self.sweep_val_timer.stop(); self.sweep_tag_timer.stop()
            self.sweep_btn.setText("▶  START SWEEP"); self.sweep_val_lbl.setText("—")
            self.sweep_pause_btn.setText("⏸  PAUSE"); self.sweep_pause_btn.setEnabled(False)
        self._stop_activity_pulse()

    def _on_stall_detected(self) -> None:
        self.log_msg("WARNING: Write stall detected — device may have stopped responding. "
                     "Check connection. Auto-reconnect will trigger if enabled.", error=True)
        self._auto_reconnect = True
        self.do_disconnect()

    # ── Operations ────────────────────────────────────────────────────────────
    def _enqueue_request(self, request: OperationRequest) -> bool:
        if not self.connected or self.command_processor is None:
            self.log_msg("Not connected", error=True); return False
        accepted = self.command_processor.enqueue(request)
        if not accepted:
            self.log_msg("Request queue full — slow down or stop current activity.", error=True)
        else:
            self._add_queue_entry(request)
        return accepted

    def _on_operation_result(self, result: OperationResult) -> None:
        quiet = self.verbose_log_combo.currentIndex() == 1
        show  = not quiet or not result.ok or not (result.request and result.request.suppress_success_log)
        if show:
            msg = result.message
            if result.request and result.request.user_text and result.request.suppress_success_log:
                msg += f"  [Tag: {result.request.user_text}]"
            self.log_msg(msg, error=not result.ok)
        if result.ok and not (result.request and result.request.suppress_success_log):
            self.write_count += 1
            self.last_write_ts = datetime.now().strftime("%H:%M:%S")
            self.sb_writes.setText(f"✓  Writes: {self.write_count}")
            self.sb_last.setText(f"⏱  Last write: {self.last_write_ts}")
        if not result.ok:
            self.error_count += 1
            self.sb_errors.setText(f"✗  Errors: {self.error_count}")
        if result.request and result.request.user_text and self.sweep_running:
            if result.ok:
                self.sweep_write_count += 1
            else:
                self.sweep_error_count += 1
                tag = result.request.user_text
                if tag not in self.sweep_tags_with_errors:
                    self.sweep_tags_with_errors.append(tag)
        if result.read_value is not None:
            self.reg_read_display.setText(result.read_value)
        if result.request is not None:
            self._consume_queue_entry()

    def _on_worker_fatal_error(self, detail: str) -> None:
        self.log_msg("Internal worker error — see console for details.", error=True)
        print(detail)
        self._clear_queue_entries()

    def _addr(self, spinbox: QSpinBox) -> int:
        """Convert Productivity Suite 6x address (400001+) to zero-based Modbus address."""
        return spinbox.value() - 400001

    def _validate_reg_range(self, start_addr: int, word_count: int) -> None:
        if start_addr < 0:
            raise ValidationError("Address must be at least 400001")
        if start_addr + word_count - 1 > 99998:
            raise ValidationError("Register range exceeds max address 499999")

    def _set_controls_enabled(self, en: bool) -> None:
        # Tabs remain accessible regardless of connection state; keep hooks for future use.
        for i in range(self.tabs.count()):
            self.tabs.setTabEnabled(i, True)
        self._update_connection_dependent_controls(en)
        if not en:
            self._set_tab_status("register", "Connect to a device to modify registers.")
            self._set_tab_status("batch", "Connect to run batch or ramp operations.")
            self._set_tab_status("sweep", "Connect to start a sweep.")
        else:
            for key in ("register", "batch", "sweep"):
                self._set_tab_status(key, "")

    def _update_connection_dependent_controls(self, connected: bool) -> None:
        buttons = [
            getattr(self, "reg_write_btn", None),
            getattr(self, "reg_read_btn", None),
            getattr(self, "reg_write_all_btn", None),
            getattr(self, "batch_write_btn", None),
            getattr(self, "ramp_btn", None),
            getattr(self, "sweep_btn", None),
            getattr(self, "sweep_pause_btn", None),
            getattr(self, "sweep_test_one_btn", None),
        ]
        for btn in buttons:
            self._set_requires_connection(btn, connected)

    def _set_tab_status(self, key: str, text: str = "", error: bool = False) -> None:
        label = getattr(self, "_tab_status_labels", {}).get(key)
        if label is None:
            return
        if not text:
            label.setVisible(False)
            return
        color = ACCENT_RED if error else ACCENT_AMBER
        label.setStyleSheet(f"color: {color}; font-size: 11px; font-weight: bold;")
        label.setText(text)
        label.setVisible(True)

    def _set_requires_connection(self, widget, enabled: bool) -> None:
        if widget is None:
            return
        orig_tip = getattr(widget, "_orig_tooltip", None)
        if orig_tip is None:
            widget._orig_tooltip = widget.toolTip()
            orig_tip = widget._orig_tooltip
        widget.setEnabled(enabled)
        if enabled:
            widget.setToolTip(orig_tip or "")
        else:
            widget.setToolTip("Connect to a Modbus device to use this action.")

    def _set_status(self, text: str, obj_name: str) -> None:
        self.status_label.setText(text)
        banner_map = {
            "status_connected":    "conn_banner_connected",
            "status_disconnected": "conn_banner_disconnected",
            "status_connecting":   "conn_banner_connecting",
        }
        self.status_label.setObjectName(banner_map.get(obj_name, obj_name))
        self.status_label.setStyle(self.status_label.style())

    def _set_queue_depth(self, depth: int) -> None:
        self.queue_label.setText(f"QUEUE: {depth}")

    def _update_reg_preview(self) -> None:
        self.reg_preview.setText(preview_pack(self.reg_dtype.currentText(), self.reg_value_edit.text()))

    def _update_type_info(self) -> None:
        info = DTYPE_INFO.get(self.reg_dtype.currentText(), ("", "", ""))
        self.reg_type_info.setText(f"{info[0]}   |   Range: {info[1]}   |   Writes: {info[2]}")

    def _update_dtype_tooltip(self) -> None:
        tt = DTYPE_TOOLTIPS.get(self.reg_dtype.currentText(), "")
        self.reg_dtype.setToolTip(tt)

    # ── Write / Read ──────────────────────────────────────────────────────────
    def write_single_register(self) -> None:
        self._set_tab_status("register", "")
        try:
            dtype = self.reg_dtype.currentText()
            words = pack_value(dtype, self.reg_value_edit.text())
            addr  = self._addr(self.reg_addr)
            self._validate_reg_range(addr, len(words))
            if len(words) == 1:
                req = OperationRequest("write_register",  addr, self.unit_spin.value(), value=words[0], user_text=self.reg_value_edit.text())
            else:
                req = OperationRequest("write_registers", addr, self.unit_spin.value(), values=words,   user_text=self.reg_value_edit.text())
            if self._enqueue_request(req):
                self._set_tab_status("register", "Write queued…")
            else:
                self._set_tab_status("register", "Queue full — try again shortly.", True)
        except Exception as exc:
            self.log_msg(f"Write error: {exc}", error=True)
            self._set_tab_status("register", f"Error: {exc}", True)

    def read_registers(self) -> None:
        self._set_tab_status("register", "")
        try:
            count = self.reg_read_count.value()
            addr  = self._addr(self.reg_addr)
            self._validate_reg_range(addr, count)
            req = OperationRequest("read_registers", addr, self.unit_spin.value(),
                                   count=count, decode_dtype=self.reg_read_dtype.currentText())
            if self._enqueue_request(req):
                self._set_tab_status("register", "Read queued…")
            else:
                self._set_tab_status("register", "Queue full — try again shortly.", True)
        except Exception as exc:
            self.log_msg(f"Read error: {exc}", error=True)
            self._set_tab_status("register", f"Error: {exc}", True)

    def write_all_quick_registers(self) -> None:
        self._set_tab_status("register", "")
        try:
            values = [sp.value() for sp in self.reg_quick_fields]
            addr   = self._addr(self.reg_base_addr)
            self._validate_reg_range(addr, len(values))
            if self._enqueue_request(OperationRequest("write_registers", addr, self.unit_spin.value(), values=values)):
                self._set_tab_status("register", "Quick write queued…")
            else:
                self._set_tab_status("register", "Queue full — try again shortly.", True)
        except Exception as exc:
            self.log_msg(f"Quick set error: {exc}", error=True)
            self._set_tab_status("register", f"Error: {exc}", True)

    def batch_write_registers(self) -> None:
        self._set_tab_status("batch", "")
        try:
            vals = [int(p.strip(), 0) for p in self.batch_reg_vals.text().split(",") if p.strip()]
            if not vals: raise ValidationError("Enter at least one value")
            if any(v < 0 or v > 65535 for v in vals): raise ValidationError("Values must be 0–65535")
            addr = self._addr(self.batch_reg_addr)
            self._validate_reg_range(addr, len(vals))
            if self._enqueue_request(OperationRequest("write_registers", addr, self.unit_spin.value(), values=vals)):
                self._set_tab_status("batch", "Batch write queued…")
            else:
                self._set_tab_status("batch", "Queue full — try again shortly.", True)
        except Exception as exc:
            self.log_msg(f"Batch write error: {exc}", error=True)
            self._set_tab_status("batch", f"Error: {exc}", True)

    # ── Ramp ──────────────────────────────────────────────────────────────────
    def _on_ramp_mode_change(self) -> None:
        self.ramp_controller.on_mode_change()

    def _update_ramp_preview(self) -> None:
        self.ramp_controller.update_preview()

    def toggle_ramp(self) -> None:
        self.ramp_controller.toggle()

    def _queue_zero_for_ramp(self) -> None:
        self.ramp_controller.queue_zero()

    def _ramp_step(self) -> None:
        self.ramp_controller.step()

    # ── Sweep ─────────────────────────────────────────────────────
    def _on_sweep_val_mode_change(self) -> None:
        self.sweep_controller.on_value_mode_change()

    def pause_sweep(self) -> None:
        self.sweep_controller.pause()

    def _test_single_tag(self) -> None:
        self.sweep_controller.test_single_tag()

    def toggle_sweep(self) -> None:
        self.sweep_controller.toggle()

    def _sweep_begin_tag(self) -> None:
        self.sweep_controller._begin_tag()

    def _sweep_val_step(self) -> None:
        self.sweep_controller.value_step()

    def _queue_zero_for_sweep(self, addr_mb: int, dtype_idx: int) -> None:
        self.sweep_controller._queue_zero(addr_mb, dtype_idx)

    def _sweep_next_tag(self) -> None:
        self.sweep_controller.next_tag()

    def _sweep_stop(self, completed: bool) -> None:
        self.sweep_controller._stop(completed)

    def _export_sweep_report(self) -> None:
        self.sweep_controller.export_report()

    # ─── Log helpers ─────────────────────────────────────────────────────────
    def log_msg(self, msg: str, error: bool = False) -> None:
        fmt = "%H:%M:%S" if hasattr(self, "st_ts_format") and self.st_ts_format.currentIndex() == 1 else "%H:%M:%S.%f"
        timestamp = datetime.now().strftime(fmt)
        if fmt.endswith("%f"):
            timestamp = timestamp[:-3]
        color = ACCENT_RED if error else ACCENT_GREEN
        icon = "✗" if error else "✓"
        self.log.append(
            f'<span style="color:{TEXT_DIM}">[{timestamp}]</span> '
            f'<span style="color:{color}">{icon} {msg}</span>'
        )
        self.log_line_count += 1
        if self.log_line_count > MAX_LOG_LINES:
            cursor = self.log.textCursor()
            cursor.movePosition(cursor.Start)
            cursor.select(cursor.BlockUnderCursor)
            cursor.removeSelectedText()
            cursor.deleteChar()
            self.log_line_count -= 1

    def _log_context_menu(self, pos) -> None:
        from PyQt5.QtWidgets import QMenu
        menu = QMenu(self)
        menu.setStyleSheet(self.styleSheet())
        copy_sel = menu.addAction("Copy Selected")
        copy_all = menu.addAction("Copy All")
        menu.addSeparator()
        clr = menu.addAction("Clear Log")
        exp = menu.addAction("Export Log…")
        action = menu.exec_(self.log.mapToGlobal(pos))
        if action == copy_sel:
            self.log.copy()
        elif action == copy_all:
            QApplication.clipboard().setText(self.log.toPlainText())
        elif action == clr:
            self._clear_log()
        elif action == exp:
            self._export_log()

    def _clear_log(self) -> None:
        self.log.clear(); self.log_line_count = 0

    def _export_log(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Export Log", f"modbus_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", "Text Files (*.txt)")
        if not path: return
        try:
            plain = self.log.toPlainText()
            Path(path).write_text(plain, encoding="utf-8")
            self.log_msg(f"Log exported to {path}")
        except Exception as exc:
            self.log_msg(f"Export failed: {exc}", error=True)

    # ── Close ─────────────────────────────────────────────────────────────────
    def closeEvent(self, event) -> None:
        if (self.ramp_running or self.sweep_running) and (not hasattr(self, "st_confirm_exit") or self.st_confirm_exit.isChecked()):
            reply = QMessageBox.question(
                self, "Confirm Exit",
                "A ramp or sweep is currently active.\nAre you sure you want to exit?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.No: event.ignore(); return
        self._save_settings()
        if self.connected:
            self.pending_shutdown = True
            self.do_disconnect(); event.ignore(); return
        self._stop_demo_mode()
        self._stop_command_processor(); event.accept()


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName(APP_COMPANY)
    win = ModbusTester()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
