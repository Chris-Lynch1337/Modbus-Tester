"""
App-wide constants, palette colours, and stylesheet.
"""
from __future__ import annotations
from pathlib import Path

# App metadata
APP_NAME      = "Modbus TCP Tester"
APP_VERSION   = "2.2.0"
APP_COMPANY   = "Your Company"
APP_CONTACT   = "support@yourcompany.com"
SETTINGS_FILE = Path.home() / ".modbus_tcp_tester.json"

# Limits
MAX_LOG_LINES   = 2000
MAX_QUEUE_DEPTH = 200
MAX_SWEEP_TAGS  = 200
RECONNECT_DELAY_MS = 3000
STALL_TIMEOUT_MS   = 8000

# Palette
BG_DARK      = "#1a1c22"
BG_MID       = "#22252e"
BG_PANEL     = "#2a2d38"
BG_FIELD     = "#1e2028"
ACCENT_BLUE  = "#3a9bd5"
ACCENT_GREEN = "#4caf7d"
ACCENT_RED   = "#e05c5c"
ACCENT_AMBER = "#f0a500"
TEXT_PRIMARY = "#e8eaf0"
TEXT_DIM     = "#7a8099"
BORDER       = "#363a4a"

APPEARANCE_THEMES = {
    "Classic Dark": {
        "bg_dark":   "#1a1c22",
        "bg_panel":  "#2a2d38",
        "bg_field":  "#1e2028",
        "accent":    "#3a9bd5",
        "text":      "#e8eaf0",
        "log_ok":    "#4caf7d",
        "log_err":   "#e05c5c",
        "font_family": "Consolas",
        "font_size": 12,
    },
    "Ocean Night": {
        "bg_dark":   "#0f1b2b",
        "bg_panel":  "#172436",
        "bg_field":  "#0e1725",
        "accent":    "#57c2ff",
        "text":      "#e4f2ff",
        "log_ok":    "#3ddc97",
        "log_err":   "#ff6b6b",
        "font_family": "Cascadia Code",
        "font_size": 13,
    },
    "Retro Amber": {
        "bg_dark":   "#1b1408",
        "bg_panel":  "#261a0b",
        "bg_field":  "#201307",
        "accent":    "#ffb347",
        "text":      "#ffe3b0",
        "log_ok":    "#9ccc65",
        "log_err":   "#ff7043",
        "font_family": "Lucida Console",
        "font_size": 12,
    },
    "Section 9": {
        "bg_dark":   "#080f16",
        "bg_panel":  "#101c28",
        "bg_field":  "#07121b",
        "accent":    "#30f0d0",
        "text":      "#d4f7ff",
        "log_ok":    "#5dfdcb",
        "log_err":   "#ff5fa2",
        "font_family": "Fira Code",
        "font_size": 13,
    },
}

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {BG_DARK};
    color: {TEXT_PRIMARY};
    font-family: 'Consolas', 'Courier New', monospace;
}}
QGroupBox {{
    background-color: {BG_PANEL};
    border: 1px solid {BORDER};
    border-radius: 4px;
    margin-top: 20px;
    padding: 14px 12px;
    font-weight: bold;
    color: {TEXT_DIM};
    letter-spacing: 1px;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 2px 8px;
    background-color: {BG_PANEL};
    color: {ACCENT_BLUE};
}}
QLineEdit, QSpinBox, QComboBox {{
    background-color: {BG_FIELD};
    border: 1px solid {BORDER};
    border-radius: 3px;
    padding: 5px 8px;
    color: {TEXT_PRIMARY};
    font-family: 'Consolas', monospace;
    font-size: 12px;
    selection-background-color: {ACCENT_BLUE};
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{ border: 1px solid {ACCENT_BLUE}; }}
QLineEdit[readOnly="true"] {{ color: {TEXT_DIM}; }}
QPushButton {{
    background-color: {BG_PANEL};
    border: 1px solid {BORDER};
    border-radius: 3px;
    padding: 6px 16px;
    color: {TEXT_PRIMARY};
    font-family: 'Consolas', monospace;
    font-size: 11px;
    font-weight: bold;
    letter-spacing: 1px;
}}
QPushButton:hover {{ background-color: #32364a; border-color: {ACCENT_BLUE}; color: {ACCENT_BLUE}; }}
QPushButton:pressed {{ background-color: #1e2230; }}
QPushButton:disabled {{ color: {TEXT_DIM}; border-color: #2a2d38; }}
QPushButton#connect_btn    {{ background-color: #1e3a28; border-color: {ACCENT_GREEN}; color: {ACCENT_GREEN}; min-width: 100px; }}
QPushButton#connect_btn:hover {{ background-color: #264830; }}
QPushButton#disconnect_btn {{ background-color: #3a1e1e; border-color: {ACCENT_RED}; color: {ACCENT_RED}; min-width: 100px; }}
QPushButton#disconnect_btn:hover {{ background-color: #4a2424; }}
QPushButton#write_btn {{ background-color: #1e2e40; border-color: {ACCENT_BLUE}; color: {ACCENT_BLUE}; }}
QPushButton#write_btn:hover {{ background-color: #253848; }}
QPushButton#read_btn  {{ background-color: #2d2a18; border-color: {ACCENT_AMBER}; color: {ACCENT_AMBER}; }}
QPushButton#read_btn:hover  {{ background-color: #3a3520; }}
QPushButton#export_btn {{ background-color: #1e2e40; border-color: {TEXT_DIM}; color: {TEXT_DIM}; }}
QPushButton#export_btn:hover {{ border-color: {ACCENT_BLUE}; color: {ACCENT_BLUE}; }}
QTextEdit {{
    background-color: {BG_FIELD};
    border: 1px solid {BORDER};
    border-radius: 3px;
    color: #a0c8a0;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 11px;
    padding: 6px;
}}
QTabWidget::pane {{ border: 1px solid {BORDER}; background-color: {BG_MID}; }}
QTabBar::tab {{
    background-color: {BG_DARK}; color: {TEXT_DIM};
    border: 1px solid {BORDER}; border-bottom: none;
    padding: 10px 28px; font-family: 'Consolas', monospace; font-size: 12px; letter-spacing: 2px; min-width: 140px;
}}
QTabBar::tab:selected {{ background-color: {BG_MID}; color: {ACCENT_BLUE}; border-bottom: 2px solid {ACCENT_BLUE}; }}
QTabBar::tab:hover:!selected {{ background-color: {BG_PANEL}; color: {TEXT_PRIMARY}; }}
QLabel#status_connected    {{ color: {ACCENT_GREEN}; font-weight: bold; }}
QLabel#status_disconnected {{ color: {ACCENT_RED};   font-weight: bold; }}
QLabel#status_connecting   {{ color: {ACCENT_AMBER}; font-weight: bold; }}
QFrame#separator {{ background-color: {BORDER}; max-height: 1px; }}
QStatusBar {{ background-color: {BG_PANEL}; color: {TEXT_DIM}; font-size: 10px; border-top: 1px solid {BORDER}; }}
QProgressBar {{
    background-color: {BG_FIELD}; border: 1px solid {BORDER}; border-radius: 3px;
    height: 18px; text-align: center; color: {TEXT_PRIMARY};
    font-family: Consolas; font-size: 10px;
}}
QProgressBar::chunk {{ background-color: {ACCENT_BLUE}; border-radius: 2px; }}
QToolTip {{
    background-color: {BG_PANEL}; color: {TEXT_PRIMARY};
    border: 1px solid {ACCENT_BLUE}; padding: 4px; font-size: 11px;
}}
QScrollBar:vertical {{ background: {BG_DARK}; width: 8px; border: none; }}
QScrollBar::handle:vertical {{ background: {BORDER}; border-radius: 4px; min-height: 20px; }}
QScrollBar::handle:vertical:hover {{ background: {ACCENT_BLUE}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QSpinBox, QLineEdit, QComboBox {{ border-left: 3px solid {BORDER}; }}
QSpinBox:focus, QLineEdit:focus, QComboBox:focus {{ border-left: 3px solid {ACCENT_BLUE}; }}
QLabel#conn_banner_connected    {{ background-color: #1a2e1a; color: {ACCENT_GREEN}; padding: 4px 14px; border-radius: 4px; font-weight: bold; font-size: 11px; letter-spacing: 1px; }}
QLabel#conn_banner_disconnected {{ background-color: #2e1a1a; color: {ACCENT_RED};   padding: 4px 14px; border-radius: 4px; font-weight: bold; font-size: 11px; letter-spacing: 1px; }}
QLabel#conn_banner_connecting   {{ background-color: #2e2a1a; color: {ACCENT_AMBER}; padding: 4px 14px; border-radius: 4px; font-weight: bold; font-size: 11px; letter-spacing: 1px; }}
"""

# ─── Data types ───────────────────────────────────────────────────────────────
DTYPE_OPTIONS = [
    "UINT16",
    "INT16",
    "UINT32  –  Hi word first  (AB CD)",
    "UINT32  –  Lo word first  (CD AB)",
    "INT32   –  Hi word first  (AB CD)",
    "INT32   –  Lo word first  (CD AB)",
    "FLOAT32 –  Hi word first  (AB CD)",
    "FLOAT32 –  Lo word first  (CD AB)",
]

DTYPE_INFO = {
    "UINT16":                             ("16-bit unsigned integer",           "0 to 65535",                "1 register  (FC06)"),
    "INT16":                              ("16-bit signed integer",             "-32768 to 32767",           "1 register  (FC06)"),
    "UINT32  –  Hi word first  (AB CD)":  ("32-bit unsigned, high word @ addr", "0 to 4294967295",           "2 registers (FC16)"),
    "UINT32  –  Lo word first  (CD AB)":  ("32-bit unsigned, low word @ addr",  "0 to 4294967295",           "2 registers (FC16)"),
    "INT32   –  Hi word first  (AB CD)":  ("32-bit signed, high word @ addr",   "-2147483648 to 2147483647", "2 registers (FC16)"),
    "INT32   –  Lo word first  (CD AB)":  ("32-bit signed, low word @ addr",    "-2147483648 to 2147483647", "2 registers (FC16)"),
    "FLOAT32 –  Hi word first  (AB CD)":  ("IEEE 754 float, high word @ addr",  "e.g. 3.14 or -0.5",        "2 registers (FC16)"),
    "FLOAT32 –  Lo word first  (CD AB)":  ("IEEE 754 float, low word @ addr",   "e.g. 3.14 or -0.5",        "2 registers (FC16)"),
}

DTYPE_TOOLTIPS = {
    "UINT16":                            "Single 16-bit register. Range 0–65535.",
    "INT16":                             "Single 16-bit signed register. Range -32768–32767.",
    "UINT32  –  Hi word first  (AB CD)": "32-bit value. High word written to the lower address.\nMost common for standard Modbus devices.",
    "UINT32  –  Lo word first  (CD AB)": "32-bit value. Low word written to the lower address.\nUsed by AutomationDirect Productivity Suite.",
    "INT32   –  Hi word first  (AB CD)": "32-bit signed. High word at lower address.",
    "INT32   –  Lo word first  (CD AB)": "32-bit signed. Low word at lower address.\nUse this for Productivity Suite INT32 tags.",
    "FLOAT32 –  Hi word first  (AB CD)": "IEEE 754 float. High word at lower address.",
    "FLOAT32 –  Lo word first  (CD AB)": "IEEE 754 float. Low word at lower address.\nUse this for Productivity Suite REAL tags.",
}
