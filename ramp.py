"""
Ramp control helper that encapsulates ramp-specific logic.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from .constants import APP_NAME, APP_VERSION
from .datatypes import OperationRequest, ValidationError

if TYPE_CHECKING:
    from .main_window import ModbusTester


class RampController:
    def __init__(self, win: "ModbusTester"):
        self.win = win

    def initialize(self) -> None:
        self.on_mode_change()
        self.update_preview()

    def on_mode_change(self) -> None:
        win = self.win
        is_linear = win.ramp_mode.currentIndex() == 0
        win.ramp_step_lbl.setVisible(is_linear)
        win.ramp_step.setVisible(is_linear)

    def update_preview(self) -> None:
        win = self.win
        is_linear = win.ramp_mode.currentIndex() == 0
        start = int(win.ramp_start.currentData())
        end = int(win.ramp_end.currentData())
        seq = []
        value = start
        while value <= end and len(seq) < 8:
            seq.append(str(value))
            value = value + win.ramp_step.value() if is_linear else value * 2
        preview = ", ".join(seq)
        if value <= end:
            preview += ", …"
        win.ramp_preview_lbl.setText(f"Seq: {preview}")

    def toggle(self) -> None:
        win = self.win
        if win.ramp_running:
            win.ramp_running = False
            win.ramp_timer.stop()
            win.ramp_btn.setText("▶  START RAMP")
            win.log_msg("Ramp stopped")
            win._stop_activity_pulse()
            win.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
            if win.command_processor:
                win.command_processor.clear_queue()
            self.queue_zero()
            win._set_tab_status("batch", "Ramp stopped.")
            return
        win._set_tab_status("batch", "")
        try:
            start = int(win.ramp_start.currentData())
            end = int(win.ramp_end.currentData())
            if start > end:
                raise ValidationError("Start must be ≤ End")
            size = 1 if win.ramp_dtype.currentIndex() == 0 else 2
            win._validate_reg_range(win._addr(win.ramp_addr), size)
        except Exception as exc:
            win.log_msg(f"Ramp config error: {exc}", error=True)
            win._set_tab_status("batch", f"Ramp error: {exc}", True)
            return
        win.ramp_current = int(win.ramp_start.currentData())
        win.ramp_running = True
        win.ramp_btn.setText("■  STOP RAMP  [Esc]")
        win.ramp_timer.start(win.ramp_delay.value())
        win._start_activity_pulse("ramp")
        win.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}  —  RAMP RUNNING")
        win.log_msg(
            f"Ramp started: addr={win.ramp_addr.value()}  "
            f"{win.ramp_start.currentData()}→{win.ramp_end.currentData()}  "
            f"mode={win.ramp_mode.currentText()}"
        )
        win._set_tab_status("batch", "Ramp running…")

    def queue_zero(self) -> None:
        win = self.win
        if not win.connected:
            return
        addr = win._addr(win.ramp_addr)
        if win.ramp_dtype.currentIndex() == 0:
            win._enqueue_request(
                OperationRequest("write_register", addr, win.unit_spin.value(), value=0, suppress_success_log=True)
            )
        else:
            win._enqueue_request(
                OperationRequest(
                    "write_registers", addr, win.unit_spin.value(), values=[0, 0], suppress_success_log=True
                )
            )

    def step(self) -> None:
        win = self.win
        if not win.connected or not win.ramp_running:
            self.toggle()
            return
        addr = win._addr(win.ramp_addr)
        value = win.ramp_current
        if win.ramp_dtype.currentIndex() == 0:
            request = OperationRequest(
                "write_register", addr, win.unit_spin.value(), value=value & 0xFFFF, suppress_success_log=True
            )
        else:
            hi = (value >> 16) & 0xFFFF
            lo = value & 0xFFFF
            words = [lo, hi] if win.ramp_dtype.currentIndex() == 1 else [hi, lo]
            request = OperationRequest(
                "write_registers", addr, win.unit_spin.value(), values=words, suppress_success_log=True
            )
        win._enqueue_request(request)
        win.ramp_current = (
            win.ramp_current + win.ramp_step.value()
            if win.ramp_mode.currentIndex() == 0
            else win.ramp_current * 2
        )
        if win.ramp_current > int(win.ramp_end.currentData()):
            win.ramp_current = int(win.ramp_start.currentData())
