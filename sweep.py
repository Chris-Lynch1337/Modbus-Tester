"""
Sweep control helper encapsulating sweep workflows.
"""
from __future__ import annotations

import csv
from datetime import datetime
from typing import TYPE_CHECKING

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QFileDialog

from .constants import APP_NAME, APP_VERSION
from .datatypes import OperationRequest, ValidationError

if TYPE_CHECKING:
    from .main_window import ModbusTester


class SweepController:
    def __init__(self, win: "ModbusTester"):
        self.win = win

    def initialize(self) -> None:
        self.on_value_mode_change()

    def on_value_mode_change(self) -> None:
        win = self.win
        is_linear = win.sweep_val_mode.currentIndex() == 0
        win.sweep_max_val_lbl.setVisible(is_linear)
        win.sweep_max_val.setVisible(is_linear)
        win.sweep_max_bit_lbl.setVisible(not is_linear)
        win.sweep_max_bit.setVisible(not is_linear)

    def pause(self) -> None:
        win = self.win
        if not win.sweep_running and not win.sweep_paused:
            return
        if win.sweep_paused:
            win.sweep_paused = False
            win.sweep_running = True
            win.sweep_pause_btn.setText("⏸  PAUSE")
            win.sweep_btn.setText("■  STOP SWEEP  [Esc]")
            win.sweep_btn.setEnabled(True)
            win._start_activity_pulse("sweep")
            win.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}  —  SWEEP RUNNING")
            win.log_msg(f"Sweep resumed at {win.sweep_tag_lbl.text()}")
            win.sweep_val_timer.start(win.sweep_val_delay.value())
            win._set_tab_status("sweep", "Sweep running…")
        else:
            win.sweep_running = False
            win.sweep_paused = True
            win.sweep_val_timer.stop()
            win.sweep_tag_timer.stop()
            win.sweep_pause_btn.setText("▶  RESUME")
            win.sweep_btn.setText("■  STOP SWEEP")
            win._stop_activity_pulse()
            win.sb_mode.setText("⏸  SWEEP PAUSED")
            win.activity_label.setText("⏸  PAUSED")
            win.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}  —  SWEEP PAUSED")
            win.log_msg(f"Sweep paused at {win.sweep_tag_lbl.text()}  value={win.sweep_val_lbl.text()}")
            win._set_tab_status("sweep", "Sweep paused.")

    def test_single_tag(self) -> None:
        win = self.win
        if win.sweep_running or win.sweep_paused:
            win.log_msg("Stop the current sweep before testing a single tag.", error=True)
            win._set_tab_status("sweep", "Stop the active sweep before testing.", True)
            return
        if not win.connected:
            win.log_msg("Not connected", error=True)
            win._set_tab_status("sweep", "Connect to a device first.", True)
            return
        win.sweep_tag_idx = win.sweep_start_from.value() - 1
        win.sweep_running = True
        win.sweep_paused = False
        win.sweep_start_time = datetime.now()
        win.sweep_error_count = 0
        win.sweep_write_count = 0
        win.sweep_tags_with_errors = []
        win._single_tag_mode = True
        win.sweep_btn.setText("■  STOP SWEEP  [Esc]")
        win.sweep_pause_btn.setEnabled(True)
        win.sweep_export_btn.setEnabled(False)
        win._start_activity_pulse("sweep")
        win.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}  —  SWEEP RUNNING")
        self._begin_tag()
        win._set_tab_status("sweep", "Testing current tag…")

    def toggle(self) -> None:
        win = self.win
        if win.sweep_running or win.sweep_paused:
            self._stop(completed=False)
            return
        win._set_tab_status("sweep", "")
        try:
            count = win.sweep_tag_count.value()
            step = win.sweep_addr_step.value()
            words = 1 if win.sweep_dtype.currentIndex() == 2 else 2
            last_mb = win.sweep_start_addr.value() + (count - 1) * step - 400001
            win._validate_reg_range(last_mb, words)
            if step < words:
                raise ValidationError("Address step too small — tags will overlap")
        except Exception as exc:
            win.log_msg(f"Sweep config error: {exc}", error=True)
            win._set_tab_status("sweep", f"Sweep error: {exc}", True)
            return
        start_from = win.sweep_start_from.value() - 1
        win.sweep_running = True
        win.sweep_paused = False
        win.sweep_tag_idx = max(0, min(start_from, win.sweep_tag_count.value() - 1))
        win.sweep_start_time = datetime.now()
        win.sweep_error_count = 0
        win.sweep_write_count = 0
        win.sweep_tags_with_errors = []
        win.sweep_btn.setText("■  STOP SWEEP  [Esc]")
        win.sweep_pause_btn.setEnabled(True)
        win.sweep_export_btn.setEnabled(False)
        self._begin_tag()
        win._set_tab_status("sweep", "Sweep running…")

    def _begin_tag(self) -> None:
        win = self.win
        if not win.sweep_running:
            return
        count = win.sweep_tag_count.value()
        if win.sweep_tag_idx >= count:
            self._stop(completed=True)
            return
        win.sweep_val = 1
        addr_4x = win.sweep_start_addr.value() + win.sweep_tag_idx * win.sweep_addr_step.value()
        prefix = win.sweep_tag_prefix.text().strip() or "tag"
        tag_name = f"{prefix}{win.sweep_tag_idx + 1:03d}"
        win.sweep_tag_lbl.setText(f"{tag_name}  ({win.sweep_tag_idx+1}/{count})")
        win.sweep_addr_lbl.setText(str(addr_4x))
        win.sweep_val_lbl.setText("1")
        win.sweep_progress.setRange(0, count)
        win.sweep_progress.setValue(win.sweep_tag_idx)
        win.sweep_progress.setFormat(f"{win.sweep_tag_idx} of {count} tags  ({win.sweep_tag_idx*100//count}%)")
        win.log_msg(f"Sweep: {tag_name}  addr={addr_4x}  MB:{addr_4x - 400001}")
        win.sweep_val_timer.start(win.sweep_val_delay.value())

    def value_step(self) -> None:
        win = self.win
        if not win.connected or not win.sweep_running:
            self._stop(completed=False)
            return
        addr_4x = win.sweep_start_addr.value() + win.sweep_tag_idx * win.sweep_addr_step.value()
        addr_mb = addr_4x - 400001
        value = win.sweep_val
        dtype_idx = win.sweep_dtype.currentIndex()
        prefix = win.sweep_tag_prefix.text().strip() or "tag"
        tag_name = f"{prefix}{win.sweep_tag_idx + 1:03d}"
        if dtype_idx == 2:
            request = OperationRequest(
                "write_register",
                addr_mb,
                win.unit_spin.value(),
                value=value & 0xFFFF,
                suppress_success_log=True,
                user_text=tag_name,
            )
        else:
            hi = (value >> 16) & 0xFFFF
            lo = value & 0xFFFF
            words = [lo, hi] if dtype_idx == 0 else [hi, lo]
            request = OperationRequest(
                "write_registers",
                addr_mb,
                win.unit_spin.value(),
                values=words,
                suppress_success_log=True,
                user_text=tag_name,
            )
        win._enqueue_request(request)
        win.sweep_val_lbl.setText(str(value))
        is_linear = win.sweep_val_mode.currentIndex() == 0
        if is_linear:
            win.sweep_val += 1
            done = win.sweep_val > win.sweep_max_val.value()
        else:
            win.sweep_val = win.sweep_val * 2 if win.sweep_val > 0 else 1
            done = win.sweep_val > int(win.sweep_max_bit.currentData())
        if done:
            win.sweep_val_timer.stop()
            addr_cap = addr_mb
            dtype_cap = dtype_idx
            QTimer.singleShot(
                win.sweep_val_delay.value(), lambda: self._queue_zero(addr_cap, dtype_cap)
            )
            win.sweep_tag_timer.start(win.sweep_tag_delay.value())

    def _queue_zero(self, addr_mb: int, dtype_idx: int) -> None:
        win = self.win
        if dtype_idx == 2:
            win._enqueue_request(
                OperationRequest("write_register", addr_mb, win.unit_spin.value(), value=0, suppress_success_log=True)
            )
        else:
            win._enqueue_request(
                OperationRequest(
                    "write_registers", addr_mb, win.unit_spin.value(), values=[0, 0], suppress_success_log=True
                )
            )

    def next_tag(self) -> None:
        win = self.win
        if not win.sweep_running:
            return
        if getattr(win, "_single_tag_mode", False):
            win._single_tag_mode = False
            self._stop(completed=True)
            return
        win.sweep_tag_idx += 1
        self._begin_tag()

    def _stop(self, completed: bool) -> None:
        win = self.win
        win.sweep_val_timer.stop()
        win.sweep_tag_timer.stop()
        win.sweep_running = False
        win.sweep_paused = False
        if win.command_processor:
            win.command_processor.clear_queue()
        win.sweep_btn.setText("▶  START SWEEP")
        win.sweep_val_lbl.setText("—")
        win.sweep_pause_btn.setText("⏸  PAUSE")
        win.sweep_pause_btn.setEnabled(False)
        win._stop_activity_pulse()
        win.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        count = win.sweep_tag_count.value()
        if completed:
            win.sweep_tag_lbl.setText("COMPLETE")
            win.sweep_addr_lbl.setText("—")
            win.sweep_progress.setRange(0, count)
            win.sweep_progress.setValue(count)
            win.sweep_progress.setFormat(f"Complete — {count} of {count} tags (100%)")
            elapsed = datetime.now() - win.sweep_start_time if win.sweep_start_time else None
            elapsed_str = str(elapsed).split(".")[0] if elapsed else "unknown"
            first_addr = win.sweep_start_addr.value()
            last_addr = win.sweep_start_addr.value() + (count - 1) * win.sweep_addr_step.value()
            is_linear = win.sweep_val_mode.currentIndex() == 0
            val_desc = (
                f"Linear 1 to {win.sweep_max_val.value()}"
                if is_linear
                else f"Powers of 2 up to {win.sweep_max_bit.currentText().strip()}"
            )
            vals_per = win.sweep_max_val.value() if is_linear else win.sweep_max_bit.currentIndex() + 1
            sep = "—" * 48
            win.log_msg(sep)
            win._last_sweep_summary = {
                "elapsed": elapsed_str,
                "tags": count,
                "first_addr": first_addr,
                "last_addr": last_addr,
                "step": win.sweep_addr_step.value(),
                "word_order": win.sweep_dtype.currentText(),
                "value_mode": val_desc,
                "vals_per_tag": vals_per,
                "total_writes": win.sweep_write_count,
                "errors": win.sweep_error_count,
                "error_tags": list(win.sweep_tags_with_errors),
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            win.sweep_export_btn.setEnabled(True)
            win.log_msg("SWEEP COMPLETE  —  SUMMARY REPORT")
            win.log_msg(sep)
            win.log_msg(f"  Total time       : {elapsed_str}")
            win.log_msg(f"  Tags tested      : {count}")
            win.log_msg(f"  Address range    : {first_addr} to {last_addr}  (step {win.sweep_addr_step.value()})")
            win.log_msg(f"  Word order       : {win.sweep_dtype.currentText()}")
            win.log_msg(f"  Value mode       : {val_desc}")
            win.log_msg(f"  Values per tag   : {vals_per}")
            win.log_msg(f"  Total writes     : {win.sweep_write_count}")
            if win.sweep_error_count == 0:
                win.log_msg("  Errors           : None")
            else:
                win.log_msg(f"  Errors           : {win.sweep_error_count}", error=True)
                win.log_msg(f"  Tags with errors : {', '.join(win.sweep_tags_with_errors)}", error=True)
            win.log_msg(sep)
            win._set_tab_status("sweep", "Sweep complete.")
        else:
            if win.connected and win.sweep_tag_idx < count:
                addr_4x = win.sweep_start_addr.value() + win.sweep_tag_idx * win.sweep_addr_step.value()
                addr_mb = addr_4x - 400001
                dtype_idx = win.sweep_dtype.currentIndex()
                self._queue_zero(addr_mb, dtype_idx)
            win.sweep_progress.setValue(0)
            win.sweep_progress.setFormat("Stopped")
            win.log_msg("Sweep stopped — register zeroed")
            win._set_tab_status("sweep", "Sweep stopped.")

    def export_report(self) -> None:
        win = self.win
        if not win._last_sweep_summary:
            return
        path, _ = QFileDialog.getSaveFileName(
            win,
            "Export Sweep Report",
            f"sweep_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV Files (*.csv)",
        )
        if not path:
            return
        try:
            summary = win._last_sweep_summary
            with open(path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(["Sweep Report", summary["timestamp"]])
                writer.writerow([])
                writer.writerow(["Field", "Value"])
                writer.writerow(["Total Time", summary["elapsed"]])
                writer.writerow(["Tags Tested", summary["tags"]])
                writer.writerow(["Address Range", f'{summary["first_addr"]} to {summary["last_addr"]}'])
                writer.writerow(["Address Step", summary["step"]])
                writer.writerow(["Word Order", summary["word_order"]])
                writer.writerow(["Value Mode", summary["value_mode"]])
                writer.writerow(["Values Per Tag", summary["vals_per_tag"]])
                writer.writerow(["Total Writes", summary["total_writes"]])
                writer.writerow(["Errors", summary["errors"]])
                if summary["error_tags"]:
                    writer.writerow(["Tags With Errors", ", ".join(summary["error_tags"])])
            win.log_msg(f"Sweep report exported to {path}")
        except Exception as exc:
            win.log_msg(f"Export failed: {exc}", error=True)
