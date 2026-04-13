"""
ConnectionWorker and CommandProcessor — background thread Modbus execution.
"""
from __future__ import annotations

import inspect
import traceback
from collections import deque
from typing import Callable, Deque

from PyQt5.QtCore import QObject, QThread, QTimer, pyqtSignal, pyqtSlot

try:
    from pymodbus.client import ModbusTcpClient
    PYMODBUS_AVAILABLE = True
except ImportError:
    ModbusTcpClient = None  # type: ignore
    PYMODBUS_AVAILABLE = False

from .constants import MAX_QUEUE_DEPTH, STALL_TIMEOUT_MS
from .datatypes import DecodeResult, OperationRequest, OperationResult, decode_words


class ConnectionWorker(QThread):
    finished_signal = pyqtSignal(bool, str, object)

    def __init__(self, host: str, port: int, timeout: int):
        super().__init__()
        self.host    = host
        self.port    = port
        self.timeout = timeout

    def run(self) -> None:
        try:
            client = ModbusTcpClient(self.host, port=self.port, timeout=self.timeout)
            if client.connect():
                self.finished_signal.emit(True, f"Connected to {self.host}:{self.port}", client)
            else:
                try: client.close()
                except Exception: pass
                self.finished_signal.emit(False, f"Connection refused by {self.host}:{self.port}", None)
        except Exception as exc:
            self.finished_signal.emit(False, str(exc), None)


# ─── Command processor ────────────────────────────────────────────────────────
class CommandProcessor(QObject):
    result_ready        = pyqtSignal(object)
    queue_depth_changed = pyqtSignal(int)
    fatal_error         = pyqtSignal(str)
    stall_detected      = pyqtSignal()

    def __init__(self) -> None:
        super().__init__()
        self._client   = None
        self._queue: Deque[OperationRequest] = deque()
        self._busy     = False
        self._stopping = False
        self._timer    = QTimer(self)
        self._timer.setInterval(5)
        self._timer.timeout.connect(self._pump)
        self._stall_timer = QTimer(self)
        self._stall_timer.setSingleShot(True)
        self._stall_timer.setInterval(STALL_TIMEOUT_MS)
        self._stall_timer.timeout.connect(self._on_stall)

    def set_client(self, client) -> None:  self._client = client
    def clear_client(self) -> None:        self._client = None

    def start(self) -> None:  self._timer.start()

    from PyQt5.QtCore import pyqtSlot
    @pyqtSlot()
    def stop(self) -> None:
        self._stopping = True
        self._timer.stop()
        self._stall_timer.stop()
        self._queue.clear()
        self.queue_depth_changed.emit(0)
        self._busy = False

    def clear_queue(self) -> None:
        self._queue.clear()
        self.queue_depth_changed.emit(0)

    def enqueue(self, request: OperationRequest) -> bool:
        if self._stopping: return False
        if len(self._queue) >= MAX_QUEUE_DEPTH: return False
        self._queue.append(request)
        self.queue_depth_changed.emit(len(self._queue))
        self._pump()
        return True

    def is_idle(self) -> bool:
        return not self._busy and not self._queue

    def _on_stall(self) -> None:
        self._busy = False
        self.stall_detected.emit()

    def _pump(self) -> None:
        if self._busy or self._stopping or not self._queue: return
        if self._client is None:
            self._queue.clear()
            self.queue_depth_changed.emit(0)
            self.result_ready.emit(OperationResult(False, "No Modbus client available"))
            return
        request = self._queue.popleft()
        self.queue_depth_changed.emit(len(self._queue))
        self._busy = True
        self._stall_timer.start()
        try:
            result = self._execute(request)
            self.result_ready.emit(result)
        except Exception as exc:
            detail = traceback.format_exc(limit=4)
            self.result_ready.emit(OperationResult(False, f"Worker error: {exc}"))
            self.fatal_error.emit(detail)
        finally:
            self._stall_timer.stop()
            self._busy = False
            if self._queue:
                QTimer.singleShot(0, self._pump)

    def _call(self, fn: Callable, request: OperationRequest, *args, **kwargs):
        params = inspect.signature(fn).parameters
        if "slave" in params: return fn(*args, slave=request.unit, **kwargs)
        if "unit"  in params: return fn(*args, unit=request.unit,  **kwargs)
        return fn(*args, **kwargs)

    def _execute(self, request: OperationRequest) -> OperationResult:
        if request.op == "write_register":
            r  = self._call(self._client.write_register, request, request.address, request.value)
            ok = not r.isError() if hasattr(r, "isError") else True
            return OperationResult(ok, f"Register {request.address} → {request.value} (0x{request.value:04X})", request=request)

        if request.op == "write_registers":
            vals = request.values or []
            r    = self._call(self._client.write_registers, request, request.address, vals)
            ok   = not r.isError() if hasattr(r, "isError") else True
            words = "  ".join(f"0x{v:04X}" for v in vals)
            return OperationResult(ok, f"Registers {request.address}–{request.address+len(vals)-1} → [{words}]", request=request)

        if request.op == "read_registers":
            r = self._call(self._client.read_holding_registers, request, request.address, count=request.count)
            if hasattr(r, "isError") and r.isError():
                return OperationResult(False, f"Read failed at {request.address}", request=request)
            regs    = list(r.registers)
            vals    = "  ".join(f"{request.address+i}={v} (0x{v:04X})" for i, v in enumerate(regs))
            decoded = decode_words(request.decode_dtype, regs)
            return OperationResult(True, f"Read Regs: {vals}", read_value=decoded.display, raw_registers=regs, request=request)

        return OperationResult(False, f"Unknown op: {request.op}", request=request)

