"""
Embedded Modbus TCP demo server for offline testing.
"""
from __future__ import annotations

import asyncio
import socket
import threading
import time
from typing import Optional

try:
    from pymodbus.datastore import (
        ModbusSequentialDataBlock,
        ModbusServerContext,
    )
    try:
        from pymodbus.datastore import ModbusSlaveContext as _DeviceContext
    except Exception:  # pragma: no cover - recent PyModbus versions
        from pymodbus.datastore import ModbusDeviceContext as _DeviceContext  # type: ignore
    try:  # PyModbus >=3.1 exposes the server directly under pymodbus.server
        from pymodbus.server import ModbusTcpServer
    except Exception:  # pragma: no cover - older releases use pymodbus.server.sync
        from pymodbus.server.sync import ModbusTcpServer  # type: ignore
    DEMO_AVAILABLE = True
except Exception:  # pragma: no cover - pymodbus optional
    ModbusTcpServer = None  # type: ignore
    _DeviceContext = ModbusServerContext = ModbusSequentialDataBlock = None  # type: ignore
    DEMO_AVAILABLE = False


class DemoServer:
    def __init__(self, host: str = "127.0.0.1", port: int = 15020):
        self.host = host
        self.port = port
        self._thread: Optional[threading.Thread] = None
        self._server: Optional[ModbusTcpServer] = None  # type: ignore[type-arg]
        self._error: Optional[Exception] = None
        self._loop: Optional[asyncio.AbstractEventLoop] = None
        self._async_mode = bool(
            DEMO_AVAILABLE and asyncio.iscoroutinefunction(getattr(ModbusTcpServer, "serve_forever", None))
        )

    @property
    def available(self) -> bool:
        return DEMO_AVAILABLE

    def is_running(self) -> bool:
        return self._thread is not None

    def start(self) -> bool:
        if not DEMO_AVAILABLE or self._thread is not None:
            return DEMO_AVAILABLE and self._thread is not None

        store = _DeviceContext(
            di=ModbusSequentialDataBlock(0, [0] * 1000),
            co=ModbusSequentialDataBlock(0, [0] * 1000),
            hr=ModbusSequentialDataBlock(0, [0] * 1000),
            ir=ModbusSequentialDataBlock(0, [0] * 1000),
        )
        try:
            context = ModbusServerContext(devices=store, single=True)
        except TypeError:
            context = ModbusServerContext(slaves=store, single=True)  # type: ignore[arg-type]

        self._error = None
        target = self._run_async_server if self._async_mode else self._run_sync_server
        self._thread = threading.Thread(target=target, args=(context,), daemon=True)
        self._thread.start()
        if not self._wait_for_ready(timeout=3.0):
            self._error = self._error or RuntimeError("Demo server did not start")
            self.stop()
            return False
        return True

    def stop(self) -> None:
        server = self._server
        if server is not None:
            if self._async_mode and self._loop is not None:
                try:
                    fut = asyncio.run_coroutine_threadsafe(server.shutdown(), self._loop)
                    fut.result(timeout=2.0)
                except Exception:
                    pass
            else:
                try:
                    server.shutdown()
                except Exception:
                    pass
                try:
                    server.server_close()
                except Exception:
                    pass
            self._server = None
        if self._thread is not None:
            self._thread.join(timeout=1.0)
            self._thread = None

    @property
    def last_error(self) -> Optional[Exception]:
        return self._error

    # Internal helpers -----------------------------------------------------
    def _run_sync_server(self, context) -> None:
        try:
            server = ModbusTcpServer(context, address=(self.host, self.port))
            self._server = server
            server.serve_forever()
        except Exception as exc:
            self._error = exc
        finally:
            self._server = None
            self._thread = None

    def _run_async_server(self, context) -> None:
        loop = asyncio.new_event_loop()
        self._loop = loop
        asyncio.set_event_loop(loop)
        async def runner():
            server = ModbusTcpServer(context, address=(self.host, self.port))
            self._server = server
            await server.serve_forever()
        try:
            loop.run_until_complete(runner())
        except Exception as exc:
            self._error = exc
        finally:
            self._server = None
            try:
                loop.run_until_complete(loop.shutdown_asyncgens())
            except Exception:
                pass
            loop.close()
            self._loop = None
            self._thread = None

    def _wait_for_ready(self, timeout: float) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if self._error is not None:
                return False
            try:
                with socket.create_connection((self.host, self.port), timeout=0.3):
                    return True
            except OSError:
                time.sleep(0.1)
        return False
