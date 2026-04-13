"""
Data structures, validation, and Modbus value packing/decoding.
"""
from __future__ import annotations

import math
import re
import struct
from dataclasses import dataclass
import math
import re
import struct
from typing import List, Optional


class ValidationError(ValueError):
    pass


@dataclass
class OperationRequest:
    op:                  str
    address:             int
    unit:                int
    count:               int = 0
    value:               int = 0
    values:              Optional[List[int]] = None
    decode_dtype:        str = "INT32   –  Lo word first  (CD AB)"
    user_text:           str = ""
    suppress_success_log: bool = False


@dataclass
class OperationResult:
    ok:            bool
    message:       str
    read_value:    Optional[str]       = None
    raw_registers: Optional[List[int]] = None
    request:       Optional[OperationRequest] = None


@dataclass
class DecodeResult:
    display:    str
    words_used: int


# ─── Connection worker ────────────────────────────────────────────────────────

def pack_value(dtype: str, text: str) -> List[int]:
    text = text.strip()
    if not text: raise ValidationError("Value is required")

    if dtype == "UINT16":
        v = int(text, 0)
        if not 0 <= v <= 0xFFFF: raise ValidationError("Out of range 0–65535")
        return [v]
    if dtype == "INT16":
        v = int(text, 0)
        if not -32768 <= v <= 32767: raise ValidationError("Out of range -32768–32767")
        return [v & 0xFFFF]
    if "UINT32" in dtype:
        v = int(text, 0)
        if not 0 <= v <= 0xFFFFFFFF: raise ValidationError("Out of range 0–4294967295")
        hi, lo = (v >> 16) & 0xFFFF, v & 0xFFFF
        return [hi, lo] if "AB CD" in dtype else [lo, hi]
    if "INT32" in dtype:
        v = int(text, 0)
        if not -2147483648 <= v <= 2147483647: raise ValidationError("Out of range")
        u = v & 0xFFFFFFFF
        hi, lo = (u >> 16) & 0xFFFF, u & 0xFFFF
        return [hi, lo] if "AB CD" in dtype else [lo, hi]
    if "FLOAT32" in dtype:
        v = float(text)
        if not math.isfinite(v): raise ValidationError("Float must be finite")
        raw = struct.pack(">f", v)
        hi  = struct.unpack(">H", raw[0:2])[0]
        lo  = struct.unpack(">H", raw[2:4])[0]
        return [hi, lo] if "AB CD" in dtype else [lo, hi]
    raise ValidationError(f"Unknown type: {dtype}")


def decode_words(dtype: str, regs: List[int]) -> DecodeResult:
    if not regs: raise ValidationError("No registers returned")
    if dtype == "UINT16": return DecodeResult(str(regs[0]), 1)
    if dtype == "INT16":
        r = regs[0]; return DecodeResult(str(r if r < 0x8000 else r - 0x10000), 1)
    if len(regs) < 2: raise ValidationError("Need 2 registers for this type")
    hi, lo = (regs[0], regs[1]) if "AB CD" in dtype else (regs[1], regs[0])
    raw32 = ((hi & 0xFFFF) << 16) | (lo & 0xFFFF)
    if "UINT32" in dtype: return DecodeResult(str(raw32), 2)
    if "INT32"  in dtype:
        v = raw32 if raw32 < 0x80000000 else raw32 - 0x100000000
        return DecodeResult(str(v), 2)
    if "FLOAT32" in dtype:
        v = struct.unpack(">f", struct.pack(">I", raw32))[0]
        return DecodeResult(f"{v:.6g}", 2)
    raise ValidationError(f"Unknown type: {dtype}")


def preview_pack(dtype: str, text: str) -> str:
    try:
        words = pack_value(dtype, text)
        return "→  " + "  ".join(f"0x{w:04X}" for w in words) + f"  ({len(words)} reg{'s' if len(words)>1 else ''})"
    except Exception:
        return ""


def validate_host(host: str) -> bool:
    """Returns True if host looks like a valid IP or hostname."""
    if not host: return False
    ip_re = re.compile(r"^(\d{1,3}\.){3}\d{1,3}$")
    if ip_re.match(host):
        parts = host.split(".")
        return all(0 <= int(p) <= 255 for p in parts)
    hostname_re = re.compile(r"^[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?)*$")
    return bool(hostname_re.match(host))

