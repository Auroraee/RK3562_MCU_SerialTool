#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RK3562 <-> MCU  UART 通信测试工具
模拟 RK3562 端，对 MCU 固件进行自测
Simulates RK3562 side for MCU firmware self-testing
"""

# ═══════════════════════════════════════════════════════════════════
#  DPI Awareness — MUST run before any GUI imports
# ═══════════════════════════════════════════════════════════════════
import sys
import ctypes
if sys.platform == 'win32':
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

try:
    import serial
    import serial.tools.list_ports
except ImportError as exc:
    raise RuntimeError("pyserial is required. Install it with: pip install pyserial") from exc

import os
import struct
import queue
import threading
import datetime
import time
import platform
import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox, filedialog, simpledialog

try:
    from PIL import Image, ImageTk
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False

APP_NAME = "RK3562 MCU UART Validation Tool"
APP_VERSION = "1.4.0"


# ═══════════════════════════════════════════════════════════════════
#  CRC-8 Calculation (matches MCU's Crc_CalculateCRC8 with SAE-J1850 settings)
# ═══════════════════════════════════════════════════════════════════
# CRC-8/SAE-J1850  Poly=0x1D  Init=0xFF  XorOut=0xFF
# Lookup table matches MCU's Cal_Crc8Tab exactly
_CRC8_TABLE = [
    0x00,0x1d,0x3a,0x27,0x74,0x69,0x4e,0x53,0xe8,0xf5,0xd2,0xcf,0x9c,0x81,0xa6,0xbb,
    0xcd,0xd0,0xf7,0xea,0xb9,0xa4,0x83,0x9e,0x25,0x38,0x1f,0x02,0x51,0x4c,0x6b,0x76,
    0x87,0x9a,0xbd,0xa0,0xf3,0xee,0xc9,0xd4,0x6f,0x72,0x55,0x48,0x1b,0x06,0x21,0x3c,
    0x4a,0x57,0x70,0x6d,0x3e,0x23,0x04,0x19,0xa2,0xbf,0x98,0x85,0xd6,0xcb,0xec,0xf1,
    0x13,0x0e,0x29,0x34,0x67,0x7a,0x5d,0x40,0xfb,0xe6,0xc1,0xdc,0x8f,0x92,0xb5,0xa8,
    0xde,0xc3,0xe4,0xf9,0xaa,0xb7,0x90,0x8d,0x36,0x2b,0x0c,0x11,0x42,0x5f,0x78,0x65,
    0x94,0x89,0xae,0xb3,0xe0,0xfd,0xda,0xc7,0x7c,0x61,0x46,0x5b,0x08,0x15,0x32,0x2f,
    0x59,0x44,0x63,0x7e,0x2d,0x30,0x17,0x0a,0xb1,0xac,0x8b,0x96,0xc5,0xd8,0xff,0xe2,
    0x26,0x3b,0x1c,0x01,0x52,0x4f,0x68,0x75,0xce,0xd3,0xf4,0xe9,0xba,0xa7,0x80,0x9d,
    0xeb,0xf6,0xd1,0xcc,0x9f,0x82,0xa5,0xb8,0x03,0x1e,0x39,0x24,0x77,0x6a,0x4d,0x50,
    0xa1,0xbc,0x9b,0x86,0xd5,0xc8,0xef,0xf2,0x49,0x54,0x73,0x6e,0x3d,0x20,0x07,0x1a,
    0x6c,0x71,0x56,0x4b,0x18,0x05,0x22,0x3f,0x84,0x99,0xbe,0xa3,0xf0,0xed,0xca,0xd7,
    0x35,0x28,0x0f,0x12,0x41,0x5c,0x7b,0x66,0xdd,0xc0,0xe7,0xfa,0xa9,0xb4,0x93,0x8e,
    0xf8,0xe5,0xc2,0xdf,0x8c,0x91,0xb6,0xab,0x10,0x0d,0x2a,0x37,0x64,0x79,0x5e,0x43,
    0xb2,0xaf,0x88,0x95,0xc6,0xdb,0xfc,0xe1,0x5a,0x47,0x60,0x7d,0x2e,0x33,0x14,0x09,
    0x7f,0x62,0x45,0x58,0x0b,0x16,0x31,0x2c,0x97,0x8a,0xad,0xb0,0xe3,0xfe,0xd9,0xc4,
]

def crc8(data: bytes) -> int:
    crc = 0xFF                          # CRC_INITVALUE8
    for byte in data:
        crc = _CRC8_TABLE[crc ^ byte]
    return crc ^ 0xFF                   # XOR CRC_XORVALUE8


# ═══════════════════════════════════════════════════════════════════
#  Protocol Constants
# ═══════════════════════════════════════════════════════════════════
FRAME_HEADER = 0x5A
FRAME_END    = 0xA5

FT_NO_ACK    = 0x01   # No ACK needed
FT_NEED_ACK  = 0x02   # Requires ACK
FT_ACK_OK    = 0x03   # ACK OK
FT_ACK_ERR   = 0x04   # ACK Error

FT_NAMES = {
    FT_NO_ACK:   "No ACK",
    FT_NEED_ACK: "Need ACK",
    FT_ACK_OK:   "ACK Ok",
    FT_ACK_ERR:  "ACK Error",
}

# (中文名称, direction, default_frame_type)
CMD_TABLE = {
    0x0000: ("查询外设状态",            "RK→MCU", FT_NO_ACK),
    0x0001: ("查询配置参数",            "RK→MCU", FT_NO_ACK),
    0x0002: ("RK3562 心跳",            "RK→MCU", FT_NO_ACK),
    0x0003: ("MCU 心跳",               "MCU→RK", FT_NO_ACK),
    0x0100: ("MCU 上报状态",           "MCU→RK", FT_NO_ACK),
    0x0101: ("设备类型",               "MCU→RK", FT_ACK_OK),
    0x0102: ("磁栅类型",               "MCU→RK", FT_ACK_OK),
    0x0103: ("人在传感器参数",          "MCU→RK", FT_ACK_OK),
    0x0104: ("霍尔传感器1 配置参数",    "MCU→RK", FT_ACK_OK),
    0x0105: ("霍尔传感器2 配置参数",    "MCU→RK", FT_ACK_OK),
    0x0106: ("设备 SN",                "MCU→RK", FT_ACK_OK),
    0x0107: ("清除使用时长",            "RK→MCU", FT_NO_ACK),
    0x0108: ("RK3562 版本号",          "RK→MCU", FT_NEED_ACK),
    0x0109: ("强制升级",               "MCU→RK", FT_ACK_OK),
    0x010A: ("LED 与护眼屏控制",       "RK→MCU", FT_NEED_ACK),
    0x010B: ("系统重启",               "MCU→RK", FT_ACK_OK),
    0x010C: ("Ultra 电机控制 (步进)",  "RK→MCU", FT_NEED_ACK),
    0x010D: ("Ultra 电机控制 (目标值)","RK→MCU", FT_NEED_ACK),
    0x010E: ("Ultra 电机调节结果",     "MCU→RK", FT_NO_ACK),
}


# ═══════════════════════════════════════════════════════════════════
#  Frame Builder
# ═══════════════════════════════════════════════════════════════════
def build_frame(frame_type: int, serial_count: int, cmd: int,
                payload: bytes = b'') -> bytes:
    """
    Frame layout (big-endian multi-byte fields):
      Header(1) | FrameType(1) | SerialCount(2) | CMD(2) |
      Length(2) | Payload(n) | CRC8(1) | End(1)
    CRC8 covers everything from Header through Payload.
    """
    hdr = struct.pack('>BBHHH',
                      FRAME_HEADER, frame_type, serial_count, cmd, len(payload))
    body = hdr + payload
    return body + bytes([crc8(body), FRAME_END])


# ═══════════════════════════════════════════════════════════════════
#  Frame Parser  (streaming buffer)
# ═══════════════════════════════════════════════════════════════════
def parse_from_buffer(buf: bytearray):
    """
    Try to extract one frame from buf.
    Returns:
        (frame_dict, n)  — valid frame, consume n bytes
        (None, n)        — n > 0: skip n garbage bytes; n == 0: need more data
    """
    # Find header
    idx = 0
    while idx < len(buf) and buf[idx] != FRAME_HEADER:
        idx += 1
    if idx > 0:
        return None, idx          # skip bytes before first 0x5A

    if len(buf) < 10:
        return None, 0            # need at least 10 bytes

    _, ft, sc, cmd, length = struct.unpack_from('>BBHHH', buf, 0)

    if length > 1024:             # sanity: max payload 1024
        return None, 1            # bad frame, skip this header byte

    total = 10 + length           # 8 header bytes + payload + crc + end
    if len(buf) < total:
        return None, 0            # wait for rest of frame

    raw = bytes(buf[:total])
    if raw[-1] != FRAME_END:
        return None, 1            # bad end byte, skip header

    payload   = raw[8 : 8 + length]
    cs_recv   = raw[8 + length]
    cs_calc   = crc8(raw[:8 + length])

    frame = {
        'ft':       ft,
        'ft_name':  FT_NAMES.get(ft, f'0x{ft:02X}'),
        'sc':       sc,
        'cmd':      cmd,
        'cmd_name': CMD_TABLE.get(cmd, (f'CMD_0x{cmd:04X}',))[0],
        'length':   length,
        'payload':  payload,
        'crc_ok':   cs_recv == cs_calc,
        'cs_recv':  cs_recv,
        'cs_calc':  cs_calc,
        'raw':      raw,
    }
    return frame, total


# ═══════════════════════════════════════════════════════════════════
#  Payload Decoder  (human-readable)
# ═══════════════════════════════════════════════════════════════════
def decode_payload(cmd: int, payload: bytes) -> str:
    """Returns a decoded description string, or empty string."""
    try:
        if cmd == 0x0100 and len(payload) >= 11:
            state_map = {
                0x01: "待机+前盖关闭",
                0x02: "待机+前盖开+平镜关",
                0x03: "待机+前盖开+平镜开",
                0x04: "工作+前盖开+平镜开(uint)",
            }
            brightness = {1: "低亮度", 2: "中亮度", 3: "高亮度"}
            color_temp  = {1: "冷光", 2: "常规", 3: "暖光"}
            raw_time    = (payload[6] << 8) | payload[7]
            person      = (raw_time >> 15) & 1
            usage_time  = raw_time & 0x7FFF
            parts = [
                f"设备状态: {state_map.get(payload[0], f'0x{payload[0]:02X}')}",
                f"LED: {'开' if payload[1] == 0x00 else '关'}",
                f"亮度: {brightness.get(payload[2], f'0x{payload[2]:02X}')}",
                f"色温: {color_temp.get(payload[3], f'0x{payload[3]:02X}')}",
                f"电机: {'开启' if payload[4] == 0x01 else ('关闭' if payload[4] == 0x00 else f'0x{payload[4]:02X}')}",
                f"像距: {payload[5]} 分米",
                f"人在: {'是' if person == 1 else '否'}",
                f"使用时长: {usage_time} 分钟",
                f"霍尔1 (LED): {'合盖/关灯' if payload[8] == 0x01 else '开盖/开灯'}",
                f"霍尔2 (遮光板): {'有磁场' if payload[9] == 0x01 else '无磁场'}",
                f"护眼屏: {'已接入' if payload[10] == 0x01 else '未接入'}",
            ]
            return "\n                ".join(parts)

        elif cmd == 0x0101 and len(payload) >= 2:
            dev_map = {0x00: "Air", 0x01: "Pro", 0x02: "Max",
                       0x03: "Ultra", 0x04: "LE Display", 0xFF: "无效"}
            gen = payload[1]
            gen_str = f"0x{gen:02X} (invalid)" if gen in (0x00, 0xFF) else f"0x{gen:02X}"
            return (f"设备类型: {dev_map.get(payload[0], f'0x{payload[0]:02X}')}  "
                    f"代际编码: {gen_str}")

        elif cmd == 0x0102 and len(payload) >= 1:
            type_map = {0x00: "对向反射膜", 0x01: "外采磁栅", 0x02: "自研磁栅"}
            return f"磁栅类型: {type_map.get(payload[0], f'0x{payload[0]:02X}')}"

        elif cmd == 0x0103 and len(payload) >= 2:
            parts = [
                f"门数: {payload[0]}",
                f"无人延时: {payload[1]} s",
            ]
            for i in range(9):
                off = 2 + i * 2
                if off + 1 >= len(payload):
                    break
                motion, static = payload[off], payload[off + 1]
                if motion == 0xFF and static == 0xFF:
                    parts.append(f"门{i}: 未使用")
                else:
                    parts.append(
                        f"门{i}: 运动灵敏度={motion}  "
                        f"静止灵敏度={static}"
                    )
            return "\n                ".join(parts)

        elif cmd in (0x0104, 0x0105) and len(payload) >= 2:
            sensor_id = "霍尔1" if cmd == 0x0104 else "霍尔2"
            return f"{sensor_id} 配置: Byte0=0x{payload[0]:02X}  Byte1=0x{payload[1]:02X}"

        elif cmd == 0x0106:
            try:
                sn = payload.rstrip(b'\x00').decode('ascii')
                return f"设备SN: {sn}"
            except Exception:
                return f"设备SN (hex): {payload.hex().upper()}"

        elif cmd == 0x0107 and len(payload) >= 1:
            return f"操作: {'清除使用时长' if payload[0] == 0xAA else f'无效值 0x{payload[0]:02X}'}"

        elif cmd == 0x0108 and len(payload) >= 3:
            return f"RK3562 版本: {payload[0]}.{payload[1]}.{payload[2]}"

        elif cmd == 0x0109 and len(payload) >= 1:
            return f"强制升级: {'进入强制升级' if payload[0] == 0xAA else f'无效值 0x{payload[0]:02X}'}"

        elif cmd == 0x010A and len(payload) >= 3:
            brightness = {1: "低", 2: "中", 3: "高"}
            color_temp  = {1: "冷光", 2: "常规", 3: "暖光"}
            return (f"LED: {'开' if payload[0] == 0x01 else '关'}  "
                    f"亮度: {brightness.get(payload[1], f'0x{payload[1]:02X}')}  "
                    f"色温: {color_temp.get(payload[2], f'0x{payload[2]:02X}')}")

        elif cmd == 0x010B and len(payload) >= 1:
            return f"系统重启: {'执行' if payload[0] == 0xAA else f'无效值 0x{payload[0]:02X}'}"

        elif cmd == 0x010C and len(payload) >= 2:
            direction = {0x01: "向上", 0x02: "向下"}
            return (f"方向: {direction.get(payload[0], f'0x{payload[0]:02X}')}  "
                    f"移动距离: {payload[1]} 分米")

        elif cmd == 0x010D and len(payload) >= 1:
            return f"目标像距: {payload[0]} 分米"

        elif cmd == 0x010E and len(payload) >= 2:
            return (f"调节结果: {'成功' if payload[0] == 0x01 else '失败'}  "
                    f"当前像距: {payload[1]} 分米")

    except Exception as exc:
        return f"[Decode error: {exc}]"

    return ""


# ═══════════════════════════════════════════════════════════════════
#  Payload Input Dialog
# ═══════════════════════════════════════════════════════════════════
class PayloadDialog(simpledialog.Dialog):
    """Generic dialog to input payload fields."""

    def __init__(self, parent, title: str,
                 fields: list):   # [(label, default, hint), ...]
        self.fields  = fields
        self.result  = None
        super().__init__(parent, title=title)

    def body(self, master):
        C = self.parent.C
        master.configure(bg=C["panel"])
        self.entries = []
        for i, (label, default, hint) in enumerate(self.fields):
            tk.Label(master, text=label, bg=C["panel"], fg="#000000",
                     font=("Microsoft YaHei UI", 9, "bold")).grid(
                row=i, column=0, sticky="w", padx=8, pady=4)
            e = tk.Entry(master, bg=C["card"], fg="#000000",
                         insertbackground="#000000",
                         font=("Microsoft YaHei UI", 10, "bold"), width=16)
            e.insert(0, default)
            e.grid(row=i, column=1, padx=6, pady=4)
            if hint:
                tk.Label(master, text=hint, bg=C["panel"], fg="#000000",
                         font=("Microsoft YaHei UI", 8, "bold")).grid(
                    row=i, column=2, padx=6, sticky="w")
            self.entries.append(e)
        return self.entries[0] if self.entries else None

    def apply(self):
        self.result = [e.get().strip() for e in self.entries]


# ═══════════════════════════════════════════════════════════════════
#  Main Application
# ═══════════════════════════════════════════════════════════════════
class App(tk.Tk):

    # ── Shared palette: keep the same Morandi light appearance in all OS themes ──
    THEME_C = {
        "bg":      "#faf8f5",   # warm white
        "panel":   "#f0ede7",   # warm light gray
        "card":    "#e6e2d9",   # greige card
        "hover":   "#dbd5ca",   # warm hover
        "fg":      "#3d3833",   # soft charcoal
        "fgdim":   "#b0a99e",   # muted warm gray
        "accent":  "#c4877b",   # dusty rose (errors / disconnect)
        "green":   "#7d9b7e",   # sage green (TX / CRC OK)
        "blue":    "#7b8fa0",   # muted steel blue (RX)
        "yellow":  "#b8a878",   # muted gold (info / warnings)
        "hb":      "#e5e1d7",   # heartbeat: almost invisible on warm white
        "hb_hex":  "#ede9e0",
    }

    # ─────────────────────────────────────────────────────────────
    @staticmethod
    def _detect_dark_mode() -> bool:
        """Return True if the OS is currently in dark mode."""
        system = platform.system()
        try:
            if system == "Windows":
                import winreg
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                winreg.CloseKey(key)
                return val == 0           # 0 = dark, 1 = light
            elif system == "Darwin":
                import subprocess
                r = subprocess.run(
                    ["defaults", "read", "-g", "AppleInterfaceStyle"],
                    capture_output=True, text=True)
                return r.stdout.strip().lower() == "dark"
            elif system == "Linux":
                import subprocess
                r = subprocess.run(
                    ["gsettings", "get",
                     "org.gnome.desktop.interface", "color-scheme"],
                    capture_output=True, text=True)
                return "dark" in r.stdout.lower()
        except Exception:
            pass
        return True   # default to dark if detection fails

    def __init__(self):
        # ── DPI scale factor (must be before super().__init__) ──
        self.dpi_scale = self._get_dpi_scale() if platform.system() == 'Windows' else 1.0

        super().__init__()
        self._apply_windows_dpi_scaling()

        self.title(f"{APP_NAME}  v{APP_VERSION}")
        self.minsize(self._s(960), self._s(640))

        # Theme — keep the same palette in both OS light/dark modes
        self._dark = self._detect_dark_mode()
        self.C = self.THEME_C.copy()
        self.configure(bg=self.C["bg"])

        self.ser: serial.Serial | None = None
        self.rx_thread: threading.Thread | None = None
        self.running  = False
        self.sc       = 0            # serial count
        self.rx_buf   = bytearray()
        self.log_q    = queue.Queue()
        self.stats    = {"tx": 0, "rx": 0, "err": 0}

        # heartbeat state
        self.hb_rk_enabled  = tk.BooleanVar(value=False)
        self.hb_rk_thread: threading.Thread | None = None
        self.hb_running = False

        self._apply_styles()
        self._load_bg_image()
        self._build_ui()
        self.update_idletasks()
        self._fit_initial_height()
        self._update_cmd_scroll_state()
        self._refresh_ports()
        self._pump_log()

    # ── Background image ─────────────────────────────────────────
    def _load_bg_image(self):
        """Load background image (bg.png) if available. Falls back silently."""
        self._bg_pil_orig = None
        self._bg_photo = None
        if not _HAS_PIL:
            return
        try:
            if getattr(sys, 'frozen', False):
                base = sys._MEIPASS
            else:
                base = os.path.dirname(os.path.abspath(__file__))
            path = os.path.join(base, "bg.png")
            if os.path.isfile(path):
                self._bg_pil_orig = Image.open(path)
        except Exception:
            self._bg_pil_orig = None

    def _get_dpi_scale(self) -> float:
        """获取主显示器的 DPI 缩放比例（1.0=100%, 2.0=200%）"""
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return dpi / 96.0
        except Exception:
            return 1.0

    def _s(self, value: int) -> int:
        """Scale：把设计像素按 DPI 比例缩放"""
        return int(value * self.dpi_scale)

    def _fit_initial_height(self):
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(self._s(1080), max(self._s(980), screen_w - 120))
        button_bottom = self.manual_send_btn.winfo_rooty() + self.manual_send_btn.winfo_height()
        window_top = self.winfo_rooty()
        frame_height = button_bottom - window_top + 20
        height = min(screen_h - 80, max(self._s(720), frame_height))
        self.geometry(f"{width}x{height}")

    def _apply_windows_dpi_scaling(self):
        if platform.system() != "Windows":
            return
        try:
            self.tk.call("tk", "scaling", self.winfo_fpixels("1i") / 72.0)
        except Exception:
            pass

    # ── Styles ────────────────────────────────────────────────────
    def _apply_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        C = self.C
        base = {"background": C["bg"], "foreground": C["fg"],
                "borderwidth": 0, "relief": "flat"}
        s.configure("TFrame",        **base)
        s.configure("TLabel",        **base)
        s.configure("Panel.TFrame",  background=C["panel"])
        s.configure("Panel.TLabel",  background=C["panel"], foreground=C["fg"])
        s.configure("TCombobox",
                    fieldbackground=C["card"], background=C["card"],
                    foreground=C["fg"], selectbackground=C["accent"],
                    selectforeground="#faf8f5", arrowcolor=C["fg"],
                    padding=(8, 6), arrowsize=18)
        s.map("TCombobox", fieldbackground=[("readonly", C["card"])],
              foreground=[("readonly", C["fg"])])
        for name, bg, fg in (
            ("Norm.TButton",   C["card"],   C["fg"]),
            ("Accent.TButton", C["accent"], "#e8e0d8"),
            ("Green.TButton",  "#7d9b7e",   "#faf8f5"),
            ("Blue.TButton",   C["blue"],   "#e8e0d8"),
        ):
            s.configure(name, background=bg, foreground=fg,
                        borderwidth=0, focusthickness=0, padding=(8, 8),
                        font=("Microsoft YaHei UI", 10))
            s.map(name,
                  background=[("active",   C["hover"]),
                               ("disabled", C["fgdim"])],
                  foreground=[("disabled",  C["fgdim"])])
        s.configure("TScrollbar", background=C["card"],
                    troughcolor=C["panel"], arrowcolor=C["fgdim"])

    # ── UI Construction ───────────────────────────────────────────
    def _build_ui(self):
        self._build_toolbar()

        body = tk.Frame(self, bg=self.C["bg"])
        body.pack(fill="both", expand=True, padx=8, pady=(4, 6))

        left = tk.Frame(body, bg=self.C["panel"], width=self._s(280))
        self.left_panel = left
        left.pack(side="left", fill="y", padx=(0, 6))
        left.pack_propagate(False)
        self._build_cmd_panel(left)

        right = tk.Frame(body, bg=self.C["bg"])
        right.pack(side="left", fill="both", expand=True)
        self._build_log_panel(right)

    # ── Toolbar ───────────────────────────────────────────────────
    def _build_toolbar(self):
        tb = tk.Frame(self, bg=self.C["panel"], height=self._s(50))
        tb.pack(fill="x", side="top")
        tb.pack_propagate(False)

        def lbl(text, **kw):
            return tk.Label(tb, text=text, bg=self.C["panel"],
                            fg=self.C["fgdim"], font=("Microsoft YaHei UI", 9), **kw)

        lbl("PORT").pack(side="left", padx=(14, 2), pady=14)
        self.port_var = tk.StringVar()
        self.port_cb  = ttk.Combobox(tb, textvariable=self.port_var,
                                      width=11, state="readonly")
        self.port_cb.pack(side="left", pady=10, padx=2)

        ttk.Button(tb, text="⟳", style="Norm.TButton", width=2,
                   command=self._refresh_ports).pack(side="left", padx=2, pady=10)

        lbl("BAUD").pack(side="left", padx=(10, 2))
        self.baud_var = tk.StringVar(value="115200")
        ttk.Combobox(tb, textvariable=self.baud_var, width=9,
                     values=["9600","19200","38400","57600",
                              "115200","230400","460800"],
                     state="readonly").pack(side="left", pady=10, padx=2)

        self.conn_btn = ttk.Button(tb, text="CONNECT",
                                    style="Green.TButton",
                                    command=self._toggle_connect)
        self.conn_btn.pack(side="left", padx=(18, 4), pady=10)

        self.dot = tk.Label(tb, text="●", bg=self.C["panel"],
                             fg="#333", font=("Microsoft YaHei UI", 18))
        self.dot.pack(side="left", padx=2)
        self.conn_lbl = tk.Label(tb, text="Disconnected",
                                  bg=self.C["panel"], fg=self.C["fgdim"],
                                  font=("Microsoft YaHei UI", 10))
        self.conn_lbl.pack(side="left", padx=4)

        # right side
        ttk.Button(tb, text="Save Log",
                   style="Norm.TButton",
                   command=self._save_log).pack(side="right", padx=4, pady=10)
        ttk.Button(tb, text="Clear Log",
                   style="Norm.TButton",
                   command=self._clear_log).pack(side="right", padx=4, pady=10)

        self.stat_lbl = tk.Label(tb, text="TX: 0  RX: 0  ERR: 0",
                                  bg=self.C["panel"], fg=self.C["fgdim"],
                                  font=("Microsoft YaHei UI", 10))
        self.stat_lbl.pack(side="right", padx=16)

    # ── Command panel ─────────────────────────────────────────────
    def _build_cmd_panel(self, parent):
        tk.Label(parent, text="COMMAND PANEL",
                 bg=self.C["panel"], fg=self.C["accent"],
                 font=("Microsoft YaHei UI", 12, "bold")).pack(
            pady=(10, 4), padx=12, anchor="w")

        # Scrollable inner frame. Scrolling is enabled only when commands overflow.
        canvas = tk.Canvas(parent, bg=self.C["panel"],
                            highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=self.C["panel"])
        self.cmd_canvas = canvas
        self.cmd_scrollbar = vsb
        self.cmd_inner = inner
        inner.bind("<Configure>", lambda e: self._update_cmd_scroll_state())
        self._cmd_canvas_win = canvas.create_window(
            (0, 0), window=inner, anchor="nw")
        # Keep inner frame width synced to canvas width
        canvas.bind("<Configure>",
                    lambda e: (
                        canvas.itemconfig(self._cmd_canvas_win, width=e.width),
                        self._update_cmd_scroll_state()))
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)

        # ─── Predefined Commands ────
        self._sh(inner, "RK3562 → MCU  (预定义命令)")

        for cmd, label in [
            (0x0000, "查询外设状态"),
            (0x0001, "查询配置参数"),
        ]:
            self._cb(inner, f"0x{cmd:04X}  {label}",
                     lambda c=cmd: self._send_simple(c))

        self._cb(inner, "0x0002  RK3562 心跳 (单次)",
                 lambda: self._send_simple(0x0002))

        # Heartbeat toggle
        hb_frame = tk.Frame(inner, bg=self.C["panel"])
        hb_frame.pack(fill="x", padx=8, pady=1)
        self._checkbutton(
            hb_frame, text="0x0002  自动心跳 (1 s)",
            variable=self.hb_rk_enabled,
            command=self._toggle_heartbeat,
            bg=self.C["panel"], fg=self.C["fg"],
        ).pack(fill="x", anchor="w", padx=2)

        self._cb(inner, "0x0107  清除使用时长",    self._send_clear_usage)
        self._cb(inner, "0x0108  上报版本号",       self._send_version)
        self._cb(inner, "0x010A  LED / 护眼屏控制", self._send_led_ctrl)
        self._cb(inner, "0x010C  电机控制 (步进)",  self._send_motor_step)
        self._cb(inner, "0x010D  电机控制 (目标值)", self._send_motor_target)

        # ─── ACK Response ────
        self._sh(inner, "发送 ACK 响应")
        self._cb(inner, "发送 ACK OK  (0x03)",  lambda: self._send_ack(FT_ACK_OK))
        self._cb(inner, "发送 ACK Err (0x04)",  lambda: self._send_ack(FT_ACK_ERR))

        # ─── Manual Frame ────
        self._sh(inner, "手动发送帧")

        pad = tk.Frame(inner, bg=self.C["panel"])
        pad.pack(fill="x", padx=8, pady=4)
        pad.grid_columnconfigure(1, weight=1)

        def row(r, label, widget):
            tk.Label(pad, text=label, bg=self.C["panel"], fg=self.C["fg"],
                     font=("Microsoft YaHei UI", 10), width=11, anchor="w").grid(
                row=r, column=0, sticky="w", pady=5, padx=(0, 8))
            widget.grid(row=r, column=1, padx=4, pady=5, sticky="ew")

        self.m_cmd = tk.Entry(pad, bg=self.C["card"], fg=self.C["green"],
                               insertbackground=self.C["green"],
                               font=("Microsoft YaHei UI", 10), width=8)
        self.m_cmd.insert(0, "0000")

        self.m_ft  = ttk.Combobox(pad,
                                   values=["01 - No ACK",
                                           "02 - Need ACK",
                                           "03 - ACK Ok",
                                           "04 - ACK Error"],
                                   width=14, height=8, state="readonly")
        self.m_ft.current(0)

        self.m_pl  = tk.Entry(pad, bg=self.C["card"], fg=self.C["green"],
                               insertbackground=self.C["green"],
                               font=("Microsoft YaHei UI", 10), width=12)

        row(0, "CMD (hex) :", self.m_cmd)
        row(1, "Frame Type:", self.m_ft)
        row(2, "Payload HEX:", self.m_pl)

        self.manual_send_btn = ttk.Button(inner, text="▶  发送手动帧",
                   style="Accent.TButton",
                   command=self._send_manual)
        self.manual_send_btn.pack(
            fill="x", padx=8, pady=(6, 18), ipady=2)

        # Bind mouse wheel to canvas and every inner widget so scrolling works
        # anywhere in the left panel, not just over the scrollbar.
        self._bind_all_cmd_mousewheel()

    def _sh(self, parent, text):
        """Section header separator."""
        f = tk.Frame(parent, bg=self.C["panel"])
        f.pack(fill="x", padx=8, pady=(10, 2))
        tk.Frame(f, bg=self.C["accent"], height=1).pack(fill="x")
        tk.Label(f, text=text, bg=self.C["panel"], fg=self.C["accent"],
                 font=("Microsoft YaHei UI", 9, "bold")).pack(anchor="w", pady=3)

    def _cmd_scroll_overflows(self) -> bool:
        if not hasattr(self, "cmd_canvas"):
            return False
        return self.cmd_inner.winfo_reqheight() > self.cmd_canvas.winfo_height()

    def _update_cmd_scroll_state(self):
        if not hasattr(self, "cmd_canvas"):
            return
        canvas = self.cmd_canvas
        canvas.configure(scrollregion=canvas.bbox("all"))
        if self._cmd_scroll_overflows():
            if not self.cmd_scrollbar.winfo_ismapped():
                canvas.pack_forget()
                self.cmd_scrollbar.pack(side="right", fill="y")
                canvas.pack(side="left", fill="both", expand=True)
            canvas.configure(yscrollcommand=self.cmd_scrollbar.set)
        else:
            if self.cmd_scrollbar.winfo_ismapped():
                self.cmd_scrollbar.pack_forget()
            canvas.yview_moveto(0)
            canvas.configure(yscrollcommand=lambda *args: None)

    def _bind_all_cmd_mousewheel(self):
        """Recursively bind mouse wheel to canvas and all descendant widgets."""
        def _recursive_bind(widget):
            widget.bind("<MouseWheel>", self._on_cmd_mousewheel)
            for child in widget.winfo_children():
                _recursive_bind(child)
        for root in (self.cmd_canvas, self.cmd_inner):
            _recursive_bind(root)

    def _on_cmd_mousewheel(self, event):
        if self._cmd_scroll_overflows():
            self.cmd_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

    def _checkbutton(self, parent, text, variable, command=None, bg=None, fg=None):
        """Large hit-target Checkbutton for toolbar and command-panel toggles."""
        bg = bg or self.C["bg"]
        fg = fg or self.C["fg"]
        return tk.Checkbutton(
            parent, text=text, variable=variable, command=command,
            bg=bg, fg=fg, selectcolor=self.C["card"],
            activebackground=bg, activeforeground=fg,
            font=("Microsoft YaHei UI", 11),
            anchor="w", justify="left",
            padx=12, pady=9,
            cursor="hand2",
        )

    def _cb(self, parent, text, cmd_func):
        """Command button."""
        b = tk.Button(
            parent, text=text, bg=self.C["card"], fg="#000000",
            activebackground=self.C["accent"], activeforeground="#000000",
            relief="flat", bd=0, cursor="hand2",
            font=("Microsoft YaHei UI", 10), anchor="w", justify="left", wraplength=self._s(232),
            padx=12, pady=8,
            command=cmd_func,
        )
        b.pack(fill="x", padx=8, pady=1)

    # ── Log panel ─────────────────────────────────────────────────
    def _build_log_panel(self, parent):
        hdr = tk.Frame(parent, bg=self.C["bg"])
        hdr.pack(fill="x", pady=(0, 4))
        tk.Label(hdr, text="COMMUNICATION LOG",
                 bg=self.C["bg"], fg=self.C["accent"],
                 font=("Microsoft YaHei UI", 12, "bold")).pack(side="left")

        self.show_tx = tk.BooleanVar(value=True)
        self.show_rx = tk.BooleanVar(value=True)
        for text, var, fg in (("TX", self.show_tx, self.C["yellow"]),
                               ("RX", self.show_rx, self.C["green"])):
            self._checkbutton(hdr, text=text, variable=var,
                              bg=self.C["bg"], fg=fg).pack(
                side="right", padx=4)

        self.log_canvas = tk.Canvas(parent, highlightthickness=0, bd=0,
                                     bg=self.C["panel"])
        vsb = ttk.Scrollbar(parent, orient="vertical",   command=self._canvas_yscroll)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=self._canvas_xscroll)
        self.log_canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.log_canvas.pack(fill="both", expand=True)

        # Canvas state
        self._log_font = tkFont.Font(family="Microsoft YaHei UI", size=10)
        self._line_height = self._log_font.metrics("linespace") + 2
        self._canvas_x = 0
        self._canvas_y = 2
        self._max_width = 0
        self._log_lines = []
        self._log_items = []
        self._log_item_lines = []   # list of lists, one per line
        self._MAX_LOG_ITEMS = 3000
        self._bg_canvas_item = None
        self._log_canvas_photo = None

        # Background image on canvas
        self._bg_resize_after = None
        if self._bg_pil_orig is not None:
            self.log_canvas.bind("<Configure>", self._schedule_bg_resize)
            self.after(1, lambda: self._on_log_canvas_resize(None))

        # Mouse wheel
        self.log_canvas.bind("<MouseWheel>", self._on_canvas_mousewheel)

        self._configure_log_tags()

    # ── Background image on log canvas ────────────────────────────
    def _schedule_bg_resize(self, event):
        """Debounce canvas resize to avoid lag during window drag."""
        if self._bg_resize_after is not None:
            self.after_cancel(self._bg_resize_after)
        self._bg_resize_after = self.after(150, lambda: self._on_log_canvas_resize(event))

    def _on_log_canvas_resize(self, event):
        """Resize background image to fit the canvas visible area."""
        w = self.log_canvas.winfo_width()
        h = self.log_canvas.winfo_height()
        if w < 2 or h < 2 or self._bg_pil_orig is None:
            return
        if hasattr(self, '_bg_cached_size') and self._bg_cached_size == (w, h):
            return
        self._bg_cached_size = (w, h)
        try:
            orig_w, orig_h = self._bg_pil_orig.size
            # Cover mode: scale to fill while keeping aspect ratio, crop overflow
            scale = max(w / orig_w, h / orig_h)
            new_w, new_h = int(orig_w * scale), int(orig_h * scale)
            resized = self._bg_pil_orig.resize((new_w, new_h), Image.Resampling.LANCZOS)
            # Center crop to target size
            left = (new_w - w) // 2
            top = (new_h - h) // 2
            resized = resized.crop((left, top, left + w, top + h))
            # Fade background for subtler appearance
            resized = resized.convert("RGBA")
            overlay = Image.new("RGBA", resized.size, (255, 255, 255, 200))
            resized = Image.alpha_composite(resized, overlay)
            self._log_canvas_photo = ImageTk.PhotoImage(resized)
            if self._bg_canvas_item is None:
                self._bg_canvas_item = self.log_canvas.create_image(
                    0, 0, anchor="nw", image=self._log_canvas_photo, tags=("bg",))
            else:
                self.log_canvas.itemconfig(self._bg_canvas_item, image=self._log_canvas_photo)
            self.log_canvas.tag_lower("bg")
            self._sync_bg_position()
        except Exception:
            pass

    def _sync_bg_position(self):
        """Keep background image fixed at the top-left of the visible viewport."""
        if self._bg_canvas_item is None:
            return
        self.log_canvas.coords(
            self._bg_canvas_item,
            self.log_canvas.canvasx(0), self.log_canvas.canvasy(0))

    # ── Canvas scrolling ──────────────────────────────────────────
    def _canvas_yscroll(self, *args):
        self.log_canvas.yview(*args)
        self._sync_bg_position()

    def _canvas_xscroll(self, *args):
        self.log_canvas.xview(*args)
        self._sync_bg_position()

    def _on_canvas_mousewheel(self, event):
        self.log_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self._sync_bg_position()
        return "break"

    def _canvas_scroll_to_end(self):
        self.log_canvas.yview_moveto(1.0)
        self._sync_bg_position()

    # ── Theme switching ───────────────────────────────────────────
    def _configure_log_tags(self):
        """Apply / re-apply all log colour tags — dark colours for light photo background."""
        self._tag_colors = {
            "ts":      "#4a4a4a",   # dark grey
            "tx_dir":  "#1b5e20",   # dark green
            "rx_dir":  "#0d47a1",   # dark blue
            "err":     "#b71c1c",   # dark red
            "info":    "#e65100",   # dark orange
            "hex":     "#4a4a4a",   # dark grey
            "crc_ok":  "#1b5e20",   # dark green
            "crc_err": "#b71c1c",   # dark red
            "decoded": "#0d47a1",   # dark blue
            "cmd":     "#1a1a1a",   # near-black
            "hb_dim":  "#bbbbbb",   # very light grey — heartbeat visually suppressed
            "hb_hex":  "#bbbbbb",   # very light grey
        }
        for item in self._log_items:
            tags = self.log_canvas.gettags(item)
            if tags and tags[0] in self._tag_colors:
                self.log_canvas.itemconfig(item, fill=self._tag_colors[tags[0]])

    def _apply_theme(self, dark: bool):
        """Keep the same palette when the OS theme changes."""
        self._dark = dark
        self.C = self.THEME_C.copy()
        self.configure(bg=self.C["bg"])

        # Re-apply ttk styles with shared colours
        self._apply_styles()

        # Re-apply log text tags
        self._configure_log_tags()

        # Update dot colour if disconnected
        if not (self.ser and self.ser.is_open):
            self.dot.configure(fg=self.C["fgdim"])

    def _retheme_widgets(self, widgets):
        """Recursively update bg/fg on classic tk widgets."""
        skip_types = (ttk.Combobox, ttk.Button, ttk.Scrollbar,
                      ttk.Frame, ttk.Label)
        for w in widgets:
            if isinstance(w, skip_types):
                self._retheme_widgets(w.winfo_children())
                continue
            self._retheme_widgets(w.winfo_children())

    # ── Port management ───────────────────────────────────────────
    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_cb["values"] = ports
        if ports and not self.port_var.get():
            self.port_var.set(ports[0])

    def _toggle_connect(self):
        if self.ser and self.ser.is_open:
            self._disconnect()
        else:
            self._connect()

    def _connect(self):
        port = self.port_var.get()
        if not port:
            messagebox.showerror("Error", "Please select a serial port")
            return
        try:
            self.ser = serial.Serial(
                port=port,
                baudrate=int(self.baud_var.get()),
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=0.05,
            )
            self.running = True
            self.rx_buf.clear()
            self.rx_thread = threading.Thread(
                target=self._rx_worker, daemon=True)
            self.rx_thread.start()

            self.conn_btn.configure(text="DISCONNECT", style="Accent.TButton")
            self.dot.configure(fg=self.C["green"])
            self.conn_lbl.configure(
                text=f"Connected  {port}  {self.baud_var.get()} 8N1",
                fg=self.C["green"])
            self._log_info(f"Connected to {port} @ {self.baud_var.get()} bps  8N1")
        except Exception as exc:
            messagebox.showerror("Connection Error", str(exc))

    def _disconnect(self):
        self.running = False
        self._stop_heartbeat()
        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None
        self.conn_btn.configure(text="CONNECT", style="Green.TButton")
        self.dot.configure(fg=self.C["fgdim"])
        self.conn_lbl.configure(text="Disconnected", fg=self.C["fgdim"])
        self._log_info("Disconnected")

    # ── RX worker (background thread) ────────────────────────────
    def _rx_worker(self):
        while self.running:
            try:
                if self.ser and self.ser.is_open:
                    waiting = self.ser.in_waiting
                    if waiting:
                        chunk = self.ser.read(waiting)
                        if chunk:
                            self.rx_buf.extend(chunk)
                            self._drain()
                else:
                    time.sleep(0.02)
            except serial.SerialException as exc:
                self.log_q.put(("err", f"Serial error: {exc}"))
                self.running = False
                break
            except Exception as exc:
                self.log_q.put(("err", f"RX error: {exc}"))

    def _drain(self):
        garbage = bytearray()
        while self.rx_buf:
            frame, n = parse_from_buffer(self.rx_buf)
            if n == 0:
                # Need more data — flush accumulated garbage first
                if garbage:
                    self.log_q.put(("raw", bytes(garbage)))
                    garbage.clear()
                break
            if frame is None:
                # Skip n unrecognisable bytes — batch them into one raw log entry
                garbage.extend(self.rx_buf[:n])
                del self.rx_buf[:n]
                continue
            # Valid frame (CRC may still be bad — _log_frame will show it)
            if garbage:
                self.log_q.put(("raw", bytes(garbage)))
                garbage.clear()
            del self.rx_buf[:n]
            self.stats["rx"] += 1
            self.log_q.put(("rx", frame))
        if garbage:
            self.log_q.put(("raw", bytes(garbage)))
            garbage.clear()

    # ── Heartbeat ─────────────────────────────────────────────────
    def _toggle_heartbeat(self):
        if self.hb_rk_enabled.get():
            if self.hb_rk_thread and self.hb_rk_thread.is_alive():
                return
            self.hb_running = True
            self.hb_rk_thread = threading.Thread(
                target=self._hb_worker, daemon=True)
            self.hb_rk_thread.start()
        else:
            self._stop_heartbeat()

    def _stop_heartbeat(self):
        self.hb_running = False
        self.hb_rk_enabled.set(False)

    def _hb_worker(self):
        while self.hb_running and self.running:
            self._send_simple(0x0002)
            time.sleep(1.0)

    # ── Frame send helpers ────────────────────────────────────────
    def _next_sc(self) -> int:
        v = self.sc
        self.sc = (self.sc + 1) & 0xFFFF
        return v

    def _send_frame(self, ft: int, cmd: int,
                    payload: bytes = b'', silent: bool = False) -> bool:
        if not self.ser or not self.ser.is_open:
            if not silent:
                messagebox.showwarning("Not Connected",
                                       "Please connect to a serial port first.")
            return False
        sc    = self._next_sc()
        frame = build_frame(ft, sc, cmd, payload)
        try:
            self.ser.write(frame)
            self.stats["tx"] += 1
            if not silent:
                self.log_q.put(("tx", {
                    "ft":       ft,
                    "ft_name":  FT_NAMES.get(ft, f"0x{ft:02X}"),
                    "sc":       sc,
                    "cmd":      cmd,
                    "cmd_name": CMD_TABLE.get(cmd, (f"CMD_0x{cmd:04X}",))[0],
                    "length":   len(payload),
                    "payload":  payload,
                    "crc_ok":   True,
                    "raw":      frame,
                }))
            return True
        except Exception as exc:
            self.log_q.put(("err", f"TX error: {exc}"))
            self.stats["err"] += 1
            return False

    # ── Predefined send actions ───────────────────────────────────
    def _send_simple(self, cmd: int, silent: bool = False):
        _, _, ft = CMD_TABLE.get(cmd, ("", "", FT_NO_ACK))
        self._send_frame(ft, cmd, silent=silent)

    def _send_clear_usage(self):
        self._send_frame(FT_NO_ACK, 0x0107, bytes([0xAA]))

    def _send_version(self):
        d = PayloadDialog(self, "RK3562 版本号  CMD 0x0108", [
            ("主版本号 (Major)", "1", "0 – 255"),
            ("次版本号 (Minor)", "0", "0 – 255"),
            ("补丁版本号 (Patch)","0", "0 – 255"),
        ])
        if d.result:
            try:
                payload = bytes([int(x, 0) for x in d.result])
                self._send_frame(FT_NEED_ACK, 0x0108, payload)
            except (ValueError, OverflowError):
                messagebox.showerror("Error", "版本号必须为 0-255 整数")

    def _send_led_ctrl(self):
        d = PayloadDialog(self, "LED 与护眼屏控制  CMD 0x010A", [
            ("Byte0  LED 开关", "1",  "0x00: 关  0x01: 开"),
            ("Byte1  亮度",     "2",  "0x01:低  0x02:中  0x03:高"),
            ("Byte2  色温",     "2",  "0x01:冷光  0x02:常规  0x03:暖光"),
        ])
        if d.result:
            try:
                payload = bytes([int(x, 0) for x in d.result])
                self._send_frame(FT_NEED_ACK, 0x010A, payload)
            except (ValueError, OverflowError):
                messagebox.showerror("Error", "输入值无效")

    def _send_motor_step(self):
        d = PayloadDialog(self, "Ultra 电机控制 (步进)  CMD 0x010C", [
            ("Byte0  方向",          "1",  "0x01:向上  0x02:向下"),
            ("Byte1  移动距离 (分米)", "10", "整数，如 10"),
        ])
        if d.result:
            try:
                payload = bytes([int(x, 0) for x in d.result])
                self._send_frame(FT_NEED_ACK, 0x010C, payload)
            except (ValueError, OverflowError):
                messagebox.showerror("Error", "输入值无效")

    def _send_motor_target(self):
        d = PayloadDialog(self, "Ultra 电机控制 (目标值)  CMD 0x010D", [
            ("Byte0  目标像距 (分米)", "15", "整数，如 15"),
        ])
        if d.result:
            try:
                payload = bytes([int(x, 0) for x in d.result])
                self._send_frame(FT_NEED_ACK, 0x010D, payload)
            except (ValueError, OverflowError):
                messagebox.showerror("Error", "输入值无效")

    def _send_ack(self, ft: int):
        """Send ACK response; user inputs SC of the frame being ACK'd."""
        sc_str = simpledialog.askstring(
            "发送 ACK",
            "输入要 ACK 的 Serial Count（十进制或 0x 十六进制）：",
            parent=self,
        )
        if sc_str is None:
            return
        cmd_str = simpledialog.askstring(
            "发送 ACK",
            "对应 CMD（十六进制，如 010A）：",
            parent=self,
        )
        if cmd_str is None:
            return
        try:
            target_sc  = int(sc_str.strip(), 0)
            target_cmd = int(cmd_str.strip(), 16)
        except ValueError:
            messagebox.showerror("Error", "输入值无效")
            return
        frame = build_frame(ft, target_sc, target_cmd)
        try:
            self.ser.write(frame)
            self.stats["tx"] += 1
            self.log_q.put(("tx", {
                "ft":       ft,
                "ft_name":  FT_NAMES.get(ft),
                "sc":       target_sc,
                "cmd":      target_cmd,
                "cmd_name": CMD_TABLE.get(target_cmd, (f"CMD_0x{target_cmd:04X}",))[0],
                "length":   0,
                "payload":  b'',
                "crc_ok":   True,
                "raw":      frame,
            }))
        except Exception as exc:
            self.log_q.put(("err", f"TX error: {exc}"))

    def _send_manual(self):
        try:
            cmd = int(self.m_cmd.get().strip(), 16)
        except ValueError:
            messagebox.showerror("Error", "CMD 格式错误，请输入十六进制，如 010A")
            return
        ft = int(self.m_ft.get().split()[0])
        raw_pl = self.m_pl.get().strip().replace(" ", "")
        try:
            payload = bytes.fromhex(raw_pl) if raw_pl else b''
        except ValueError:
            messagebox.showerror("Error", "Payload 十六进制格式错误")
            return
        self._send_frame(ft, cmd, payload)

    # ── Log / display ─────────────────────────────────────────────
    def _pump_log(self):
        """Drain log queue in main thread."""
        try:
            while True:
                kind, data = self.log_q.get_nowait()
                if kind == "tx":
                    self._log_frame(data, tx=True)
                elif kind == "rx":
                    self._log_frame(data, tx=False)
                elif kind == "raw":
                    self._log_raw(data)
                elif kind == "info":
                    self._log_info(data)
                self._update_stats()
        except queue.Empty:
            pass
        self.after(40, self._pump_log)

    @staticmethod
    def _ts() -> str:
        return datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]

    def _write(self, *parts):
        """parts: [(text, tag), ...]"""
        line_y = self._canvas_y
        line_text = ""
        line_items = []
        for text, tag in parts:
            if not text:
                continue
            fill = self._tag_colors.get(tag, self.C["fg"])
            item = self.log_canvas.create_text(
                self._canvas_x, line_y,
                text=text, anchor="nw",
                font=self._log_font, fill=fill,
                tags=(tag,)
            )
            self._log_items.append(item)
            line_items.append(item)
            text_width = self._log_font.measure(text)
            self._canvas_x += text_width
            self._max_width = max(self._max_width, self._canvas_x)
            line_text += text
        self._canvas_y += self._line_height
        self._canvas_x = 0
        self.log_canvas.configure(
            scrollregion=(0, 0, max(self._max_width + 20, self.log_canvas.winfo_width()),
                          self._canvas_y + 20)
        )
        self._canvas_scroll_to_end()
        self._log_lines.append(line_text)
        self._log_item_lines.append(line_items)
        # Prune oldest lines when item count exceeds limit
        while len(self._log_items) > self._MAX_LOG_ITEMS and self._log_item_lines:
            old_line = self._log_item_lines.pop(0)
            self._log_lines.pop(0)
            for it in old_line:
                self.log_canvas.delete(it)
                self._log_items.remove(it)

    def _log_frame(self, f: dict, tx: bool):
        if tx and not self.show_tx.get():
            return
        if not tx and not self.show_rx.get():
            return

        is_hb    = f["cmd"] in (0x0002, 0x0003)   # heartbeat — visually dimmed
        dir_str  = "TX ▶" if tx else "◀ RX"
        crc_str  = "✓ CRC OK" if f["crc_ok"] else "✗ CRC ERR"
        raw_hex  = " ".join(f'{b:02X}' for b in f["raw"])
        pl_hex   = f["payload"].hex().upper() if f["payload"] else "(empty)"

        if is_hb:
            # Single compact line, all in dim color — heartbeats should not dominate
            self._write(
                (f"[{self._ts()}] ", "hb_dim"),
                (f"[{dir_str}] ",   "hb_dim"),
                (f"CMD:0x{f['cmd']:04X} {f['cmd_name']}  "
                 f"SC:{f['sc']}  {crc_str}  ", "hb_dim"),
                (raw_hex, "hb_hex"),
            )
            return

        dir_tag  = "tx_dir" if tx else "rx_dir"
        crc_tag  = "crc_ok" if f["crc_ok"] else "crc_err"

        # Header line
        self._write(
            (f"[{self._ts()}] ", "ts"),
            (f"[{dir_str}] ", dir_tag),
            (f"CMD:0x{f['cmd']:04X} ", "cmd"),
            (f"{f['cmd_name']}  ", dir_tag),
            (f"Type:{f['ft_name']}  SC:{f['sc']}  Len:{f['length']}  ", "ts"),
            (crc_str, crc_tag),
        )

        if f["payload"]:
            self._write(
                ("         Payload : ", "ts"),
                (pl_hex,               "hex"),
            )
            decoded = decode_payload(f["cmd"], f["payload"])
            if decoded:
                for line in decoded.split("\n"):
                    self._write(
                        ("         Decoded : ", "ts"),
                        (line.strip(),          "decoded"),
                    )

        self._write(
            ("         Raw HEX : ", "ts"),
            (raw_hex,               "hex"),
        )
        self._write(("", ""))   # blank separator line

    def _log_raw(self, data: bytes):
        """Log raw bytes that could not be parsed into a valid frame."""
        if not self.show_rx.get():
            return
        raw_hex = " ".join(f"{b:02X}" for b in data)
        self._write(
            (f"[{self._ts()}] ", "ts"),
            ("◀ RX] ", "rx_dir"),
            ("[RAW / UNPARSED]  ", "err"),
            (f"{len(data)} bytes: ", "ts"),
            (raw_hex, "hex"),
        )
        self._write(("", ""))

    def _log_info(self, msg: str):
        self._write(
            (f"[{self._ts()}] ", "ts"),
            (f"[INFO]  {msg}",   "info"),
        )

    def _log_err(self, msg: str):
        self._write(
            (f"[{self._ts()}] ", "ts"),
            (f"[ERROR] {msg}",   "err"),
        )

    def _update_stats(self):
        self.stat_lbl.configure(
            text=f"TX: {self.stats['tx']}  "
                 f"RX: {self.stats['rx']}  "
                 f"ERR: {self.stats['err']}")

    def _clear_log(self):
        for item in self._log_items:
            self.log_canvas.delete(item)
        self._log_items.clear()
        self._log_item_lines.clear()
        self._log_lines.clear()
        self._canvas_x = 0
        self._canvas_y = 2
        self._max_width = 0
        self.log_canvas.configure(scrollregion=(0, 0, 1, 1))
        self._sync_bg_position()

    def _save_log(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"uart_log_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt",
        )
        if not path:
            return
        content = "\n".join(self._log_lines) + "\n"
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(content)
            self._log_info(f"Log saved to: {path}")
        except Exception as exc:
            messagebox.showerror("Save Error", str(exc))

    # ── Clean exit ────────────────────────────────────────────────
    def on_close(self):
        self._disconnect()
        self.destroy()


# ═══════════════════════════════════════════════════════════════════
#  Entry point
# ═══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
