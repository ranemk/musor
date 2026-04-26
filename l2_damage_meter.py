from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
import sys
import threading
import time
import ctypes
from collections import Counter
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from ctypes import wintypes


def _configure_tcl_paths() -> None:
    def tcl_path(path: Path) -> str:
        value = str(path.resolve())
        if os.name == "nt" and not value.startswith("\\\\?\\"):
            return "\\\\?\\" + value
        return value

    bases: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bases.append(Path(meipass))
    if getattr(sys, "frozen", False):
        bases.append(Path(sys.executable).resolve().parent)
    else:
        bases.append(Path(__file__).resolve().parent)

    for base in bases:
        tcl_root = base / "tcl"
        tcl_dir = tcl_root / "tcl8.6"
        tk_dir = tcl_root / "tk8.6"
        if tcl_dir.exists() and tk_dir.exists():
            os.environ["TCL_LIBRARY"] = tcl_path(tcl_dir)
            os.environ["TK_LIBRARY"] = tcl_path(tk_dir)
            return


_configure_tcl_paths()

from tkinter import BOTH, BOTTOM, END, LEFT, RIGHT, TOP, X, Button, Entry, Frame, Label, Listbox, StringVar, Tk, Toplevel, messagebox

from PIL import Image, ImageEnhance, ImageGrab, ImageTk


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    RESOURCE_DIR = APP_DIR
CONFIG_PATH = APP_DIR / "l2_damage_meter_config.json"
CAPTURE_PATH = APP_DIR / "last_chat_capture.png"
OCR_TEXT_PATH = APP_DIR / "last_ocr_text.txt"
WINDOWS_OCR_SCRIPT = RESOURCE_DIR / "win_ocr.ps1"
ASSET_DIR = RESOURCE_DIR / "assets"
APP_ICON_PATH = ASSET_DIR / "parashaoly_circlet_x.ico"
BARRIER_ICON_PATH = ASSET_DIR / "skill1496_servitor_barrier.png"
EMPOWER_ICON_PATH = ASSET_DIR / "skill1299_servitor_empowerment.png"

CREATE_NO_WINDOW = 0x08000000
GA_ROOT = 2

THEME = {
    "bg_dark": "#0f0d0a",
    "bg_panel": "#1c1812",
    "bg_button": "#2a2218",
    "bg_button_active": "#4a3a22",
    "bg_listbox": "#0a0806",
    "fg_text": "#d4c89a",
    "fg_title": "#c9a961",
    "fg_dim": "#8a7a5a",
    "fg_select": "#f0d878",
    "border": "#5a4a2a",
    "accent_red": "#a83232",
    "accent_green": "#5a8a3a",
}
TITLE_FONT = ("Book Antiqua", 11, "bold")

POLL_SECONDS = 0.8
MIN_POLL_SECONDS = 0.15
MAX_POLL_SECONDS = 5.0
MIN_OVERLAY_FPS = 1
MAX_OVERLAY_FPS = 30
BASELINE_SECONDS = 3.0
WINDOW_SECONDS = (5, 10, 30)
SERVITOR_TIMERS = {
    "barrier": {"name": "Barrier", "duration": 30, "icon": BARRIER_ICON_PATH},
    "empower": {"name": "Empower", "duration": 60, "icon": EMPOWER_ICON_PATH},
}


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


USER32 = ctypes.windll.user32
USER32.GetForegroundWindow.restype = wintypes.HWND
USER32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
USER32.GetAncestor.restype = wintypes.HWND
USER32.IsChild.argtypes = [wintypes.HWND, wintypes.HWND]
USER32.IsChild.restype = wintypes.BOOL
USER32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
USER32.GetWindowTextLengthW.restype = ctypes.c_int
USER32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
USER32.GetWindowTextW.restype = ctypes.c_int
USER32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(RECT)]
USER32.GetWindowRect.restype = wintypes.BOOL
USER32.IsWindow.argtypes = [wintypes.HWND]
USER32.IsWindow.restype = wintypes.BOOL
USER32.IsWindowVisible.argtypes = [wintypes.HWND]
USER32.IsWindowVisible.restype = wintypes.BOOL
USER32.GetWindowDC.argtypes = [wintypes.HWND]
USER32.GetWindowDC.restype = wintypes.HDC
USER32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
USER32.ReleaseDC.restype = ctypes.c_int
USER32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
USER32.PrintWindow.restype = wintypes.BOOL

GDI32 = ctypes.windll.gdi32
GDI32.CreateCompatibleDC.argtypes = [wintypes.HDC]
GDI32.CreateCompatibleDC.restype = wintypes.HDC
GDI32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
GDI32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
GDI32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
GDI32.SelectObject.restype = wintypes.HGDIOBJ
GDI32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
GDI32.DeleteObject.restype = wintypes.BOOL
GDI32.DeleteDC.argtypes = [wintypes.HDC]
GDI32.DeleteDC.restype = wintypes.BOOL
GDI32.GetDIBits.argtypes = [
    wintypes.HDC,
    wintypes.HBITMAP,
    wintypes.UINT,
    wintypes.UINT,
    wintypes.LPVOID,
    wintypes.LPVOID,
    wintypes.UINT,
]
GDI32.GetDIBits.restype = ctypes.c_int


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", wintypes.DWORD * 3),
    ]


def get_window_title(hwnd: int) -> str:
    length = USER32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buffer = ctypes.create_unicode_buffer(length + 1)
    USER32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value.strip()


def get_root_window(hwnd: int) -> int:
    if not hwnd:
        return 0
    root = int(USER32.GetAncestor(hwnd, GA_ROOT))
    return root or hwnd


def get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    if not hwnd or not USER32.IsWindow(hwnd):
        return None
    rect = RECT()
    if not USER32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None
    return (rect.left, rect.top, rect.right, rect.bottom)


def find_window_by_title(title: str) -> int | None:
    matches = find_windows_by_title(title)
    return matches[0] if matches else None


def find_windows_by_title(title: str) -> list[int]:
    if not title:
        return []

    found: list[int] = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def enum_proc(hwnd: int, _lparam: int) -> bool:
        if USER32.IsWindowVisible(hwnd) and get_window_title(hwnd) == title:
            found.append(get_root_window(hwnd))
        return True

    USER32.EnumWindows(enum_proc, 0)
    return list(dict.fromkeys(found))


def rect_distance(a: tuple[int, int, int, int] | None, b: tuple[int, int, int, int] | None) -> int:
    if not a or not b:
        return 1_000_000_000
    return sum(abs(a[index] - b[index]) for index in range(4))


def rectangles_close(a: tuple[int, int, int, int] | None, b: tuple[int, int, int, int] | None, tolerance: int = 40) -> bool:
    return rect_distance(a, b) <= tolerance


def rectangles_intersect(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return a[0] < b[2] and a[2] > b[0] and a[1] < b[3] and a[3] > b[1]


def capture_window_image(hwnd: int) -> Image.Image | None:
    rect = get_window_rect(hwnd)
    if not rect:
        return None

    left, top, right, bottom = rect
    width = max(1, right - left)
    height = max(1, bottom - top)

    window_dc = USER32.GetWindowDC(hwnd)
    if not window_dc:
        return None

    memory_dc = GDI32.CreateCompatibleDC(window_dc)
    bitmap = GDI32.CreateCompatibleBitmap(window_dc, width, height)
    previous_object = None

    try:
        previous_object = GDI32.SelectObject(memory_dc, bitmap)
        rendered = USER32.PrintWindow(hwnd, memory_dc, 2)
        if not rendered:
            rendered = USER32.PrintWindow(hwnd, memory_dc, 0)
        if not rendered:
            return None

        bitmap_info = BITMAPINFO()
        bitmap_info.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bitmap_info.bmiHeader.biWidth = width
        bitmap_info.bmiHeader.biHeight = -height
        bitmap_info.bmiHeader.biPlanes = 1
        bitmap_info.bmiHeader.biBitCount = 32
        bitmap_info.bmiHeader.biCompression = 0

        buffer_size = width * height * 4
        buffer = ctypes.create_string_buffer(buffer_size)
        lines = GDI32.GetDIBits(memory_dc, bitmap, 0, height, buffer, ctypes.byref(bitmap_info), 0)
        if lines == 0:
            return None

        image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1)
        return image.convert("RGB")
    finally:
        if previous_object:
            GDI32.SelectObject(memory_dc, previous_object)
        if bitmap:
            GDI32.DeleteObject(bitmap)
        if memory_dc:
            GDI32.DeleteDC(memory_dc)
        USER32.ReleaseDC(hwnd, window_dc)


@dataclass(frozen=True)
class DamageEvent:
    timestamp: float
    amount: int
    line: str
    direction: str
    target: str | None = None


class DamageParser:
    AMOUNT_TOKEN = r"([0-9oOIl|sSbB]{1,12})"
    DAMAGE_PATTERNS = [
        ("out", re.compile(rf"вы\s+нанесли\s+{AMOUNT_TOKEN}\s+урона", re.IGNORECASE)),
        ("out", re.compile(rf"нанесли\s+{AMOUNT_TOKEN}\s+урона", re.IGNORECASE)),
        ("out", re.compile(rf"you\s+(?:have\s+)?(?:dealt|inflicted)\s+{AMOUNT_TOKEN}\s+damage", re.IGNORECASE)),
        ("out", re.compile(rf"\b\w+\s+has\s+given\s+{AMOUNT_TOKEN}\s+damage\s+of\s+([^.\r\n]+)", re.IGNORECASE)),
        ("in", re.compile(rf"\b\w+\s+has\s+received\s+{AMOUNT_TOKEN}\s+damage\s+from\s+([^.\r\n]+)", re.IGNORECASE)),
    ]

    @staticmethod
    def normalize_line(line: str) -> str:
        line = line.strip()
        line = line.replace("|", "I")
        line = re.sub(r"\s+", " ", line)
        return line

    def parse_lines(self, text: str) -> list[tuple[str, int, str, str | None]]:
        parsed: list[tuple[str, int, str, str | None]] = []
        normalized_text = self.normalize_line(text)

        for direction, pattern in self.DAMAGE_PATTERNS:
            for match in pattern.finditer(normalized_text):
                phrase = self.normalize_line(match.group(0))
                target = match.group(2).strip() if match.lastindex and match.lastindex >= 2 else None
                amount = self.normalize_amount(match.group(1))
                if amount is not None:
                    parsed.append((phrase, amount, direction, target))

        if parsed:
            return sorted(parsed, key=lambda item: normalized_text.find(item[0]))

        for raw_line in text.splitlines():
            line = self.normalize_line(raw_line)
            if not line:
                continue

            for direction, pattern in self.DAMAGE_PATTERNS:
                match = pattern.search(line)
                if match:
                    target = match.group(2).strip() if match.lastindex and match.lastindex >= 2 else None
                    amount = self.normalize_amount(match.group(1))
                    if amount is not None:
                        parsed.append((line, amount, direction, target))
                    break
        return parsed

    @staticmethod
    def normalize_amount(raw: str) -> int | None:
        translation = str.maketrans({
            "o": "0",
            "O": "0",
            "I": "1",
            "l": "1",
            "|": "1",
            "s": "5",
            "S": "5",
            "b": "6",
            "B": "8",
        })
        normalized = raw.translate(translation)
        normalized = re.sub(r"\D", "", normalized)
        if not normalized:
            return None
        return int(normalized)


class TesseractOcr:
    def __init__(self) -> None:
        self.path = self._find_tesseract()

    @staticmethod
    def _find_tesseract() -> str | None:
        configured = os.environ.get("TESSERACT_CMD")
        if configured and Path(configured).exists():
            return configured

        found = shutil.which("tesseract")
        if found:
            return found

        candidates = [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
        for candidate in candidates:
            if candidate.exists():
                return str(candidate)
        return None

    def available(self) -> bool:
        return self.path is not None

    def read(self, image: Image.Image) -> str:
        if not self.path:
            raise RuntimeError(
                "Tesseract OCR was not found. Install Tesseract OCR with the Russian language pack, "
                "or set TESSERACT_CMD to tesseract.exe."
            )

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        try:
            self._preprocess(image).save(tmp_path)
            result = subprocess.run(
                [
                    self.path,
                    str(tmp_path),
                    "stdout",
                    "-l",
                    "rus+eng",
                    "--psm",
                    "6",
                ],
                capture_output=True,
                text=True,
                timeout=8,
                check=False,
                creationflags=CREATE_NO_WINDOW,
            )

            if result.returncode != 0:
                fallback = subprocess.run(
                    [self.path, str(tmp_path), "stdout", "-l", "eng", "--psm", "6"],
                    capture_output=True,
                    text=True,
                    timeout=8,
                    check=False,
                    creationflags=CREATE_NO_WINDOW,
                )
                if fallback.returncode == 0:
                    return fallback.stdout
                raise RuntimeError((result.stderr or fallback.stderr or "OCR failed").strip())

            return result.stdout
        finally:
            tmp_path.unlink(missing_ok=True)

    @staticmethod
    def _preprocess(image: Image.Image) -> Image.Image:
        image = image.convert("RGB")
        scale = 3
        image = image.resize((image.width * scale, image.height * scale), Image.Resampling.LANCZOS)
        image = ImageEnhance.Contrast(image).enhance(1.8)
        image = ImageEnhance.Sharpness(image).enhance(1.6)
        return image


class WindowsOcr:
    def __init__(self) -> None:
        self.path = shutil.which("powershell.exe") or str(Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe")

    def available(self) -> bool:
        return Path(self.path).exists() and WINDOWS_OCR_SCRIPT.exists()

    def read(self, image: Image.Image) -> str:
        if not self.available():
            raise RuntimeError("Windows OCR is not available.")

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        try:
            self._preprocess(image).save(tmp_path)
            result = subprocess.run(
                [
                    self.path,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(WINDOWS_OCR_SCRIPT),
                    "-ImagePath",
                    str(tmp_path),
                    "-Language",
                    "ru",
                ],
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                timeout=10,
                check=False,
                creationflags=CREATE_NO_WINDOW,
            )
            if result.returncode != 0:
                raise RuntimeError((result.stderr or "Windows OCR failed").strip())
            return result.stdout
        finally:
            tmp_path.unlink(missing_ok=True)

    @staticmethod
    def _preprocess(image: Image.Image) -> Image.Image:
        image = image.convert("RGB")
        scale = 2
        image = image.resize((image.width * scale, image.height * scale), Image.Resampling.LANCZOS)
        image = ImageEnhance.Contrast(image).enhance(1.6)
        image = ImageEnhance.Sharpness(image).enhance(1.4)
        return image


class OcrReader:
    def __init__(self) -> None:
        self.backends = [WindowsOcr(), TesseractOcr()]
        self.active_name = "none"

    def available(self) -> bool:
        return any(backend.available() for backend in self.backends)

    def read(self, image: Image.Image) -> str:
        errors = []
        for backend in self.backends:
            if not backend.available():
                continue
            try:
                self.active_name = backend.__class__.__name__.replace("Ocr", " OCR")
                return backend.read(image)
            except Exception as exc:  # noqa: BLE001 - fallback to the next OCR backend.
                errors.append(f"{backend.__class__.__name__}: {exc}")

        raise RuntimeError("No OCR backend succeeded. " + " | ".join(errors))


class DamageMeter:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("parashaoly")
        self._set_window_icon(self.root)
        self.root.geometry("660x650")
        self.root.minsize(560, 520)
        self._apply_theme()

        self.parser = DamageParser()
        self.ocr = OcrReader()
        self.region: tuple[int, int, int, int] | None = None
        self.events: deque[DamageEvent] = deque(maxlen=5000)
        self.previous_visible_counts: Counter[tuple[str, int]] = Counter()
        self.previous_visible_keys: list[tuple[str, int]] = []
        self.visible_damage_max_counts: Counter[tuple[str, int]] = Counter()
        self.counted_damage_contexts: set[tuple[tuple[str, int] | None, tuple[str, int], tuple[str, int] | None]] = set()
        self.pending_damage_items: list[tuple[str, int, str, str | None]] = []
        self.pending_damage_keys: list[tuple[str, int]] = []
        self.pending_stable_scans = 0
        self.baseline_next_tick = False
        self.baseline_until = 0.0
        self.count_mode = "safe"
        self.running = False
        self.worker: threading.Thread | None = None
        self.overlay: Toplevel | None = None
        self.overlay_drag_start: tuple[int, int] | None = None
        self.overlay_refresh_job: str | None = None
        self.overlay_position: tuple[int, int] = (80, 80)
        self.overlay_relative_position: tuple[int, int] | None = None
        self.poll_seconds = POLL_SECONDS
        self.overlay_fps = 15
        self.view_regions: list[tuple[int, int, int, int] | None] = [None, None]
        self.view_relative = [False, False]
        self.view_visible = [True, True]
        self.view_zoom = [1.0, 1.0]
        self.name_regions: list[tuple[int, int, int, int] | None] = [None, None]
        self.name_relative = [False, False]
        self.damage_counter_visible = True
        self.damage_counter_slot: Frame | None = None
        self.damage_counter_frame: Frame | None = None
        self.damage_actions_slot: Frame | None = None
        self.damage_actions_frame: Frame | None = None
        self.damage_counter_button: Button | None = None
        self.overlay_transparent = True
        self.overlay_bg_button: Button | None = None
        self.servitor_timers_visible = True
        self.servitor_timer_slot: Frame | None = None
        self.servitor_timer_frame: Frame | None = None
        self.servitor_timer_button: Button | None = None
        self.servitor_timer_end_times = {"barrier": 0.0, "empower": 0.0}
        self.servitor_timer_job: str | None = None
        self.servitor_timer_vars = {
            "barrier": StringVar(value="Barrier: ready"),
            "empower": StringVar(value="Empower: ready"),
        }
        self.servitor_timer_photos: dict[str, ImageTk.PhotoImage] = {}
        self.view_frames: list[Frame | None] = [None, None]
        self.view_labels: list[Label | None] = [None, None]
        self.view_photos: list[ImageTk.PhotoImage | None] = [None, None]
        self.view_buttons: list[Button | None] = [None, None]
        self.view_zoom_buttons: list[Button | None] = [None, None]
        self.name_frames: list[Frame | None] = [None, None]
        self.name_labels: list[Label | None] = [None, None]
        self.name_photos: list[ImageTk.PhotoImage | None] = [None, None]
        self.source_hwnd: int | None = None
        self.source_title = ""
        self.source_rect_hint: tuple[int, int, int, int] | None = None
        self.damage_hwnd: int | None = None
        self.damage_title = ""
        self.damage_rect_hint: tuple[int, int, int, int] | None = None
        self.overlay_active_window = "damage"
        self.overlay_manual_target: str | None = None

        self.status = StringVar()
        self.given_total_var = StringVar(value="Given: 0")
        self.received_total_var = StringVar(value="Received: 0")
        self.dps_var = StringVar(value="Given DPS 5s: 0.0   10s: 0.0   30s: 0.0")
        self.overlay_dps_var = StringVar(value="DPS 5s: 0.0")
        self.count_mode_var = StringVar(value="Mode: Safe")
        self.damage_counter_var = StringVar(value="Damage in overlay: On")
        self.servitor_timer_var = StringVar(value="Servitor timers: On")
        self.overlay_bg_var = StringVar(value="Overlay BG: Transparent")
        self.region_var = StringVar(value="Region: not selected")
        self.source_var = StringVar(value="Zone window: none")
        self.damage_var = StringVar(value="Damage window: none")
        self.zone_var = StringVar(value="Zones: not selected")
        self.interval_var = StringVar(value=f"{POLL_SECONDS:.2f}")
        self.overlay_fps_var = StringVar(value=str(self.overlay_fps))
        self.overlay_window_var = StringVar(value="Win: Auto")

        self._load_config()
        self._build_ui()
        self._refresh_status()
        self._refresh_servitor_timers()

    def _set_window_icon(self, window: Tk | Toplevel) -> None:
        if not APP_ICON_PATH.exists():
            return
        try:
            window.iconbitmap(str(APP_ICON_PATH))
        except Exception:
            pass

    def _apply_theme(self) -> None:
        t = THEME
        self.root.configure(bg=t["bg_dark"])
        self._enable_dark_titlebar(self.root)
        opt = self.root.option_add
        opt("*Frame.background", t["bg_panel"])
        opt("*LabelFrame.background", t["bg_panel"])
        opt("*LabelFrame.foreground", t["fg_title"])
        opt("*LabelFrame.font", "{Book Antiqua} 11 bold")
        opt("*LabelFrame.relief", "flat")
        opt("*LabelFrame.borderWidth", 0)
        opt("*LabelFrame.highlightThickness", 1)
        opt("*LabelFrame.highlightBackground", t["border"])
        opt("*LabelFrame.highlightColor", t["border"])
        opt("*LabelFrame.padX", 10)
        opt("*LabelFrame.padY", 8)
        opt("*Label.background", t["bg_panel"])
        opt("*Label.foreground", t["fg_text"])
        opt("*Label.font", "{Segoe UI} 9")
        opt("*Button.background", t["bg_button"])
        opt("*Button.foreground", t["fg_text"])
        opt("*Button.activeBackground", t["bg_button_active"])
        opt("*Button.activeForeground", t["fg_select"])
        opt("*Button.relief", "ridge")
        opt("*Button.borderWidth", 1)
        opt("*Button.font", "{Segoe UI} 9 bold")
        opt("*Button.padX", 6)
        opt("*Button.padY", 2)
        opt("*Entry.background", t["bg_dark"])
        opt("*Entry.foreground", t["fg_text"])
        opt("*Entry.insertBackground", t["fg_select"])
        opt("*Entry.relief", "flat")
        opt("*Entry.borderWidth", 1)
        opt("*Entry.highlightBackground", t["border"])
        opt("*Entry.highlightColor", t["fg_title"])
        opt("*Entry.highlightThickness", 1)
        opt("*Listbox.background", t["bg_listbox"])
        opt("*Listbox.foreground", t["fg_text"])
        opt("*Listbox.selectBackground", t["bg_button_active"])
        opt("*Listbox.selectForeground", t["fg_select"])
        opt("*Listbox.relief", "flat")
        opt("*Listbox.borderWidth", 0)
        opt("*Listbox.font", "{Consolas} 9")

    @staticmethod
    def _enable_dark_titlebar(window) -> None:  # type: ignore[no-untyped-def]
        try:
            window.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            if not hwnd:
                hwnd = window.winfo_id()
            value = ctypes.c_int(1)
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(value), ctypes.sizeof(value)
            )
            if result != 0:
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, 19, ctypes.byref(value), ctypes.sizeof(value)
                )
        except (OSError, AttributeError):
            pass

    def _section(self, parent: Frame, title: str, *, expand: bool = False) -> Frame:
        t = THEME
        outer = Frame(parent, bg=t["border"], padx=1, pady=1)
        outer.pack(fill=BOTH if expand else X, expand=expand, pady=(0, 10))

        title_bar = Frame(outer, bg=t["bg_dark"])
        title_bar.pack(fill=X)
        Label(
            title_bar,
            text=title,
            anchor="w",
            bg=t["bg_dark"],
            fg=t["fg_title"],
            font=TITLE_FONT,
            padx=6,
            pady=1,
        ).pack(fill=X)

        body = Frame(outer, bg=t["bg_panel"], padx=10, pady=8)
        body.pack(fill=BOTH if expand else X, expand=expand)
        return body

    def _build_ui(self) -> None:
        t = THEME
        container = Frame(self.root, bg=t["bg_dark"], padx=8, pady=8)
        container.pack(side=TOP, fill=BOTH, expand=True)

        setup = self._section(container, "Window Setup")
        setup_buttons = Frame(setup)
        setup_buttons.pack(fill=X)
        Button(setup_buttons, text="Pick main window", command=self.pick_damage_window).pack(side=LEFT)
        Button(setup_buttons, text="Pick buff window", command=self.pick_source_window).pack(side=LEFT, padx=(8, 0))
        Button(setup_buttons, text="Open damage overlay", command=self.open_damage_overlay).pack(side=LEFT, padx=(12, 0))
        Button(setup_buttons, text="Open windows-only overlay", command=self.open_windows_overlay).pack(side=LEFT, padx=(8, 0))

        overlay_buttons = Frame(setup, pady=6)
        overlay_buttons.pack(fill=X)
        Button(overlay_buttons, textvariable=self.damage_counter_var, command=self.toggle_damage_counter_visibility).pack(side=LEFT)
        Button(overlay_buttons, textvariable=self.overlay_bg_var, command=self.toggle_overlay_transparent).pack(side=LEFT, padx=(8, 0))

        Label(setup, textvariable=self.damage_var, anchor="w").pack(fill=X)
        Label(setup, textvariable=self.source_var, anchor="w").pack(fill=X)

        damage = self._section(container, "Damage Counter")

        damage_buttons = Frame(damage)
        damage_buttons.pack(fill=X)
        Button(damage_buttons, text="Select chat region", command=self.select_region).pack(side=LEFT)
        Button(damage_buttons, text="Start", command=self.start).pack(side=LEFT, padx=(8, 0))
        Button(damage_buttons, text="Stop", command=self.stop).pack(side=LEFT, padx=(8, 0))
        Button(damage_buttons, text="Reset stats", command=self.reset).pack(side=LEFT, padx=(8, 0))

        parser_controls = Frame(damage, pady=6)
        parser_controls.pack(fill=X)
        Label(parser_controls, text="Parse interval:", anchor="w").pack(side=LEFT)
        Entry(parser_controls, textvariable=self.interval_var, width=6).pack(side=LEFT, padx=(6, 0))
        Button(parser_controls, text="Apply", command=self.apply_interval).pack(side=LEFT, padx=(8, 0))
        Button(parser_controls, textvariable=self.count_mode_var, command=self.toggle_count_mode).pack(side=LEFT, padx=(12, 0))

        Label(damage, textvariable=self.region_var, anchor="w").pack(fill=X)

        stats = Frame(damage, pady=4)
        stats.pack(fill=X)
        Label(stats, textvariable=self.given_total_var, anchor="w", fg="#00b050", font=("Segoe UI", 13, "bold")).pack(side=LEFT)
        Label(stats, textvariable=self.received_total_var, anchor="w", fg="#d90000", font=("Segoe UI", 13, "bold")).pack(side=LEFT, padx=(20, 0))
        Label(damage, textvariable=self.dps_var, anchor="w").pack(fill=X)

        timers = self._section(container, "Servitor Timers")
        timer_buttons = Frame(timers)
        timer_buttons.pack(fill=X)
        Button(timer_buttons, text="Start Barrier 30s", command=lambda: self.start_servitor_timer("barrier")).pack(side=LEFT)
        Button(timer_buttons, text="Start Empowerment 60s", command=lambda: self.start_servitor_timer("empower")).pack(side=LEFT, padx=(8, 0))
        Button(timer_buttons, textvariable=self.servitor_timer_var, command=self.toggle_servitor_timers_visibility).pack(side=LEFT, padx=(8, 0))
        Label(timers, textvariable=self.servitor_timer_vars["barrier"], anchor="w", fg="#377dff").pack(side=LEFT, padx=(0, 18))
        Label(timers, textvariable=self.servitor_timer_vars["empower"], anchor="w", fg="#b85cff").pack(side=LEFT)

        mirror = self._section(container, "Window Overlay")

        mirror_buttons = Frame(mirror)
        mirror_buttons.pack(fill=X)
        Button(mirror_buttons, text="Select zone 1", command=lambda: self.select_view_region(0)).pack(side=LEFT)
        Button(mirror_buttons, text="Select zone 2", command=lambda: self.select_view_region(1)).pack(side=LEFT, padx=(8, 0))
        Button(mirror_buttons, text="Select name 1", command=lambda: self.select_name_region(0)).pack(side=LEFT, padx=(14, 0))
        Button(mirror_buttons, text="Select name 2", command=lambda: self.select_name_region(1)).pack(side=LEFT, padx=(8, 0))

        mirror_controls = Frame(mirror, pady=6)
        mirror_controls.pack(fill=X)
        Label(mirror_controls, text="Overlay FPS:", anchor="w").pack(side=LEFT)
        Entry(mirror_controls, textvariable=self.overlay_fps_var, width=4).pack(side=LEFT, padx=(6, 0))
        Button(mirror_controls, text="Apply FPS", command=self.apply_overlay_fps).pack(side=LEFT, padx=(8, 0))

        Label(mirror, textvariable=self.zone_var, anchor="w").pack(fill=X)

        Label(container, textvariable=self.status, anchor="w", bg=t["bg_dark"], fg=t["fg_text"], pady=2).pack(fill=X)

        log_frame = self._section(container, "Damage Log", expand=True)

        self.log = Listbox(log_frame)
        self.log.pack(side=TOP, fill=BOTH, expand=True)

        Label(
            self.root,
            text="Passive screen capture only. No client memory, packets, or automation.",
            anchor="w",
            bg=t["bg_dark"],
            fg=t["fg_dim"],
            padx=10,
            pady=6,
        ).pack(side=BOTTOM, fill=X)

    def _load_config(self) -> None:
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            region = data.get("region")
            if isinstance(region, list) and len(region) == 4:
                self.region = tuple(int(value) for value in region)
            poll_seconds = data.get("poll_seconds")
            if isinstance(poll_seconds, (int, float)):
                self.poll_seconds = min(MAX_POLL_SECONDS, max(MIN_POLL_SECONDS, float(poll_seconds)))
                self.interval_var.set(f"{self.poll_seconds:.2f}")
            overlay_fps = data.get("overlay_fps")
            if isinstance(overlay_fps, (int, float)):
                self.overlay_fps = min(MAX_OVERLAY_FPS, max(MIN_OVERLAY_FPS, int(overlay_fps)))
                self.overlay_fps_var.set(str(self.overlay_fps))
            count_mode = data.get("count_mode")
            if count_mode in ("safe", "fast"):
                self.count_mode = count_mode
                self.count_mode_var.set(f"Mode: {self.count_mode.title()}")
            damage_counter_visible = data.get("damage_counter_visible")
            if isinstance(damage_counter_visible, bool):
                self.damage_counter_visible = damage_counter_visible
                self.damage_counter_var.set(f"Damage in overlay: {'On' if self.damage_counter_visible else 'Off'}")
            servitor_timers_visible = data.get("servitor_timers_visible")
            if isinstance(servitor_timers_visible, bool):
                self.servitor_timers_visible = servitor_timers_visible
                self.servitor_timer_var.set(f"Servitor timers: {'On' if self.servitor_timers_visible else 'Off'}")
            overlay_transparent = data.get("overlay_transparent")
            if isinstance(overlay_transparent, bool):
                self.overlay_transparent = overlay_transparent
                self.overlay_bg_var.set(
                    f"Overlay BG: {'Transparent' if self.overlay_transparent else 'Solid'}"
                )
            view_regions = data.get("view_regions")
            if isinstance(view_regions, list):
                for index, region_data in enumerate(view_regions[:2]):
                    if isinstance(region_data, list) and len(region_data) == 4:
                        self.view_regions[index] = tuple(int(value) for value in region_data)
            view_relative = data.get("view_relative")
            if isinstance(view_relative, list):
                for index, value in enumerate(view_relative[:2]):
                    self.view_relative[index] = bool(value)
            view_visible = data.get("view_visible")
            if isinstance(view_visible, list):
                for index, value in enumerate(view_visible[:2]):
                    self.view_visible[index] = bool(value)
            view_zoom = data.get("view_zoom")
            if isinstance(view_zoom, list):
                for index, value in enumerate(view_zoom[:2]):
                    if isinstance(value, (int, float)):
                        self.view_zoom[index] = min(2.5, max(0.5, float(value)))
            name_regions = data.get("name_regions")
            if isinstance(name_regions, list):
                for index, region_data in enumerate(name_regions[:2]):
                    if isinstance(region_data, list) and len(region_data) == 4:
                        self.name_regions[index] = tuple(int(value) for value in region_data)
            name_relative = data.get("name_relative")
            if isinstance(name_relative, list):
                for index, value in enumerate(name_relative[:2]):
                    self.name_relative[index] = bool(value)
            source_rect_hint = data.get("source_rect_hint")
            if isinstance(source_rect_hint, list) and len(source_rect_hint) == 4:
                self.source_rect_hint = tuple(int(value) for value in source_rect_hint)
            damage_rect_hint = data.get("damage_rect_hint")
            if isinstance(damage_rect_hint, list) and len(damage_rect_hint) == 4:
                self.damage_rect_hint = tuple(int(value) for value in damage_rect_hint)
            source_title = data.get("source_title")
            if isinstance(source_title, str):
                self.source_title = source_title
                self.source_hwnd = self._resolve_saved_window(None, source_title, self.source_rect_hint)
            damage_title = data.get("damage_title")
            if isinstance(damage_title, str):
                self.damage_title = damage_title
                self.damage_hwnd = self._resolve_saved_window(None, damage_title, self.damage_rect_hint)
            overlay_relative_position = data.get("overlay_relative_position")
            if isinstance(overlay_relative_position, list) and len(overlay_relative_position) == 2:
                self.overlay_relative_position = (
                    int(overlay_relative_position[0]),
                    int(overlay_relative_position[1]),
                )
        except (OSError, ValueError, TypeError):
            self.region = None

    def _save_config(self) -> None:
        CONFIG_PATH.write_text(
            json.dumps(
                {
                    "region": self.region,
                    "poll_seconds": self.poll_seconds,
                    "overlay_fps": self.overlay_fps,
                    "count_mode": self.count_mode,
                    "damage_counter_visible": self.damage_counter_visible,
                    "servitor_timers_visible": self.servitor_timers_visible,
                    "overlay_transparent": self.overlay_transparent,
                    "view_regions": self.view_regions,
                    "view_relative": self.view_relative,
                    "view_visible": self.view_visible,
                    "view_zoom": self.view_zoom,
                    "name_regions": self.name_regions,
                    "name_relative": self.name_relative,
                    "source_title": self.source_title,
                    "damage_title": self.damage_title,
                    "source_rect_hint": self.source_rect_hint,
                    "damage_rect_hint": self.damage_rect_hint,
                    "overlay_relative_position": self.overlay_relative_position,
                },
                indent=2,
            ),
            encoding="utf-8",
        )

    def _refresh_status(self, error: str | None = None) -> None:
        if self.region:
            left, top, right, bottom = self.region
            self.region_var.set(f"Region: left={left}, top={top}, width={right - left}, height={bottom - top}")
        else:
            self.region_var.set("Region: not selected")

        if self.damage_hwnd and USER32.IsWindow(self.damage_hwnd):
            self.damage_var.set(f"Damage window: {self.damage_title}")
        elif self.damage_title:
            self.damage_var.set(f"Damage window: {self.damage_title} (not found)")
        else:
            self.damage_var.set("Damage window: none")

        if self.source_hwnd and USER32.IsWindow(self.source_hwnd):
            self.source_var.set(f"Zone window: {self.source_title}")
        elif self.source_title:
            self.source_var.set(f"Zone window: {self.source_title} (not found)")
        else:
            self.source_var.set("Zone window: none")

        zone_parts = []
        for index, region in enumerate(self.view_regions, start=1):
            if region:
                left, top, right, bottom = region
                mode = "window" if self.view_relative[index - 1] else "screen"
                zone_parts.append(f"Z{index}: {right - left}x{bottom - top} {mode}")
            else:
                zone_parts.append(f"Z{index}: none")
        for index, region in enumerate(self.name_regions, start=1):
            if region:
                left, top, right, bottom = region
                mode = "window" if self.name_relative[index - 1] else "screen"
                zone_parts.append(f"N{index}: {right - left}x{bottom - top} {mode}")
            else:
                zone_parts.append(f"N{index}: none")
        self.zone_var.set("Zones: " + "   ".join(zone_parts))

        if error:
            self.status.set(error)
        elif not self.ocr.available():
            self.status.set("OCR missing: Windows OCR or Tesseract OCR is required.")
        elif self.running and time.time() < self.baseline_until:
            remaining = max(0.0, self.baseline_until - time.time())
            self.status.set(f"Learning visible chat for {remaining:.1f}s...")
        elif self.running:
            self.status.set(f"Watching chat with {self.ocr.active_name}...")
        else:
            self.status.set("Ready.")

        self._refresh_stats()

    def apply_interval(self) -> None:
        raw = self.interval_var.get().strip().replace(",", ".")
        try:
            value = float(raw)
        except ValueError:
            messagebox.showerror("Parse interval", "Enter a number like 1.0 or 0.2.")
            self.interval_var.set(f"{self.poll_seconds:.2f}")
            return

        if value < MIN_POLL_SECONDS or value > MAX_POLL_SECONDS:
            messagebox.showerror(
                "Parse interval",
                f"Use a value from {MIN_POLL_SECONDS:.2f} to {MAX_POLL_SECONDS:.1f} seconds.",
            )
            self.interval_var.set(f"{self.poll_seconds:.2f}")
            return

        self.poll_seconds = value
        self.interval_var.set(f"{self.poll_seconds:.2f}")
        self._save_config()
        self.status.set(f"Parse interval set to {self.poll_seconds:.2f}s.")

    def apply_overlay_fps(self) -> None:
        raw = self.overlay_fps_var.get().strip()
        try:
            value = int(float(raw))
        except ValueError:
            messagebox.showerror("Overlay FPS", "Enter a number like 10, 15, or 30.")
            self.overlay_fps_var.set(str(self.overlay_fps))
            return

        if value < MIN_OVERLAY_FPS or value > MAX_OVERLAY_FPS:
            messagebox.showerror("Overlay FPS", f"Use a value from {MIN_OVERLAY_FPS} to {MAX_OVERLAY_FPS}.")
            self.overlay_fps_var.set(str(self.overlay_fps))
            return

        self.overlay_fps = value
        self.overlay_fps_var.set(str(self.overlay_fps))
        self._save_config()
        self.status.set(f"Overlay FPS set to {self.overlay_fps}.")

    def toggle_count_mode(self) -> None:
        self.count_mode = "fast" if self.count_mode == "safe" else "safe"
        self.count_mode_var.set(f"Mode: {self.count_mode.title()}")
        self._clear_pending_damage()
        self.baseline_next_tick = self.running
        self.baseline_until = time.time() + self._baseline_seconds() if self.running else 0.0
        self._save_config()
        self.status.set(f"Damage count mode set to {self.count_mode}.")

    def _baseline_seconds(self) -> float:
        return 0.8 if self.count_mode == "fast" else BASELINE_SECONDS

    def _refresh_stats(self) -> None:
        now = time.time()
        given_total = sum(event.amount for event in self.events if event.direction == "out")
        received_total = sum(event.amount for event in self.events if event.direction == "in")
        dps_parts = []
        for seconds in WINDOW_SECONDS:
            window_damage = sum(
                event.amount
                for event in self.events
                if event.direction == "out" and now - event.timestamp <= seconds
            )
            dps_parts.append(f"{seconds}s: {window_damage / seconds:.1f}")

        self.given_total_var.set(f"Given: {given_total}")
        self.received_total_var.set(f"Received: {received_total}")
        self.dps_var.set("Given DPS " + "   ".join(dps_parts))
        self.overlay_dps_var.set(f"DPS 5s: {dps_parts[0].split(': ')[1]}")

    def toggle_overlay(self) -> None:
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.destroy()
            self.overlay = None
            return

        overlay = Toplevel(self.root)
        self.overlay = overlay
        overlay.title("parashaoly")
        self._set_window_icon(overlay)
        self.overlay_position = self._initial_overlay_position()
        overlay.geometry(f"460x220+{self.overlay_position[0]}+{self.overlay_position[1]}")
        overlay.configure(bg="#111111")
        overlay.attributes("-topmost", True)
        overlay.minsize(180, 80)
        overlay.overrideredirect(True)
        self._apply_overlay_transparency()

        panel = Frame(overlay, bg="#111111", padx=8, pady=6)
        panel.pack(fill=BOTH, expand=True)

        header = Frame(panel, bg="#1a1a1a", height=22)
        header.pack(fill=X)
        Label(header, text=" ≡ ", bg="#1a1a1a", fg="#888888", font=("Segoe UI", 9, "bold"), cursor="fleur").pack(side=LEFT)
        Button(
            header,
            text="x",
            command=self.toggle_overlay,
            bg="#3a1a1a",
            fg="#ffdddd",
            activebackground="#552222",
            activeforeground="#ffffff",
            relief="flat",
            width=2,
        ).pack(side=RIGHT)
        Button(
            header,
            textvariable=self.overlay_window_var,
            command=self.cycle_overlay_window_mode,
            bg="#252525",
            fg="#dddddd",
            activebackground="#333333",
            activeforeground="#ffffff",
            relief="flat",
            width=10,
        ).pack(side=RIGHT, padx=(0, 5))
        self.damage_counter_button = None
        self.servitor_timer_button = None
        self.overlay_bg_button = None

        self.damage_counter_slot = Frame(panel, bg="#111111")
        self.damage_counter_slot.pack(fill=X)
        self.damage_counter_frame = Frame(self.damage_counter_slot, bg="#111111")
        Label(self.damage_counter_frame, textvariable=self.given_total_var, bg="#111111", fg="#00ff55", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        Label(self.damage_counter_frame, textvariable=self.received_total_var, bg="#111111", fg="#ff3333", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        Label(self.damage_counter_frame, textvariable=self.overlay_dps_var, bg="#111111", fg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        self.damage_actions_slot = Frame(panel, bg="#111111")
        self.damage_actions_slot.pack(fill=X)
        actions = Frame(self.damage_actions_slot, bg="#111111")
        self.damage_actions_frame = actions
        Button(
            actions,
            text="Start",
            command=self.start,
            bg="#164d24",
            fg="#ffffff",
            activebackground="#207436",
            activeforeground="#ffffff",
            relief="flat",
            width=6,
        ).pack(side=LEFT, padx=(0, 6))
        Button(
            actions,
            text="Reset",
            command=self.reset,
            bg="#333333",
            fg="#ffffff",
            activebackground="#444444",
            activeforeground="#ffffff",
            relief="flat",
            width=6,
        ).pack(side=LEFT)
        Button(
            actions,
            textvariable=self.count_mode_var,
            command=self.toggle_count_mode,
            bg="#222222",
            fg="#ffffff",
            activebackground="#333333",
            activeforeground="#ffffff",
            relief="flat",
            width=10,
        ).pack(side=LEFT, padx=(6, 0))

        self.servitor_timer_slot = Frame(panel, bg="#111111")
        self.servitor_timer_slot.pack(fill=X)
        self.servitor_timer_frame = Frame(self.servitor_timer_slot, bg="#111111")
        for key in ("barrier", "empower"):
            self._build_overlay_timer_button(self.servitor_timer_frame, key)

        toggles = Frame(panel, bg="#111111")
        toggles.pack(fill=X, pady=(5, 3))
        for index in range(2):
            button = Button(
                toggles,
                text=f"Z{index + 1}",
                command=lambda zone_index=index: self.toggle_view_zone(zone_index),
                bg="#13a85a" if self.view_visible[index] else "#262a2f",
                fg="#ffffff" if self.view_visible[index] else "#9aa3ad",
                activebackground="#19c86d",
                activeforeground="#ffffff",
                relief="flat",
                bd=0,
                highlightthickness=1,
                highlightbackground="#1ee077" if self.view_visible[index] else "#3b4148",
                highlightcolor="#1ee077",
                font=("Segoe UI", 9, "bold"),
                padx=10,
                pady=4,
                cursor="hand2",
            )
            button.pack(side=LEFT, padx=(0, 7))
            self.view_buttons[index] = button

        for index in range(2):
            frame = Frame(panel, bg="#111111")
            zone_header = Frame(frame, bg="#111111")
            zone_header.pack(fill=X)
            name_frame = Frame(
                zone_header,
                bg="#070707",
                padx=4,
                pady=3,
                highlightthickness=1,
                highlightbackground="#3b3324",
            )
            name_frame.pack(side=LEFT)
            name_label = Label(
                name_frame,
                text=f"Z{index + 1}",
                bg="#070707",
                fg="#d8d8d8",
                font=("Segoe UI", 8, "bold"),
                width=4,
            )
            name_label.pack()
            name_frame.bind("<ButtonPress-1>", self._start_overlay_drag)
            name_frame.bind("<B1-Motion>", self._drag_overlay)
            name_label.bind("<ButtonPress-1>", self._start_overlay_drag)
            name_label.bind("<B1-Motion>", self._drag_overlay)
            self.name_frames[index] = name_frame
            self.name_labels[index] = name_label
            Button(
                zone_header,
                text="Pick",
                command=lambda zone_index=index: self.select_view_region(zone_index),
                bg="#1f2a38",
                fg="#cfe8ff",
                activebackground="#2c3d52",
                activeforeground="#ffffff",
                relief="flat",
                font=("Segoe UI", 8, "bold"),
                padx=6,
                pady=1,
                cursor="hand2",
            ).pack(side=LEFT, padx=(6, 0))
            zoom_button = Button(
                zone_header,
                text=f"{self.view_zoom[index]:.1f}x",
                command=lambda zone_index=index: self.cycle_view_zoom(zone_index),
                bg="#222222",
                fg="#dddddd",
                activebackground="#333333",
                activeforeground="#ffffff",
                relief="flat",
                width=5,
            )
            zoom_button.pack(side=RIGHT, padx=(4, 0))
            Button(
                zone_header,
                text="x",
                command=lambda zone_index=index: self.hide_view_zone(zone_index),
                bg="#331111",
                fg="#ffdddd",
                activebackground="#552222",
                activeforeground="#ffffff",
                relief="flat",
                width=2,
            ).pack(side=RIGHT)
            label = Label(frame, text=f"Zone {index + 1} not selected", bg="#111111", fg="#888888")
            label.pack(anchor="w")
            frame.bind("<ButtonPress-1>", self._start_overlay_drag)
            frame.bind("<B1-Motion>", self._drag_overlay)
            label.bind("<ButtonPress-1>", self._start_overlay_drag)
            label.bind("<B1-Motion>", self._drag_overlay)
            self.view_frames[index] = frame
            self.view_labels[index] = label
            self.view_zoom_buttons[index] = zoom_button

        grip = Label(overlay, text="◢", bg="#1a1a1a", fg="#888888", font=("Segoe UI", 10, "bold"), cursor="size_nw_se")
        grip.place(relx=1.0, rely=1.0, anchor="se")
        grip.bind("<ButtonPress-1>", self._start_overlay_resize)
        grip.bind("<B1-Motion>", self._do_overlay_resize)

        self._refresh_damage_counter_visibility()
        self._refresh_servitor_timers_visibility()
        self._refresh_overlay_name_visibility()
        self._refresh_overlay_zone_visibility()
        self._refresh_overlay_zones()
        self._bind_overlay_drag_recursive(panel)
        self._schedule_overlay_refresh()
        self._fit_overlay_to_content()

    def open_damage_overlay(self) -> None:
        self.set_damage_counter_visibility(True)
        self._show_overlay()

    def open_windows_overlay(self) -> None:
        self.set_damage_counter_visibility(False)
        self._show_overlay()

    def _show_overlay(self) -> None:
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.lift()
            self.overlay.attributes("-topmost", True)
            return
        self.toggle_overlay()

    def _start_overlay_drag(self, event) -> None:  # type: ignore[no-untyped-def]
        self.overlay_drag_start = (event.x_root, event.y_root)

    def _drag_overlay(self, event) -> None:  # type: ignore[no-untyped-def]
        if not self.overlay or not self.overlay_drag_start:
            return

        start_x, start_y = self.overlay_drag_start
        dx = event.x_root - start_x
        dy = event.y_root - start_y
        x = self.overlay.winfo_x() + dx
        y = self.overlay.winfo_y() + dy
        self.overlay.geometry(f"+{x}+{y}")
        self.overlay_position = (x, y)
        self._save_overlay_relative_position()
        self.overlay_drag_start = (event.x_root, event.y_root)

    def toggle_overlay_transparent(self) -> None:
        self.overlay_transparent = not self.overlay_transparent
        self._save_config()
        self._apply_overlay_transparency()

    def _apply_overlay_transparency(self) -> None:
        self.overlay_bg_var.set(
            f"Overlay BG: {'Transparent' if self.overlay_transparent else 'Solid'}"
        )
        if not self.overlay or not self.overlay.winfo_exists():
            return
        if self.overlay_transparent:
            self.overlay.attributes("-transparentcolor", "#111111")
            self.overlay.attributes("-alpha", 1.0)
        else:
            self.overlay.attributes("-transparentcolor", "")
            self.overlay.attributes("-alpha", 0.88)

    def _start_overlay_resize(self, event) -> None:  # type: ignore[no-untyped-def]
        if not self.overlay:
            return
        self.overlay_resize_start = (
            event.x_root,
            event.y_root,
            self.overlay.winfo_width(),
            self.overlay.winfo_height(),
        )

    def _do_overlay_resize(self, event) -> None:  # type: ignore[no-untyped-def]
        if not self.overlay or not getattr(self, "overlay_resize_start", None):
            return
        start_x, start_y, start_w, start_h = self.overlay_resize_start
        new_w = max(180, start_w + (event.x_root - start_x))
        new_h = max(80, start_h + (event.y_root - start_y))
        x, y = self.overlay_position
        self.overlay.geometry(f"{new_w}x{new_h}+{x}+{y}")

    def _bind_overlay_drag_recursive(self, widget) -> None:  # type: ignore[no-untyped-def]
        if not isinstance(widget, Button):
            widget.bind("<ButtonPress-1>", self._start_overlay_drag)
            widget.bind("<B1-Motion>", self._drag_overlay)
        for child in widget.winfo_children():
            self._bind_overlay_drag_recursive(child)

    def toggle_damage_counter_visibility(self) -> None:
        self.set_damage_counter_visibility(not self.damage_counter_visible)

    def cycle_overlay_window_mode(self) -> None:
        if self.overlay_manual_target is None:
            self.overlay_manual_target = "source"
        elif self.overlay_manual_target == "source":
            self.overlay_manual_target = "damage"
        else:
            self.overlay_manual_target = None

        self._refresh_overlay_window_mode()
        self._follow_source_window()
        self._refresh_overlay_zones()

    def _refresh_overlay_window_mode(self) -> None:
        label = {
            None: "Win: Auto",
            "source": "Win: Buff",
            "damage": "Win: Main",
        }[self.overlay_manual_target]
        self.overlay_window_var.set(label)

    def set_damage_counter_visibility(self, visible: bool) -> None:
        self.damage_counter_visible = visible
        self._save_config()
        self._refresh_damage_counter_visibility()

    def _refresh_damage_counter_visibility(self) -> None:
        self.damage_counter_var.set(f"Damage in overlay: {'On' if self.damage_counter_visible else 'Off'}")
        if self.damage_counter_button:
            self.damage_counter_button.configure(bg="#00a040" if self.damage_counter_visible else "#333333")
        if self.damage_counter_slot:
            if self.damage_counter_visible:
                self.damage_counter_slot.pack(fill=X)
            else:
                self.damage_counter_slot.pack_forget()
        if self.damage_counter_frame:
            if self.damage_counter_visible:
                self.damage_counter_frame.pack(fill=X)
            else:
                self.damage_counter_frame.pack_forget()
        if self.damage_actions_slot:
            if self.damage_counter_visible:
                self.damage_actions_slot.pack(fill=X)
            else:
                self.damage_actions_slot.pack_forget()
        if self.damage_actions_frame:
            if self.damage_counter_visible:
                self.damage_actions_frame.pack(fill=X, pady=(5, 2))
            else:
                self.damage_actions_frame.pack_forget()
        self._fit_overlay_to_content()

    def _build_overlay_timer_button(self, parent: Frame, key: str) -> None:
        info = SERVITOR_TIMERS[key]
        row = Frame(parent, bg="#111111")
        row.pack(side=LEFT, padx=(0, 10))

        photo = self._load_timer_icon(key)
        button_kwargs = {
            "command": lambda timer_key=key: self.start_servitor_timer(timer_key),
            "bg": "#222222",
            "fg": "#ffffff",
            "activebackground": "#333333",
            "activeforeground": "#ffffff",
            "relief": "flat",
            "width": 42,
            "height": 38,
        }
        if photo:
            Button(row, image=photo, **button_kwargs).pack(side=LEFT)
        else:
            Button(row, text=info["name"][:3], **button_kwargs).pack(side=LEFT)
        Label(row, textvariable=self.servitor_timer_vars[key], bg="#111111", fg="#ffffff", font=("Segoe UI", 10, "bold")).pack(side=LEFT, padx=(5, 0))

    def _load_timer_icon(self, key: str) -> ImageTk.PhotoImage | None:
        if key in self.servitor_timer_photos:
            return self.servitor_timer_photos[key]
        icon_path = SERVITOR_TIMERS[key]["icon"]
        if not icon_path.exists():
            return None
        try:
            image = Image.open(icon_path).convert("RGBA").resize((32, 32), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
        except Exception:
            return None
        self.servitor_timer_photos[key] = photo
        return photo

    def start_servitor_timer(self, key: str) -> None:
        self.servitor_timer_end_times[key] = time.time() + int(SERVITOR_TIMERS[key]["duration"])
        self._refresh_servitor_timers()

    def toggle_servitor_timers_visibility(self) -> None:
        self.servitor_timers_visible = not self.servitor_timers_visible
        self._save_config()
        self._refresh_servitor_timers_visibility()

    def _refresh_servitor_timers_visibility(self) -> None:
        self.servitor_timer_var.set(f"Servitor timers: {'On' if self.servitor_timers_visible else 'Off'}")
        if self.servitor_timer_button:
            self.servitor_timer_button.configure(bg="#00a040" if self.servitor_timers_visible else "#333333")
        if self.servitor_timer_slot:
            if self.servitor_timers_visible:
                self.servitor_timer_slot.pack(fill=X)
            else:
                self.servitor_timer_slot.pack_forget()
        if self.servitor_timer_frame:
            if self.servitor_timers_visible:
                self.servitor_timer_frame.pack(fill=X, pady=(5, 2))
            else:
                self.servitor_timer_frame.pack_forget()
        self._fit_overlay_to_content()

    def _refresh_servitor_timers(self) -> None:
        if self.servitor_timer_job:
            try:
                self.root.after_cancel(self.servitor_timer_job)
            except Exception:
                pass
        self.servitor_timer_job = None
        now = time.time()
        for key, info in SERVITOR_TIMERS.items():
            remaining = max(0, int(self.servitor_timer_end_times[key] - now + 0.999))
            if remaining > 0:
                self.servitor_timer_vars[key].set(f"{info['name']}: {remaining}s")
            else:
                self.servitor_timer_vars[key].set(f"{info['name']}: ready")
        self.servitor_timer_job = self.root.after(250, self._refresh_servitor_timers)

    def toggle_view_zone(self, index: int) -> None:
        self.view_visible[index] = not self.view_visible[index]
        self._save_config()
        self._refresh_overlay_zone_visibility()
        self._refresh_overlay_zones()

    def hide_view_zone(self, index: int) -> None:
        self.view_visible[index] = False
        self._save_config()
        self._refresh_overlay_zone_visibility()

    def cycle_view_zoom(self, index: int) -> None:
        zoom_steps = [0.75, 1.0, 1.5, 2.0, 2.5]
        current = self.view_zoom[index]
        next_zoom = zoom_steps[0]
        for zoom in zoom_steps:
            if zoom > current + 0.01:
                next_zoom = zoom
                break
        self.view_zoom[index] = next_zoom
        button = self.view_zoom_buttons[index]
        if button:
            button.configure(text=f"{next_zoom:.1f}x")
        self._save_config()
        self._refresh_overlay_zones()

    def _refresh_overlay_zone_visibility(self) -> None:
        if not self.overlay or not self.overlay.winfo_exists():
            return

        for index, frame in enumerate(self.view_frames):
            button = self.view_buttons[index]
            if button:
                button.configure(
                    bg="#13a85a" if self.view_visible[index] else "#262a2f",
                    fg="#ffffff" if self.view_visible[index] else "#9aa3ad",
                    activebackground="#19c86d" if self.view_visible[index] else "#343b44",
                    highlightbackground="#1ee077" if self.view_visible[index] else "#3b4148",
                )
            zoom_button = self.view_zoom_buttons[index]
            if zoom_button:
                zoom_button.configure(text=f"{self.view_zoom[index]:.1f}x")
            if not frame:
                continue
            if self.view_visible[index] and self.view_regions[index]:
                frame.pack(fill=X, pady=(3, 0))
            else:
                frame.pack_forget()
        self._fit_overlay_to_content()

    def _refresh_overlay_name_visibility(self) -> None:
        if not self.overlay or not self.overlay.winfo_exists():
            return

        for index, frame in enumerate(self.name_frames):
            if not frame:
                continue
            frame.pack(side=LEFT)
        self._fit_overlay_to_content()

    def _refresh_overlay_nameplates(self) -> None:
        if not self.overlay or not self.overlay.winfo_exists():
            return

        overlay_rect = (
            self.overlay.winfo_rootx(),
            self.overlay.winfo_rooty(),
            self.overlay.winfo_rootx() + self.overlay.winfo_width(),
            self.overlay.winfo_rooty() + self.overlay.winfo_height(),
        )
        for index, region in enumerate(self.name_regions):
            label = self.name_labels[index]
            if not label:
                continue
            if not region:
                label.configure(text=f"Z{index + 1}", image="", width=4, height=1)
                self.name_photos[index] = None
                continue

            capture_region = self._resolve_name_region(index)
            if not capture_region:
                label.configure(text=f"Z{index + 1}", image="", width=4, height=1)
                self.name_photos[index] = None
                continue
            if not self.name_relative[index] and rectangles_intersect(overlay_rect, capture_region):
                continue

            try:
                image = self._capture_name_image(index, capture_region)
                if image is None:
                    raise RuntimeError("capture failed")
                image = self._resize_name_image(image)
                photo = ImageTk.PhotoImage(image)
            except Exception:
                label.configure(text=f"Name {index + 1}", image="")
                self.name_photos[index] = None
                continue

            self.name_photos[index] = photo
            label.configure(image=photo, text="", width=image.width, height=image.height)

    def _refresh_overlay_zones(self) -> None:
        if not self.overlay or not self.overlay.winfo_exists():
            return

        self._refresh_overlay_nameplates()

        capture_jobs: list[tuple[int, Label, tuple[int, int, int, int]]] = []
        for index, region in enumerate(self.view_regions):
            label = self.view_labels[index]
            if not label or not self.view_visible[index] or not region:
                continue

            capture_region = self._resolve_view_region(index)
            if not capture_region:
                label.configure(text=f"Zone {index + 1}: source window not found", image="")
                self.view_photos[index] = None
                continue
            capture_jobs.append((index, label, capture_region))

        if not capture_jobs:
            return

        overlay_rect = (
            self.overlay.winfo_rootx(),
            self.overlay.winfo_rooty(),
            self.overlay.winfo_rootx() + self.overlay.winfo_width(),
            self.overlay.winfo_rooty() + self.overlay.winfo_height(),
        )
        for index, label, capture_region in capture_jobs:
            if not self.view_relative[index] and rectangles_intersect(overlay_rect, capture_region):
                if self.view_photos[index] is None:
                    label.configure(text=f"Zone {index + 1}: move overlay away from source", image="")
                continue

            try:
                image = self._capture_view_image(index, capture_region)
                if image is None:
                    raise RuntimeError("capture failed")
                image = self._resize_zone_image(index, image)
                photo = ImageTk.PhotoImage(image)
            except Exception:
                label.configure(text=f"Zone {index + 1} capture failed", image="")
                self.view_photos[index] = None
                continue

            self.view_photos[index] = photo
            label.configure(image=photo, text="")

        if self.overlay:
            self.overlay.update_idletasks()
            self._fit_overlay_to_content()

    def _resize_name_image(self, image: Image.Image) -> Image.Image:
        max_width = 190
        max_height = 54
        scale = min(max_width / image.width, max_height / image.height, 1.0)
        width = max(1, int(image.width * scale))
        height = max(1, int(image.height * scale))
        return image.resize((width, height), Image.Resampling.LANCZOS)

    def _resize_zone_image(self, index: int, image: Image.Image) -> Image.Image:
        zoom = self.view_zoom[index]
        width = max(1, int(image.width * zoom))
        height = max(1, int(image.height * zoom))

        max_width = 430
        max_height = 135
        scale = min(max_width / width, max_height / height, 1.0)
        width = max(1, int(width * scale))
        height = max(1, int(height * scale))
        return image.resize((width, height), Image.Resampling.LANCZOS)

    def _fit_overlay_to_content(self) -> None:
        if not self.overlay or not self.overlay.winfo_exists():
            return

        self.overlay.update_idletasks()
        current_width = self.overlay.winfo_width()
        current_height = self.overlay.winfo_height()
        width = max(180, self.overlay.winfo_reqwidth(), current_width)
        height = max(80, self.overlay.winfo_reqheight(), current_height)
        x, y = self.overlay_position
        self.overlay.geometry(f"{width}x{height}+{x}+{y}")

    def _schedule_overlay_refresh(self) -> None:
        if not self.overlay or not self.overlay.winfo_exists():
            self.overlay_refresh_job = None
            return

        self._follow_source_window()
        self._refresh_overlay_zones()
        delay_ms = max(33, int(1000 / max(1, self.overlay_fps)))
        self.overlay_refresh_job = self.root.after(delay_ms, self._schedule_overlay_refresh)

    def _initial_overlay_position(self) -> tuple[int, int]:
        target_rect = self._get_overlay_target_rect()
        if target_rect:
            left, top, _right, _bottom = target_rect
            if self.overlay_relative_position:
                rel_x, rel_y = self.overlay_relative_position
            else:
                rel_x, rel_y = 20, 60
                self.overlay_relative_position = (rel_x, rel_y)
            return (left + rel_x, top + rel_y)
        return self.overlay_position

    def _save_overlay_relative_position(self) -> None:
        target_rect = self._get_overlay_target_rect()
        if not target_rect:
            return
        source_left, source_top, _source_right, _source_bottom = target_rect
        x, y = self.overlay_position
        self.overlay_relative_position = (x - source_left, y - source_top)
        self._save_config()

    def _follow_source_window(self) -> None:
        if not self.overlay or not self.overlay.winfo_exists() or not self.overlay_relative_position:
            return
        target_rect = self._get_overlay_target_rect()
        if not target_rect:
            return

        source_left, source_top, _source_right, _source_bottom = target_rect
        rel_x, rel_y = self.overlay_relative_position
        x = source_left + rel_x
        y = source_top + rel_y
        if (x, y) != self.overlay_position:
            self.overlay_position = (x, y)
            self.overlay.geometry(f"+{x}+{y}")

    def _capture_view_image(self, index: int, fallback_region: tuple[int, int, int, int]) -> Image.Image | None:
        return self._capture_overlay_region(self.view_regions[index], self.view_relative[index], fallback_region)

    def _capture_name_image(self, index: int, fallback_region: tuple[int, int, int, int]) -> Image.Image | None:
        return self._capture_overlay_region(self.name_regions[index], self.name_relative[index], fallback_region)

    def _capture_overlay_region(
        self,
        relative_region: tuple[int, int, int, int] | None,
        is_relative: bool,
        fallback_region: tuple[int, int, int, int],
    ) -> Image.Image | None:
        if is_relative:
            hwnd = self._get_view_source_hwnd()
            if hwnd:
                source_image = capture_window_image(hwnd)
                if source_image and relative_region:
                    left, top, right, bottom = relative_region
                    left = max(0, min(source_image.width, left))
                    top = max(0, min(source_image.height, top))
                    right = max(left + 1, min(source_image.width, right))
                    bottom = max(top + 1, min(source_image.height, bottom))
                    return source_image.crop((left, top, right, bottom))

        return ImageGrab.grab(bbox=fallback_region).convert("RGB")

    def _resolve_view_region(self, index: int) -> tuple[int, int, int, int] | None:
        return self._resolve_overlay_region(self.view_regions[index], self.view_relative[index])

    def _resolve_name_region(self, index: int) -> tuple[int, int, int, int] | None:
        return self._resolve_overlay_region(self.name_regions[index], self.name_relative[index])

    def _resolve_overlay_region(
        self,
        region: tuple[int, int, int, int] | None,
        is_relative: bool,
    ) -> tuple[int, int, int, int] | None:
        if not region:
            return None
        if not is_relative:
            return region

        window_rect = self._get_view_source_rect()
        if not window_rect:
            return None

        window_left, window_top, _window_right, _window_bottom = window_rect
        left, top, right, bottom = region
        return (
            window_left + left,
            window_top + top,
            window_left + right,
            window_top + bottom,
        )

    def _get_source_rect(self) -> tuple[int, int, int, int] | None:
        hwnd = self._get_source_hwnd()
        if hwnd:
            return get_window_rect(hwnd)
        return None

    def _get_overlay_window_pair(self) -> tuple[int | None, int | None]:
        damage_hwnd = self._get_damage_hwnd()
        source_hwnd = self._get_source_hwnd()
        if damage_hwnd and source_hwnd:
            if self.overlay_manual_target == "source":
                return source_hwnd, damage_hwnd
            if self.overlay_manual_target == "damage":
                return damage_hwnd, source_hwnd

        active_hwnd = get_root_window(int(USER32.GetForegroundWindow()))
        if self.overlay and self.overlay.winfo_exists() and active_hwnd == get_root_window(self.overlay.winfo_id()):
            active_hwnd = 0

        if damage_hwnd and source_hwnd and active_hwnd == get_root_window(source_hwnd):
            return source_hwnd, damage_hwnd
        return damage_hwnd, source_hwnd

    def _get_overlay_target_hwnd(self) -> int | None:
        target_hwnd, _view_source_hwnd = self._get_overlay_window_pair()
        return target_hwnd

    def _get_overlay_target_rect(self) -> tuple[int, int, int, int] | None:
        hwnd = self._get_overlay_target_hwnd()
        if hwnd:
            return get_window_rect(hwnd)
        return None

    def _get_view_source_hwnd(self) -> int | None:
        _target_hwnd, view_source_hwnd = self._get_overlay_window_pair()
        return view_source_hwnd

    def _get_view_source_rect(self) -> tuple[int, int, int, int] | None:
        hwnd = self._get_view_source_hwnd()
        if hwnd:
            return get_window_rect(hwnd)
        return None

    def _get_source_hwnd(self) -> int | None:
        self.source_hwnd = self._resolve_saved_window(
            self.source_hwnd,
            self.source_title,
            self.source_rect_hint,
        )
        if self.source_hwnd:
            if not self.source_rect_hint:
                self.source_rect_hint = get_window_rect(self.source_hwnd)
        return self.source_hwnd

    def _get_damage_rect(self) -> tuple[int, int, int, int] | None:
        hwnd = self._get_damage_hwnd()
        if hwnd:
            return get_window_rect(hwnd)
        return None

    def _get_damage_hwnd(self) -> int | None:
        self.damage_hwnd = self._resolve_saved_window(
            self.damage_hwnd,
            self.damage_title,
            self.damage_rect_hint,
        )
        if self.damage_hwnd:
            if not self.damage_rect_hint:
                self.damage_rect_hint = get_window_rect(self.damage_hwnd)
        return self.damage_hwnd

    def _resolve_saved_window(
        self,
        hwnd: int | None,
        title: str,
        rect_hint: tuple[int, int, int, int] | None,
    ) -> int | None:
        if hwnd and USER32.IsWindow(hwnd):
            return get_root_window(hwnd)

        if not title:
            return hwnd if hwnd and USER32.IsWindow(hwnd) else None

        matches = find_windows_by_title(title)
        if not matches:
            return hwnd if hwnd and USER32.IsWindow(hwnd) else None
        if rect_hint:
            return min(matches, key=lambda candidate: rect_distance(get_window_rect(candidate), rect_hint))
        return matches[0]

    def pick_source_window(self) -> None:
        self.status.set("Click the zone source window now. Picking active window in 3 seconds...")
        self.root.after(3000, lambda: self._capture_foreground_window("zone"))

    def pick_damage_window(self) -> None:
        self.status.set("Click the damage overlay window now. Picking active window in 3 seconds...")
        self.root.after(3000, lambda: self._capture_foreground_window("damage"))

    def _capture_foreground_window(self, kind: str) -> None:
        hwnd = get_root_window(int(USER32.GetForegroundWindow()))
        title = get_window_title(hwnd)
        rect = get_window_rect(hwnd)
        if not hwnd or not title or not rect:
            messagebox.showerror("Source window", "Could not read the active window. Try again.")
            return

        if hwnd == self.root.winfo_id():
            messagebox.showinfo("Window picker", "The meter window is active. Click the target window after pressing the pick button.")
            return

        if kind == "damage":
            self.damage_hwnd = hwnd
            self.damage_title = title
            self.damage_rect_hint = rect
            self.overlay_active_window = "damage"
            self.overlay_relative_position = None
            message = f"Damage window picked: {title}"
        else:
            self.source_hwnd = hwnd
            self.source_title = title
            self.source_rect_hint = rect
            self.overlay_active_window = "source"
            message = f"Zone window picked: {title}"
        self._save_config()
        self._refresh_status()
        self.status.set(message)

    def select_region(self) -> None:
        self.stop()
        selector = RegionSelector(self.root)
        region = selector.select()
        if region:
            self.region = region
            self.previous_visible_counts.clear()
            self.previous_visible_keys.clear()
            self.visible_damage_max_counts.clear()
            self.counted_damage_contexts.clear()
            self._clear_pending_damage()
            self._save_config()
        self._refresh_status()

    def select_view_region(self, index: int) -> None:
        selector = RegionSelector(self.root, f"Drag over overlay zone {index + 1}. Press Esc to cancel.")
        region = selector.select()
        if region:
            source_rect = self._get_source_rect()
            if source_rect:
                window_left, window_top, _window_right, _window_bottom = source_rect
                left, top, right, bottom = region
                self.view_regions[index] = (
                    left - window_left,
                    top - window_top,
                    right - window_left,
                    bottom - window_top,
                )
                self.view_relative[index] = True
            else:
                self.view_regions[index] = region
                self.view_relative[index] = False
            self.view_visible[index] = True
            self._save_config()
            self._refresh_overlay_zone_visibility()
            self._refresh_overlay_zones()
        self._refresh_status()

    def select_name_region(self, index: int) -> None:
        selector = RegionSelector(self.root, f"Drag over nickname zone {index + 1}. Press Esc to cancel.")
        region = selector.select()
        if region:
            source_rect = self._get_source_rect()
            if source_rect:
                window_left, window_top, _window_right, _window_bottom = source_rect
                left, top, right, bottom = region
                self.name_regions[index] = (
                    left - window_left,
                    top - window_top,
                    right - window_left,
                    bottom - window_top,
                )
                self.name_relative[index] = True
            else:
                self.name_regions[index] = region
                self.name_relative[index] = False
            self._save_config()
            self._refresh_overlay_name_visibility()
            self._refresh_overlay_zones()
        self._refresh_status()

    def start(self) -> None:
        if self.running:
            return
        if not self.region:
            messagebox.showinfo("Select region", "Select the Lineage 2 chat area first.")
            return
        if not self.ocr.available():
            self._refresh_status()
            messagebox.showerror(
                "OCR missing",
                "No OCR backend was found.\n\nWindows OCR should work on Windows 10/11. "
                "As a fallback, install Tesseract OCR with Russian language data, "
                "or set TESSERACT_CMD to the full path of tesseract.exe.",
            )
            return

        self.running = True
        self.baseline_next_tick = True
        self.baseline_until = time.time() + self._baseline_seconds()
        self.worker = threading.Thread(target=self._watch_loop, daemon=True)
        self.worker.start()
        self._refresh_status()

    def stop(self) -> None:
        self.running = False
        self._refresh_status()

    def reset(self) -> None:
        self.events.clear()
        self.previous_visible_counts.clear()
        self.previous_visible_keys.clear()
        self.visible_damage_max_counts.clear()
        self.counted_damage_contexts.clear()
        self._clear_pending_damage()
        self.baseline_next_tick = self.running
        self.baseline_until = time.time() + self._baseline_seconds() if self.running else 0.0
        self.log.delete(0, END)
        self._refresh_status()

    def _watch_loop(self) -> None:
        while self.running:
            try:
                self._tick()
                self.root.after(0, self._refresh_status)
            except Exception as exc:  # noqa: BLE001 - keep GUI alive and report the real issue.
                self.running = False
                self.root.after(0, lambda message=str(exc): self._refresh_status(message))
            time.sleep(self.poll_seconds)

    def _tick(self) -> None:
        assert self.region is not None
        image = ImageGrab.grab(bbox=self.region)
        image.save(CAPTURE_PATH)
        text = self.ocr.read(image)
        OCR_TEXT_PATH.write_text(text, encoding="utf-8", errors="replace")
        damage_lines = self.parser.parse_lines(text)
        new_lines = self._new_visible_damage_lines(damage_lines)

        now = time.time()
        for line, amount, direction, target in new_lines:
            event = DamageEvent(now, amount, line, direction, target)
            self.events.append(event)
            self.root.after(0, self._append_log, event)
        self.root.after(0, self._refresh_overlay_zones)

    def _new_visible_damage_lines(
        self,
        damage_lines: list[tuple[str, int, str, str | None]],
    ) -> list[tuple[str, int, str, str | None]]:
        if not damage_lines:
            return []

        current_keys = [self._damage_key(item) for item in damage_lines]
        current_counts = Counter(current_keys)
        if self.baseline_next_tick or time.time() < self.baseline_until:
            self._remember_visible_damage(damage_lines, current_keys, current_counts)
            self._clear_pending_damage()
            self.baseline_next_tick = False
            return []

        if self.count_mode == "fast":
            return self._new_visible_damage_lines_fast(damage_lines, current_keys, current_counts)

        overlap = min(len(self.previous_visible_keys), len(current_keys))
        while overlap > 0 and self.previous_visible_keys[-overlap:] != current_keys[:overlap]:
            overlap -= 1

        if not self.previous_visible_keys:
            candidate_indexes = []
        elif overlap > 0:
            suffix_items = damage_lines[overlap:]
            suffix_keys = current_keys[overlap:]
            if overlap == len(self.previous_visible_keys):
                return self._handle_pending_suffix(damage_lines, current_keys, current_counts, suffix_items, suffix_keys)

            new_items = self._release_pending_damage()
            candidate_indexes = list(range(overlap, len(damage_lines)))
            for index in candidate_indexes:
                item = damage_lines[index]
                context_keys = self._damage_context_keys(current_keys, index)
                if any(context in self.counted_damage_contexts for context in context_keys):
                    continue
                new_items.append(item)
            self._remember_visible_damage(damage_lines, current_keys, current_counts)
            return new_items
        else:
            needed = current_counts - self.visible_damage_max_counts
            if not needed:
                self._remember_visible_damage(damage_lines, current_keys, current_counts)
                return []
            self._remember_visible_damage(damage_lines, current_keys, current_counts)
            return []

        new_items: list[tuple[str, int, str, str | None]] = []
        for index in candidate_indexes:
            item = damage_lines[index]
            key = self._damage_key(item)
            context_keys = self._damage_context_keys(current_keys, index)
            if current_counts[key] <= self.visible_damage_max_counts[key]:
                continue
            if self._full_damage_context_key(current_keys, index) in self.counted_damage_contexts:
                continue

            new_items.append(item)

        self._remember_visible_damage(damage_lines, current_keys, current_counts)
        return new_items

    def _new_visible_damage_lines_fast(
        self,
        damage_lines: list[tuple[str, int, str, str | None]],
        current_keys: list[tuple[str, int]],
        current_counts: Counter[tuple[str, int]],
    ) -> list[tuple[str, int, str, str | None]]:
        overlap = min(len(self.previous_visible_keys), len(current_keys))
        while overlap > 0 and self.previous_visible_keys[-overlap:] != current_keys[:overlap]:
            overlap -= 1

        if not self.previous_visible_keys:
            candidate_indexes: list[int] = []
        elif overlap > 0:
            candidate_indexes = list(range(overlap, len(damage_lines)))
        else:
            needed = current_counts - self.visible_damage_max_counts
            if not needed:
                self._remember_visible_damage(damage_lines, current_keys, current_counts)
                return []

            candidate_indexes = []
            occurrence = Counter()
            remaining = Counter(needed)
            for index, item in enumerate(damage_lines):
                key = self._damage_key(item)
                occurrence[key] += 1
                if occurrence[key] <= self.visible_damage_max_counts[key] or remaining[key] <= 0:
                    continue
                candidate_indexes.append(index)
                remaining[key] -= 1

        new_items: list[tuple[str, int, str, str | None]] = []
        for index in candidate_indexes:
            context_key = self._full_damage_context_key(current_keys, index)
            if context_key in self.counted_damage_contexts:
                continue
            new_items.append(damage_lines[index])

        self._clear_pending_damage()
        self._remember_visible_damage(damage_lines, current_keys, current_counts)
        return new_items

    def _handle_pending_suffix(
        self,
        damage_lines: list[tuple[str, int, str, str | None]],
        current_keys: list[tuple[str, int]],
        current_counts: Counter[tuple[str, int]],
        suffix_items: list[tuple[str, int, str, str | None]],
        suffix_keys: list[tuple[str, int]],
    ) -> list[tuple[str, int, str, str | None]]:
        if not suffix_items:
            self.pending_stable_scans += 1
            if self.pending_stable_scans >= 2:
                self._clear_pending_damage()
            self._remember_visible_damage(damage_lines, current_keys, current_counts)
            return []

        new_items = self._release_pending_damage()
        self.pending_damage_items = list(suffix_items)
        self.pending_damage_keys = list(suffix_keys)
        self.pending_stable_scans = 0
        self._remember_visible_damage(damage_lines, current_keys, current_counts)
        return new_items

    def _release_pending_damage(self) -> list[tuple[str, int, str, str | None]]:
        released = list(self.pending_damage_items)
        self._clear_pending_damage()
        return released

    def _clear_pending_damage(self) -> None:
        self.pending_damage_items = []
        self.pending_damage_keys = []
        self.pending_stable_scans = 0

    def _remember_visible_damage(
        self,
        damage_lines: list[tuple[str, int, str, str | None]],
        current_keys: list[tuple[str, int]],
        current_counts: Counter[tuple[str, int]],
    ) -> None:
        self.previous_visible_keys = current_keys[-120:]
        self.previous_visible_counts = current_counts
        for key, count in current_counts.items():
            if count > self.visible_damage_max_counts[key]:
                self.visible_damage_max_counts[key] = count
        for index in range(len(damage_lines)):
            self.counted_damage_contexts.update(self._damage_context_keys(current_keys, index))

    @staticmethod
    def _damage_context_keys(
        current_keys: list[tuple[str, int]],
        index: int,
    ) -> set[tuple[tuple[str, int] | None, tuple[str, int], tuple[str, int] | None]]:
        full_key = DamageMeter._full_damage_context_key(current_keys, index)
        previous_key, key, next_key = full_key
        return {
            full_key,
            (previous_key, key, None),
            (None, key, next_key),
        }

    @staticmethod
    def _full_damage_context_key(
        current_keys: list[tuple[str, int]],
        index: int,
    ) -> tuple[tuple[str, int] | None, tuple[str, int], tuple[str, int] | None]:
        key = current_keys[index]
        previous_key = current_keys[index - 1] if index > 0 else None
        next_key = current_keys[index + 1] if index + 1 < len(current_keys) else None
        return (previous_key, key, next_key)

    @staticmethod
    def _damage_key(item: tuple[str, int, str, str | None]) -> tuple[str, int]:
        _line, amount, direction, target = item
        return (direction, amount)

    def _append_log(self, event: DamageEvent) -> None:
        clock = time.strftime("%H:%M:%S", time.localtime(event.timestamp))
        marker = "GIVEN" if event.direction == "out" else "TAKEN"
        target = f"   {event.target}" if event.target else ""
        self.log.insert(0, f"{clock}   {marker}   {event.amount}{target}   {event.line}")
        self.log.itemconfig(0, fg="#00ff55" if event.direction == "out" else "#ff3333")
        self._refresh_stats()

    def run(self) -> None:
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.mainloop()

    def _close(self) -> None:
        self.stop()
        if self.servitor_timer_job:
            try:
                self.root.after_cancel(self.servitor_timer_job)
            except Exception:
                pass
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.destroy()
        self.overlay = None
        self.root.destroy()


class RegionSelector:
    def __init__(self, parent: Tk, hint: str = "Drag around the chat log. Press Esc to cancel.") -> None:
        self.parent = parent
        self.hint = hint
        self.result: tuple[int, int, int, int] | None = None
        self.start_x = 0
        self.start_y = 0
        self.rect_id: int | None = None

    def select(self) -> tuple[int, int, int, int] | None:
        overlay = Toplevel(self.parent)
        overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.28)
        overlay.attributes("-topmost", True)
        overlay.configure(bg="black")
        overlay.title("Drag over Lineage 2 chat")

        from tkinter import Canvas

        canvas = Canvas(overlay, cursor="cross", bg="black", highlightthickness=0)
        canvas.pack(fill=BOTH, expand=True)

        hint = canvas.create_text(
            20,
            20,
            anchor="nw",
            text=self.hint,
            fill="white",
            font=("Segoe UI", 16, "bold"),
        )
        canvas.tag_raise(hint)

        def on_press(event) -> None:  # type: ignore[no-untyped-def]
            self.start_x = event.x_root
            self.start_y = event.y_root
            self.rect_id = canvas.create_rectangle(event.x, event.y, event.x, event.y, outline="#00ffff", width=3)

        def on_drag(event) -> None:  # type: ignore[no-untyped-def]
            if self.rect_id is None:
                return
            canvas.coords(
                self.rect_id,
                self.start_x,
                self.start_y,
                event.x_root,
                event.y_root,
            )

        def on_release(event) -> None:  # type: ignore[no-untyped-def]
            left = min(self.start_x, event.x_root)
            top = min(self.start_y, event.y_root)
            right = max(self.start_x, event.x_root)
            bottom = max(self.start_y, event.y_root)
            if right - left >= 40 and bottom - top >= 20:
                self.result = (left, top, right, bottom)
            overlay.destroy()

        def cancel(_event=None) -> None:  # type: ignore[no-untyped-def]
            self.result = None
            overlay.destroy()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        overlay.bind("<Escape>", cancel)
        overlay.wait_window()
        return self.result


if __name__ == "__main__":
    DamageMeter().run()
