# =================================================================
# Automatic Writing Assistant v3.0 - Premium Edition
# Pure tkinter + ttk.  Zero extra dependencies on Windows.
# Simulates natural keyboard typing into any focused input field.
# =================================================================

import json
import math
import os
import platform
import random
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, font as tkfont, messagebox, filedialog

try:
    import winsound
except ImportError:
    winsound = None

SYSTEM = platform.system()
APP_DIR = os.path.dirname(os.path.abspath(__file__))
PRESETS_FILE = os.path.join(APP_DIR, "presets.json")
DRAFT_FILE = os.path.join(APP_DIR, ".draft.json")
HISTORY_FILE = os.path.join(APP_DIR, ".history.json")
SETTINGS_FILE = os.path.join(APP_DIR, ".settings.json")


# ==================================================================
# Windows SendInput
# ==================================================================
if SYSTEM == "Windows":
    import ctypes
    import ctypes.wintypes

    _u32 = ctypes.windll.user32
    _IK = 1
    _KU = 0x0002
    _KUN = 0x0004
    _VS = 0x10       # Shift
    _VC = 0x11       # Ctrl
    _VM = 0x12       # Alt
    _VR = 0x0D       # Enter
    _VF9 = 0x78      # F9

    class _KBI(ctypes.Structure):
        """KEYBDINPUT"""
        _fields_ = [
            ("wVk", ctypes.c_ushort),
            ("wScan", ctypes.c_ushort),
            ("dwFlags", ctypes.c_ulong),
            ("time", ctypes.c_ulong),
            ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
        ]

    class _MI(ctypes.Structure):
        """MOUSEINPUT - needed so the union has the correct size."""
        _fields_ = [
            ("dx", ctypes.c_long),
            ("dy", ctypes.c_long),
            ("mouseData", ctypes.c_ulong),
            ("dwFlags", ctypes.c_ulong),
            ("time", ctypes.c_ulong),
            ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
        ]

    class _INP(ctypes.Structure):
        """INPUT struct - union must include MOUSEINPUT for correct sizeof."""
        class _U(ctypes.Union):
            _fields_ = [("ki", _KBI), ("mi", _MI)]
        _anonymous_ = ("u",)
        _fields_ = [("type", ctypes.c_ulong), ("u", _U)]


# ==================================================================
# Typing Backends
# ==================================================================
class TypingBackend:
    hotkey_label = "F9"
    hotkey_supported = False
    shift_enter = True

    def type_char(self, ch):
        raise NotImplementedError

    def stop_requested(self):
        return False

    def shutdown(self):
        pass


class WindowsBackend(TypingBackend):
    hotkey_supported = True

    def _send(self, vk=0, scan=0, flags=0):
        i = _INP(type=_IK)
        i.ki = _KBI(wVk=vk, wScan=scan, dwFlags=flags,
                     time=0, dwExtraInfo=None)
        _u32.SendInput(1, ctypes.byref(i), ctypes.sizeof(_INP))

    def _tap(self, vk):
        self._send(vk=vk)
        self._send(vk=vk, flags=_KU)

    def type_char(self, ch):
        if ch == "\n":
            if self.shift_enter:
                self._send(vk=_VS)
                self._tap(_VR)
                self._send(vk=_VS, flags=_KU)
            else:
                self._tap(_VR)
            return
        c = ord(ch)
        # Emoji / chars above U+FFFF  ->  UTF-16 surrogate pairs
        if c > 0xFFFF:
            hi = 0xD800 + ((c - 0x10000) >> 10)
            lo = 0xDC00 + ((c - 0x10000) & 0x3FF)
            for surrogate in (hi, lo):
                self._send(scan=surrogate, flags=_KUN)
                self._send(scan=surrogate, flags=_KUN | _KU)
            return
        m = _u32.VkKeyScanW(c)
        if m in (-1, 0xFFFF):
            self._send(scan=c, flags=_KUN)
            self._send(scan=c, flags=_KUN | _KU)
            return
        vk, sh = m & 0xFF, (m >> 8) & 0xFF
        mods = []
        if sh & 1:
            mods.append(_VS)
        if sh & 2:
            mods.append(_VC)
        if sh & 4:
            mods.append(_VM)
        for mod in mods:
            self._send(vk=mod)
        self._tap(vk)
        for mod in reversed(mods):
            self._send(vk=mod, flags=_KU)

    def stop_requested(self):
        return bool(_u32.GetAsyncKeyState(_VF9) & 0x8000)


class PynputBackend(TypingBackend):
    hotkey_supported = True

    def __init__(self):
        from pynput import keyboard as kb
        self._kb = kb
        self._ctrl = kb.Controller()
        self._flag = threading.Event()
        self._lis = kb.Listener(on_press=self._on_press)
        self._lis.start()

    def _on_press(self, key):
        if key == self._kb.Key.f9:
            self._flag.set()

    def type_char(self, ch):
        if ch == "\n":
            if self.shift_enter:
                self._ctrl.press(self._kb.Key.shift)
                self._ctrl.press(self._kb.Key.enter)
                self._ctrl.release(self._kb.Key.enter)
                self._ctrl.release(self._kb.Key.shift)
            else:
                self._ctrl.press(self._kb.Key.enter)
                self._ctrl.release(self._kb.Key.enter)
        else:
            self._ctrl.type(ch)

    def stop_requested(self):
        if self._flag.is_set():
            self._flag.clear()
            return True
        return False

    def shutdown(self):
        self._lis.stop()


def _make_backend():
    if SYSTEM == "Windows":
        return WindowsBackend()
    if SYSTEM in ("Darwin", "Linux"):
        return PynputBackend()
    raise RuntimeError("Unsupported OS: " + SYSTEM)


# ==================================================================
# Theme System
# ==================================================================
THEMES = {
    "dark": {
        "BG":      "#0c0c18",
        "BG2":     "#10102a",
        "CARD":    "#171733",
        "CARD2":   "#1c1c40",
        "CARD3":   "#22224a",
        "BORDER":  "#2a2a52",
        "ACCENT":  "#7b6fff",
        "ACCENT2": "#9d93ff",
        "ACCENT3": "#5b4fcc",
        "CYAN":    "#00d4e8",
        "GREEN":   "#00e090",
        "GREEN2":  "#30ffaa",
        "YELLOW":  "#ffc848",
        "ORANGE":  "#ff9050",
        "RED":     "#ff5858",
        "RED2":    "#ff7878",
        "FG":      "#eaeaf8",
        "FG2":     "#a8a8cc",
        "FG3":     "#686898",
        "INP_BG":  "#0e0e22",
        "SEL_BG":  "#3c3c80",
        "TAB_SEL": "#1e1e48",
    },
    "light": {
        "BG":      "#f0f1f6",
        "BG2":     "#e6e8f0",
        "CARD":    "#ffffff",
        "CARD2":   "#f6f6fc",
        "CARD3":   "#eeeeF6",
        "BORDER":  "#d0d2e0",
        "ACCENT":  "#6055e8",
        "ACCENT2": "#7a70ff",
        "ACCENT3": "#4840b8",
        "CYAN":    "#00a0b0",
        "GREEN":   "#00b070",
        "GREEN2":  "#00cc85",
        "YELLOW":  "#d4a000",
        "ORANGE":  "#e07030",
        "RED":     "#e04040",
        "RED2":    "#f05050",
        "FG":      "#1a1a30",
        "FG2":     "#505070",
        "FG3":     "#8888a8",
        "INP_BG":  "#f6f7fc",
        "SEL_BG":  "#c0c0e0",
        "TAB_SEL": "#e8e8f8",
    },
}

_current_theme = "light"


def _t(key):
    """Return a colour from the active theme."""
    return THEMES[_current_theme][key]


def _refresh_globals():
    """Push current theme colours into module-level names."""
    g = globals()
    for k, v in THEMES[_current_theme].items():
        g[k] = v


# Initial export
_refresh_globals()


# ==================================================================
# Data Managers
# ==================================================================
class PresetManager:
    def __init__(self):
        self.presets = {}
        self._load()

    def _load(self):
        try:
            if os.path.exists(PRESETS_FILE):
                with open(PRESETS_FILE, "r", encoding="utf-8") as fh:
                    self.presets = json.load(fh)
        except Exception:
            self.presets = {}

    def _save(self):
        try:
            with open(PRESETS_FILE, "w", encoding="utf-8") as fh:
                json.dump(self.presets, fh, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def names(self):
        return sorted(self.presets.keys())

    def get(self, name):
        return self.presets.get(name, "")

    def save(self, name, text):
        self.presets[name] = text
        self._save()

    def delete(self, name):
        self.presets.pop(name, None)
        self._save()


class DraftManager:
    """Auto-save / restore the editor text between sessions."""

    def __init__(self):
        self._path = DRAFT_FILE

    def save(self, text):
        try:
            with open(self._path, "w", encoding="utf-8") as fh:
                json.dump({"text": text}, fh, ensure_ascii=False)
        except Exception:
            pass

    def load(self):
        try:
            if os.path.exists(self._path):
                with open(self._path, "r", encoding="utf-8") as fh:
                    return json.load(fh).get("text", "")
        except Exception:
            pass
        return ""

    def clear(self):
        try:
            if os.path.exists(self._path):
                os.remove(self._path)
        except Exception:
            pass


class HistoryManager:
    """Track typing session history and lifetime statistics."""

    def __init__(self):
        self._path = HISTORY_FILE
        self.lifetime = {"sessions": 0, "chars": 0, "time_sec": 0.0}
        self.sessions = []
        self._load()

    def _load(self):
        try:
            if os.path.exists(self._path):
                with open(self._path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                self.lifetime = data.get("lifetime", self.lifetime)
                self.sessions = data.get("sessions", [])
        except Exception:
            pass

    def _save(self):
        try:
            with open(self._path, "w", encoding="utf-8") as fh:
                json.dump({
                    "lifetime": self.lifetime,
                    "sessions": self.sessions[-200:],  # keep last 200
                }, fh, indent=2)
        except Exception:
            pass

    def record(self, chars, elapsed, mode, repeat):
        self.lifetime["sessions"] += 1
        self.lifetime["chars"] += chars
        self.lifetime["time_sec"] += elapsed
        entry = {
            "date": time.strftime("%Y-%m-%d %H:%M"),
            "chars": chars,
            "time": round(elapsed, 1),
            "mode": mode,
            "repeat": repeat,
        }
        self.sessions.append(entry)
        self._save()


class AppSettings:
    """Persist user preferences between sessions."""

    DEFAULTS = {
        "theme": "light",
        "countdown": 5,
        "delay": 30,
        "randomness": 40,
        "mode": "normal",
        "repeat": 1,
        "minimize": True,
        "notify": True,
        "skip_empty": False,
        "shift_enter": True,
        "trim_trailing": False,
        "restore_window": True,
        "wrap": True,
        "font_size": 12,
        "win_w": 1060,
        "win_h": 800,
        "recent_files": [],
    }

    def __init__(self):
        self._path = SETTINGS_FILE
        self.data = dict(self.DEFAULTS)
        self._load()

    def _load(self):
        try:
            if os.path.exists(self._path):
                with open(self._path, "r", encoding="utf-8") as fh:
                    saved = json.load(fh)
                for k, v in saved.items():
                    if k in self.DEFAULTS:
                        self.data[k] = v
        except Exception:
            pass

    def save(self):
        try:
            with open(self._path, "w", encoding="utf-8") as fh:
                json.dump(self.data, fh, indent=2)
        except Exception:
            pass

    def __getitem__(self, key):
        return self.data.get(key, self.DEFAULTS.get(key))

    def __setitem__(self, key, val):
        self.data[key] = val


# ==================================================================
# Main Application
# ==================================================================
class App:
    VERSION = "3.0"

    def __init__(self, root, backend):
        self.root = root
        self.backend = backend
        self.presets = PresetManager()
        self.drafts = DraftManager()
        self.history = HistoryManager()
        self.settings = AppSettings()
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.worker = None
        self._typing = False
        self._paused = False
        self._chars_typed = 0
        self._total_chars = 0
        self._start_time = 0
        self._find_visible = False

        # Apply saved theme
        global _current_theme
        _current_theme = self.settings["theme"]
        _refresh_globals()

        root.title("Automatic Writing Assistant v" + self.VERSION)
        root.configure(bg=_t("BG"))
        root.protocol("WM_DELETE_WINDOW", self._quit)

        fam = tkfont.families()
        self._bf = "Segoe UI" if "Segoe UI" in fam else "Helvetica"
        self._mf = "Consolas" if "Consolas" in fam else "Courier"
        self._font_size = self.settings["font_size"]

        self._setup_styles()
        self._build_ui()
        self._bind_shortcuts()

        # Restore draft
        draft = self.drafts.load()
        if draft:
            self.textbox.insert("1.0", draft)
            self._update_stats()

        # Restore settings into UI
        self._apply_saved_settings()

    # ============================================================
    # Apply Saved Settings
    # ============================================================
    def _apply_saved_settings(self):
        s = self.settings
        self._cd_var.set(s["countdown"])
        self._cd_lbl.configure(text=str(s["countdown"]) + " s")
        self._sp_var.set(s["delay"])
        self._sp_lbl.configure(text=str(s["delay"]) + " ms")
        self._rand_var.set(s["randomness"])
        self._rand_lbl.configure(text=str(s["randomness"]) + "%")
        self._mode_var.set(s["mode"])
        self._repeat_var.set(s["repeat"])
        self._minimize_var.set(s["minimize"])
        self._notify_var.set(s["notify"])
        self._skip_nl_var.set(s["skip_empty"])
        self._shift_enter_var.set(s["shift_enter"])
        self._trim_var.set(s["trim_trailing"])
        self._restore_var.set(s["restore_window"])
        self._wrap_var.set(s["wrap"])
        self._apply_wrap()
        self._refresh_recent_menu()
        self._refresh_stats_tab()

    # ============================================================
    # Styles
    # ============================================================
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        bf = self._bf

        _bg = _t("BG"); _bg2 = _t("BG2"); _card = _t("CARD")
        _card2 = _t("CARD2"); _card3 = _t("CARD3"); _border = _t("BORDER")
        _accent = _t("ACCENT"); _accent2 = _t("ACCENT2")
        _cyan = _t("CYAN"); _green = _t("GREEN"); _green2 = _t("GREEN2")
        _yellow = _t("YELLOW"); _orange = _t("ORANGE")
        _red = _t("RED"); _red2 = _t("RED2")
        _fg = _t("FG"); _fg2 = _t("FG2"); _fg3 = _t("FG3")
        _inp = _t("INP_BG"); _sel = _t("SEL_BG"); _tab = _t("TAB_SEL")
        _dis_bg = "#c0c0d0" if _current_theme == "light" else "#2a2a44"
        _card_act = "#dddde8" if _current_theme == "light" else "#282858"

        # -- Frames --
        s.configure("BG.TFrame",    background=_bg)
        s.configure("BG2.TFrame",   background=_bg2)
        s.configure("Card.TFrame",  background=_card)
        s.configure("Card2.TFrame", background=_card2)
        s.configure("Card3.TFrame", background=_card3)
        s.configure("Accent.TFrame", background=_accent)

        # -- Labels --
        label_defs = [
            ("Hero.TLabel",     _bg,   _accent, (bf, 20, "bold")),
            ("Title.TLabel",    _bg,   _fg,     (bf, 16, "bold")),
            ("Sub.TLabel",      _bg,   _fg3,    (bf, 10)),
            ("Head.TLabel",     _card, _fg,     (bf, 12, "bold")),
            ("Head2.TLabel",    _card, _accent, (bf, 12, "bold")),
            ("HeadBG.TLabel",   _bg,   _fg,     (bf, 12, "bold")),
            ("Body.TLabel",     _card, _fg2,    (bf, 10)),
            ("BodyBG.TLabel",   _bg,   _fg2,    (bf, 10)),
            ("Body2.TLabel",    _card2, _fg2,   (bf, 10)),
            ("Val.TLabel",      _card, _accent2, (bf, 11, "bold")),
            ("ValG.TLabel",     _card, _green,   (bf, 11, "bold")),
            ("ValO.TLabel",     _card, _orange,  (bf, 11, "bold")),
            ("Stat.TLabel",     _card, _fg2,     (bf, 10)),
            ("StatBG.TLabel",   _bg,   _fg2,     (bf, 10)),
            ("StatBG2.TLabel",  _bg2,  _fg2,     (bf, 10)),
            ("Foot.TLabel",     _bg,   _fg3,     (bf, 9)),
            ("Cnt.TLabel",      _card, _fg3,     (bf, 9)),
            ("CntBG.TLabel",    _bg,   _fg3,     (bf, 9)),
            ("Big.TLabel",      _card, _fg,      (bf, 28, "bold")),
            ("BigSub.TLabel",   _card, _fg2,     (bf, 11)),
            ("StatNum.TLabel",  _card, _accent,  (bf, 22, "bold")),
            ("StatNum2.TLabel", _card, _green,   (bf, 22, "bold")),
            ("StatNum3.TLabel", _card, _cyan,    (bf, 22, "bold")),
            ("StatNum4.TLabel", _card, _orange,  (bf, 22, "bold")),
            ("StatUnit.TLabel", _card, _fg3,     (bf, 9)),
            ("Find.TLabel",     _card2, _fg2,    (bf, 10)),
        ]
        for name, bg_c, fg_c, fnt in label_defs:
            s.configure(name, background=bg_c, foreground=fg_c, font=fnt)

        # -- Buttons --
        btn_defs = [
            ("Accent.TButton",  _accent, "#fff",   _accent2, (bf, 10, "bold"), (16, 10)),
            ("Green.TButton",   _green,  "#fff",   _green2,  (bf, 10, "bold"), (14, 8)),
            ("Red.TButton",     _red,    "#fff",   _red2,    (bf, 10, "bold"), (14, 8)),
            ("Orange.TButton",  _orange, "#fff",   "#ffaa66",(bf, 10, "bold"), (14, 8)),
            ("Cyan.TButton",    _cyan,   "#fff",   "#28e0f0",(bf, 10, "bold"), (14, 8)),
            ("Card.TButton",    _card2,  _fg2,     _card_act,(bf, 10),         (12, 7)),
            ("CardSm.TButton",  _card2,  _fg2,     _card_act,(bf, 9),          (8, 5)),
            ("CardAcc.TButton", _card3,  _accent,  _card_act,(bf, 10, "bold"), (12, 7)),
            ("Pill.TButton",    _card3,  _fg2,     _card_act,(bf, 9),          (10, 4)),
        ]
        for name, bg_c, fg_c, act_bg, fnt, pad in btn_defs:
            s.configure(name, background=bg_c, foreground=fg_c,
                        font=fnt, padding=pad, borderwidth=0)
            s.map(name,
                  background=[("active", act_bg), ("disabled", _dis_bg)],
                  foreground=[("disabled", "#888")])

        # -- Scales --
        for prefix in ("A", "G", "O"):
            sn = prefix + ".Horizontal.TScale"
            s.configure(sn, background=_card, troughcolor=_border,
                        sliderthickness=18, borderwidth=0)

        # -- Progressbar --
        s.configure("pointed.Horizontal.TProgressbar",
                    troughcolor=_border, background=_accent, thickness=8)
        s.configure("green.Horizontal.TProgressbar",
                    troughcolor=_border, background=_green, thickness=8)

        # -- Combobox --
        s.configure("Dark.TCombobox", fieldbackground=_inp, background=_card,
                    foreground=_fg, selectbackground=_sel,
                    font=(bf, 10), padding=6)
        s.map("Dark.TCombobox",
              fieldbackground=[("readonly", _inp)],
              foreground=[("readonly", _fg)])

        # -- Checkbutton --
        s.configure("Dark.TCheckbutton", background=_card, foreground=_fg2,
                    font=(bf, 10), indicatorcolor=_border)
        s.map("Dark.TCheckbutton",
              background=[("active", _card)],
              indicatorcolor=[("selected", _accent)])

        s.configure("BG.TCheckbutton", background=_bg, foreground=_fg2,
                    font=(bf, 10), indicatorcolor=_border)
        s.map("BG.TCheckbutton",
              background=[("active", _bg)],
              indicatorcolor=[("selected", _accent)])

        # -- Radiobutton --
        s.configure("Dark.TRadiobutton", background=_card, foreground=_fg2,
                    font=(bf, 10), indicatorcolor=_border)
        s.map("Dark.TRadiobutton",
              background=[("active", _card)],
              indicatorcolor=[("selected", _accent)])

        # -- Notebook --
        s.configure("Dark.TNotebook", background=_bg, borderwidth=0)
        s.configure("Dark.TNotebook.Tab", background=_card2, foreground=_fg3,
                    font=(bf, 10, "bold"), padding=(20, 10))
        s.map("Dark.TNotebook.Tab",
              background=[("selected", _tab)],
              foreground=[("selected", _fg)])

        # -- Separator --
        s.configure("Dark.TSeparator", background=_border)
        s.configure("Accent.TSeparator", background=_accent)

        # -- Labelframe --
        s.configure("Card.TLabelframe", background=_card, foreground=_fg,
                    font=(bf, 10, "bold"), borderwidth=1, relief="solid")
        s.configure("Card.TLabelframe.Label", background=_card,
                    foreground=_fg, font=(bf, 10, "bold"))

    # ============================================================
    # Build UI
    # ============================================================
    def _build_ui(self):
        # ---------- Header ----------
        hdr = ttk.Frame(self.root, style="BG.TFrame")
        hdr.pack(fill="x", padx=24, pady=(16, 0))

        # Accent bar
        ttk.Frame(hdr, style="Accent.TFrame", height=3).pack(fill="x")

        hdr2 = ttk.Frame(self.root, style="BG.TFrame")
        hdr2.pack(fill="x", padx=24, pady=(8, 0))

        ttk.Label(hdr2, text="Automatic Writing Assistant",
                  style="Hero.TLabel").pack(side="left")
        ttk.Label(hdr2, text="  v" + self.VERSION,
                  style="Sub.TLabel").pack(side="left", pady=(8, 0))

        # Header buttons (right side)
        btn_row = ttk.Frame(hdr2, style="BG.TFrame")
        btn_row.pack(side="right")

        self._aot_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(btn_row, text=" Pin",
                        variable=self._aot_var,
                        style="BG.TCheckbutton",
                        command=self._toggle_aot).pack(side="right", padx=(8, 0))

        self._theme_btn = ttk.Button(
            btn_row,
            text="Light Mode" if _current_theme == "dark" else "Dark Mode",
            style="Pill.TButton", command=self._switch_theme)
        self._theme_btn.pack(side="right", padx=(8, 0))

        # ---------- Notebook ----------
        self.nb = ttk.Notebook(self.root, style="Dark.TNotebook")
        self.nb.pack(fill="both", expand=True, padx=24, pady=(12, 0))

        self._tab_editor   = ttk.Frame(self.nb, style="BG.TFrame")
        self._tab_settings = ttk.Frame(self.nb, style="BG.TFrame")
        self._tab_log      = ttk.Frame(self.nb, style="BG.TFrame")
        self._tab_stats    = ttk.Frame(self.nb, style="BG.TFrame")
        self._tab_help     = ttk.Frame(self.nb, style="BG.TFrame")

        self.nb.add(self._tab_editor,   text="   Editor   ")
        self.nb.add(self._tab_settings, text="   Settings   ")
        self.nb.add(self._tab_log,      text="   Live Log   ")
        self.nb.add(self._tab_stats,    text="   Statistics   ")
        self.nb.add(self._tab_help,     text="   How to Use   ")

        self._build_editor_tab()
        self._build_settings_tab()
        self._build_log_tab()
        self._build_stats_tab()
        self._build_help_tab()
        self._build_bottom_bar()

    # ============================================================
    # Editor Tab
    # ============================================================
    def _build_editor_tab(self):
        tab = self._tab_editor

        # --- Toolbar ---
        tb = ttk.Frame(tab, style="BG2.TFrame")
        tb.pack(fill="x", padx=8, pady=(10, 0))

        # File group
        ttk.Button(tb, text="Open", style="Card.TButton",
                   command=self._open_file).pack(side="left", padx=2)
        ttk.Button(tb, text="Save", style="Card.TButton",
                   command=self._save_file).pack(side="left", padx=2)

        # Recent files menu
        self._recent_mb = ttk.Menubutton(tb, text="Recent",
                                         style="Card.TButton")
        self._recent_menu = tk.Menu(self._recent_mb, tearoff=0,
                                    bg=_t("CARD"), fg=_t("FG"),
                                    activebackground=_t("ACCENT"),
                                    activeforeground="#fff",
                                    font=(self._bf, 10))
        self._recent_mb["menu"] = self._recent_menu
        self._recent_mb.pack(side="left", padx=2)

        self._sep(tb)

        ttk.Button(tb, text="Paste", style="Card.TButton",
                   command=self._paste_clip).pack(side="left", padx=2)
        ttk.Button(tb, text="Clear", style="Card.TButton",
                   command=self._clear_text).pack(side="left", padx=2)
        ttk.Button(tb, text="Undo", style="CardSm.TButton",
                   command=lambda: self.textbox.event_generate("<<Undo>>")
                   ).pack(side="left", padx=2)
        ttk.Button(tb, text="Redo", style="CardSm.TButton",
                   command=lambda: self.textbox.event_generate("<<Redo>>")
                   ).pack(side="left", padx=2)

        self._sep(tb)

        # Presets
        ttk.Label(tb, text="Presets:", style="StatBG2.TLabel"
                  ).pack(side="left", padx=(4, 4))
        self._preset_var = tk.StringVar()
        self._preset_cb = ttk.Combobox(tb, textvariable=self._preset_var,
                                       style="Dark.TCombobox", width=16,
                                       state="readonly")
        self._preset_cb.pack(side="left", padx=2)
        self._refresh_presets()
        ttk.Button(tb, text="Load", style="Cyan.TButton",
                   command=self._load_preset).pack(side="left", padx=2)
        ttk.Button(tb, text="Save As", style="Green.TButton",
                   command=self._save_preset).pack(side="left", padx=2)
        ttk.Button(tb, text="Del", style="Red.TButton",
                   command=self._del_preset).pack(side="left", padx=2)

        self._sep(tb)

        # Transform menu
        self._xform_mb = ttk.Menubutton(tb, text="Transform",
                                        style="CardAcc.TButton")
        self._xform_menu = tk.Menu(self._xform_mb, tearoff=0,
                                   bg=_t("CARD"), fg=_t("FG"),
                                   activebackground=_t("ACCENT"),
                                   activeforeground="#fff",
                                   font=(self._bf, 10))
        for label, cmd in [
            ("UPPERCASE",            self._xform_upper),
            ("lowercase",            self._xform_lower),
            ("Title Case",           self._xform_title),
            ("Sentence case",        self._xform_sentence),
            ("---", None),
            ("Sort Lines A-Z",       self._xform_sort_az),
            ("Sort Lines Z-A",       self._xform_sort_za),
            ("Reverse Line Order",   self._xform_reverse),
            ("Remove Duplicate Lines", self._xform_dedupe),
            ("---", None),
            ("Number Lines",         self._xform_number),
            ("Remove Empty Lines",   self._xform_remove_empty),
            ("Trim Whitespace",      self._xform_trim),
            ("Remove Extra Spaces",  self._xform_squeeze),
        ]:
            if label == "---":
                self._xform_menu.add_separator()
            else:
                self._xform_menu.add_command(label=label, command=cmd)
        self._xform_mb["menu"] = self._xform_menu
        self._xform_mb.pack(side="left", padx=2)

        # Right side: find + zoom
        ttk.Button(tb, text="Find", style="CardSm.TButton",
                   command=self._toggle_find).pack(side="right", padx=2)

        self._zoom_lbl = ttk.Label(tb, text=str(self._font_size) + "px",
                                   style="CntBG.TLabel")
        self._zoom_lbl.pack(side="right", padx=(0, 4))
        ttk.Button(tb, text="+", style="CardSm.TButton",
                   command=self._zoom_in).pack(side="right")
        ttk.Button(tb, text="-", style="CardSm.TButton",
                   command=self._zoom_out).pack(side="right", padx=(0, 2))

        self._wrap_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(tb, text="Wrap", variable=self._wrap_var,
                        style="Dark.TCheckbutton",
                        command=self._apply_wrap).pack(side="right", padx=(0, 6))

        # --- Find & Replace Bar (hidden by default) ---
        self._find_frame = tk.Frame(tab, bg=_t("CARD2"),
                                    highlightbackground=_t("BORDER"),
                                    highlightthickness=1)
        # (not packed yet - shown on Ctrl+F)

        fi = ttk.Frame(self._find_frame, style="Card2.TFrame")
        fi.pack(fill="x", padx=8, pady=6)

        ttk.Label(fi, text="Find:", style="Find.TLabel").pack(side="left")
        self._find_var = tk.StringVar()
        self._find_entry = tk.Entry(
            fi, textvariable=self._find_var, font=(self._bf, 10),
            bg=_t("INP_BG"), fg=_t("FG"), insertbackground=_t("ACCENT"),
            relief="flat", highlightthickness=1,
            highlightbackground=_t("BORDER"), width=20)
        self._find_entry.pack(side="left", padx=(4, 8))
        self._find_entry.bind("<Return>", lambda e: self._find_next())
        self._find_entry.bind("<Escape>", lambda e: self._toggle_find())

        ttk.Button(fi, text="<", style="CardSm.TButton",
                   command=self._find_prev).pack(side="left", padx=1)
        ttk.Button(fi, text=">", style="CardSm.TButton",
                   command=self._find_next).pack(side="left", padx=1)

        self._find_count_lbl = ttk.Label(fi, text="", style="Find.TLabel")
        self._find_count_lbl.pack(side="left", padx=(6, 8))

        ttk.Label(fi, text="Replace:", style="Find.TLabel"
                  ).pack(side="left", padx=(8, 0))
        self._replace_var = tk.StringVar()
        self._replace_entry = tk.Entry(
            fi, textvariable=self._replace_var, font=(self._bf, 10),
            bg=_t("INP_BG"), fg=_t("FG"), insertbackground=_t("ACCENT"),
            relief="flat", highlightthickness=1,
            highlightbackground=_t("BORDER"), width=16)
        self._replace_entry.pack(side="left", padx=(4, 4))

        ttk.Button(fi, text="Replace", style="CardSm.TButton",
                   command=self._replace_one).pack(side="left", padx=1)
        ttk.Button(fi, text="All", style="CardSm.TButton",
                   command=self._replace_all).pack(side="left", padx=1)

        self._find_case_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(fi, text="Aa", variable=self._find_case_var,
                        style="Dark.TCheckbutton",
                        command=self._find_highlight_all
                        ).pack(side="left", padx=(8, 0))

        ttk.Button(fi, text="X", style="CardSm.TButton",
                   command=self._toggle_find).pack(side="right")

        # --- Text Card ---
        self._tcard = tk.Frame(tab, bg=_t("CARD"),
                               highlightbackground=_t("BORDER"),
                               highlightthickness=1)
        self._tcard.pack(fill="both", expand=True, padx=8, pady=(8, 0))
        tcard = self._tcard

        # Text header
        thdr = ttk.Frame(tcard, style="Card.TFrame")
        thdr.pack(fill="x", padx=16, pady=(12, 2))

        ttk.Label(thdr, text="Text to Type", style="Head.TLabel"
                  ).pack(side="left")

        self._char_lbl = ttk.Label(thdr,
                                   text="0 chars | 0 words | 0 lines",
                                   style="Cnt.TLabel")
        self._char_lbl.pack(side="right")

        self._pos_lbl = ttk.Label(thdr, text="Ln 1, Col 1",
                                  style="Cnt.TLabel")
        self._pos_lbl.pack(side="right", padx=(0, 12))

        # Text widget with scrollbar
        twrap = tk.Frame(tcard, bg=_t("INP_BG"),
                         highlightbackground=_t("BORDER"),
                         highlightthickness=1)
        twrap.pack(fill="both", expand=True, padx=16, pady=(4, 14))

        self.textbox = tk.Text(
            twrap, font=(self._mf, self._font_size),
            bg=_t("INP_BG"), fg=_t("FG"),
            insertbackground=_t("ACCENT"),
            selectbackground=_t("SEL_BG"),
            selectforeground="#ffffff",
            relief="flat", wrap="word",
            padx=14, pady=12, undo=True, maxundo=80,
            blockcursor=False, spacing1=3, spacing3=3,
        )
        sb = tk.Scrollbar(twrap, command=self.textbox.yview,
                          bg=_t("CARD"), troughcolor=_t("INP_BG"),
                          highlightthickness=0, bd=0, width=10)
        self.textbox.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.textbox.pack(side="left", fill="both", expand=True)

        self.textbox.bind("<KeyRelease>", self._on_text_key)
        self.textbox.bind("<<Paste>>",
                          lambda e: self.root.after(50, self._update_stats))

        # Find highlight tag
        self.textbox.tag_configure("found",
                                   background=_t("YELLOW"),
                                   foreground="#000")
        self.textbox.tag_configure("found_current",
                                   background=_t("ACCENT"),
                                   foreground="#fff")

        # Context menu
        self._ctx_menu = tk.Menu(self.textbox, tearoff=0,
                                 bg=_t("CARD"), fg=_t("FG"),
                                 activebackground=_t("ACCENT"),
                                 activeforeground="#fff", bd=0,
                                 font=(self._bf, 10))
        for lbl, cmd in [
            ("Cut          Ctrl+X",  lambda: self.textbox.event_generate("<<Cut>>")),
            ("Copy         Ctrl+C",  lambda: self.textbox.event_generate("<<Copy>>")),
            ("Paste        Ctrl+V",  lambda: self.textbox.event_generate("<<Paste>>")),
            (None, None),
            ("Select All   Ctrl+A",  self._select_all),
            ("Clear All",            self._clear_text),
            (None, None),
            ("Find         Ctrl+F",  self._toggle_find),
            (None, None),
            ("Undo         Ctrl+Z",  lambda: self.textbox.event_generate("<<Undo>>")),
            ("Redo         Ctrl+Y",  lambda: self.textbox.event_generate("<<Redo>>")),
        ]:
            if lbl is None:
                self._ctx_menu.add_separator()
            else:
                self._ctx_menu.add_command(label=lbl, command=cmd)
        self.textbox.bind("<Button-3>", self._show_ctx_menu)

        # --- Stats row ---
        scard = tk.Frame(tab, bg=_t("CARD"),
                         highlightbackground=_t("BORDER"),
                         highlightthickness=1)
        scard.pack(fill="x", padx=8, pady=(6, 8))

        si = ttk.Frame(scard, style="Card.TFrame")
        si.pack(fill="x", padx=16, pady=10)

        self._stat_frames = {}
        for key, label, sty in [
            ("chars", "Characters", "StatNum.TLabel"),
            ("words", "Words",      "StatNum2.TLabel"),
            ("lines", "Lines",      "StatNum3.TLabel"),
            ("eta",   "Est. Time",  "StatNum4.TLabel"),
        ]:
            box = ttk.Frame(si, style="Card.TFrame")
            box.pack(side="left", fill="x", expand=True)
            val_l = ttk.Label(box, text="0", style=sty)
            val_l.pack()
            ttk.Label(box, text=label, style="StatUnit.TLabel").pack()
            self._stat_frames[key] = val_l

    # ============================================================
    # Settings Tab
    # ============================================================
    def _build_settings_tab(self):
        tab = self._tab_settings

        # Scrollable frame
        canvas = tk.Canvas(tab, bg=_t("BG"), highlightthickness=0, bd=0)
        vsb = tk.Scrollbar(tab, orient="vertical", command=canvas.yview,
                           bg=_t("CARD"), troughcolor=_t("BG"),
                           highlightthickness=0, bd=0, width=10)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", pady=8)
        canvas.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        wrapper = ttk.Frame(canvas, style="BG.TFrame")
        canvas.create_window((0, 0), window=wrapper, anchor="nw")

        def _cfg(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Match width
            canvas.itemconfigure(canvas.find_all()[0], width=canvas.winfo_width())
        wrapper.bind("<Configure>", _cfg)
        canvas.bind("<Configure>", _cfg)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * (e.delta / 120)), "units"), add="+")

        # ---- Typing Mode ----
        self._card_header(wrapper, "Typing Mode")
        mode_card = self._make_card(wrapper)
        mi = ttk.Frame(mode_card, style="Card.TFrame")
        mi.pack(fill="x", padx=20, pady=14)

        mode_row = ttk.Frame(mi, style="Card.TFrame")
        mode_row.pack(fill="x")

        self._mode_var = tk.StringVar(value="normal")
        for val, txt, desc in [
            ("normal", "Constant",   "Fixed delay between each keystroke"),
            ("human",  "Human-like", "Natural variance, pauses at punctuation"),
            ("burst",  "Burst",      "Fast bursts with micro-pauses between"),
        ]:
            rf = ttk.Frame(mode_row, style="Card.TFrame")
            rf.pack(side="left", padx=(0, 24))
            ttk.Radiobutton(rf, text=txt, value=val,
                            variable=self._mode_var,
                            style="Dark.TRadiobutton").pack(anchor="w")
            ttk.Label(rf, text=desc, style="Cnt.TLabel").pack(anchor="w")

        # ---- Speed Presets ----
        self._card_header(wrapper, "Quick Speed Presets")
        sp_card = self._make_card(wrapper)
        sp_inner = ttk.Frame(sp_card, style="Card.TFrame")
        sp_inner.pack(fill="x", padx=20, pady=14)

        sp_row = ttk.Frame(sp_inner, style="Card.TFrame")
        sp_row.pack(fill="x")

        presets = [
            ("Slow  (80ms)",    80,  "Orange.TButton"),
            ("Normal  (30ms)",  30,  "Cyan.TButton"),
            ("Fast  (15ms)",    15,  "Green.TButton"),
            ("Blazing  (5ms)",  5,   "Accent.TButton"),
        ]
        for txt, val, sty in presets:
            ttk.Button(sp_row, text=txt, style=sty,
                       command=lambda v=val: self._apply_speed_preset(v)
                       ).pack(side="left", padx=(0, 10))

        ttk.Label(sp_inner,
                  text="Click a preset to instantly set typing delay.",
                  style="Cnt.TLabel").pack(anchor="w", pady=(8, 0))

        # ---- Speed & Timing ----
        self._card_header(wrapper, "Speed & Timing")
        st_card = self._make_card(wrapper)
        spi = ttk.Frame(st_card, style="Card.TFrame")
        spi.pack(fill="x", padx=20, pady=14)

        grid = ttk.Frame(spi, style="Card.TFrame")
        grid.pack(fill="x")

        # Countdown
        c1 = ttk.Frame(grid, style="Card.TFrame")
        c1.pack(side="left", fill="x", expand=True)
        ttk.Label(c1, text="Countdown (seconds)",
                  style="Body.TLabel").pack(anchor="w")
        cr = ttk.Frame(c1, style="Card.TFrame")
        cr.pack(anchor="w", pady=(6, 0))
        self._cd_var = tk.IntVar(value=5)
        ttk.Scale(cr, from_=1, to=30, variable=self._cd_var,
                  orient="horizontal", length=200,
                  style="A.Horizontal.TScale",
                  command=self._on_cd).pack(side="left")
        self._cd_lbl = ttk.Label(cr, text="5 s",
                                 style="Val.TLabel", width=5)
        self._cd_lbl.pack(side="left", padx=(10, 0))

        # Delay
        c2 = ttk.Frame(grid, style="Card.TFrame")
        c2.pack(side="left", fill="x", expand=True)
        ttk.Label(c2, text="Typing Delay (ms / char)",
                  style="Body.TLabel").pack(anchor="w")
        sr = ttk.Frame(c2, style="Card.TFrame")
        sr.pack(anchor="w", pady=(6, 0))
        self._sp_var = tk.IntVar(value=30)
        ttk.Scale(sr, from_=5, to=300, variable=self._sp_var,
                  orient="horizontal", length=200,
                  style="G.Horizontal.TScale",
                  command=self._on_sp).pack(side="left")
        self._sp_lbl = ttk.Label(sr, text="30 ms",
                                 style="ValG.TLabel", width=7)
        self._sp_lbl.pack(side="left", padx=(10, 0))

        # Randomness
        c3 = ttk.Frame(grid, style="Card.TFrame")
        c3.pack(side="left", fill="x", expand=True)
        ttk.Label(c3, text="Randomness %",
                  style="Body.TLabel").pack(anchor="w")
        rr = ttk.Frame(c3, style="Card.TFrame")
        rr.pack(anchor="w", pady=(6, 0))
        self._rand_var = tk.IntVar(value=40)
        ttk.Scale(rr, from_=0, to=100, variable=self._rand_var,
                  orient="horizontal", length=200,
                  style="O.Horizontal.TScale",
                  command=self._on_rand).pack(side="left")
        self._rand_lbl = ttk.Label(rr, text="40%",
                                   style="ValO.TLabel", width=5)
        self._rand_lbl.pack(side="left", padx=(10, 0))

        # ---- Options ----
        self._card_header(wrapper, "Behavior Options")
        opt_card = self._make_card(wrapper)
        oi = ttk.Frame(opt_card, style="Card.TFrame")
        oi.pack(fill="x", padx=20, pady=14)

        orow = ttk.Frame(oi, style="Card.TFrame")
        orow.pack(fill="x")

        # Col 1
        oc1 = ttk.Frame(orow, style="Card.TFrame")
        oc1.pack(side="left", fill="x", expand=True)

        self._minimize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc1, text="Auto-minimize before typing",
                        variable=self._minimize_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        self._notify_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc1, text="Sound notification on complete",
                        variable=self._notify_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        self._skip_nl_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(oc1, text="Skip empty lines",
                        variable=self._skip_nl_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        self._shift_enter_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc1, text="Shift+Enter for newlines (chat apps)",
                        variable=self._shift_enter_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        # Col 2
        oc2 = ttk.Frame(orow, style="Card.TFrame")
        oc2.pack(side="left", fill="x", expand=True)

        self._trim_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(oc2, text="Trim trailing spaces per line",
                        variable=self._trim_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        self._restore_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc2, text="Restore window after typing",
                        variable=self._restore_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        self._auto_draft_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc2, text="Auto-save draft on exit",
                        variable=self._auto_draft_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        self._progress_title_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc2, text="Show progress in title bar",
                        variable=self._progress_title_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=3)

        # Repeat row
        rep_row = ttk.Frame(oi, style="Card.TFrame")
        rep_row.pack(anchor="w", pady=(10, 0))
        ttk.Label(rep_row, text="Repeat:", style="Body.TLabel"
                  ).pack(side="left")
        self._repeat_var = tk.IntVar(value=1)
        rep_spin = tk.Spinbox(rep_row, from_=1, to=99, width=4,
                              textvariable=self._repeat_var,
                              font=(self._bf, 10),
                              bg=_t("INP_BG"), fg=_t("FG"),
                              buttonbackground=_t("CARD"),
                              insertbackground=_t("ACCENT"),
                              highlightthickness=1,
                              highlightbackground=_t("BORDER"),
                              relief="flat")
        rep_spin.pack(side="left", padx=(8, 4))
        ttk.Label(rep_row, text="time(s)", style="Body.TLabel"
                  ).pack(side="left")

    # ============================================================
    # Live Log Tab
    # ============================================================
    def _build_log_tab(self):
        tab = self._tab_log

        lcard = tk.Frame(tab, bg=_t("CARD"),
                         highlightbackground=_t("BORDER"),
                         highlightthickness=1)
        lcard.pack(fill="both", expand=True, padx=8, pady=8)

        lhdr = ttk.Frame(lcard, style="Card.TFrame")
        lhdr.pack(fill="x", padx=14, pady=(12, 6))
        ttk.Label(lhdr, text="Typing Log", style="Head.TLabel"
                  ).pack(side="left")
        ttk.Button(lhdr, text="Export", style="Card.TButton",
                   command=self._export_log).pack(side="right", padx=(4, 0))
        ttk.Button(lhdr, text="Clear", style="Card.TButton",
                   command=self._clear_log).pack(side="right")

        lwrap = tk.Frame(lcard, bg=_t("INP_BG"),
                         highlightbackground=_t("BORDER"),
                         highlightthickness=1)
        lwrap.pack(fill="both", expand=True, padx=14, pady=(2, 14))

        self._log = tk.Text(
            lwrap, font=(self._mf, 10), bg=_t("INP_BG"), fg=_t("FG2"),
            insertbackground=_t("ACCENT"), relief="flat", wrap="word",
            padx=12, pady=10, state="disabled")
        lsb = tk.Scrollbar(lwrap, command=self._log.yview,
                           bg=_t("CARD"), troughcolor=_t("INP_BG"),
                           highlightthickness=0, bd=0, width=10)
        self._log.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)

        # Colour tags
        for tag, col_key in [
            ("info",    "FG2"),
            ("success", "GREEN"),
            ("warn",    "YELLOW"),
            ("error",   "RED"),
            ("accent",  "ACCENT2"),
            ("dim",     "FG3"),
            ("cyan",    "CYAN"),
        ]:
            self._log.tag_configure(tag, foreground=_t(col_key))

    # ============================================================
    # Statistics Tab
    # ============================================================
    def _build_stats_tab(self):
        tab = self._tab_stats

        # Scrollable
        canvas = tk.Canvas(tab, bg=_t("BG"), highlightthickness=0, bd=0)
        vsb = tk.Scrollbar(tab, orient="vertical", command=canvas.yview,
                           bg=_t("CARD"), troughcolor=_t("BG"),
                           highlightthickness=0, bd=0, width=10)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", pady=8)
        canvas.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        inner = ttk.Frame(canvas, style="BG.TFrame")
        canvas.create_window((0, 0), window=inner, anchor="nw")

        def _cfg(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(canvas.find_all()[0], width=canvas.winfo_width())
        inner.bind("<Configure>", _cfg)
        canvas.bind("<Configure>", _cfg)

        self._stats_inner = inner

        # -- Lifetime Stats Row --
        ttk.Label(inner, text="Lifetime Statistics",
                  style="HeadBG.TLabel").pack(anchor="w", padx=4, pady=(4, 8))

        row = ttk.Frame(inner, style="BG.TFrame")
        row.pack(fill="x", padx=4, pady=(0, 12))

        self._lt_cards = {}
        for key, label, sty in [
            ("sessions", "Total Sessions",  "StatNum.TLabel"),
            ("chars",    "Characters Typed", "StatNum2.TLabel"),
            ("time",     "Total Time",       "StatNum3.TLabel"),
            ("speed",    "Avg Speed",        "StatNum4.TLabel"),
        ]:
            card = tk.Frame(row, bg=_t("CARD"),
                            highlightbackground=_t("BORDER"),
                            highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ci = ttk.Frame(card, style="Card.TFrame")
            ci.pack(fill="x", padx=16, pady=14)
            vl = ttk.Label(ci, text="0", style=sty)
            vl.pack()
            ttk.Label(ci, text=label, style="StatUnit.TLabel").pack()
            self._lt_cards[key] = vl

        # -- Session History --
        ttk.Label(inner, text="Session History",
                  style="HeadBG.TLabel").pack(anchor="w", padx=4, pady=(4, 8))

        hist_card = tk.Frame(inner, bg=_t("CARD"),
                             highlightbackground=_t("BORDER"),
                             highlightthickness=1)
        hist_card.pack(fill="x", padx=4, pady=(0, 8))

        # Header row
        hrow = ttk.Frame(hist_card, style="Card.TFrame")
        hrow.pack(fill="x", padx=16, pady=(12, 4))
        for text, w in [("Date", 18), ("Characters", 12),
                        ("Time", 10), ("Mode", 10), ("Repeats", 8)]:
            ttk.Label(hrow, text=text, style="Head.TLabel",
                      width=w).pack(side="left")

        ttk.Separator(hist_card, orient="horizontal",
                      style="Dark.TSeparator").pack(fill="x", padx=16, pady=2)

        # Scrollable session list
        self._hist_list_frame = ttk.Frame(hist_card, style="Card.TFrame")
        self._hist_list_frame.pack(fill="x", padx=16, pady=(0, 12))

        # Clear history button
        bf = ttk.Frame(hist_card, style="Card.TFrame")
        bf.pack(fill="x", padx=16, pady=(0, 12))
        ttk.Button(bf, text="Clear History", style="Red.TButton",
                   command=self._clear_history).pack(side="right")

    def _refresh_stats_tab(self):
        """Update the stats tab with latest data."""
        lt = self.history.lifetime
        self._lt_cards["sessions"].configure(text=str(lt["sessions"]))
        self._lt_cards["chars"].configure(text=self._fmt_number(lt["chars"]))
        self._lt_cards["time"].configure(text=self._fmt_time(lt["time_sec"]))
        if lt["time_sec"] > 0:
            cps = lt["chars"] / lt["time_sec"]
            self._lt_cards["speed"].configure(
                text=str(int(cps)) + " c/s")
        else:
            self._lt_cards["speed"].configure(text="--")

        # Session rows
        for child in self._hist_list_frame.winfo_children():
            child.destroy()

        for sess in reversed(self.history.sessions[-50:]):
            r = ttk.Frame(self._hist_list_frame, style="Card.TFrame")
            r.pack(fill="x", pady=1)
            ttk.Label(r, text=sess.get("date", ""),
                      style="Body.TLabel", width=18).pack(side="left")
            ttk.Label(r, text=str(sess.get("chars", 0)),
                      style="Body.TLabel", width=12).pack(side="left")
            ttk.Label(r, text=self._fmt_time(sess.get("time", 0)),
                      style="Body.TLabel", width=10).pack(side="left")
            ttk.Label(r, text=sess.get("mode", ""),
                      style="Body.TLabel", width=10).pack(side="left")
            ttk.Label(r, text=str(sess.get("repeat", 1)),
                      style="Body.TLabel", width=8).pack(side="left")

    def _clear_history(self):
        if messagebox.askyesno("Clear History",
                               "Delete all session history?"):
            self.history.lifetime = {"sessions": 0, "chars": 0,
                                     "time_sec": 0.0}
            self.history.sessions = []
            self.history._save()
            self._refresh_stats_tab()
            self._log_msg("Session history cleared.", "warn")

    # ============================================================
    # How to Use Tab
    # ============================================================
    def _build_help_tab(self):
        tab = self._tab_help

        canvas = tk.Canvas(tab, bg=_t("BG"), highlightthickness=0, bd=0)
        vsb = tk.Scrollbar(tab, orient="vertical", command=canvas.yview,
                           bg=_t("CARD"), troughcolor=_t("BG"),
                           highlightthickness=0, bd=0, width=10)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", pady=8)
        canvas.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        inner = ttk.Frame(canvas, style="BG.TFrame")
        canvas.create_window((0, 0), window=inner, anchor="nw")

        def _cfg(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(canvas.find_all()[0], width=canvas.winfo_width())
        inner.bind("<Configure>", _cfg)
        canvas.bind("<Configure>", _cfg)

        sections = [
            ("Getting Started",
             "Welcome to the Automatic Writing Assistant! This app simulates\n"
             "keyboard typing into any focused input field.\n\n"
             "Quick start in 3 steps:\n"
             "  1. Paste or type your text in the Editor tab\n"
             "  2. Click 'Start Typing' at the bottom\n"
             "  3. Quickly click into the target field during the countdown"),

            ("Enter Your Text",
             "Go to the Editor tab. You have several ways to add text:\n\n"
             "  - Type directly into the text area\n"
             "  - Click 'Paste' to paste from your clipboard\n"
             "  - Click 'Open' to load a text file (Ctrl+O)\n"
             "  - Load a saved preset from the Presets dropdown\n"
             "  - Use 'Recent' to quickly re-open previous files\n\n"
             "The stats bar shows character/word/line count and ETA."),

            ("Configure Settings",
             "Go to the Settings tab to customise:\n\n"
             "  Typing Mode:\n"
             "    Constant  -- Fixed delay between each character\n"
             "    Human-like -- Random variation with pauses at punctuation\n"
             "    Burst -- Fast bursts of 3-8 chars with micro-pauses\n\n"
             "  Speed Presets: Click Slow/Normal/Fast/Blazing for one-click setup.\n\n"
             "  Countdown: Time before typing starts (1-30 seconds).\n"
             "  Typing Delay: Milliseconds per character (5-300ms).\n"
             "  Randomness: Variation in Human-like mode (0-100%)."),

            ("Start Typing",
             "Click 'Start Typing' at the bottom (or Ctrl+Enter).\n\n"
             "During the countdown:\n"
             "  - Switch to the target window\n"
             "  - Click on the target text field\n"
             "  - The cursor must be where you want text to appear\n\n"
             "The app auto-minimises if that option is enabled."),

            ("Control Typing",
             "While typing is in progress:\n\n"
             "  Pause / Resume: Click the orange Pause button\n"
             "  Stop: Click the red Stop button\n"
             "  Emergency Stop: Press F9 on your keyboard\n\n"
             "Progress, elapsed time, and ETA are displayed.\n"
             "WPM (words per minute) is shown in real time."),

            ("Find & Replace",
             "Press Ctrl+F to open the Find & Replace bar.\n\n"
             "  - Type a search term and press Enter or click > to find next\n"
             "  - Click < to go to previous match\n"
             "  - Check 'Aa' for case-sensitive search\n"
             "  - Enter replacement text and click Replace/All\n"
             "  - Press Escape or X to close the bar"),

            ("Text Transforms",
             "Click 'Transform' in the toolbar for text manipulation:\n\n"
             "  UPPERCASE / lowercase / Title Case / Sentence case\n"
             "  Sort Lines A-Z or Z-A\n"
             "  Reverse Line Order\n"
             "  Remove Duplicate Lines\n"
             "  Number Lines\n"
             "  Remove Empty Lines\n"
             "  Trim Whitespace / Remove Extra Spaces"),

            ("Presets",
             "Save frequently used text as presets:\n\n"
             "  1. Enter text in the Editor\n"
             "  2. Click 'Save As' in the toolbar\n"
             "  3. Enter a name for the preset\n"
             "  4. Click Save\n\n"
             "Load a preset from the dropdown. Delete with 'Del'.\n"
             "Presets are saved in presets.json."),

            ("Options Reference",
             "  Auto-minimize: Minimise window when typing starts\n"
             "  Restore window: Bring app back after typing finishes\n"
             "  Sound notification: Beep on completion (Windows)\n"
             "  Skip empty lines: Ignore blank lines\n"
             "  Trim trailing spaces: Remove trailing spaces per line\n"
             "  Shift+Enter: Use Shift+Enter for newlines (chat apps)\n"
             "  Auto-save draft: Restore your last text on next startup\n"
             "  Show progress in title: Update title bar during typing\n"
             "  Repeat: Type the same text multiple times (1-99)"),

            ("Keyboard Shortcuts",
             "  Ctrl + O         Open a text file\n"
             "  Ctrl + S         Save text to a file\n"
             "  Ctrl + Enter     Start typing\n"
             "  Ctrl + F         Find & Replace\n"
             "  Ctrl + A         Select all text\n"
             "  Ctrl + Z / Y     Undo / Redo\n"
             "  Ctrl + Plus      Zoom in editor font\n"
             "  Ctrl + Minus     Zoom out editor font\n"
             "  F9               Emergency stop while typing\n"
             "  Right-click      Context menu in editor"),

            ("Statistics Tab",
             "The Statistics tab tracks your usage over time:\n\n"
             "  - Total sessions, characters typed, total time\n"
             "  - Average typing speed (characters per second)\n"
             "  - Full session history with date, chars, time, mode\n\n"
             "Data is saved between sessions in .history.json."),

            ("Tips & Best Practices",
             "  Use 5-10 second countdown for switching windows.\n"
             "  Human-like mode looks more natural for long text.\n"
             "  Use 'Pin' to keep the app visible while working.\n"
             "  Save frequently-used text as presets.\n"
             "  Test with a small sample before a big paste.\n"
             "  Press F9 immediately if typing goes wrong!\n"
             "  Check the Live Log if something seems off.\n"
             "  Use Transform tools to clean up text before typing."),

            ("Troubleshooting",
             "  Nothing gets typed:\n"
             "    Make sure the cursor is in a text field before countdown\n"
             "    Some apps may block simulated input\n\n"
             "  Characters wrong or garbled:\n"
             "    Check keyboard layout matches text language\n"
             "    Try slower typing speed (higher ms value)\n\n"
             "  App does not open:\n"
             "    Make sure Python 3.10+ is installed\n"
             "    On macOS/Linux: pip install pynput\n\n"
             "  F9 hotkey does not work:\n"
             "    macOS: Grant Accessibility permissions\n"
             "    Linux: Ensure input access is available"),
        ]
        for title, body in sections:
            self._help_section(inner, title, body)

    def _help_section(self, parent, title, body):
        card = tk.Frame(parent, bg=_t("CARD"),
                        highlightbackground=_t("BORDER"),
                        highlightthickness=1)
        card.pack(fill="x", pady=(0, 8), padx=4)

        ci = ttk.Frame(card, style="Card.TFrame")
        ci.pack(fill="x", padx=18, pady=14)

        ttk.Label(ci, text=title, style="Head2.TLabel").pack(anchor="w")
        ttk.Separator(ci, orient="horizontal",
                      style="Dark.TSeparator").pack(fill="x", pady=(6, 8))
        lbl = tk.Label(ci, text=body, font=(self._bf, 10),
                       fg=_t("FG2"), bg=_t("CARD"),
                       justify="left", anchor="nw", wraplength=800)
        lbl.pack(anchor="w", fill="x")

    # ============================================================
    # Bottom Bar
    # ============================================================
    def _build_bottom_bar(self):
        bar = tk.Frame(self.root, bg=_t("CARD"),
                       highlightbackground=_t("BORDER"),
                       highlightthickness=1)
        bar.pack(fill="x", padx=24, pady=(8, 12))

        inner = ttk.Frame(bar, style="Card.TFrame")
        inner.pack(fill="x", padx=14, pady=10)

        # Buttons
        bf = ttk.Frame(inner, style="Card.TFrame")
        bf.pack(side="left")

        self._start_btn = ttk.Button(bf, text="  Start Typing  ",
                                     style="Accent.TButton",
                                     command=self._start)
        self._start_btn.pack(side="left")

        self._pause_btn = ttk.Button(bf, text="  Pause  ",
                                     style="Orange.TButton",
                                     command=self._pause_resume)
        self._pause_btn.pack(side="left", padx=(8, 0))
        self._pause_btn.state(["disabled"])

        self._stop_btn = ttk.Button(bf, text="  Stop  ",
                                    style="Red.TButton",
                                    command=self._stop)
        self._stop_btn.pack(side="left", padx=(8, 0))

        # Progress area
        pf = ttk.Frame(inner, style="Card.TFrame")
        pf.pack(side="right")

        self._wpm_lbl = ttk.Label(pf, text="", style="Stat.TLabel")
        self._wpm_lbl.pack(side="right", padx=(10, 0))

        self._pct_lbl = ttk.Label(pf, text="0%",
                                  style="Val.TLabel", width=5)
        self._pct_lbl.pack(side="right", padx=(8, 0))

        self._progress = ttk.Progressbar(
            pf, orient="horizontal", length=220, mode="determinate",
            style="pointed.Horizontal.TProgressbar")
        self._progress.pack(side="right")

        self._elapsed_lbl = ttk.Label(pf, text="", style="Stat.TLabel")
        self._elapsed_lbl.pack(side="right", padx=(0, 12))

        # Status row
        srow = ttk.Frame(bar, style="Card.TFrame")
        srow.pack(fill="x", padx=14, pady=(0, 6))

        self._dot = tk.Label(srow, text="*", font=(self._bf, 12),
                             fg=_t("ACCENT2"), bg=_t("CARD"))
        self._dot.pack(side="left")

        hk = ""
        if self.backend.hotkey_supported:
            hk = "  |  Press " + self.backend.hotkey_label + " to emergency-stop"

        self._status_lbl = ttk.Label(
            srow,
            text="Ready. Paste text, configure settings, click Start." + hk,
            style="Stat.TLabel")
        self._status_lbl.pack(side="left", padx=(4, 0), fill="x", expand=True)

        # Footer
        ttk.Label(
            self.root,
            text="Use responsibly. This tool does not bypass any site or app policies.",
            style="Foot.TLabel"
        ).pack(anchor="w", padx=28, pady=(0, 6))

    # ============================================================
    # UI Helpers
    # ============================================================
    def _sep(self, parent):
        """Vertical separator in a toolbar."""
        ttk.Separator(parent, orient="vertical",
                      style="Dark.TSeparator"
                      ).pack(side="left", padx=8, fill="y", pady=4)

    def _card_header(self, parent, text):
        """Section header above a card."""
        ttk.Label(parent, text=text, style="HeadBG.TLabel"
                  ).pack(anchor="w", padx=4, pady=(12, 4))

    def _make_card(self, parent):
        """Create and return a card frame."""
        card = tk.Frame(parent, bg=_t("CARD"),
                        highlightbackground=_t("BORDER"),
                        highlightthickness=1)
        card.pack(fill="x", padx=4, pady=(0, 4))
        return card

    # ============================================================
    # Keyboard Shortcuts
    # ============================================================
    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self._open_file())
        self.root.bind("<Control-O>", lambda e: self._open_file())
        self.root.bind("<Control-s>", lambda e: self._save_file())
        self.root.bind("<Control-S>", lambda e: self._save_file())
        self.root.bind("<Control-Return>", lambda e: self._start())
        self.root.bind("<Control-f>", lambda e: self._toggle_find())
        self.root.bind("<Control-F>", lambda e: self._toggle_find())
        self.root.bind("<Control-plus>", lambda e: self._zoom_in())
        self.root.bind("<Control-equal>", lambda e: self._zoom_in())
        self.root.bind("<Control-minus>", lambda e: self._zoom_out())

    # ============================================================
    # Callbacks
    # ============================================================
    def _on_cd(self, val):
        v = int(float(val))
        self._cd_var.set(v)
        self._cd_lbl.configure(text=str(v) + " s")
        self._update_eta()

    def _on_sp(self, val):
        v = int(float(val))
        self._sp_var.set(v)
        self._sp_lbl.configure(text=str(v) + " ms")
        self._update_eta()

    def _on_rand(self, val):
        v = int(float(val))
        self._rand_var.set(v)
        self._rand_lbl.configure(text=str(v) + "%")

    def _toggle_aot(self):
        self.root.attributes("-topmost", self._aot_var.get())

    def _apply_speed_preset(self, ms):
        """Apply a speed preset value."""
        self._sp_var.set(ms)
        self._sp_lbl.configure(text=str(ms) + " ms")
        self._update_eta()
        self._status("Speed: " + str(ms) + " ms/char", _t("CYAN"))

    def _on_text_key(self, _event=None):
        """Handle key release in editor: update stats + cursor pos."""
        self._update_stats()
        self._update_cursor_pos()

    def _update_cursor_pos(self):
        try:
            pos = self.textbox.index("insert")
            ln, col = pos.split(".")
            self._pos_lbl.configure(
                text="Ln " + ln + ", Col " + str(int(col) + 1))
        except Exception:
            pass

    # ============================================================
    # Find & Replace
    # ============================================================
    def _toggle_find(self):
        if self._find_visible:
            self._find_frame.pack_forget()
            self._find_visible = False
            self.textbox.tag_remove("found", "1.0", "end")
            self.textbox.tag_remove("found_current", "1.0", "end")
        else:
            self._find_frame.pack(fill="x", padx=8, pady=(6, 0),
                                  before=self._tcard)
            self._find_visible = True
            self._find_entry.focus_set()
            # Select all text in find entry
            self._find_entry.select_range(0, "end")

    def _find_highlight_all(self):
        """Highlight all occurrences of search term."""
        self.textbox.tag_remove("found", "1.0", "end")
        self.textbox.tag_remove("found_current", "1.0", "end")
        term = self._find_var.get()
        if not term:
            self._find_count_lbl.configure(text="")
            return 0
        nocase = not self._find_case_var.get()
        count = 0
        start = "1.0"
        while True:
            pos = self.textbox.search(term, start, stopindex="end",
                                      nocase=nocase)
            if not pos:
                break
            end = pos + "+" + str(len(term)) + "c"
            self.textbox.tag_add("found", pos, end)
            start = end
            count += 1
        self._find_count_lbl.configure(text=str(count) + " found")
        return count

    def _find_next(self):
        """Jump to next occurrence."""
        count = self._find_highlight_all()
        if count == 0:
            return
        term = self._find_var.get()
        nocase = not self._find_case_var.get()
        # Search from current cursor
        start = self.textbox.index("insert")
        pos = self.textbox.search(term, start, stopindex="end",
                                  nocase=nocase)
        if not pos:
            # Wrap around
            pos = self.textbox.search(term, "1.0", stopindex="end",
                                      nocase=nocase)
        if pos:
            end = pos + "+" + str(len(term)) + "c"
            self.textbox.tag_remove("found_current", "1.0", "end")
            self.textbox.tag_add("found_current", pos, end)
            self.textbox.mark_set("insert", end)
            self.textbox.see(pos)

    def _find_prev(self):
        """Jump to previous occurrence."""
        count = self._find_highlight_all()
        if count == 0:
            return
        term = self._find_var.get()
        nocase = not self._find_case_var.get()
        start = self.textbox.index("insert")
        # Go back 1 char to avoid finding current
        try:
            start = self.textbox.index(start + "-1c")
        except Exception:
            start = "end"
        pos = self.textbox.search(term, start, stopindex="1.0",
                                  nocase=nocase, backwards=True)
        if not pos:
            pos = self.textbox.search(term, "end", stopindex="1.0",
                                      nocase=nocase, backwards=True)
        if pos:
            end = pos + "+" + str(len(term)) + "c"
            self.textbox.tag_remove("found_current", "1.0", "end")
            self.textbox.tag_add("found_current", pos, end)
            self.textbox.mark_set("insert", pos)
            self.textbox.see(pos)

    def _replace_one(self):
        """Replace the current match."""
        try:
            ranges = self.textbox.tag_ranges("found_current")
            if ranges:
                self.textbox.delete(ranges[0], ranges[1])
                self.textbox.insert(ranges[0], self._replace_var.get())
                self._find_next()
        except Exception:
            pass

    def _replace_all(self):
        """Replace all occurrences."""
        term = self._find_var.get()
        repl = self._replace_var.get()
        if not term:
            return
        text = self.textbox.get("1.0", "end-1c")
        if self._find_case_var.get():
            new_text = text.replace(term, repl)
        else:
            new_text = re.sub(re.escape(term), repl, text, flags=re.IGNORECASE)
        count = (len(text) - len(new_text.replace(repl, ""))) // max(len(term), 1)
        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", new_text)
        self._find_highlight_all()
        self._update_stats()
        self._status("Replaced occurrences.", _t("GREEN"))

    # ============================================================
    # Text Transforms
    # ============================================================
    def _xform_apply(self, fn):
        """Apply a transform function to the text."""
        sel = False
        try:
            start = self.textbox.index("sel.first")
            end = self.textbox.index("sel.last")
            text = self.textbox.get(start, end)
            sel = True
        except tk.TclError:
            start = "1.0"
            end = "end-1c"
            text = self.textbox.get(start, end)

        result = fn(text)
        self.textbox.delete(start, end)
        self.textbox.insert(start, result)
        self._update_stats()

    def _xform_upper(self):
        self._xform_apply(lambda t: t.upper())

    def _xform_lower(self):
        self._xform_apply(lambda t: t.lower())

    def _xform_title(self):
        self._xform_apply(lambda t: t.title())

    def _xform_sentence(self):
        def fn(text):
            result = []
            capitalize = True
            for ch in text:
                if capitalize and ch.isalpha():
                    result.append(ch.upper())
                    capitalize = False
                else:
                    result.append(ch)
                if ch in ".!?\n":
                    capitalize = True
            return "".join(result)
        self._xform_apply(fn)

    def _xform_sort_az(self):
        self._xform_apply(
            lambda t: "\n".join(sorted(t.split("\n"),
                                       key=lambda s: s.lower())))

    def _xform_sort_za(self):
        self._xform_apply(
            lambda t: "\n".join(sorted(t.split("\n"),
                                       key=lambda s: s.lower(),
                                       reverse=True)))

    def _xform_reverse(self):
        self._xform_apply(
            lambda t: "\n".join(reversed(t.split("\n"))))

    def _xform_dedupe(self):
        def fn(text):
            seen = set()
            out = []
            for line in text.split("\n"):
                key = line.strip().lower()
                if key not in seen:
                    seen.add(key)
                    out.append(line)
            return "\n".join(out)
        self._xform_apply(fn)

    def _xform_number(self):
        def fn(text):
            lines = text.split("\n")
            width = len(str(len(lines)))
            return "\n".join(
                str(i + 1).rjust(width) + "  " + ln
                for i, ln in enumerate(lines))
        self._xform_apply(fn)

    def _xform_remove_empty(self):
        self._xform_apply(
            lambda t: "\n".join(l for l in t.split("\n") if l.strip()))

    def _xform_trim(self):
        self._xform_apply(
            lambda t: "\n".join(l.strip() for l in t.split("\n")))

    def _xform_squeeze(self):
        self._xform_apply(
            lambda t: "\n".join(
                " ".join(l.split()) for l in t.split("\n")))

    # ============================================================
    # Zoom & Wrap
    # ============================================================
    def _zoom_in(self):
        if self._font_size < 24:
            self._font_size += 1
            self.textbox.configure(font=(self._mf, self._font_size))
            self._zoom_lbl.configure(text=str(self._font_size) + "px")

    def _zoom_out(self):
        if self._font_size > 8:
            self._font_size -= 1
            self.textbox.configure(font=(self._mf, self._font_size))
            self._zoom_lbl.configure(text=str(self._font_size) + "px")

    def _apply_wrap(self):
        mode = "word" if self._wrap_var.get() else "none"
        self.textbox.configure(wrap=mode)

    # ============================================================
    # Text Stats
    # ============================================================
    def _update_stats(self, _event=None):
        text = self.textbox.get("1.0", "end-1c")
        chars = len(text)
        words = len(text.split()) if text.strip() else 0
        lines = text.count("\n") + 1 if text else 0

        self._char_lbl.configure(
            text=(str(chars) + " chars | " + str(words)
                  + " words | " + str(lines) + " lines"))
        self._stat_frames["chars"].configure(text=self._fmt_number(chars))
        self._stat_frames["words"].configure(text=self._fmt_number(words))
        self._stat_frames["lines"].configure(text=str(lines))
        self._update_eta()

    def _update_eta(self):
        text = self.textbox.get("1.0", "end-1c")
        chars = len(text)
        if chars == 0:
            self._stat_frames["eta"].configure(text="--")
            return
        delay_ms = self._sp_var.get()
        repeat = self._repeat_var.get()
        total_s = (chars * delay_ms / 1000.0) * repeat + self._cd_var.get()
        self._stat_frames["eta"].configure(text=self._fmt_time(total_s))

    def _fmt_time(self, seconds):
        seconds = float(seconds)
        if seconds < 60:
            return str(int(seconds)) + "s"
        m = int(seconds) // 60
        sc = int(seconds) % 60
        if m < 60:
            return str(m) + "m " + str(sc) + "s"
        h = m // 60
        m = m % 60
        return str(h) + "h " + str(m) + "m"

    def _fmt_number(self, n):
        """Format a number with thousand separators."""
        if n < 1000:
            return str(n)
        if n < 1000000:
            return str(n // 1000) + "," + str(n % 1000).zfill(3)
        return str(n // 1000000) + "," + str((n // 1000) % 1000).zfill(3) + "," + str(n % 1000).zfill(3)

    # ============================================================
    # File I/O
    # ============================================================
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Open Text File",
            filetypes=[("Text files", "*.txt"),
                       ("Markdown", "*.md"),
                       ("All files", "*.*")])
        if path:
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    content = fh.read()
                self.textbox.delete("1.0", "end")
                self.textbox.insert("1.0", content)
                self._update_stats()
                self._log_msg("Opened: " + os.path.basename(path), "info")
                self._status("Loaded " + os.path.basename(path), _t("CYAN"))
                self._add_recent(path)
            except Exception as exc:
                messagebox.showerror("Error",
                                     "Could not open file:\n" + str(exc))

    def _save_file(self):
        path = filedialog.asksaveasfilename(
            title="Save Text File",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"),
                       ("All files", "*.*")])
        if path:
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(self.textbox.get("1.0", "end-1c"))
                self._log_msg("Saved: " + os.path.basename(path), "info")
                self._status("Saved " + os.path.basename(path), _t("GREEN"))
                self._add_recent(path)
            except Exception as exc:
                messagebox.showerror("Error",
                                     "Could not save file:\n" + str(exc))

    def _add_recent(self, path):
        """Add a file to the recent files list."""
        recent = self.settings["recent_files"]
        if path in recent:
            recent.remove(path)
        recent.insert(0, path)
        self.settings["recent_files"] = recent[:8]
        self._refresh_recent_menu()

    def _refresh_recent_menu(self):
        self._recent_menu.delete(0, "end")
        recent = self.settings["recent_files"]
        if not recent:
            self._recent_menu.add_command(label="(no recent files)",
                                          state="disabled")
            return
        for path in recent:
            name = os.path.basename(path)
            self._recent_menu.add_command(
                label=name,
                command=lambda p=path: self._open_recent(p))

    def _open_recent(self, path):
        if not os.path.exists(path):
            self._status("File not found: " + path, _t("RED"))
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                content = fh.read()
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", content)
            self._update_stats()
            self._status("Loaded " + os.path.basename(path), _t("CYAN"))
        except Exception as exc:
            messagebox.showerror("Error", "Could not open file:\n" + str(exc))

    def _paste_clip(self):
        try:
            clip = self.root.clipboard_get()
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", clip)
            self._update_stats()
            self._status("Pasted from clipboard.", _t("CYAN"))
        except tk.TclError:
            self._status("Clipboard is empty.", _t("YELLOW"))

    def _clear_text(self):
        text = self.textbox.get("1.0", "end-1c")
        if text.strip() and len(text) > 50:
            if not messagebox.askyesno("Clear?",
                                       "Clear all text? This cannot be undone."):
                return
        self.textbox.delete("1.0", "end")
        self._update_stats()
        self._set_progress(0)
        self._status("Cleared.", _t("FG3"))

    # ============================================================
    # Presets
    # ============================================================
    def _refresh_presets(self):
        names = self.presets.names()
        self._preset_cb.configure(values=names)
        if names:
            self._preset_var.set(names[0])

    def _load_preset(self):
        name = self._preset_var.get()
        if not name:
            self._status("No preset selected.", _t("YELLOW"))
            return
        text = self.presets.get(name)
        if text:
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", text)
            self._update_stats()
            self._log_msg("Loaded preset: " + name, "info")
            self._status("Loaded preset: " + name, _t("CYAN"))

    def _save_preset(self):
        text = self.textbox.get("1.0", "end-1c")
        if not text.strip():
            self._status("Nothing to save.", _t("YELLOW"))
            return

        win = tk.Toplevel(self.root)
        win.title("Save Preset")
        win.geometry("340x140")
        win.configure(bg=_t("CARD"))
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(win, text="Preset name:", style="Head.TLabel"
                  ).pack(pady=(18, 6))
        entry = tk.Entry(win, font=(self._bf, 11),
                         bg=_t("INP_BG"), fg=_t("FG"),
                         insertbackground=_t("ACCENT"), relief="flat",
                         highlightthickness=1,
                         highlightbackground=_t("BORDER"))
        entry.pack(padx=24, fill="x")
        entry.focus_set()

        result = [None]

        def on_ok(_=None):
            result[0] = entry.get().strip()
            win.destroy()

        entry.bind("<Return>", on_ok)
        ttk.Button(win, text="Save", style="Green.TButton",
                   command=on_ok).pack(pady=12)
        win.wait_window()
        name = result[0]

        if name:
            self.presets.save(name, text)
            self._refresh_presets()
            self._preset_var.set(name)
            self._log_msg("Saved preset: " + name, "success")
            self._status("Preset saved: " + name, _t("GREEN"))

    def _del_preset(self):
        name = self._preset_var.get()
        if not name:
            self._status("No preset selected.", _t("YELLOW"))
            return
        if messagebox.askyesno("Delete Preset",
                               "Delete preset '" + name + "'?"):
            self.presets.delete(name)
            self._refresh_presets()
            self._log_msg("Deleted preset: " + name, "warn")
            self._status("Deleted preset: " + name, _t("RED"))

    # ============================================================
    # Log
    # ============================================================
    def _log_msg(self, msg, tag="info"):
        def _do():
            self._log.configure(state="normal")
            ts = time.strftime("[%H:%M:%S] ")
            self._log.insert("end", ts, "dim")
            self._log.insert("end", msg + "\n", tag)
            self._log.see("end")
            self._log.configure(state="disabled")
        self.root.after(0, _do)

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _export_log(self):
        path = filedialog.asksaveasfilename(
            title="Export Log",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")])
        if path:
            try:
                text = self._log.get("1.0", "end-1c")
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(text)
                self._status("Log exported.", _t("GREEN"))
            except Exception as exc:
                messagebox.showerror("Error",
                                     "Could not export log:\n" + str(exc))

    # ============================================================
    # Status / Progress
    # ============================================================
    def _status(self, msg, color=None):
        self.root.after(0, lambda: self._status_lbl.configure(text=msg))
        if color:
            self.root.after(0, lambda: self._dot.configure(fg=color))

    def _set_progress(self, pct):
        def _do():
            self._progress.configure(value=pct)
            self._pct_lbl.configure(text=str(int(pct)) + "%")
            if self._progress_title_var.get() and self._typing:
                self.root.title(
                    str(int(pct)) + "% - Automatic Writing Assistant")
        self.root.after(0, _do)

    def _set_elapsed(self, text):
        self.root.after(0, lambda: self._elapsed_lbl.configure(text=text))

    def _set_wpm(self, chars, elapsed):
        if elapsed > 0:
            words = chars / 5.0  # standard: 5 chars = 1 word
            wpm = int(words / (elapsed / 60.0))
            self.root.after(0,
                            lambda: self._wpm_lbl.configure(
                                text=str(wpm) + " WPM"))

    # ============================================================
    # Context Menu / Selection
    # ============================================================
    def _show_ctx_menu(self, event):
        try:
            self._ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._ctx_menu.grab_release()

    def _select_all(self):
        self.textbox.tag_add("sel", "1.0", "end")
        self.textbox.mark_set("insert", "end")

    # ============================================================
    # Theme Switching
    # ============================================================
    def _switch_theme(self):
        global _current_theme
        _current_theme = "light" if _current_theme == "dark" else "dark"
        _refresh_globals()

        lbl = "Light Mode" if _current_theme == "dark" else "Dark Mode"
        self._theme_btn.configure(text=lbl)

        self._setup_styles()
        self.root.configure(bg=_t("BG"))
        self._recolor_all(self.root)

        # Log tags
        for tag, col_key in [
            ("info", "FG2"), ("success", "GREEN"), ("warn", "YELLOW"),
            ("error", "RED"), ("accent", "ACCENT2"), ("dim", "FG3"),
            ("cyan", "CYAN"),
        ]:
            self._log.tag_configure(tag, foreground=_t(col_key))

        # Find tags
        self.textbox.tag_configure("found",
                                   background=_t("YELLOW"), foreground="#000")
        self.textbox.tag_configure("found_current",
                                   background=_t("ACCENT"), foreground="#fff")

        # Menus
        for menu in (self._ctx_menu, self._xform_menu, self._recent_menu):
            menu.configure(bg=_t("CARD"), fg=_t("FG"),
                           activebackground=_t("ACCENT"))

        self._dot.configure(bg=_t("CARD"))

        # Title bar colour on Windows
        if SYSTEM == "Windows":
            try:
                use_dark = 1 if _current_theme == "dark" else 0
                hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, 20, ctypes.byref(ctypes.c_int(use_dark)),
                    ctypes.sizeof(ctypes.c_int))
                # Force redraw
                self.root.withdraw()
                self.root.deiconify()
            except Exception:
                pass

    def _recolor_all(self, widget):
        """Recursively update bg/fg of plain tk widgets for the new theme."""
        cls = widget.winfo_class()
        try:
            if cls == "Frame":
                old_bg = str(widget.cget("bg")).lower()
                for key in ("BG", "BG2", "CARD", "CARD2", "CARD3", "INP_BG"):
                    for tn in THEMES:
                        if old_bg == THEMES[tn][key].lower():
                            widget.configure(bg=_t(key))
                            break
                try:
                    old_hl = str(widget.cget("highlightbackground")).lower()
                    for tn in THEMES:
                        if old_hl == THEMES[tn]["BORDER"].lower():
                            widget.configure(highlightbackground=_t("BORDER"))
                            break
                except Exception:
                    pass
            elif cls in ("Text", "Entry"):
                widget.configure(bg=_t("INP_BG"), fg=_t("FG"),
                                 insertbackground=_t("ACCENT"),
                                 selectbackground=_t("SEL_BG"))
            elif cls == "Label":
                old_bg = str(widget.cget("bg")).lower()
                for key in ("BG", "CARD", "CARD2", "CARD3", "INP_BG"):
                    for tn in THEMES:
                        if old_bg == THEMES[tn][key].lower():
                            widget.configure(bg=_t(key))
                            break
                old_fg = str(widget.cget("fg")).lower()
                for key in ("FG", "FG2", "FG3", "ACCENT", "ACCENT2",
                            "GREEN", "RED", "YELLOW", "ORANGE", "CYAN"):
                    for tn in THEMES:
                        if old_fg == THEMES[tn][key].lower():
                            widget.configure(fg=_t(key))
                            break
            elif cls == "Scrollbar":
                widget.configure(bg=_t("CARD"), troughcolor=_t("INP_BG"))
            elif cls == "Canvas":
                old_bg = str(widget.cget("bg")).lower()
                for key in ("BG", "CARD", "INP_BG"):
                    for tn in THEMES:
                        if old_bg == THEMES[tn][key].lower():
                            widget.configure(bg=_t(key))
                            break
            elif cls == "Spinbox":
                widget.configure(bg=_t("INP_BG"), fg=_t("FG"),
                                 buttonbackground=_t("CARD"),
                                 insertbackground=_t("ACCENT"),
                                 highlightbackground=_t("BORDER"))
        except Exception:
            pass
        for child in widget.winfo_children():
            self._recolor_all(child)

    # ============================================================
    # Actions
    # ============================================================
    def _start(self):
        if self._typing:
            self._status("Already typing!", _t("YELLOW"))
            return

        raw = self.textbox.get("1.0", "end").rstrip("\n")
        if not raw.strip():
            self._status("Please enter or paste some text first.", _t("YELLOW"))
            self._log_msg("Start aborted: no text.", "warn")
            return

        text = raw
        if self._trim_var.get():
            text = "\n".join(line.rstrip() for line in text.split("\n"))
        if self._skip_nl_var.get():
            text = "\n".join(l for l in text.split("\n") if l.strip())

        countdown = self._cd_var.get()
        delay_ms = self._sp_var.get()
        mode = self._mode_var.get()
        randomness = self._rand_var.get() / 100.0
        repeat = max(1, self._repeat_var.get())

        self.stop_event.clear()
        self.pause_event.clear()
        self._typing = True
        self._paused = False
        self._start_btn.state(["disabled"])
        self._pause_btn.state(["!disabled"])
        self._set_progress(0)

        self.backend.shift_enter = self._shift_enter_var.get()

        self._log_msg("Starting typing session", "accent")
        self._log_msg("  Mode: " + mode + " | Delay: " + str(delay_ms)
                      + "ms | Rand: " + str(int(randomness * 100))
                      + "% | Repeat: " + str(repeat) + "x", "dim")
        self._log_msg("  Text: " + str(len(text)) + " chars", "dim")
        if self.backend.shift_enter:
            self._log_msg("  Newlines: Shift+Enter (chat-safe)", "dim")
        else:
            self._log_msg("  Newlines: plain Enter", "dim")

        self.worker = threading.Thread(
            target=self._type_job,
            args=(text, countdown, delay_ms / 1000.0,
                  mode, randomness, repeat),
            daemon=True)
        self.worker.start()

    def _stop(self):
        self.stop_event.set()
        if self._paused:
            self.pause_event.set()
        self._status("Stop requested...", _t("RED"))
        self._log_msg("Stop requested by user.", "error")

    def _pause_resume(self):
        if not self._typing:
            return
        if self._paused:
            self._paused = False
            self.pause_event.set()
            self._pause_btn.configure(text="  Pause  ")
            self._status("Resumed.", _t("GREEN"))
            self._log_msg("Resumed.", "success")
        else:
            self._paused = True
            self.pause_event.clear()
            self._pause_btn.configure(text="  Resume  ")
            self._status("Paused. Click Resume to continue.", _t("ORANGE"))
            self._log_msg("Paused.", "warn")

    def _quit(self):
        # Save draft
        if self._auto_draft_var.get():
            text = self.textbox.get("1.0", "end-1c")
            if text.strip():
                self.drafts.save(text)
            else:
                self.drafts.clear()

        # Save settings
        self.settings["theme"] = _current_theme
        self.settings["countdown"] = self._cd_var.get()
        self.settings["delay"] = self._sp_var.get()
        self.settings["randomness"] = self._rand_var.get()
        self.settings["mode"] = self._mode_var.get()
        self.settings["repeat"] = self._repeat_var.get()
        self.settings["minimize"] = self._minimize_var.get()
        self.settings["notify"] = self._notify_var.get()
        self.settings["skip_empty"] = self._skip_nl_var.get()
        self.settings["shift_enter"] = self._shift_enter_var.get()
        self.settings["trim_trailing"] = self._trim_var.get()
        self.settings["restore_window"] = self._restore_var.get()
        self.settings["wrap"] = self._wrap_var.get()
        self.settings["font_size"] = self._font_size
        try:
            geo = self.root.geometry()
            parts = geo.split("x")
            w = int(parts[0])
            h = int(parts[1].split("+")[0])
            self.settings["win_w"] = w
            self.settings["win_h"] = h
        except Exception:
            pass
        self.settings.save()

        self.stop_event.set()
        if self._paused:
            self.pause_event.set()
        self.backend.shutdown()
        self.root.destroy()

    # ============================================================
    # Typing Worker Thread
    # ============================================================
    def _type_job(self, text, countdown, base_delay, mode, randomness, repeat):
        try:
            if self._minimize_var.get():
                self.root.after(0, self.root.iconify)

            self._log_msg("Countdown: " + str(countdown) + "s", "warn")
            for sec in range(countdown, 0, -1):
                if self.stop_event.is_set():
                    self._finish("Cancelled.", _t("RED"), 0, mode, repeat)
                    return
                self._status("Starting in " + str(sec)
                             + "s -- focus the target!", _t("YELLOW"))
                self._set_progress(int((countdown - sec) / countdown * 5))
                time.sleep(1)

            total = len(text) * repeat
            typed = 0
            self._start_time = time.time()

            for rep in range(repeat):
                rep_lbl = ""
                if repeat > 1:
                    rep_lbl = " [" + str(rep + 1) + "/" + str(repeat) + "]"
                    self._log_msg("Repeat " + str(rep + 1) + "/"
                                  + str(repeat), "cyan")

                hk = ""
                if self.backend.hotkey_supported:
                    hk = "  " + self.backend.hotkey_label + " to stop."
                self._status("Typing..." + rep_lbl + hk, _t("GREEN"))

                for i, ch in enumerate(text):
                    if self.stop_event.is_set():
                        self._finish("Stopped after " + str(typed) + " chars.",
                                     _t("RED"), typed, mode, repeat)
                        return
                    if self.backend.stop_requested():
                        self.stop_event.set()
                        self._finish("Stopped by " + self.backend.hotkey_label
                                     + " after " + str(typed) + " chars.",
                                     _t("RED"), typed, mode, repeat)
                        return

                    if self._paused:
                        self.pause_event.wait()
                        if self.stop_event.is_set():
                            self._finish("Stopped while paused.",
                                         _t("RED"), typed, mode, repeat)
                            return

                    self.backend.type_char(ch)
                    typed += 1

                    pct = 5 + int(typed / total * 95)
                    self._set_progress(pct)

                    elapsed = time.time() - self._start_time
                    remaining = (elapsed / typed) * (total - typed) if typed else 0
                    self._set_elapsed(
                        self._fmt_time(elapsed) + " / ~"
                        + self._fmt_time(remaining) + " left")
                    self._set_wpm(typed, elapsed)

                    delay = self._calc_delay(ch, base_delay, mode, randomness)
                    if delay > 0:
                        time.sleep(delay)

                if rep < repeat - 1:
                    self._log_msg("Waiting 1s before next repeat...", "dim")
                    time.sleep(1.0)

            elapsed = time.time() - self._start_time
            self._set_progress(100)
            msg = ("Done! " + str(typed) + " characters typed in "
                   + self._fmt_time(elapsed) + ".")
            self._finish(msg, _t("ACCENT2"), typed, mode, repeat)

            if self._notify_var.get() and SYSTEM == "Windows":
                try:
                    import winsound as ws
                    ws.MessageBeep(ws.MB_OK)
                except Exception:
                    pass

        except Exception as exc:
            self._finish("Error: " + str(exc), _t("RED"), 0, mode, repeat)

    def _calc_delay(self, ch, base, mode, rand_pct):
        if mode == "normal":
            return base
        elif mode == "human":
            variance = base * rand_pct
            delay = base + random.uniform(-variance, variance)
            if ch in ".!?":
                delay += base * random.uniform(2, 5)
            elif ch in ",;:":
                delay += base * random.uniform(0.5, 2)
            elif ch == "\n":
                delay += base * random.uniform(1, 3)
            elif ch == " ":
                delay += base * random.uniform(0, 0.5)
            return max(0, delay)
        elif mode == "burst":
            if not hasattr(self, "_burst_count"):
                self._burst_count = 0
                self._burst_size = random.randint(3, 8)
            self._burst_count += 1
            if self._burst_count >= self._burst_size:
                self._burst_count = 0
                self._burst_size = random.randint(3, 8)
                return base * random.uniform(3, 7)
            return base * 0.3
        return base

    def _finish(self, msg, color, chars_typed=0, mode="", repeat=1):
        self._typing = False
        self._paused = False
        tag = "success" if color == _t("ACCENT2") else "error"
        self._status(msg, color)
        self._log_msg(msg, tag)
        self.root.after(0, lambda: self._start_btn.state(["!disabled"]))
        self.root.after(0, lambda: self._pause_btn.state(["disabled"]))
        self.root.after(0, lambda: self._pause_btn.configure(text="  Pause  "))
        self.root.after(0, lambda: self._wpm_lbl.configure(text=""))
        self.root.after(0, lambda: self.root.title(
            "Automatic Writing Assistant v" + self.VERSION))

        # Record to history
        if chars_typed > 0:
            elapsed = time.time() - self._start_time
            self.history.record(chars_typed, elapsed, mode, repeat)
            self.root.after(0, self._refresh_stats_tab)

        if self._restore_var.get():
            self.root.after(0, self.root.deiconify)
            self.root.after(100, self.root.lift)


# ==================================================================
# Main
# ==================================================================
def main():
    root = tk.Tk()
    root.withdraw()

    # Window icon
    try:
        base_dir = getattr(sys, '_MEIPASS',
                           os.path.dirname(os.path.abspath(__file__)))
        if SYSTEM == "Windows":
            ico = os.path.join(base_dir, "icon.ico")
            if os.path.exists(ico):
                root.iconbitmap(ico)
        else:
            png = os.path.join(base_dir, "icon.png")
            if os.path.exists(png):
                img = tk.PhotoImage(file=png)
                root.iconphoto(True, img)
    except Exception:
        pass

    # Dark title bar on Windows
    if SYSTEM == "Windows":
        try:
            root.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
            use_dark = 1 if _current_theme == "dark" else 0
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 20, ctypes.byref(ctypes.c_int(use_dark)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

    try:
        backend = _make_backend()
    except Exception as exc:
        messagebox.showerror("Startup Error", str(exc))
        root.destroy()
        return

    app = App(root, backend)

    # Window geometry from settings
    w = app.settings["win_w"]
    h = app.settings["win_h"]
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry("{}x{}+{}+{}".format(w, h, (sw - w) // 2, (sh - h) // 2))
    root.minsize(860, 680)

    root.deiconify()
    root.lift()
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
