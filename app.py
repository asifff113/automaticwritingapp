# Automatic Writing Assistant - Feature-Rich Edition
# Pure tkinter + ttk. Zero extra dependencies on Windows.
# Simulates natural typing into any focused input field.

import json
import math
import os
import platform
import random
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

# ---------------------------------------------------------------------------
# Windows SendInput
# ---------------------------------------------------------------------------
if SYSTEM == "Windows":
    import ctypes, ctypes.wintypes
    _u32 = ctypes.windll.user32
    _IK = 1; _KU = 0x0002; _KUN = 0x0004
    _VS, _VC, _VM, _VR, _VF9 = 0x10, 0x11, 0x12, 0x0D, 0x78

    class _KBI(ctypes.Structure):
        """KEYBDINPUT"""
        _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
                     ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
                     ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

    class _MI(ctypes.Structure):
        """MOUSEINPUT - needed so the union has the correct size."""
        _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long),
                     ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
                     ("time", ctypes.c_ulong),
                     ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

    class _INP(ctypes.Structure):
        """INPUT struct - union must include MOUSEINPUT for correct sizeof."""
        class _U(ctypes.Union):
            _fields_ = [("ki", _KBI), ("mi", _MI)]
        _anonymous_ = ("u",)
        _fields_ = [("type", ctypes.c_ulong), ("u", _U)]


# ---------------------------------------------------------------------------
# Typing Backends
# ---------------------------------------------------------------------------
class TypingBackend:
    hotkey_label = "F9"
    hotkey_supported = False
    shift_enter = True  # Use Shift+Enter for newlines (chat apps)
    def type_char(self, ch): raise NotImplementedError
    def stop_requested(self): return False
    def shutdown(self): pass


class WindowsBackend(TypingBackend):
    hotkey_supported = True

    def _send(self, vk=0, scan=0, flags=0):
        i = _INP(type=_IK)
        i.ki = _KBI(wVk=vk, wScan=scan, dwFlags=flags, time=0, dwExtraInfo=None)
        _u32.SendInput(1, ctypes.byref(i), ctypes.sizeof(_INP))

    def _tap(self, vk):
        self._send(vk=vk); self._send(vk=vk, flags=_KU)

    def type_char(self, ch):
        if ch == "\n":
            if self.shift_enter:
                # Shift+Enter = newline without sending in chat apps
                self._send(vk=_VS)          # Shift down
                self._tap(_VR)              # Enter tap
                self._send(vk=_VS, flags=_KU)  # Shift up
            else:
                self._tap(_VR)
            return
        m = _u32.VkKeyScanW(ord(ch))
        if m in (-1, 0xFFFF):
            c = ord(ch)
            self._send(scan=c, flags=_KUN)
            self._send(scan=c, flags=_KUN | _KU)
            return
        vk, sh = m & 0xFF, (m >> 8) & 0xFF
        mods = []
        if sh & 1: mods.append(_VS)
        if sh & 2: mods.append(_VC)
        if sh & 4: mods.append(_VM)
        for mod in mods: self._send(vk=mod)
        self._tap(vk)
        for mod in reversed(mods): self._send(vk=mod, flags=_KU)

    def stop_requested(self):
        return bool(_u32.GetAsyncKeyState(_VF9) & 0x8000)


class PynputBackend(TypingBackend):
    hotkey_supported = True
    def __init__(self):
        from pynput import keyboard as kb
        self._kb = kb; self._ctrl = kb.Controller()
        self._flag = threading.Event()
        self._lis = kb.Listener(on_press=self._on_press)
        self._lis.start()

    def _on_press(self, key):
        if key == self._kb.Key.f9: self._flag.set()

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
            self._flag.clear(); return True
        return False

    def shutdown(self): self._lis.stop()


def _make_backend():
    if SYSTEM == "Windows": return WindowsBackend()
    if SYSTEM in ("Darwin", "Linux"): return PynputBackend()
    raise RuntimeError("Unsupported OS: " + SYSTEM)


# ---------------------------------------------------------------------------
# Theme Colors
# ---------------------------------------------------------------------------
BG       = "#0d0d1a"
BG2      = "#111126"
CARD     = "#161630"
CARD2    = "#1a1a3a"
BORDER   = "#282855"
ACCENT   = "#6c63ff"
ACCENT2  = "#8b83ff"
ACCENT3  = "#4a44cc"
CYAN     = "#00c9db"
GREEN    = "#00d68f"
GREEN2   = "#00f0a0"
YELLOW   = "#ffc048"
ORANGE   = "#ff8c42"
RED      = "#ff5c5c"
RED2     = "#ff7878"
FG       = "#e8e8f4"
FG2      = "#a0a0c0"
FG3      = "#6a6a90"
INP_BG   = "#0f0f24"
SEL_BG   = "#3d3d80"
TAB_SEL  = "#1e1e42"


# ---------------------------------------------------------------------------
# Preset Manager
# ---------------------------------------------------------------------------
class PresetManager:
    def __init__(self):
        self.presets = {}
        self._load()

    def _load(self):
        try:
            if os.path.exists(PRESETS_FILE):
                with open(PRESETS_FILE, "r", encoding="utf-8") as f:
                    self.presets = json.load(f)
        except Exception:
            self.presets = {}

    def _save(self):
        try:
            with open(PRESETS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.presets, f, indent=2, ensure_ascii=False)
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


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------
class App:
    VERSION = "2.0"

    def __init__(self, root, backend):
        self.root = root
        self.backend = backend
        self.presets = PresetManager()
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.worker = None
        self._typing = False
        self._paused = False
        self._chars_typed = 0
        self._total_chars = 0
        self._start_time = 0

        root.title("Automatic Writing Assistant v" + self.VERSION)
        root.configure(bg=BG)
        root.protocol("WM_DELETE_WINDOW", self._quit)

        fam = tkfont.families()
        self._bf = "Segoe UI" if "Segoe UI" in fam else "Helvetica"
        self._mf = "Consolas" if "Consolas" in fam else "Courier"

        self._setup_styles()
        self._build_ui()
        self._bind_shortcuts()

    # ---------------------------------------------------------------
    # Styles
    # ---------------------------------------------------------------
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        bf = self._bf

        # Frames
        s.configure("BG.TFrame", background=BG)
        s.configure("BG2.TFrame", background=BG2)
        s.configure("Card.TFrame", background=CARD)
        s.configure("Card2.TFrame", background=CARD2)

        # Labels
        for name, bg, fg, font_spec in [
            ("Title.TLabel",    BG,   FG,  (bf, 18, "bold")),
            ("Sub.TLabel",      BG,   FG3, (bf, 10)),
            ("Head.TLabel",     CARD, FG,  (bf, 11, "bold")),
            ("HeadBG.TLabel",   BG,   FG,  (bf, 11, "bold")),
            ("Body.TLabel",     CARD, FG2, (bf, 10)),
            ("BodyBG.TLabel",   BG,   FG2, (bf, 10)),
            ("Val.TLabel",      CARD, ACCENT2, (bf, 11, "bold")),
            ("ValG.TLabel",     CARD, GREEN, (bf, 11, "bold")),
            ("ValO.TLabel",     CARD, ORANGE, (bf, 11, "bold")),
            ("Stat.TLabel",     CARD, FG2, (bf, 10)),
            ("StatBG.TLabel",   BG,   FG2, (bf, 10)),
            ("Foot.TLabel",     BG,   FG3, (bf, 9)),
            ("Cnt.TLabel",      CARD, FG3, (bf, 9)),
            ("Big.TLabel",      CARD, FG,  (bf, 26, "bold")),
            ("BigSub.TLabel",   CARD, FG2, (bf, 11)),
            ("StatNum.TLabel",  CARD, ACCENT, (bf, 16, "bold")),
            ("StatUnit.TLabel", CARD, FG3, (bf, 9)),
        ]:
            s.configure(name, background=bg, foreground=fg, font=font_spec)

        # Buttons
        for name, bg_c, fg_c, act_bg in [
            ("Accent.TButton", ACCENT, "#fff", ACCENT2),
            ("Green.TButton",  GREEN,  "#fff", GREEN2),
            ("Red.TButton",    RED,    "#fff", RED2),
            ("Orange.TButton", ORANGE, "#fff", "#ffaa66"),
            ("Card.TButton",   CARD2,  FG2,   "#252550"),
            ("Cyan.TButton",   CYAN,   "#fff", "#33dde8"),
        ]:
            s.configure(name, background=bg_c, foreground=fg_c,
                        font=(bf, 10, "bold"), padding=(14, 8), borderwidth=0)
            s.map(name, background=[("active", act_bg), ("disabled", "#2a2a44")],
                  foreground=[("disabled", "#555")])

        # Scales
        s.configure("A.Horizontal.TScale", background=CARD, troughcolor=BORDER,
                    sliderthickness=16, borderwidth=0)
        s.configure("G.Horizontal.TScale", background=CARD, troughcolor=BORDER,
                    sliderthickness=16, borderwidth=0)
        s.configure("O.Horizontal.TScale", background=CARD, troughcolor=BORDER,
                    sliderthickness=16, borderwidth=0)

        # Progressbar
        s.configure("pointed.Horizontal.TProgressbar",
                    troughcolor=BORDER, background=ACCENT, thickness=6)
        s.configure("green.Horizontal.TProgressbar",
                    troughcolor=BORDER, background=GREEN, thickness=6)

        # Combobox
        s.configure("Dark.TCombobox", fieldbackground=INP_BG, background=CARD,
                    foreground=FG, selectbackground=SEL_BG,
                    font=(bf, 10), padding=6)
        s.map("Dark.TCombobox",
              fieldbackground=[("readonly", INP_BG)],
              foreground=[("readonly", FG)])

        # Checkbutton
        s.configure("Dark.TCheckbutton", background=CARD, foreground=FG2,
                    font=(bf, 10), indicatorcolor=BORDER)
        s.map("Dark.TCheckbutton",
              background=[("active", CARD)],
              indicatorcolor=[("selected", ACCENT)])

        # Radiobutton
        s.configure("Dark.TRadiobutton", background=CARD, foreground=FG2,
                    font=(bf, 10), indicatorcolor=BORDER)
        s.map("Dark.TRadiobutton",
              background=[("active", CARD)],
              indicatorcolor=[("selected", ACCENT)])

        # Notebook
        s.configure("Dark.TNotebook", background=BG, borderwidth=0)
        s.configure("Dark.TNotebook.Tab", background=CARD, foreground=FG3,
                    font=(bf, 10, "bold"), padding=(16, 8))
        s.map("Dark.TNotebook.Tab",
              background=[("selected", TAB_SEL)],
              foreground=[("selected", FG)])

        # Separator
        s.configure("Dark.TSeparator", background=BORDER)

        # LabelFrame
        s.configure("Card.TLabelframe", background=CARD, foreground=FG,
                    font=(bf, 10, "bold"), borderwidth=1, relief="solid")
        s.configure("Card.TLabelframe.Label", background=CARD, foreground=FG,
                    font=(bf, 10, "bold"))

    # ---------------------------------------------------------------
    # Build UI
    # ---------------------------------------------------------------
    def _build_ui(self):
        # Top bar
        top = ttk.Frame(self.root, style="BG.TFrame")
        top.pack(fill="x", padx=20, pady=(14, 0))

        ttk.Label(top, text="Automatic Writing Assistant",
                  style="Title.TLabel").pack(side="left")

        ver_lbl = ttk.Label(top, text="v" + self.VERSION,
                            style="Sub.TLabel")
        ver_lbl.pack(side="left", padx=(8, 0), pady=(6, 0))

        # Always on top toggle
        self._aot_var = tk.BooleanVar(value=False)
        aot_btn = ttk.Checkbutton(top, text="Pin on Top",
                                  variable=self._aot_var,
                                  style="Dark.TCheckbutton",
                                  command=self._toggle_aot)
        aot_btn.pack(side="right")

        # Help button
        help_btn = ttk.Button(top, text="? How to Use", style="Cyan.TButton",
                              command=self._show_help_tab)
        help_btn.pack(side="right", padx=(0, 10))

        # Notebook with tabs
        self.nb = ttk.Notebook(self.root, style="Dark.TNotebook")
        self.nb.pack(fill="both", expand=True, padx=20, pady=(10, 0))

        # Tab 1: Editor
        self._tab_editor = ttk.Frame(self.nb, style="BG.TFrame")
        self.nb.add(self._tab_editor, text="  Editor  ")

        # Tab 2: Settings
        self._tab_settings = ttk.Frame(self.nb, style="BG.TFrame")
        self.nb.add(self._tab_settings, text="  Settings  ")

        # Tab 3: Log
        self._tab_log = ttk.Frame(self.nb, style="BG.TFrame")
        self.nb.add(self._tab_log, text="  Live Log  ")

        # Tab 4: How to Use
        self._tab_help = ttk.Frame(self.nb, style="BG.TFrame")
        self.nb.add(self._tab_help, text="  How to Use  ")

        self._build_editor_tab()
        self._build_settings_tab()
        self._build_log_tab()
        self._build_help_tab()

        # Bottom bar (always visible)
        self._build_bottom_bar()

    # ---------------------------------------------------------------
    # Editor Tab
    # ---------------------------------------------------------------
    def _build_editor_tab(self):
        tab = self._tab_editor

        # Toolbar row
        toolbar = ttk.Frame(tab, style="BG2.TFrame")
        toolbar.pack(fill="x", padx=8, pady=(10, 0))

        ttk.Button(toolbar, text="Open File", style="Card.TButton",
                   command=self._open_file).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Save to File", style="Card.TButton",
                   command=self._save_file).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical",
                      style="Dark.TSeparator").pack(side="left", padx=8, fill="y", pady=4)

        ttk.Button(toolbar, text="Paste", style="Card.TButton",
                   command=self._paste_clip).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Clear", style="Card.TButton",
                   command=self._clear_text).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical",
                      style="Dark.TSeparator").pack(side="left", padx=8, fill="y", pady=4)

        # Presets
        ttk.Label(toolbar, text="Presets:", style="StatBG.TLabel").pack(side="left", padx=(4, 4))
        self._preset_var = tk.StringVar()
        self._preset_cb = ttk.Combobox(toolbar, textvariable=self._preset_var,
                                       style="Dark.TCombobox", width=18,
                                       state="readonly")
        self._preset_cb.pack(side="left", padx=2)
        self._refresh_presets()

        ttk.Button(toolbar, text="Load", style="Cyan.TButton",
                   command=self._load_preset).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Save As", style="Green.TButton",
                   command=self._save_preset).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Delete", style="Red.TButton",
                   command=self._del_preset).pack(side="left", padx=2)

        # Text area card
        tcard = tk.Frame(tab, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1)
        tcard.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        # Text header
        thdr = ttk.Frame(tcard, style="Card.TFrame")
        thdr.pack(fill="x", padx=14, pady=(10, 2))

        ttk.Label(thdr, text="Text to Type", style="Head.TLabel").pack(side="left")

        self._char_lbl = ttk.Label(thdr, text="0 chars | 0 words | 0 lines",
                                   style="Cnt.TLabel")
        self._char_lbl.pack(side="right")

        # Text widget
        twrap = tk.Frame(tcard, bg=INP_BG, highlightbackground=BORDER,
                         highlightthickness=1)
        twrap.pack(fill="both", expand=True, padx=14, pady=(2, 10))

        self.textbox = tk.Text(
            twrap, font=(self._mf, 11), bg=INP_BG, fg=FG,
            insertbackground=ACCENT, selectbackground=SEL_BG,
            selectforeground="#ffffff", relief="flat", wrap="word",
            padx=12, pady=10, undo=True, maxundo=50,
            blockcursor=False, spacing1=2, spacing3=2,
        )
        sb = tk.Scrollbar(twrap, command=self.textbox.yview,
                          bg=CARD, troughcolor=INP_BG,
                          highlightthickness=0, bd=0, width=10)
        self.textbox.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.textbox.pack(side="left", fill="both", expand=True)

        self.textbox.bind("<KeyRelease>", self._update_stats)
        self.textbox.bind("<<Paste>>",
                          lambda e: self.root.after(50, self._update_stats))

        # Right-click context menu
        self._ctx_menu = tk.Menu(self.textbox, tearoff=0,
                                 bg=CARD, fg=FG, activebackground=ACCENT,
                                 activeforeground="#fff", bd=0,
                                 font=(self._bf, 10))
        self._ctx_menu.add_command(label="Cut         Ctrl+X",
                                   command=lambda: self.textbox.event_generate("<<Cut>>"))
        self._ctx_menu.add_command(label="Copy        Ctrl+C",
                                   command=lambda: self.textbox.event_generate("<<Copy>>"))
        self._ctx_menu.add_command(label="Paste       Ctrl+V",
                                   command=lambda: self.textbox.event_generate("<<Paste>>"))
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="Select All   Ctrl+A",
                                   command=self._select_all)
        self._ctx_menu.add_command(label="Clear All",
                                   command=self._clear_text)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="Undo        Ctrl+Z",
                                   command=lambda: self.textbox.event_generate("<<Undo>>"))
        self._ctx_menu.add_command(label="Redo        Ctrl+Y",
                                   command=lambda: self.textbox.event_generate("<<Redo>>"))

        self.textbox.bind("<Button-3>", self._show_ctx_menu)

        # Stats row
        stats = tk.Frame(tab, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1)
        stats.pack(fill="x", padx=8, pady=(6, 8))

        si = ttk.Frame(stats, style="Card.TFrame")
        si.pack(fill="x", padx=12, pady=8)

        # Stat boxes
        self._stat_frames = {}
        for key, label in [("chars", "Characters"), ("words", "Words"),
                           ("lines", "Lines"), ("eta", "Est. Time")]:
            box = ttk.Frame(si, style="Card.TFrame")
            box.pack(side="left", fill="x", expand=True)
            val_l = ttk.Label(box, text="0", style="StatNum.TLabel")
            val_l.pack()
            ttk.Label(box, text=label, style="StatUnit.TLabel").pack()
            self._stat_frames[key] = val_l

    # ---------------------------------------------------------------
    # Settings Tab
    # ---------------------------------------------------------------
    def _build_settings_tab(self):
        tab = self._tab_settings

        wrapper = ttk.Frame(tab, style="BG.TFrame")
        wrapper.pack(fill="both", expand=True, padx=8, pady=8)

        # -- Typing Mode --
        mode_card = tk.Frame(wrapper, bg=CARD, highlightbackground=BORDER,
                             highlightthickness=1)
        mode_card.pack(fill="x", pady=(0, 8))
        mi = ttk.Frame(mode_card, style="Card.TFrame")
        mi.pack(fill="x", padx=16, pady=12)

        ttk.Label(mi, text="Typing Mode", style="Head.TLabel").pack(anchor="w")

        mode_row = ttk.Frame(mi, style="Card.TFrame")
        mode_row.pack(anchor="w", pady=(8, 0))

        self._mode_var = tk.StringVar(value="normal")
        for val, txt, desc in [
            ("normal", "Constant", "Fixed delay between each character"),
            ("human",  "Human-like", "Random delays, longer pauses at punctuation"),
            ("burst",  "Burst", "Type in fast bursts with small pauses between"),
        ]:
            rf = ttk.Frame(mode_row, style="Card.TFrame")
            rf.pack(side="left", padx=(0, 20))
            ttk.Radiobutton(rf, text=txt, value=val,
                            variable=self._mode_var,
                            style="Dark.TRadiobutton").pack(anchor="w")
            ttk.Label(rf, text=desc, style="Cnt.TLabel").pack(anchor="w")

        # -- Speed & Timing --
        speed_card = tk.Frame(wrapper, bg=CARD, highlightbackground=BORDER,
                              highlightthickness=1)
        speed_card.pack(fill="x", pady=(0, 8))
        spi = ttk.Frame(speed_card, style="Card.TFrame")
        spi.pack(fill="x", padx=16, pady=12)

        ttk.Label(spi, text="Speed & Timing", style="Head.TLabel").pack(anchor="w")

        grid = ttk.Frame(spi, style="Card.TFrame")
        grid.pack(fill="x", pady=(8, 0))

        # Countdown
        c1 = ttk.Frame(grid, style="Card.TFrame")
        c1.pack(side="left", fill="x", expand=True)
        ttk.Label(c1, text="Countdown", style="Body.TLabel").pack(anchor="w")
        cr = ttk.Frame(c1, style="Card.TFrame")
        cr.pack(anchor="w", pady=(4, 0))
        self._cd_var = tk.IntVar(value=5)
        ttk.Scale(cr, from_=1, to=30, variable=self._cd_var,
                  orient="horizontal", length=180,
                  style="A.Horizontal.TScale",
                  command=self._on_cd).pack(side="left")
        self._cd_lbl = ttk.Label(cr, text="5 s", style="Val.TLabel", width=5)
        self._cd_lbl.pack(side="left", padx=(8, 0))

        # Base speed
        c2 = ttk.Frame(grid, style="Card.TFrame")
        c2.pack(side="left", fill="x", expand=True)
        ttk.Label(c2, text="Typing Delay (ms/char)", style="Body.TLabel").pack(anchor="w")
        sr = ttk.Frame(c2, style="Card.TFrame")
        sr.pack(anchor="w", pady=(4, 0))
        self._sp_var = tk.IntVar(value=30)
        ttk.Scale(sr, from_=5, to=300, variable=self._sp_var,
                  orient="horizontal", length=180,
                  style="G.Horizontal.TScale",
                  command=self._on_sp).pack(side="left")
        self._sp_lbl = ttk.Label(sr, text="30 ms", style="ValG.TLabel", width=7)
        self._sp_lbl.pack(side="left", padx=(8, 0))

        # Randomness (for human mode)
        c3 = ttk.Frame(grid, style="Card.TFrame")
        c3.pack(side="left", fill="x", expand=True)
        ttk.Label(c3, text="Randomness %", style="Body.TLabel").pack(anchor="w")
        rr = ttk.Frame(c3, style="Card.TFrame")
        rr.pack(anchor="w", pady=(4, 0))
        self._rand_var = tk.IntVar(value=40)
        ttk.Scale(rr, from_=0, to=100, variable=self._rand_var,
                  orient="horizontal", length=180,
                  style="O.Horizontal.TScale",
                  command=self._on_rand).pack(side="left")
        self._rand_lbl = ttk.Label(rr, text="40%", style="ValO.TLabel", width=5)
        self._rand_lbl.pack(side="left", padx=(8, 0))

        # -- Options --
        opt_card = tk.Frame(wrapper, bg=CARD, highlightbackground=BORDER,
                            highlightthickness=1)
        opt_card.pack(fill="x", pady=(0, 8))
        oi = ttk.Frame(opt_card, style="Card.TFrame")
        oi.pack(fill="x", padx=16, pady=12)

        ttk.Label(oi, text="Options", style="Head.TLabel").pack(anchor="w")

        orow = ttk.Frame(oi, style="Card.TFrame")
        orow.pack(fill="x", pady=(8, 0))

        # Column 1
        oc1 = ttk.Frame(orow, style="Card.TFrame")
        oc1.pack(side="left", fill="x", expand=True)

        self._minimize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc1, text="Auto-minimize before typing",
                        variable=self._minimize_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=2)

        self._notify_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc1, text="Sound notification on complete",
                        variable=self._notify_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=2)

        self._skip_nl_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(oc1, text="Skip empty lines",
                        variable=self._skip_nl_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=2)

        self._shift_enter_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc1, text="Shift+Enter for newlines (chat apps)",
                        variable=self._shift_enter_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=2)

        # Column 2
        oc2 = ttk.Frame(orow, style="Card.TFrame")
        oc2.pack(side="left", fill="x", expand=True)

        self._trim_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(oc2, text="Trim trailing spaces per line",
                        variable=self._trim_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=2)

        self._restore_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(oc2, text="Restore window after typing",
                        variable=self._restore_var,
                        style="Dark.TCheckbutton").pack(anchor="w", pady=2)

        # Repeat
        rep_row = ttk.Frame(oi, style="Card.TFrame")
        rep_row.pack(anchor="w", pady=(8, 0))

        ttk.Label(rep_row, text="Repeat:", style="Body.TLabel").pack(side="left")
        self._repeat_var = tk.IntVar(value=1)
        rep_spin = tk.Spinbox(rep_row, from_=1, to=99, width=4,
                              textvariable=self._repeat_var,
                              font=(self._bf, 10), bg=INP_BG, fg=FG,
                              buttonbackground=CARD, insertbackground=ACCENT,
                              highlightthickness=1, highlightbackground=BORDER,
                              relief="flat")
        rep_spin.pack(side="left", padx=(8, 4))
        ttk.Label(rep_row, text="time(s)", style="Body.TLabel").pack(side="left")

    # ---------------------------------------------------------------
    # Live Log Tab
    # ---------------------------------------------------------------
    def _build_log_tab(self):
        tab = self._tab_log

        lcard = tk.Frame(tab, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1)
        lcard.pack(fill="both", expand=True, padx=8, pady=8)

        lhdr = ttk.Frame(lcard, style="Card.TFrame")
        lhdr.pack(fill="x", padx=12, pady=(10, 4))
        ttk.Label(lhdr, text="Typing Log", style="Head.TLabel").pack(side="left")
        ttk.Button(lhdr, text="Clear Log", style="Card.TButton",
                   command=self._clear_log).pack(side="right")

        lwrap = tk.Frame(lcard, bg=INP_BG, highlightbackground=BORDER,
                         highlightthickness=1)
        lwrap.pack(fill="both", expand=True, padx=12, pady=(2, 10))

        self._log = tk.Text(
            lwrap, font=(self._mf, 10), bg=INP_BG, fg=FG2,
            insertbackground=ACCENT, relief="flat", wrap="word",
            padx=10, pady=8, state="disabled",
        )
        lsb = tk.Scrollbar(lwrap, command=self._log.yview,
                           bg=CARD, troughcolor=INP_BG,
                           highlightthickness=0, bd=0, width=10)
        self._log.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)

        # Color tags
        self._log.tag_configure("info", foreground=FG2)
        self._log.tag_configure("success", foreground=GREEN)
        self._log.tag_configure("warn", foreground=YELLOW)
        self._log.tag_configure("error", foreground=RED)
        self._log.tag_configure("accent", foreground=ACCENT2)
        self._log.tag_configure("dim", foreground=FG3)
        self._log.tag_configure("cyan", foreground=CYAN)

    # ---------------------------------------------------------------
    # How to Use Tab
    # ---------------------------------------------------------------
    def _build_help_tab(self):
        tab = self._tab_help

        # Scrollable frame via canvas
        canvas = tk.Canvas(tab, bg=BG, highlightthickness=0, bd=0)
        vsb = tk.Scrollbar(tab, orient="vertical", command=canvas.yview,
                           bg=CARD, troughcolor=BG, highlightthickness=0,
                           bd=0, width=10)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", pady=8)
        canvas.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        inner = ttk.Frame(canvas, style="BG.TFrame")
        canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_configure(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_configure)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")

        # ---- Content ----
        self._help_section(inner, "Getting Started",
            "Welcome to the Automatic Writing Assistant! This app simulates\n"
            "keyboard typing into any focused input field. It is perfect for\n"
            "filling forms or text fields that block copy-paste.\n\n"
            "Quick start in 3 steps:\n"
            "  1. Paste or type your text in the Editor tab\n"
            "  2. Click 'Start Typing' at the bottom\n"
            "  3. Quickly click into the target field during the countdown")

        self._help_section(inner, "Step 1: Enter Your Text",
            "Go to the Editor tab. You have several ways to add text:\n\n"
            "  - Type directly into the text area\n"
            "  - Click 'Paste' to paste from your clipboard\n"
            "  - Click 'Open File' to load a .txt file (Ctrl+O)\n"
            "  - Load a saved preset from the Presets dropdown\n\n"
            "The stats bar at the bottom shows character count, word count,\n"
            "line count, and the estimated typing time.")

        self._help_section(inner, "Step 2: Configure Settings",
            "Go to the Settings tab to customize behavior:\n\n"
            "  Typing Mode:\n"
            "    - Constant: Fixed delay between each character\n"
            "    - Human-like: Random variation with pauses at punctuation\n"
            "    - Burst: Fast bursts of 3-8 chars with pauses between\n\n"
            "  Countdown: Time (1-30s) before typing starts. Use this\n"
            "  to switch to your target window.\n\n"
            "  Typing Delay: Milliseconds between each character (5-300ms).\n"
            "  Lower = faster typing. 30ms is a good default.\n\n"
            "  Randomness: How much the delay varies in Human-like mode.\n"
            "  Higher = more natural but less predictable speed.")

        self._help_section(inner, "Step 3: Start Typing",
            "Click the 'Start Typing' button at the bottom (or Ctrl+Enter).\n\n"
            "During the countdown:\n"
            "  - Click into the target text field or input box\n"
            "  - Make sure the cursor is where you want text to appear\n"
            "  - The app will auto-minimize if that option is enabled\n\n"
            "The typing will begin after the countdown reaches zero.")

        self._help_section(inner, "Controlling Typing",
            "While typing is in progress:\n\n"
            "  - Pause / Resume: Click the orange Pause button\n"
            "  - Stop: Click the red Stop button\n"
            "  - Emergency Stop: Press F9 on your keyboard (anytime!)\n\n"
            "The progress bar and percentage show how far along you are.\n"
            "Elapsed time and estimated remaining time are displayed.")

        self._help_section(inner, "Saving Presets",
            "If you type the same text often, save it as a preset:\n\n"
            "  1. Type or paste text in the Editor\n"
            "  2. Click 'Save As' in the toolbar\n"
            "  3. Enter a name for the preset\n"
            "  4. Click Save\n\n"
            "To load a preset later, select it from the dropdown and\n"
            "click 'Load'. Presets are saved in a presets.json file\n"
            "next to the app.")

        self._help_section(inner, "Options Explained",
            "  Auto-minimize: The app window minimizes when typing starts\n"
            "  so it doesn't block your target field.\n\n"
            "  Restore window: The app window comes back after typing is done.\n\n"
            "  Sound notification: Plays a beep when typing completes (Windows).\n\n"
            "  Skip empty lines: Ignores blank lines in your text.\n\n"
            "  Trim trailing spaces: Removes extra spaces at end of each line.\n\n"
            "  Repeat: Type the same text multiple times (1-99). Useful for\n"
            "  repetitive data entry.")

        self._help_section(inner, "Keyboard Shortcuts",
            "  Ctrl + O         Open a text file\n"
            "  Ctrl + S         Save text to a file\n"
            "  Ctrl + Enter     Start typing\n"
            "  Ctrl + A         Select all text (in editor)\n"
            "  Ctrl + Z         Undo\n"
            "  Ctrl + Y         Redo\n"
            "  F9               Emergency stop while typing\n"
            "  Right-click      Context menu in editor")

        self._help_section(inner, "Live Log",
            "The Live Log tab shows a real-time, color-coded log of\n"
            "everything happening during a typing session:\n\n"
            "  - Session start, mode, and speed info\n"
            "  - Countdown progress\n"
            "  - Repeat tracking (if repeating)\n"
            "  - Pause, resume, stop events\n"
            "  - Completion messages and errors\n\n"
            "Click 'Clear Log' to reset the log.")

        self._help_section(inner, "Tips & Best Practices",
            "  - Use 5-10 second countdown to give yourself time to switch\n"
            "    to the target window and click the right field.\n\n"
            "  - For long text, use Human-like mode for a natural look.\n\n"
            "  - Use 'Pin on Top' if you need the app visible while working.\n\n"
            "  - Save frequently used text as presets to avoid re-pasting.\n\n"
            "  - Always test with a small sample first before a big paste.\n\n"
            "  - If typing goes into the wrong field, press F9 immediately!\n\n"
            "  - Check the Live Log if something seems wrong - it shows\n"
            "    detailed info about what happened.")

        self._help_section(inner, "Troubleshooting",
            "  Nothing gets typed:\n"
            "    - Make sure the cursor is in a text field before countdown ends\n"
            "    - Some apps may block simulated input (try a different field)\n\n"
            "  Characters are wrong or garbled:\n"
            "    - Check your keyboard layout matches the text language\n"
            "    - Try slower typing speed (higher ms value)\n\n"
            "  App does not open:\n"
            "    - Make sure Python 3.10+ is installed\n"
            "    - On macOS/Linux, install pynput: pip install pynput\n\n"
            "  F9 hotkey does not work:\n"
            "    - On macOS, grant Accessibility permissions\n"
            "    - On Linux, ensure input access is available")

    def _help_section(self, parent, title, body):
        """Create a styled help section card."""
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                        highlightthickness=1)
        card.pack(fill="x", pady=(0, 8), padx=4)

        inner = ttk.Frame(card, style="Card.TFrame")
        inner.pack(fill="x", padx=16, pady=12)

        ttk.Label(inner, text=title, style="Head.TLabel").pack(anchor="w")

        ttk.Separator(inner, orient="horizontal",
                      style="Dark.TSeparator").pack(fill="x", pady=(6, 8))

        body_lbl = tk.Label(inner, text=body, font=(self._bf, 10),
                            fg=FG2, bg=CARD, justify="left", anchor="nw",
                            wraplength=800)
        body_lbl.pack(anchor="w", fill="x")

    def _show_help_tab(self):
        """Switch to the How to Use tab."""
        self.nb.select(self._tab_help)

    # ---------------------------------------------------------------
    # Bottom Bar (always visible)
    # ---------------------------------------------------------------
    def _build_bottom_bar(self):
        bar = tk.Frame(self.root, bg=CARD, highlightbackground=BORDER,
                       highlightthickness=1)
        bar.pack(fill="x", padx=20, pady=(6, 12))

        inner = ttk.Frame(bar, style="Card.TFrame")
        inner.pack(fill="x", padx=12, pady=8)

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
        self._pause_btn.pack(side="left", padx=(6, 0))
        self._pause_btn.state(["disabled"])

        self._stop_btn = ttk.Button(bf, text="  Stop  ",
                                    style="Red.TButton",
                                    command=self._stop)
        self._stop_btn.pack(side="left", padx=(6, 0))

        # Progress
        pf = ttk.Frame(inner, style="Card.TFrame")
        pf.pack(side="right")

        self._pct_lbl = ttk.Label(pf, text="0%", style="Val.TLabel", width=5)
        self._pct_lbl.pack(side="right", padx=(8, 0))

        self._progress = ttk.Progressbar(
            pf, orient="horizontal", length=200, mode="determinate",
            style="pointed.Horizontal.TProgressbar")
        self._progress.pack(side="right")

        self._elapsed_lbl = ttk.Label(pf, text="", style="Stat.TLabel")
        self._elapsed_lbl.pack(side="right", padx=(0, 12))

        # Status row
        srow = ttk.Frame(bar, style="Card.TFrame")
        srow.pack(fill="x", padx=12, pady=(0, 6))

        self._dot = tk.Label(srow, text="*", font=(self._bf, 12),
                             fg=ACCENT2, bg=CARD)
        self._dot.pack(side="left")

        hk_hint = ""
        if self.backend.hotkey_supported:
            hk_hint = "  |  Press " + self.backend.hotkey_label + " to emergency-stop"

        self._status_lbl = ttk.Label(
            srow,
            text="Ready. Paste text, configure settings, click Start." + hk_hint,
            style="Stat.TLabel")
        self._status_lbl.pack(side="left", padx=(4, 0), fill="x", expand=True)

        # Footer
        ttk.Label(
            self.root,
            text="Use responsibly. This tool does not bypass any website or application policies.",
            style="Foot.TLabel",
        ).pack(anchor="w", padx=24, pady=(0, 6))

    # ---------------------------------------------------------------
    # Keyboard Shortcuts
    # ---------------------------------------------------------------
    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self._open_file())
        self.root.bind("<Control-O>", lambda e: self._open_file())
        self.root.bind("<Control-s>", lambda e: self._save_file())
        self.root.bind("<Control-S>", lambda e: self._save_file())
        self.root.bind("<Control-Return>", lambda e: self._start())

    # ---------------------------------------------------------------
    # Callbacks
    # ---------------------------------------------------------------
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

    def _show_ctx_menu(self, event):
        try:
            self._ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._ctx_menu.grab_release()

    def _select_all(self):
        self.textbox.tag_add("sel", "1.0", "end")
        self.textbox.mark_set("insert", "end")

    # ---------------------------------------------------------------
    # Text Stats
    # ---------------------------------------------------------------
    def _update_stats(self, _event=None):
        text = self.textbox.get("1.0", "end-1c")
        chars = len(text)
        words = len(text.split()) if text.strip() else 0
        lines = text.count("\n") + 1 if text else 0

        self._char_lbl.configure(
            text=str(chars) + " chars | " + str(words) + " words | " + str(lines) + " lines"
        )
        self._stat_frames["chars"].configure(text=str(chars))
        self._stat_frames["words"].configure(text=str(words))
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
        if seconds < 60:
            return str(int(seconds)) + "s"
        m = int(seconds) // 60
        s = int(seconds) % 60
        if m < 60:
            return str(m) + "m " + str(s) + "s"
        h = m // 60
        m = m % 60
        return str(h) + "h " + str(m) + "m"

    # ---------------------------------------------------------------
    # File I/O
    # ---------------------------------------------------------------
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Open Text File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.textbox.delete("1.0", "end")
                self.textbox.insert("1.0", content)
                self._update_stats()
                self._log_msg("Opened file: " + os.path.basename(path), "info")
                self._status("Loaded " + os.path.basename(path), CYAN)
            except Exception as e:
                messagebox.showerror("Error", "Could not open file:\n" + str(e))

    def _save_file(self):
        path = filedialog.asksaveasfilename(
            title="Save Text File",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self.textbox.get("1.0", "end-1c"))
                self._log_msg("Saved to: " + os.path.basename(path), "info")
                self._status("Saved " + os.path.basename(path), GREEN)
            except Exception as e:
                messagebox.showerror("Error", "Could not save file:\n" + str(e))

    def _paste_clip(self):
        try:
            clip = self.root.clipboard_get()
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", clip)
            self._update_stats()
            self._status("Pasted from clipboard.", CYAN)
        except tk.TclError:
            self._status("Clipboard is empty.", YELLOW)

    def _clear_text(self):
        self.textbox.delete("1.0", "end")
        self._update_stats()
        self._set_progress(0)
        self._status("Cleared.", FG3)

    # ---------------------------------------------------------------
    # Presets
    # ---------------------------------------------------------------
    def _refresh_presets(self):
        names = self.presets.names()
        self._preset_cb.configure(values=names)
        if names:
            self._preset_var.set(names[0])

    def _load_preset(self):
        name = self._preset_var.get()
        if not name:
            self._status("No preset selected.", YELLOW); return
        text = self.presets.get(name)
        if text:
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", text)
            self._update_stats()
            self._log_msg("Loaded preset: " + name, "info")
            self._status("Loaded preset: " + name, CYAN)

    def _save_preset(self):
        text = self.textbox.get("1.0", "end-1c")
        if not text.strip():
            self._status("Nothing to save.", YELLOW); return

        # Simple preset name dialog
        win = tk.Toplevel(self.root)
        win.title("Save Preset")
        win.geometry("320x130")
        win.configure(bg=CARD)
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(win, text="Preset name:", style="Head.TLabel").pack(pady=(16, 4))
        entry = tk.Entry(win, font=(self._bf, 11), bg=INP_BG, fg=FG,
                         insertbackground=ACCENT, relief="flat",
                         highlightthickness=1, highlightbackground=BORDER)
        entry.pack(padx=20, fill="x")
        entry.focus_set()

        result = [None]
        def on_ok(_=None):
            result[0] = entry.get().strip()
            win.destroy()
        entry.bind("<Return>", on_ok)
        ttk.Button(win, text="Save", style="Green.TButton",
                   command=on_ok).pack(pady=10)
        win.wait_window()
        name = result[0]

        if name:
            self.presets.save(name, text)
            self._refresh_presets()
            self._preset_var.set(name)
            self._log_msg("Saved preset: " + name, "success")
            self._status("Preset saved: " + name, GREEN)

    def _del_preset(self):
        name = self._preset_var.get()
        if not name:
            self._status("No preset selected.", YELLOW); return
        if messagebox.askyesno("Delete Preset",
                               "Delete preset '" + name + "'?"):
            self.presets.delete(name)
            self._refresh_presets()
            self._log_msg("Deleted preset: " + name, "warn")
            self._status("Deleted preset: " + name, RED)

    # ---------------------------------------------------------------
    # Log
    # ---------------------------------------------------------------
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

    # ---------------------------------------------------------------
    # Status / Progress
    # ---------------------------------------------------------------
    def _status(self, msg, color=None):
        self.root.after(0, lambda: self._status_lbl.configure(text=msg))
        if color:
            self.root.after(0, lambda: self._dot.configure(fg=color))

    def _set_progress(self, pct):
        def _do():
            self._progress.configure(value=pct)
            self._pct_lbl.configure(text=str(int(pct)) + "%")
        self.root.after(0, _do)

    def _set_elapsed(self, text):
        self.root.after(0, lambda: self._elapsed_lbl.configure(text=text))

    # ---------------------------------------------------------------
    # Actions
    # ---------------------------------------------------------------
    def _start(self):
        if self._typing:
            self._status("Already typing!", YELLOW); return

        raw = self.textbox.get("1.0", "end").rstrip("\n")
        if not raw.strip():
            self._status("Please enter or paste some text first.", YELLOW)
            self._log_msg("Start aborted: no text.", "warn")
            return

        # Preprocess text
        text = raw
        if self._trim_var.get():
            text = "\n".join(line.rstrip() for line in text.split("\n"))
        if self._skip_nl_var.get():
            lines = text.split("\n")
            text = "\n".join(l for l in lines if l.strip())

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

        # Apply Shift+Enter setting to backend
        self.backend.shift_enter = self._shift_enter_var.get()

        self._log_msg("Starting typing session", "accent")
        self._log_msg("  Mode: " + mode + " | Delay: " + str(delay_ms)
                      + "ms | Randomness: " + str(int(randomness * 100))
                      + "% | Repeat: " + str(repeat) + "x", "dim")
        self._log_msg("  Text length: " + str(len(text)) + " chars", "dim")
        if self.backend.shift_enter:
            self._log_msg("  Newlines: Shift+Enter (chat-app safe)", "dim")
        else:
            self._log_msg("  Newlines: plain Enter", "dim")

        self.worker = threading.Thread(
            target=self._type_job,
            args=(text, countdown, delay_ms / 1000.0, mode, randomness, repeat),
            daemon=True)
        self.worker.start()

    def _stop(self):
        self.stop_event.set()
        if self._paused:
            self.pause_event.set()  # unpause so thread can exit
        self._status("Stop requested...", RED)
        self._log_msg("Stop requested by user.", "error")

    def _pause_resume(self):
        if not self._typing:
            return
        if self._paused:
            self._paused = False
            self.pause_event.set()
            self._pause_btn.configure(text="  Pause  ")
            self._status("Resumed.", GREEN)
            self._log_msg("Resumed.", "success")
        else:
            self._paused = True
            self.pause_event.clear()
            self._pause_btn.configure(text="  Resume  ")
            self._status("Paused. Click Resume to continue.", ORANGE)
            self._log_msg("Paused.", "warn")

    def _quit(self):
        self.stop_event.set()
        if self._paused:
            self.pause_event.set()
        self.backend.shutdown()
        self.root.destroy()

    # ---------------------------------------------------------------
    # Typing Worker Thread
    # ---------------------------------------------------------------
    def _type_job(self, text, countdown, base_delay, mode, randomness, repeat):
        try:
            # Auto-minimize
            if self._minimize_var.get():
                self.root.after(0, self.root.iconify)

            # Countdown
            self._log_msg("Countdown: " + str(countdown) + "s", "warn")
            for sec in range(countdown, 0, -1):
                if self.stop_event.is_set():
                    self._finish("Cancelled before start.", RED)
                    return
                self._status("Starting in " + str(sec) + "s - focus the target!",
                             YELLOW)
                self._set_progress(int((countdown - sec) / countdown * 5))
                time.sleep(1)

            total = len(text) * repeat
            typed = 0
            self._start_time = time.time()

            for rep in range(repeat):
                rep_label = ""
                if repeat > 1:
                    rep_label = " [" + str(rep + 1) + "/" + str(repeat) + "]"
                    self._log_msg("Repeat " + str(rep + 1) + "/" + str(repeat), "cyan")

                hk = ""
                if self.backend.hotkey_supported:
                    hk = "  " + self.backend.hotkey_label + " to stop."
                self._status("Typing..." + rep_label + hk, GREEN)

                for i, ch in enumerate(text):
                    # Check stop
                    if self.stop_event.is_set():
                        self._finish("Stopped after " + str(typed) + " chars.", RED)
                        return
                    if self.backend.stop_requested():
                        self.stop_event.set()
                        self._finish("Stopped by " + self.backend.hotkey_label
                                     + " after " + str(typed) + " chars.", RED)
                        return

                    # Check pause
                    if self._paused:
                        self.pause_event.wait()
                        if self.stop_event.is_set():
                            self._finish("Stopped while paused.", RED)
                            return

                    # Type the character
                    self.backend.type_char(ch)
                    typed += 1

                    # Progress
                    pct = 5 + int(typed / total * 95)
                    self._set_progress(pct)

                    # Elapsed time
                    elapsed = time.time() - self._start_time
                    remaining = (elapsed / typed) * (total - typed) if typed else 0
                    self._set_elapsed(self._fmt_time(elapsed) + " / ~"
                                      + self._fmt_time(remaining) + " left")

                    # Delay calculation
                    delay = self._calc_delay(ch, base_delay, mode, randomness)
                    if delay > 0:
                        time.sleep(delay)

                # Gap between repeats
                if rep < repeat - 1:
                    self._log_msg("Waiting 1s before next repeat...", "dim")
                    time.sleep(1.0)

            elapsed = time.time() - self._start_time
            self._set_progress(100)
            msg = ("Done! " + str(typed) + " characters typed in "
                   + self._fmt_time(elapsed) + ".")
            self._finish(msg, ACCENT2)

            # Sound notification
            if self._notify_var.get() and SYSTEM == "Windows":
                try:
                    import winsound
                    winsound.MessageBeep(winsound.MB_OK)
                except Exception:
                    pass

        except Exception as exc:
            self._finish("Error: " + str(exc), RED)

    def _calc_delay(self, ch, base, mode, rand_pct):
        if mode == "normal":
            return base
        elif mode == "human":
            # Vary delay randomly
            variance = base * rand_pct
            delay = base + random.uniform(-variance, variance)
            # Longer pause at sentence-enders
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
            # Type 3-8 chars fast, then pause
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

    def _finish(self, msg, color):
        self._typing = False
        self._paused = False
        self._status(msg, color)
        self._log_msg(msg, "success" if color == ACCENT2 else "error")
        self.root.after(0, lambda: self._start_btn.state(["!disabled"]))
        self.root.after(0, lambda: self._pause_btn.state(["disabled"]))
        self.root.after(0, lambda: self._pause_btn.configure(text="  Pause  "))
        # Restore window
        if self._restore_var.get():
            self.root.after(0, self.root.deiconify)
            self.root.after(100, self.root.lift)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    root = tk.Tk()
    root.withdraw()

    # Set window icon (cross-platform)
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
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 20, ctypes.byref(ctypes.c_int(1)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

    try:
        backend = _make_backend()
    except Exception as exc:
        messagebox.showerror("Startup Error", str(exc))
        root.destroy()
        return

    App(root, backend)

    root.update_idletasks()
    w, h = 980, 780
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry("{}x{}+{}+{}".format(w, h, (sw - w) // 2, (sh - h) // 2))
    root.minsize(800, 650)

    root.deiconify()
    root.lift()
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
