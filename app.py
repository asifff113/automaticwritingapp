import platform
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk


SYSTEM_NAME = platform.system()


if SYSTEM_NAME == "Windows":
    import ctypes

    user32 = ctypes.windll.user32

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002
    KEYEVENTF_UNICODE = 0x0004
    VK_SHIFT = 0x10
    VK_CONTROL = 0x11
    VK_MENU = 0x12
    VK_RETURN = 0x0D
    VK_F9 = 0x78

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", ctypes.c_ushort),
            ("wScan", ctypes.c_ushort),
            ("dwFlags", ctypes.c_ulong),
            ("time", ctypes.c_ulong),
            ("dwExtraInfo", ctypes.c_ulonglong),
        ]

    class INPUT_UNION(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _anonymous_ = ("u",)
        _fields_ = [("type", ctypes.c_ulong), ("u", INPUT_UNION)]


class TypingBackend:
    hotkey_label = "F9"
    hotkey_supported = False

    def type_char(self, ch):
        raise NotImplementedError

    def stop_requested_by_hotkey(self):
        return False

    def shutdown(self):
        return


class WindowsTypingBackend(TypingBackend):
    hotkey_supported = True

    def send_input(self, vk=0, scan=0, flags=0):
        inp = INPUT(type=INPUT_KEYBOARD)
        inp.ki = KEYBDINPUT(wVk=vk, wScan=scan, dwFlags=flags, time=0, dwExtraInfo=0)
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

    def key_down(self, vk):
        self.send_input(vk=vk, flags=0)

    def key_up(self, vk):
        self.send_input(vk=vk, flags=KEYEVENTF_KEYUP)

    def tap_vk(self, vk):
        self.key_down(vk)
        self.key_up(vk)

    def type_unicode_char(self, ch):
        code = ord(ch)
        self.send_input(scan=code, flags=KEYEVENTF_UNICODE)
        self.send_input(scan=code, flags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)

    def type_char(self, ch):
        if ch == "\n":
            self.tap_vk(VK_RETURN)
            return

        mapped = user32.VkKeyScanW(ord(ch))
        if mapped in (-1, 0xFFFF):
            self.type_unicode_char(ch)
            return

        vk = mapped & 0xFF
        shift_state = (mapped >> 8) & 0xFF

        modifiers = []
        if shift_state & 1:
            modifiers.append(VK_SHIFT)
        if shift_state & 2:
            modifiers.append(VK_CONTROL)
        if shift_state & 4:
            modifiers.append(VK_MENU)

        for mod in modifiers:
            self.key_down(mod)
        self.tap_vk(vk)
        for mod in reversed(modifiers):
            self.key_up(mod)

    def stop_requested_by_hotkey(self):
        return bool(user32.GetAsyncKeyState(VK_F9) & 0x8000)


class PynputTypingBackend(TypingBackend):
    hotkey_supported = True

    def __init__(self):
        from pynput import keyboard

        self.keyboard = keyboard
        self.controller = keyboard.Controller()
        self.hotkey_pressed = threading.Event()
        self.listener = None

        try:
            self.listener = keyboard.Listener(on_press=self.on_press)
            self.listener.start()
        except Exception:
            self.hotkey_supported = False

    def on_press(self, key):
        if key == self.keyboard.Key.f9:
            self.hotkey_pressed.set()

    def type_char(self, ch):
        if ch == "\n":
            self.controller.press(self.keyboard.Key.enter)
            self.controller.release(self.keyboard.Key.enter)
            return
        self.controller.type(ch)

    def stop_requested_by_hotkey(self):
        if self.hotkey_pressed.is_set():
            self.hotkey_pressed.clear()
            return True
        return False

    def shutdown(self):
        if self.listener is not None:
            self.listener.stop()


def create_typing_backend():
    if SYSTEM_NAME == "Windows":
        return WindowsTypingBackend()
    if SYSTEM_NAME in ("Darwin", "Linux"):
        try:
            return PynputTypingBackend()
        except Exception as exc:
            raise RuntimeError(
                "pynput is required on macOS/Linux. Install dependencies from requirements.txt."
            ) from exc
    raise RuntimeError(f"Unsupported operating system: {SYSTEM_NAME}")


class AutoWriterApp:
    def __init__(self, root, backend):
        self.root = root
        self.backend = backend
        self.root.title("Automatic Writing Assistant")
        self.root.geometry("760x520")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.stop_event = threading.Event()
        self.worker = None

        if self.backend.hotkey_supported:
            default_status = (
                "Paste text, set options, click Start. "
                f"Press {self.backend.hotkey_label} anytime to stop."
            )
        else:
            default_status = "Paste text, set options, click Start. Use Stop in the app."

        self.start_delay_var = tk.StringVar(value="5")
        self.char_delay_var = tk.StringVar(value="0.03")
        self.status_var = tk.StringVar(value=default_status)

        self.build_ui()

    def build_ui(self):
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(
            container,
            text="Automatic Writing Assistant",
            font=("Segoe UI", 16, "bold"),
        )
        title.pack(anchor="w")

        subtitle = ttk.Label(
            container,
            text="This tool types text into whichever window is focused after the countdown.",
        )
        subtitle.pack(anchor="w", pady=(2, 8))

        self.text_box = tk.Text(container, wrap="word", font=("Consolas", 11), height=18)
        self.text_box.pack(fill=tk.BOTH, expand=True)

        options = ttk.Frame(container)
        options.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(options, text="Start countdown (sec):").pack(side=tk.LEFT)
        ttk.Entry(options, textvariable=self.start_delay_var, width=8).pack(
            side=tk.LEFT, padx=(6, 14)
        )

        ttk.Label(options, text="Delay per character (sec):").pack(side=tk.LEFT)
        ttk.Entry(options, textvariable=self.char_delay_var, width=8).pack(
            side=tk.LEFT, padx=(6, 14)
        )

        buttons = ttk.Frame(container)
        buttons.pack(fill=tk.X, pady=(10, 0))

        self.start_btn = ttk.Button(buttons, text="Start Typing", command=self.start_typing)
        self.start_btn.pack(side=tk.LEFT)

        ttk.Button(buttons, text="Stop", command=self.stop_typing).pack(side=tk.LEFT, padx=8)
        ttk.Button(buttons, text="Clear Text", command=self.clear_text).pack(side=tk.LEFT)

        status = ttk.Label(
            container,
            textvariable=self.status_var,
            foreground="#0b4a6f",
            wraplength=720,
        )
        status.pack(anchor="w", pady=(12, 0))

        caution = ttk.Label(
            container,
            text="Use only where automation is allowed. The app does not bypass website policies.",
            foreground="#7a3e00",
        )
        caution.pack(anchor="w", pady=(4, 0))

    def set_status(self, text):
        self.root.after(0, lambda: self.status_var.set(text))

    def set_start_button_enabled(self, enabled):
        state = tk.NORMAL if enabled else tk.DISABLED
        self.root.after(0, lambda: self.start_btn.config(state=state))

    def clear_text(self):
        self.text_box.delete("1.0", tk.END)
        self.set_status("Text cleared.")

    def start_typing(self):
        if self.worker and self.worker.is_alive():
            self.set_status("Typing is already running.")
            return

        text = self.text_box.get("1.0", tk.END).rstrip("\n")
        if not text.strip():
            self.set_status("Please paste some text first.")
            return

        try:
            start_delay = int(self.start_delay_var.get())
            char_delay = float(self.char_delay_var.get())
        except ValueError:
            self.set_status("Start countdown must be integer, character delay must be decimal.")
            return

        if start_delay < 0 or char_delay < 0:
            self.set_status("Countdown and character delay must be non-negative.")
            return

        self.stop_event.clear()
        self.set_start_button_enabled(False)
        self.worker = threading.Thread(
            target=self.run_typing_job,
            args=(text, start_delay, char_delay),
            daemon=True,
        )
        self.worker.start()

    def stop_typing(self):
        self.stop_event.set()
        self.set_status("Stop requested. Finishing current key event.")

    def on_close(self):
        self.stop_event.set()
        self.backend.shutdown()
        self.root.destroy()

    def run_typing_job(self, text, start_delay, char_delay):
        try:
            for sec in range(start_delay, 0, -1):
                if self.stop_event.is_set():
                    self.set_status("Typing canceled before start.")
                    return
                self.set_status(f"Typing starts in {sec}s. Focus the target input field now.")
                time.sleep(1)

            if self.backend.hotkey_supported:
                self.set_status(
                    f"Typing in progress... Press {self.backend.hotkey_label} to stop."
                )
            else:
                self.set_status("Typing in progress... Use Stop in this app to stop.")
            for ch in text:
                if self.stop_event.is_set():
                    self.set_status("Typing stopped by user.")
                    return

                if self.backend.stop_requested_by_hotkey():
                    self.stop_event.set()
                    self.set_status("Typing stopped by hotkey.")
                    return

                self.backend.type_char(ch)
                if char_delay > 0:
                    time.sleep(char_delay)

            self.set_status("Typing completed.")
        except Exception as exc:
            self.set_status(f"Typing failed: {exc}")
        finally:
            self.set_start_button_enabled(True)


def main():
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    try:
        backend = create_typing_backend()
    except Exception as exc:
        messagebox.showerror("Startup Error", str(exc))
        root.destroy()
        return

    app = AutoWriterApp(root, backend)
    root.mainloop()


if __name__ == "__main__":
    main()
