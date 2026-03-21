import ctypes
import json
import threading
import time
import tkinter as tk
from functools import lru_cache
from ctypes import wintypes
from tkinter import font as tkfont
from tkinter import scrolledtext
from urllib.parse import quote

import pyperclip
import requests

MOD_CONTROL = 0x0002
VK_L = 0x4C
HOTKEY_ID = 1
WM_HOTKEY = 0x0312
WM_QUIT = 0x0012

COPY_DELAY_SEC = 0.06
CLIPBOARD_STABLE_WAIT = 0.03
MAX_TEXT_LENGTH = 120


class TranslatorApp:
    def __init__(self) -> None:
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32
        self.session = requests.Session()

        self.root: tk.Tk | None = None
        self.log_text: scrolledtext.ScrolledText | None = None
        # Tk 变量必须在创建 root 之后绑定，否则报错：Too early to create variable
        self.status_var: tk.StringVar | None = None
        self.enable_var: tk.BooleanVar | None = None
        self.floating_var: tk.BooleanVar | None = None

        self._enabled_lock = threading.Lock()
        self._translate_enabled = True

        self.floating_win: tk.Toplevel | None = None
        self.floating_label: tk.Label | None = None
        self.floating_timer_id: str | None = None
        self.last_floating_msg = ""

        self.hotkey_thread: threading.Thread | None = None
        self.hotkey_thread_id: int | None = None
        self._closing = False

    @staticmethod
    def clean_text(raw: str) -> str:
        text = raw.strip().replace("\r", " ").replace("\n", " ")
        while "  " in text:
            text = text.replace("  ", " ")
        return text

    @lru_cache(maxsize=2048)
    def translate(self, text: str) -> str:
        encoded = quote(text)
        url = (
            "https://translate.googleapis.com/translate_a/single"
            f"?client=gtx&sl=auto&tl=zh-CN&dt=t&q={encoded}"
        )
        resp = self.session.get(url, timeout=2.5)
        resp.raise_for_status()
        data = json.loads(resp.text)
        translated = "".join(part[0] for part in data[0] if part and part[0])
        return translated.strip() if translated.strip() else "(无翻译结果)"

    def get_cursor_pos(self) -> tuple[int, int]:
        point = wintypes.POINT()
        self.user32.GetCursorPos(ctypes.byref(point))
        return point.x, point.y

    def _is_translate_enabled(self) -> bool:
        with self._enabled_lock:
            return self._translate_enabled

    def _on_enable_toggle(self) -> None:
        if self.enable_var is None or self.status_var is None:
            return
        with self._enabled_lock:
            self._translate_enabled = bool(self.enable_var.get())
        if self._translate_enabled:
            self.status_var.set("已开启 — 选中文本后按 Ctrl + L")
        else:
            self.status_var.set("已关闭 — 不会响应 Ctrl + L")

    def _append_log(self, original: str, result: str) -> None:
        if self.log_text is None:
            return
        ts = time.strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{ts}] 原文：{original}\n", "orig")
        self.log_text.insert("end", f"     译文：{result}\n\n", "trans")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _show_floating_near_cursor(self, original: str, translated: str, duration_ms: int = 2200) -> None:
        if self.floating_var is None or not self.floating_var.get():
            return
        if self.root is None:
            return

        msg = f"{original}\n=> {translated}"
        if msg == self.last_floating_msg:
            return
        self.last_floating_msg = msg

        if self.floating_win is None:
            top = tk.Toplevel(self.root)
            top.overrideredirect(True)
            top.attributes("-topmost", True)
            top.attributes("-alpha", 0.94)
            top.configure(bg="#1F2937")
            lbl = tk.Label(
                top,
                text=msg,
                justify="left",
                anchor="w",
                padx=12,
                pady=10,
                bg="#1F2937",
                fg="#F9FAFB",
                wraplength=480,
                font=tkfont.Font(family="Microsoft YaHei UI", size=10),
            )
            lbl.pack(fill="both", expand=True)
            self.floating_win = top
            self.floating_label = lbl
        else:
            assert self.floating_label is not None
            self.floating_label.configure(text=msg)

        assert self.floating_win is not None
        popup_w, popup_h = 500, 110
        cursor_x, cursor_y = self.get_cursor_pos()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = min(max(10, cursor_x + 16), max(10, screen_w - popup_w - 10))
        y = min(max(10, cursor_y + 16), max(10, screen_h - popup_h - 10))
        self.floating_win.geometry(f"{popup_w}x{popup_h}+{x}+{y}")
        self.floating_win.deiconify()
        self.floating_win.lift()

        if self.floating_timer_id is not None:
            self.floating_win.after_cancel(self.floating_timer_id)
        self.floating_timer_id = self.floating_win.after(duration_ms, self.floating_win.withdraw)

    def _ui_show_result(self, original: str, result: str) -> None:
        self._append_log(original, result)
        self._show_floating_near_cursor(original, result)

    def _ui_show_error(self, original: str, err: str) -> None:
        self._append_log(original, err)
        if self.floating_var is not None and self.floating_var.get():
            self._show_floating_near_cursor(original, err, duration_ms=2800)

    def copy_selected_text(self) -> str:
        before = pyperclip.paste()
        self.user32.keybd_event(0x11, 0, 0, 0)
        self.user32.keybd_event(0x43, 0, 0, 0)
        self.user32.keybd_event(0x43, 0, 2, 0)
        self.user32.keybd_event(0x11, 0, 2, 0)
        time.sleep(COPY_DELAY_SEC)
        copied = pyperclip.paste()
        if copied == before:
            time.sleep(CLIPBOARD_STABLE_WAIT)
            copied = pyperclip.paste()
        return self.clean_text(copied)

    def _do_translate_job(self) -> None:
        if not self._is_translate_enabled():
            return
        text = self.copy_selected_text()
        if not text:
            if self.root is not None:
                self.root.after(
                    0,
                    lambda: self._ui_show_error("提示", "未检测到选中文本，请先划词再按 Ctrl + L"),
                )
            return
        if len(text) > MAX_TEXT_LENGTH:
            text = text[:MAX_TEXT_LENGTH] + "..."
        try:
            translated = self.translate(text)
            print(f"[{time.strftime('%H:%M:%S')}] {text} => {translated}")
            if self.root is not None:
                self.root.after(0, lambda t=text, tr=translated: self._ui_show_result(t, tr))
        except Exception as exc:
            if self.root is not None:
                self.root.after(0, lambda t=text, e=exc: self._ui_show_error(t, f"翻译失败: {e}"))

    def hotkey_loop(self) -> None:
        self.hotkey_thread_id = self.kernel32.GetCurrentThreadId()
        if not self.user32.RegisterHotKey(None, HOTKEY_ID, MOD_CONTROL, VK_L):
            if self.root is not None:

                def _show_hotkey_fail() -> None:
                    if self.status_var is not None:
                        self.status_var.set("错误：Ctrl + L 注册失败，可能被占用")

                self.root.after(0, _show_hotkey_fail)
            return

        msg = wintypes.MSG()
        try:
            while not self._closing:
                ret = self.user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if ret == 0 or ret == -1:
                    break
                if msg.message == WM_HOTKEY and msg.wParam == HOTKEY_ID:
                    threading.Thread(target=self._do_translate_job, daemon=True).start()
                self.user32.TranslateMessage(ctypes.byref(msg))
                self.user32.DispatchMessageW(ctypes.byref(msg))
        finally:
            self.user32.UnregisterHotKey(None, HOTKEY_ID)

    def _clear_log(self) -> None:
        if self.log_text is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _build_ui(self) -> None:
        self.root = tk.Tk()
        # 必须在 root 存在后再创建 Variable，否则会 RuntimeError
        self.status_var = tk.StringVar(master=self.root, value="就绪")
        self.enable_var = tk.BooleanVar(master=self.root, value=True)
        self.floating_var = tk.BooleanVar(master=self.root, value=True)

        self.root.title("SnapTranslate")
        self.root.geometry("560x480")
        self.root.minsize(420, 320)
        self.root.configure(bg="#F3F4F6")

        top = tk.Frame(self.root, bg="#F3F4F6", padx=12, pady=10)
        top.pack(fill="x")

        tk.Label(
            top,
            text="SnapTranslate",
            font=tkfont.Font(family="Microsoft YaHei UI", size=14, weight="bold"),
            bg="#F3F4F6",
            fg="#111827",
        ).pack(anchor="w")

        row = tk.Frame(top, bg="#F3F4F6")
        row.pack(fill="x", pady=(8, 4))

        tk.Checkbutton(
            row,
            text="启用划词翻译（Ctrl + L）",
            variable=self.enable_var,
            command=self._on_enable_toggle,
            bg="#F3F4F6",
            fg="#111827",
            activebackground="#F3F4F6",
            font=tkfont.Font(family="Microsoft YaHei UI", size=10),
        ).pack(side="left")

        tk.Checkbutton(
            row,
            text="鼠标旁悬浮提示",
            variable=self.floating_var,
            bg="#F3F4F6",
            fg="#111827",
            activebackground="#F3F4F6",
            font=tkfont.Font(family="Microsoft YaHei UI", size=10),
        ).pack(side="left", padx=(16, 0))

        tk.Button(
            top,
            text="清空记录",
            command=self._clear_log,
            font=tkfont.Font(family="Microsoft YaHei UI", size=9),
        ).pack(anchor="w", pady=(4, 0))

        tk.Label(
            top,
            textvariable=self.status_var,
            bg="#F3F4F6",
            fg="#4B5563",
            font=tkfont.Font(family="Microsoft YaHei UI", size=9),
        ).pack(anchor="w", pady=(6, 0))

        mid = tk.Frame(self.root, bg="#E5E7EB", padx=10, pady=8)
        mid.pack(fill="both", expand=True)

        tk.Label(
            mid,
            text="划词与翻译记录",
            bg="#E5E7EB",
            fg="#374151",
            font=tkfont.Font(family="Microsoft YaHei UI", size=10, weight="bold"),
        ).pack(anchor="w", pady=(0, 6))

        self.log_text = scrolledtext.ScrolledText(
            mid,
            wrap="word",
            state="disabled",
            font=tkfont.Font(family="Microsoft YaHei UI", size=10),
            bg="#FFFFFF",
            fg="#111827",
            insertbackground="#111827",
            relief="flat",
            padx=8,
            pady=8,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_configure("orig", foreground="#1D4ED8")
        self.log_text.tag_configure("trans", foreground="#047857")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self) -> None:
        self._closing = True
        tid = self.hotkey_thread_id
        if tid:
            self.user32.PostThreadMessageW(tid, WM_QUIT, 0, 0)
        if self.hotkey_thread is not None:
            self.hotkey_thread.join(timeout=1.5)
        if self.root is not None:
            self.root.destroy()

    def run(self) -> None:
        self._build_ui()
        assert self.root is not None

        self.hotkey_thread = threading.Thread(target=self.hotkey_loop, daemon=False)
        self.hotkey_thread.start()

        if self.status_var is not None:
            self.status_var.set("已开启 — 选中文本后按 Ctrl + L")
        self.root.mainloop()


def main() -> None:
    app = TranslatorApp()
    app.run()


if __name__ == "__main__":
    main()
