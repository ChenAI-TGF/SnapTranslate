import ctypes
import json
import os
import re
import subprocess
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

# 翻译请求：弱网/跨境链路易超时，略拉长并做有限次重试
TRANSLATE_RETRIES = 3
TRANSLATE_TIMEOUT = (12, 30)  # (连接超时秒数, 读取超时秒数)

HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) SnapTranslate/1.0",
}


@lru_cache(maxsize=2048)
def _translate_google_gtx(text: str) -> str:
    """Google 公开 gtx 接口，质量较好，国内常需可访问 Google 的网络。"""
    encoded = quote(text)
    url = (
        "https://translate.googleapis.com/translate_a/single"
        f"?client=gtx&sl=auto&tl=zh-CN&dt=t&q={encoded}"
    )
    for attempt in range(TRANSLATE_RETRIES):
        try:
            resp = requests.get(url, timeout=TRANSLATE_TIMEOUT, headers=HTTP_HEADERS)
            resp.raise_for_status()
            data = json.loads(resp.text)
            translated = "".join(part[0] for part in data[0] if part and part[0])
            return translated.strip() if translated.strip() else "(无翻译结果)"
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt + 1 < TRANSLATE_RETRIES:
                time.sleep(0.75 * (attempt + 1))
                continue
            raise


def _mymemory_parse(resp: requests.Response) -> str:
    data = resp.json()
    block = data.get("responseData") or {}
    out = (block.get("translatedText") or "").strip()
    if not out:
        return ""
    upper = out.upper()
    if "MYMEMORY WARNING" in upper or ("QUOTA" in upper and "EXCEED" in upper):
        raise RuntimeError(out)
    return out


def _mymemory_langpairs(text: str) -> tuple[str, ...]:
    """Autodetect 对纯英文有时返回原文，含拉丁字母时优先走 en|zh-CN。"""
    if re.search(r"[A-Za-z]", text):
        return ("en|zh-CN", "Autodetect|zh-CN")
    return ("Autodetect|zh-CN", "en|zh-CN")


@lru_cache(maxsize=2048)
def _translate_mymemory(text: str) -> str:
    """
    MyMemory Translated.net 免费接口：无需 API Key，国内多数网络可直连。
    有每日免费额度，超限会在译文里返回提示文案。
    """
    for langpair in _mymemory_langpairs(text):
        for attempt in range(TRANSLATE_RETRIES):
            try:
                resp = requests.get(
                    "https://api.mymemory.translated.net/get",
                    params={"q": text, "langpair": langpair},
                    timeout=TRANSLATE_TIMEOUT,
                    headers=HTTP_HEADERS,
                )
                resp.raise_for_status()
                out = _mymemory_parse(resp)
                if out:
                    return out
                break
            except RuntimeError:
                raise
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
                if attempt + 1 < TRANSLATE_RETRIES:
                    time.sleep(0.75 * (attempt + 1))
                    continue
                break
            except requests.exceptions.HTTPError:
                break
            except (json.JSONDecodeError, KeyError, ValueError):
                break
    return "(无翻译结果)"

# 界面主题（与 vocab_review 一致）
UI_BG = "#eef1f6"
UI_CARD = "#ffffff"
UI_BORDER = "#e2e8f0"
UI_TEXT = "#0f172a"
UI_TEXT_MUTED = "#64748b"
UI_ACCENT = "#4f46e5"
UI_ACCENT_HOVER = "#4338ca"
UI_CHIP = "#e0e7ff"
UI_STATUS_BG = "#f1f5f9"
UI_LOG_BG = "#f8fafc"
UI_FLOAT_BG = "#1e293b"
UI_FLOAT_FG = "#f1f5f9"
UI_FLOAT_MUTED = "#94a3b8"
UI_FLOAT_BTN = "#6366f1"
UI_FLOAT_BTN_H = "#4f46e5"

FONT_FAMILY = "Microsoft YaHei UI"


class TranslatorApp:
    def __init__(self) -> None:
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32

        self.root: tk.Tk | None = None
        self.log_text: scrolledtext.ScrolledText | None = None
        # Tk 变量必须在创建 root 之后绑定，否则报错：Too early to create variable
        self.status_var: tk.StringVar | None = None
        self.enable_var: tk.BooleanVar | None = None
        self.floating_var: tk.BooleanVar | None = None
        self.translate_source_var: tk.StringVar | None = None

        self._enabled_lock = threading.Lock()
        self._translate_enabled = True

        self.floating_win: tk.Toplevel | None = None
        self.floating_label: tk.Label | None = None
        self.floating_save_btn: tk.Button | None = None
        self.floating_timer_id: str | None = None
        self.last_floating_msg = ""
        self._floating_original = ""
        self._floating_translated = ""

        self.hotkey_thread: threading.Thread | None = None
        self.hotkey_thread_id: int | None = None
        self._closing = False

        self._last_lock = threading.Lock()
        self._last_original: str | None = None
        self._last_translated: str | None = None

        self.vocab_path = os.path.join(os.path.dirname(__file__), "vocab.json")

    @staticmethod
    def clean_text(raw: str) -> str:
        text = raw.strip().replace("\r", " ").replace("\n", " ")
        while "  " in text:
            text = text.replace("  ", " ")
        return text

    def translate(self, text: str) -> str:
        src = "google"
        if self.translate_source_var is not None:
            src = self.translate_source_var.get()
        if src == "mymemory":
            return _translate_mymemory(text)
        return _translate_google_gtx(text)

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
            self.status_var.set("已开启 — Ctrl + L 翻译")
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
        self.last_floating_msg = msg
        self._floating_original = original
        self._floating_translated = translated

        if self.floating_win is None:
            top = tk.Toplevel(self.root)
            top.overrideredirect(True)
            top.attributes("-topmost", True)
            top.attributes("-alpha", 0.97)
            top.configure(bg=UI_FLOAT_BG, highlightbackground="#334155", highlightthickness=1)

            body = tk.Frame(top, bg=UI_FLOAT_BG, padx=14, pady=12)
            body.pack(fill="both", expand=True)

            lbl = tk.Label(
                body,
                text=msg,
                justify="left",
                anchor="w",
                padx=0,
                pady=0,
                bg=UI_FLOAT_BG,
                fg=UI_FLOAT_FG,
                wraplength=460,
                font=tkfont.Font(family=FONT_FAMILY, size=10),
            )
            # Label 的 pady 在部分 Tcl 下不支持 (a,b) 元组，间距交给 pack
            lbl.pack(fill="both", expand=True, pady=(0, 10))

            btn = tk.Button(
                body,
                text="收录生词本",
                command=self._on_floating_save_click,
                relief="flat",
                bd=0,
                padx=14,
                pady=6,
                bg=UI_FLOAT_BTN,
                fg="#ffffff",
                activebackground=UI_FLOAT_BTN_H,
                activeforeground="#ffffff",
                font=tkfont.Font(family=FONT_FAMILY, size=9, weight="bold"),
                cursor="hand2",
            )
            btn.pack(anchor="e")
            self.floating_win = top
            self.floating_label = lbl
            self.floating_save_btn = btn
        else:
            assert self.floating_label is not None
            self.floating_label.configure(text=msg)

        assert self.floating_win is not None
        popup_w, popup_h = 500, 140
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
        with self._last_lock:
            self._last_original = original
            self._last_translated = result
        self._append_log(original, result)
        self._show_floating_near_cursor(original, result)

    def _ui_show_error(self, original: str, err: str) -> None:
        self._append_log(original, err)
        if self.floating_var is not None and self.floating_var.get():
            self._show_floating_near_cursor(original, err, duration_ms=2800)

    def _ui_vocab_feedback(self, title: str, msg: str, *, floating: bool = True) -> None:
        self._append_log(title, msg)
        if self.status_var is not None:
            self.status_var.set(msg)
        if floating and self.floating_var is not None and self.floating_var.get():
            self._show_floating_near_cursor(title, msg, duration_ms=2000)

    def _on_floating_save_click(self) -> None:
        word = self.clean_text(self._floating_original)
        meaning = self.clean_text(self._floating_translated)
        threading.Thread(target=self._do_save_vocab_job, args=(word, meaning), daemon=True).start()

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

    @staticmethod
    def _is_likely_english(text: str) -> bool:
        letters = re.findall(r"[A-Za-z]", text)
        return len(letters) >= 2

    def _speak_english_text(self, text: str) -> None:
        # 使用 Windows 内置 System.Speech 朗读，避免新增 Python 依赖。
        escaped = text.replace("'", "''")
        script = (
            "Add-Type -AssemblyName System.Speech; "
            "$s=New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            "$voice=$s.GetInstalledVoices() | Where-Object {$_.VoiceInfo.Culture.Name -like 'en-*'} | Select-Object -First 1; "
            "if($voice){$s.SelectVoice($voice.VoiceInfo.Name)}; "
            f"$s.Speak('{escaped}')"
        )
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=8,
                check=False,
            )
        except Exception:
            pass

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
        if self._is_likely_english(text):
            threading.Thread(target=self._speak_english_text, args=(text,), daemon=True).start()
        try:
            translated = self.translate(text)
            print(f"[{time.strftime('%H:%M:%S')}] {text} => {translated}")
            if self.root is not None:
                self.root.after(0, lambda t=text, tr=translated: self._ui_show_result(t, tr))
        except Exception as exc:
            if self.root is not None:
                self.root.after(0, lambda t=text, e=exc: self._ui_show_error(t, f"翻译失败: {e}"))

    def _load_vocab(self) -> list[dict]:
        try:
            with open(self.vocab_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        except FileNotFoundError:
            return []
        except Exception:
            return []

    def _save_vocab(self, items: list[dict]) -> None:
        tmp = self.vocab_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.vocab_path)

    def _do_save_vocab_job(self, word: str, meaning: str) -> None:
        if not word or not meaning:
            if self.root is not None:
                self.root.after(0, lambda: self._ui_vocab_feedback("生词本", "暂无可记录内容：请先 Ctrl + L 翻译一次", floating=False))
            return

        items = self._load_vocab()
        exists = any(isinstance(it, dict) and it.get("word") == word for it in items)
        if exists:
            if self.root is not None:
                self.root.after(0, lambda w=word: self._ui_vocab_feedback("生词本", f"已存在于生词本：{w}"))
            return

        items.append(
            {
                "word": word,
                "meaning": meaning,
                "example": "",
                "example_zh": "",
                "score": 50.0,
                "reviews": 0,
            }
        )
        try:
            self._save_vocab(items)
            if self.root is not None:
                self.root.after(0, lambda w=word: self._ui_vocab_feedback("生词本", f"已记录到生词本：{w}"))
        except Exception:
            if self.root is not None:
                self.root.after(0, lambda: self._ui_vocab_feedback("生词本", "记录失败：写入生词本出错"))

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
        self.translate_source_var = tk.StringVar(master=self.root, value="mymemory")

        self.root.title("SnapTranslate")
        self.root.geometry("580x500")
        self.root.minsize(440, 340)
        self.root.configure(bg=UI_BG)

        outer = tk.Frame(self.root, bg=UI_BG, padx=18, pady=16)
        outer.pack(fill="both", expand=True)

        hdr = tk.Frame(outer, bg=UI_BG)
        hdr.pack(fill="x", pady=(0, 14))
        tk.Label(
            hdr,
            text="SnapTranslate",
            font=tkfont.Font(family=FONT_FAMILY, size=20, weight="bold"),
            bg=UI_BG,
            fg=UI_TEXT,
        ).pack(anchor="w")
        tk.Label(
            hdr,
            text="选中文字后按 Ctrl + L · 简体中文翻译 · 可朗读英文",
            font=tkfont.Font(family=FONT_FAMILY, size=10),
            bg=UI_BG,
            fg=UI_TEXT_MUTED,
        ).pack(anchor="w", pady=(4, 0))

        ctrl_card = tk.Frame(
            outer,
            bg=UI_CARD,
            highlightbackground=UI_BORDER,
            highlightthickness=1,
            padx=16,
            pady=14,
        )
        ctrl_card.pack(fill="x", pady=(0, 12))

        row = tk.Frame(ctrl_card, bg=UI_CARD)
        row.pack(fill="x")

        chk_kw = dict(
            bg=UI_CARD,
            fg=UI_TEXT,
            activebackground=UI_CARD,
            activeforeground=UI_TEXT,
            selectcolor=UI_CHIP,
            font=tkfont.Font(family=FONT_FAMILY, size=10),
        )
        tk.Checkbutton(
            row,
            text="启用划词翻译（Ctrl + L）",
            variable=self.enable_var,
            command=self._on_enable_toggle,
            **chk_kw,
        ).pack(side="left")

        tk.Checkbutton(
            row,
            text="鼠标旁悬浮提示",
            variable=self.floating_var,
            **chk_kw,
        ).pack(side="left", padx=(20, 0))

        src_wrap = tk.Frame(ctrl_card, bg=UI_CARD)
        src_wrap.pack(fill="x", pady=(14, 0))
        tk.Label(
            src_wrap,
            text="翻译源",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w")
        src_row = tk.Frame(src_wrap, bg=UI_CARD)
        src_row.pack(fill="x", pady=(6, 0))
        rb_kw = dict(
            bg=UI_CARD,
            activebackground=UI_CARD,
            fg=UI_TEXT,
            selectcolor=UI_CHIP,
            font=tkfont.Font(family=FONT_FAMILY, size=10),
        )
        assert self.translate_source_var is not None
        tk.Radiobutton(
            src_row,
            text="MyMemory（免梯 · 免密钥 · 有每日免费限额）",
            variable=self.translate_source_var,
            value="mymemory",
            **rb_kw,
        ).pack(side="left", padx=(0, 18))
        tk.Radiobutton(
            src_row,
            text="Google（质量通常更好，国内多数网络需代理）",
            variable=self.translate_source_var,
            value="google",
            **rb_kw,
        ).pack(side="left")
        tk.Label(
            src_wrap,
            text="备选接口来自 api.mymemory.translated.net，用量大时可能提示配额用尽。",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=8),
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        row2 = tk.Frame(ctrl_card, bg=UI_CARD)
        row2.pack(fill="x", pady=(12, 0))

        tk.Button(
            row2,
            text="清空记录",
            command=self._clear_log,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
            bg=UI_CARD,
            fg=UI_ACCENT,
            activebackground=UI_CHIP,
            activeforeground=UI_ACCENT_HOVER,
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            padx=14,
            pady=6,
            cursor="hand2",
        ).pack(side="left")

        stat_wrap = tk.Frame(ctrl_card, bg=UI_STATUS_BG, padx=12, pady=10)
        stat_wrap.pack(fill="x", pady=(14, 0))
        tk.Label(
            stat_wrap,
            textvariable=self.status_var,
            bg=UI_STATUS_BG,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
            anchor="w",
            justify="left",
        ).pack(fill="x")

        log_card = tk.Frame(
            outer,
            bg=UI_CARD,
            highlightbackground=UI_BORDER,
            highlightthickness=1,
            padx=14,
            pady=14,
        )
        log_card.pack(fill="both", expand=True)

        tk.Label(
            log_card,
            text="翻译记录",
            bg=UI_CARD,
            fg=UI_TEXT,
            font=tkfont.Font(family=FONT_FAMILY, size=11, weight="bold"),
        ).pack(anchor="w", pady=(0, 10))

        self.log_text = scrolledtext.ScrolledText(
            log_card,
            wrap="word",
            state="disabled",
            font=tkfont.Font(family=FONT_FAMILY, size=10),
            bg=UI_LOG_BG,
            fg=UI_TEXT,
            insertbackground=UI_TEXT,
            relief="flat",
            bd=0,
            padx=12,
            pady=12,
            highlightthickness=0,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_configure("orig", foreground=UI_ACCENT)
        self.log_text.tag_configure("trans", foreground="#0d9488")

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
            self.status_var.set("已开启 — Ctrl + L 翻译")
        self.root.mainloop()


def main() -> None:
    app = TranslatorApp()
    app.run()


if __name__ == "__main__":
    main()
