import ctypes
import json
import os
import re
import shutil
import subprocess
import threading
import time
import tkinter as tk
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
from ctypes import wintypes
from tkinter import font as tkfont
from tkinter import scrolledtext
from urllib.parse import quote

import pyperclip
import requests
try:
    from PIL import ImageGrab
except Exception:  # pragma: no cover
    ImageGrab = None

try:
    import pytesseract
except Exception:  # pragma: no cover
    pytesseract = None

MOD_CONTROL = 0x0002
VK_L = 0x4C
VK_Q = 0x51
VK_W = 0x57
VK_E = 0x45
VK_CTRL = 0x11
VK_TAB = 0x09
VK_SHIFT = 0x10
VK_ALT = 0x12
HOTKEY_ID = 1
WM_HOTKEY = 0x0312
WM_QUIT = 0x0012

COPY_DELAY_SEC = 0.06
CLIPBOARD_STABLE_WAIT = 0.03
MAX_TEXT_LENGTH = 120

# 翻译请求：连接阶段过长会让界面“卡住”；失败后会自动尝试 MyMemory
TRANSLATE_RETRIES = 2
TRANSLATE_TIMEOUT = (6, 22)  # (连接超时秒数, 读取超时秒数)

# Lingva：开源 Google 翻译前端；多实例并发，谁先可用用谁（部分实例在国内可能 403）
LINGVA_BASES: tuple[str, ...] = (
    "https://lingva.ml",
    "https://translate.plausibility.cloud",
)

# 仅缓存「确有译文」的结果，避免把“(无翻译结果)”或偶发空结果钉死
_TRANS_OK_CACHE: OrderedDict[tuple[str, str], str] = OrderedDict()
_TRANS_OK_CACHE_MAX = 2048


def _cache_translate_get(engine: str, text: str) -> str | None:
    key = (engine, text)
    if key not in _TRANS_OK_CACHE:
        return None
    _TRANS_OK_CACHE.move_to_end(key)
    return _TRANS_OK_CACHE[key]


def _cache_translate_put(engine: str, text: str, result: str) -> None:
    if not result or result == "(无翻译结果)":
        return
    key = (engine, text)
    _TRANS_OK_CACHE[key] = result
    _TRANS_OK_CACHE.move_to_end(key)
    while len(_TRANS_OK_CACHE) > _TRANS_OK_CACHE_MAX:
        _TRANS_OK_CACHE.popitem(last=False)
OCR_LANG = "eng+chi_sim"
# OCR 内存保护：截图区域过大时先降采样，降低峰值内存与 Tesseract 压力
OCR_MAX_PIXELS = 2_400_000
# 截图遮罩：过低 alpha 在 Win32 上易被合成器“吃没”；映射后再设一次更稳定
SNIP_OVERLAY_ALPHA = 0.5
SNIP_OVERLAY_CANVAS_BG = "#0f172a"
TESSERACT_CANDIDATE_DIRS = (
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
)
# 若此路径存在，则强制优先使用，避免被 PATH 中旧版本（如 E:\tes）劫持
FORCE_TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
DEFAULT_SETTINGS_FILE = "main_settings.json"
BACKUP_DIR_NAME = "backups"
DEFAULT_HOTKEYS: dict[str, str] = {
    "translate": "ctrl+l",
    "snip": "tab+q",
    "save_last": "tab+e",
}

HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) SnapTranslate/1.0",
}


def _translate_google_gtx(text: str) -> str:
    """Google 公开 gtx 接口，质量较好，国内常需可访问 Google 的网络。"""
    hit = _cache_translate_get("google", text)
    if hit is not None:
        return hit
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
            out = translated.strip() if translated.strip() else "(无翻译结果)"
            if out != "(无翻译结果)":
                _cache_translate_put("google", text, out)
            return out
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


def _translate_mymemory(text: str) -> str:
    """
    MyMemory Translated.net 免费接口：无需 API Key，国内多数网络可直连。
    有每日免费额度，超限会在译文里返回提示文案。
    """
    hit = _cache_translate_get("mymemory", text)
    if hit is not None:
        return hit
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
                    _cache_translate_put("mymemory", text, out)
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


def _parse_google_clients5_payload(data: object) -> str:
    """解析 clients5.google.com translate_a/t 返回的 JSON（多为 [[译文, 源语言], ...]）。"""
    if isinstance(data, str):
        return data.strip()
    if not isinstance(data, list):
        return ""
    parts: list[str] = []
    for row in data:
        if isinstance(row, (list, tuple)) and row:
            cell = row[0]
            if isinstance(cell, str) and cell:
                parts.append(cell)
        elif isinstance(row, str) and row:
            parts.append(row)
    return "".join(parts).strip()


def _translate_google_clients5(text: str) -> str:
    """Google 另一公开入口（clients5），与 translate.googleapis.com 不同主机，便于并发择优。"""
    hit = _cache_translate_get("google_c5", text)
    if hit is not None:
        return hit
    encoded = quote(text, safe="")
    url = (
        "https://clients5.google.com/translate_a/t"
        f"?client=dict-chrome-ex&sl=auto&tl=zh-CN&q={encoded}"
    )
    for attempt in range(TRANSLATE_RETRIES):
        try:
            resp = requests.get(url, timeout=TRANSLATE_TIMEOUT, headers=HTTP_HEADERS)
            resp.raise_for_status()
            data = json.loads(resp.text)
            translated = _parse_google_clients5_payload(data)
            out = translated if translated else "(无翻译结果)"
            if out != "(无翻译结果)":
                _cache_translate_put("google_c5", text, out)
            return out
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt + 1 < TRANSLATE_RETRIES:
                time.sleep(0.75 * (attempt + 1))
                continue
            raise
        except (json.JSONDecodeError, requests.exceptions.HTTPError):
            return "(无翻译结果)"


def _translate_lingva_mirror(text: str, base: str) -> str:
    """Lingva GET /api/v1/auto/zh/{query}，与 google 直连不同，部分网络下仍可访问。"""
    shared = _cache_translate_get("lingva", text)
    if shared is not None:
        return shared
    seg = quote(text, safe="")
    if len(seg) > 1400:
        return "(无翻译结果)"
    url = f"{base.rstrip('/')}/api/v1/auto/zh/{seg}"
    for attempt in range(TRANSLATE_RETRIES):
        try:
            resp = requests.get(url, timeout=TRANSLATE_TIMEOUT, headers=HTTP_HEADERS)
            resp.raise_for_status()
            data = resp.json()
            if not isinstance(data, dict):
                return "(无翻译结果)"
            out = (data.get("translation") or "").strip()
            if not out:
                return "(无翻译结果)"
            _cache_translate_put("lingva", text, out)
            return out
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt + 1 < TRANSLATE_RETRIES:
                time.sleep(0.75 * (attempt + 1))
                continue
            return "(无翻译结果)"
        except (json.JSONDecodeError, requests.exceptions.RequestException, ValueError, TypeError):
            return "(无翻译结果)"
    return "(无翻译结果)"


def _translate_parallel_race_zh(text: str) -> tuple[str, str | None]:
    """
    同时请求多条免费线路，谁先返回有效中文译文即用谁（其余请求在后台自然结束）。
    命中任一引擎的成功缓存时直接返回，不发起网络请求。
    """
    for key in ("google", "google_c5", "mymemory", "lingva"):
        hit = _cache_translate_get(key, text)
        if hit is not None:
            return hit, None

    max_workers = min(32, 3 + len(LINGVA_BASES))
    executor = ThreadPoolExecutor(max_workers=max_workers)
    future_to_label: dict[object, str] = {}
    try:
        fg = executor.submit(_translate_google_gtx, text)
        future_to_label[fg] = "Google"
        fc = executor.submit(_translate_google_clients5, text)
        future_to_label[fc] = "Google（clients5）"
        fm = executor.submit(_translate_mymemory, text)
        future_to_label[fm] = "MyMemory"
        for base in LINGVA_BASES:
            host = base.split("//", 1)[-1].split("/", 1)[0]
            fl = executor.submit(_translate_lingva_mirror, text, base)
            future_to_label[fl] = f"Lingva（{host}）"

        errors: list[BaseException] = []
        for fut in as_completed(future_to_label):
            exn = fut.exception()
            if exn is not None:
                errors.append(exn)
                continue
            out = fut.result()
            if out and out != "(无翻译结果)":
                label = future_to_label[fut]
                return out, f"（{label} 最快返回）"
        if errors:
            for err in errors:
                if isinstance(err, RuntimeError):
                    raise err
            raise errors[0]
        return "(无翻译结果)", None
    finally:
        executor.shutdown(wait=False, cancel_futures=True)


def _format_translate_failure(exc: BaseException) -> str:
    raw = str(exc)
    low = raw.lower()
    if (
        "failed to resolve" in low
        or "getaddrinfo failed" in low
        or "nameresolutionerror" in low
        or "name resolution" in low
        or "name or service not known" in low
    ):
        return (
            "部分翻译线路 DNS 解析失败（常见于国内网络）。可改用「仅 MyMemory」以省并发；"
            "若使用自动竞速，请检查代理/VPN 是否对 Python 生效。"
        )
    if isinstance(exc, requests.exceptions.RequestException):
        if len(raw) > 260:
            return f"翻译请求失败（{type(exc).__name__}）：{raw[:260]}…"
        return f"翻译请求失败：{raw}"
    return raw if len(raw) <= 400 else f"{raw[:400]}…"


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
UI_CURSOR_STATUS_BG = "#0f172a"
UI_CURSOR_STATUS_FG = "#e2e8f0"

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
        self.tts_volume_var: tk.IntVar | None = None
        self.hotkey_translate_var: tk.StringVar | None = None
        self.hotkey_snip_var: tk.StringVar | None = None
        self.hotkey_save_var: tk.StringVar | None = None
        self.hotkey_hint_var: tk.StringVar | None = None
        self.recent_vars: list[tk.StringVar] = []
        self.recent_items: list[tuple[str, str]] = []
        self.recent_saved_vars: list[tk.StringVar] = []
        self.recent_saved_words: list[str] = []

        self._enabled_lock = threading.Lock()
        self._translate_enabled = True
        self._ocr_lock = threading.Lock()
        self._ocr_running = False
        self._ocr_langs_cache: set[str] | None = None

        self.floating_win: tk.Toplevel | None = None
        self.floating_label: tk.Label | None = None
        self.floating_save_btn: tk.Button | None = None
        self.floating_timer_id: str | None = None
        self.cursor_status_win: tk.Toplevel | None = None
        self.cursor_status_label: tk.Label | None = None
        self.cursor_status_timer_id: str | None = None
        self.cursor_status_follow_id: str | None = None
        self.last_floating_msg = ""
        self._floating_original = ""
        self._floating_translated = ""

        self.hotkey_thread: threading.Thread | None = None
        self.hotkey_thread_id: int | None = None
        self.tab_q_thread: threading.Thread | None = None
        self._closing = False

        self._last_lock = threading.Lock()
        self._last_original: str | None = None
        self._last_translated: str | None = None

        self.vocab_path = os.path.join(os.path.dirname(__file__), "vocab.json")
        self.settings_path = os.path.join(os.path.dirname(__file__), DEFAULT_SETTINGS_FILE)
        self.backup_dir = os.path.join(os.path.dirname(__file__), BACKUP_DIR_NAME)
        self.hotkeys = self._load_hotkeys()
        self._tts_volume_default = self._load_tts_volume()
        self._snip_overlay: tk.Toplevel | None = None
        self._snip_canvas: tk.Canvas | None = None
        self._snip_start: tuple[int, int] | None = None
        self._snip_rect_id: int | None = None
        self._snip_info_id: int | None = None
        self._snip_busy = False

    @staticmethod
    def clean_text(raw: str) -> str:
        text = raw.strip().replace("\r", " ").replace("\n", " ")
        while "  " in text:
            text = text.replace("  ", " ")
        return text

    def _load_settings(self) -> dict:
        try:
            with open(self.settings_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
        return {}

    def _save_settings(self, patch: dict) -> None:
        data = self._load_settings()
        if not isinstance(data, dict):
            data = {}
        data.update(patch)
        tmp = self.settings_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.settings_path)

    def _load_tts_volume(self) -> int:
        data = self._load_settings()
        try:
            v = int(data.get("tts_volume", 100))
        except Exception:
            v = 100
        return max(0, min(100, v))

    def _save_tts_volume(self, volume: int) -> None:
        v = max(0, min(100, int(volume)))
        try:
            self._save_settings({"tts_volume": v})
        except Exception:
            pass

    def _on_tts_volume_change(self, *_: object) -> None:
        if self.tts_volume_var is None:
            return
        try:
            self._save_tts_volume(int(self.tts_volume_var.get()))
        except (tk.TclError, TypeError, ValueError):
            return

    def _normalize_hotkey(self, combo: str) -> str:
        return (combo or "").strip().lower().replace(" ", "")

    def _load_hotkeys(self) -> dict[str, str]:
        data = self._load_settings()
        got = data.get("hotkeys")
        if not isinstance(got, dict):
            return dict(DEFAULT_HOTKEYS)
        out = dict(DEFAULT_HOTKEYS)
        for k in out:
            v = got.get(k)
            if isinstance(v, str) and self._parse_hotkey(v) is not None:
                out[k] = self._normalize_hotkey(v)
        return out

    def _save_hotkeys(self) -> None:
        try:
            self._save_settings({"hotkeys": self.hotkeys})
        except Exception:
            pass

    def _parse_hotkey(self, combo: str) -> tuple[str, str] | None:
        s = self._normalize_hotkey(combo)
        if "+" not in s:
            return None
        mod, key = s.split("+", 1)
        if mod not in {"ctrl", "tab", "shift", "alt"}:
            return None
        if not key:
            return None
        if self._vk_from_key_token(key) is None:
            return None
        return mod, key

    def _vk_from_key_token(self, key: str) -> int | None:
        k = key.strip().upper()
        if len(k) == 1 and "A" <= k <= "Z":
            return ord(k)
        if len(k) == 1 and "0" <= k <= "9":
            return ord(k)
        if k.startswith("F") and k[1:].isdigit():
            n = int(k[1:])
            if 1 <= n <= 12:
                return 0x70 + n - 1
        named = {"TAB": VK_TAB, "SPACE": 0x20}
        return named.get(k)

    def _is_down(self, vk: int) -> bool:
        return bool(self.user32.GetAsyncKeyState(vk) & 0x8000)

    def _is_hotkey_pressed(self, combo: str) -> bool:
        parsed = self._parse_hotkey(combo)
        if parsed is None:
            return False
        mod, key = parsed
        key_vk = self._vk_from_key_token(key)
        if key_vk is None:
            return False
        mod_vk = {"ctrl": VK_CTRL, "tab": VK_TAB, "shift": VK_SHIFT, "alt": VK_ALT}[mod]
        return self._is_down(mod_vk) and self._is_down(key_vk)

    def _tab_mod_key_fire_edge(self, combo: str, prev_tab: bool, prev_key: bool) -> tuple[bool, bool, bool]:
        """
        Tab+X 组合：不要求同一轮询里两键同时按下。
        支持先 Tab 后按 X、先 X 后 Tab、或同时按下，避免「要掐得很准才触发」。
        返回 (本周期是否应触发一次, 新的 prev_tab, 新的 prev_key)。
        """
        parsed = self._parse_hotkey(combo)
        if parsed is None or parsed[0] != "tab":
            return False, False, False
        key_vk = self._vk_from_key_token(parsed[1])
        if key_vk is None:
            return False, False, False
        t = self._is_down(VK_TAB)
        k = self._is_down(key_vk)
        fire = (
            (k and not prev_key and t)
            or (t and not prev_tab and k)
            or ((t and k) and not (prev_tab and prev_key))
        )
        return fire, t, k

    def _hotkey_label(self, key: str) -> str:
        parsed = self._parse_hotkey(self.hotkeys.get(key, ""))
        if parsed is None:
            return "（未设置）"
        mod, k = parsed
        return f"{mod.upper()}+{k.upper()}"

    def _status_enabled_text(self) -> str:
        return (
            f"已开启 — {self._hotkey_label('translate')} 划词翻译，"
            f"{self._hotkey_label('snip')} 截图 OCR，{self._hotkey_label('save_last')} 收录最近一条"
        )

    def _status_disabled_text(self) -> str:
        return (
            f"已关闭 — 不会响应 {self._hotkey_label('translate')} / "
            f"{self._hotkey_label('snip')} / {self._hotkey_label('save_last')}"
        )

    def _refresh_hotkey_hint(self) -> None:
        if self.hotkey_hint_var is None:
            return
        self.hotkey_hint_var.set(
            f"划词翻译：{self._hotkey_label('translate')}  |  截图 OCR：{self._hotkey_label('snip')}  |  收录：{self._hotkey_label('save_last')}"
        )

    def _backup_vocab_on_startup(self) -> tuple[bool, str]:
        if not os.path.isfile(self.vocab_path):
            return False, "未找到 vocab.json，跳过备份"
        items = self._load_vocab()
        count = len(items)
        stamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        os.makedirs(self.backup_dir, exist_ok=True)
        backup_name = f"vocab_backup_{stamp}_entries-{count}.json"
        backup_path = os.path.join(self.backup_dir, backup_name)
        try:
            shutil.copy2(self.vocab_path, backup_path)
        except Exception as exc:
            return False, f"备份失败：{exc}"
        return True, backup_path

    def translate(self, text: str) -> str:
        src = "google"
        if self.translate_source_var is not None:
            src = self.translate_source_var.get()
        if src == "mymemory":
            return _translate_mymemory(text)
        return _translate_parallel_race_zh(text)[0]

    def _set_status_safe(self, msg: str) -> None:
        if self.root is None:
            return

        def apply() -> None:
            if self.status_var is not None:
                self.status_var.set(msg)

        self.root.after(0, apply)

    def _translate_primary_source(self) -> str:
        src = "google"
        if self.translate_source_var is not None:
            try:
                src = self.translate_source_var.get()
            except tk.TclError:
                src = "google"
        return src

    def _translate_resilient(self, text: str) -> tuple[str, str | None]:
        """选 Google 时：多线路并发，谁先返回有效译文用谁。MyMemory 模式仍为单接口。"""
        if self._translate_primary_source() == "mymemory":
            return _translate_mymemory(text), None
        return _translate_parallel_race_zh(text)

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
            self.status_var.set(self._status_enabled_text())
        else:
            self.status_var.set(self._status_disabled_text())

    def _on_apply_hotkeys(self) -> None:
        if self.hotkey_translate_var is None or self.hotkey_snip_var is None or self.hotkey_save_var is None:
            return
        pending = {
            "translate": self._normalize_hotkey(self.hotkey_translate_var.get()),
            "snip": self._normalize_hotkey(self.hotkey_snip_var.get()),
            "save_last": self._normalize_hotkey(self.hotkey_save_var.get()),
        }
        for k, v in pending.items():
            if self._parse_hotkey(v) is None:
                if self.status_var is not None:
                    self.status_var.set(f"快捷键格式错误：{k}={v}（示例：ctrl+l / tab+q）")
                return
        if len(set(pending.values())) < 3:
            if self.status_var is not None:
                self.status_var.set("快捷键不能重复，请设置 3 组不同组合")
            return
        self.hotkeys = pending
        self._save_hotkeys()
        self._refresh_hotkey_hint()
        if self.status_var is not None:
            self.status_var.set(self._status_enabled_text() if self._is_translate_enabled() else self._status_disabled_text())

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

    def _show_cursor_status_near_cursor(self, msg: str, *, duration_ms: int | None = None) -> None:
        if self.root is None:
            return
        if self.floating_var is not None and not self.floating_var.get():
            return
        text = self.clean_text(msg)
        if not text:
            return
        if self.cursor_status_win is None:
            top = tk.Toplevel(self.root)
            top.overrideredirect(True)
            top.attributes("-topmost", True)
            top.attributes("-alpha", 0.92)
            top.configure(bg=UI_CURSOR_STATUS_BG, highlightbackground="#334155", highlightthickness=1)
            lbl = tk.Label(
                top,
                text=text,
                justify="left",
                anchor="w",
                bg=UI_CURSOR_STATUS_BG,
                fg=UI_CURSOR_STATUS_FG,
                padx=8,
                pady=5,
                font=tkfont.Font(family=FONT_FAMILY, size=9, weight="bold"),
            )
            lbl.pack(fill="both", expand=True)
            self.cursor_status_win = top
            self.cursor_status_label = lbl
        else:
            assert self.cursor_status_label is not None
            self.cursor_status_label.configure(text=text)

        assert self.cursor_status_win is not None
        self.cursor_status_win.update_idletasks()
        w = max(100, self.cursor_status_win.winfo_reqwidth())
        h = max(26, self.cursor_status_win.winfo_reqheight())
        cx, cy = self.get_cursor_pos()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = min(max(8, cx + 14), max(8, sw - w - 8))
        y = min(max(8, cy + 18), max(8, sh - h - 8))
        self.cursor_status_win.geometry(f"{w}x{h}+{x}+{y}")
        self.cursor_status_win.deiconify()
        self.cursor_status_win.lift()

        if self.cursor_status_timer_id is not None:
            self.cursor_status_win.after_cancel(self.cursor_status_timer_id)
            self.cursor_status_timer_id = None
        self._start_cursor_status_follow()
        if duration_ms is not None:
            self.cursor_status_timer_id = self.cursor_status_win.after(duration_ms, self._hide_cursor_status)

    def _start_cursor_status_follow(self) -> None:
        if self.cursor_status_win is None or self.root is None:
            return
        if self.cursor_status_follow_id is not None:
            try:
                self.cursor_status_win.after_cancel(self.cursor_status_follow_id)
            except Exception:
                pass
            self.cursor_status_follow_id = None

        def tick() -> None:
            if self.cursor_status_win is None or self.root is None:
                self.cursor_status_follow_id = None
                return
            if not self.cursor_status_win.winfo_viewable():
                self.cursor_status_follow_id = None
                return
            try:
                self.cursor_status_win.update_idletasks()
                w = max(100, self.cursor_status_win.winfo_reqwidth())
                h = max(26, self.cursor_status_win.winfo_reqheight())
                cx, cy = self.get_cursor_pos()
                sw = self.root.winfo_screenwidth()
                sh = self.root.winfo_screenheight()
                x = min(max(8, cx + 14), max(8, sw - w - 8))
                y = min(max(8, cy + 18), max(8, sh - h - 8))
                self.cursor_status_win.geometry(f"{w}x{h}+{x}+{y}")
            except Exception:
                self.cursor_status_follow_id = None
                return
            self.cursor_status_follow_id = self.cursor_status_win.after(70, tick)

        self.cursor_status_follow_id = self.cursor_status_win.after(70, tick)

    def _hide_cursor_status(self) -> None:
        if self.cursor_status_win is None:
            return
        if self.cursor_status_follow_id is not None:
            try:
                self.cursor_status_win.after_cancel(self.cursor_status_follow_id)
            except Exception:
                pass
            self.cursor_status_follow_id = None
        if self.cursor_status_timer_id is not None:
            try:
                self.cursor_status_win.after_cancel(self.cursor_status_timer_id)
            except Exception:
                pass
            self.cursor_status_timer_id = None
        try:
            self.cursor_status_win.withdraw()
        except Exception:
            pass

    def _set_cursor_status_safe(self, msg: str, *, duration_ms: int | None = None) -> None:
        if self.root is None:
            return
        self.root.after(0, lambda m=msg, d=duration_ms: self._show_cursor_status_near_cursor(m, duration_ms=d))

    def _ui_show_result(self, original: str, result: str, *, save_translation: str | None = None) -> None:
        stored = save_translation if save_translation is not None else result
        with self._last_lock:
            self._last_original = original
            self._last_translated = stored
        self._append_log(original, result)
        self._push_recent_translation(original, stored)
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

    def _push_recent_translation(self, original: str, translated: str) -> None:
        pair = (self.clean_text(original), self.clean_text(translated))
        if not pair[0] or not pair[1]:
            return
        self.recent_items.insert(0, pair)
        self.recent_items = self.recent_items[:3]
        self._refresh_recent_ui()

    def _refresh_recent_ui(self) -> None:
        if not self.recent_vars:
            return
        for i in range(3):
            if i < len(self.recent_items):
                w, m = self.recent_items[i]
                self.recent_vars[i].set(f"{w} => {m}")
            else:
                self.recent_vars[i].set("（暂无）")

    def _load_recent_saved_words(self) -> list[str]:
        items = self._load_vocab()
        words: list[str] = []
        for it in items:
            if not isinstance(it, dict):
                continue
            w = self.clean_text(str(it.get("word", "")))
            if w:
                words.append(w)
        words = list(dict.fromkeys(words))
        words.reverse()
        return words[:5]

    def _refresh_recent_saved_ui(self) -> None:
        if not self.recent_saved_vars:
            return
        self.recent_saved_words = self._load_recent_saved_words()
        for i in range(5):
            if i < len(self.recent_saved_words):
                self.recent_saved_vars[i].set(self.recent_saved_words[i])
            else:
                self.recent_saved_vars[i].set("（暂无）")

    def _delete_saved_word(self, idx: int) -> None:
        if idx < 0 or idx >= len(self.recent_saved_words):
            if self.status_var is not None:
                self.status_var.set("该条记录为空")
            return
        target = self.recent_saved_words[idx]
        items = self._load_vocab()
        new_items = [it for it in items if not (isinstance(it, dict) and self.clean_text(str(it.get("word", ""))) == target)]
        if len(new_items) == len(items):
            if self.status_var is not None:
                self.status_var.set(f"未找到词条：{target}")
            return
        try:
            self._save_vocab(new_items)
            self._refresh_recent_saved_ui()
            self._ui_vocab_feedback("生词本", f"已删除：{target}", floating=False)
        except Exception:
            if self.status_var is not None:
                self.status_var.set(f"删除失败：{target}")

    def _on_recent_save_click(self, idx: int) -> None:
        if idx < 0 or idx >= len(self.recent_items):
            if self.status_var is not None:
                self.status_var.set("该条记录为空")
            return
        word, meaning = self.recent_items[idx]
        threading.Thread(target=self._do_save_vocab_job, args=(word, meaning), daemon=True).start()

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
            # 某些应用复制响应慢，短暂轮询几次，降低误判“未选中”的概率。
            for _ in range(8):
                time.sleep(CLIPBOARD_STABLE_WAIT)
                copied = pyperclip.paste()
                if copied != before:
                    break
        return self.clean_text(copied)

    def _translate_text_job(self, text: str, *, no_text_hint: str) -> None:
        text = self.clean_text(text)
        if not text:
            if self.root is not None:
                self.root.after(0, lambda h=no_text_hint: self._ui_show_error("提示", h))
                self._set_cursor_status_safe("未检测到可翻译文本", duration_ms=1300)
                self._set_status_safe(
                    self._status_enabled_text() if self._is_translate_enabled() else self._status_disabled_text()
                )
            return
        if len(text) > MAX_TEXT_LENGTH:
            text = text[:MAX_TEXT_LENGTH] + "..."
        if self._is_likely_english(text):
            speak_vol = 100
            if self.tts_volume_var is not None:
                try:
                    speak_vol = max(0, min(100, int(self.tts_volume_var.get())))
                except (tk.TclError, TypeError, ValueError):
                    speak_vol = 100
            threading.Thread(target=self._speak_english_text, args=(text, speak_vol), daemon=True).start()
        try:
            self._set_cursor_status_safe("并发翻译中…")
            translated, fb_note = self._translate_resilient(text)
            display = f"{translated}\n{fb_note}" if fb_note else translated
            print(f"[{time.strftime('%H:%M:%S')}] {text} => {display}")
            if self.root is not None:

                def _done() -> None:
                    self._ui_show_result(text, display, save_translation=translated)
                    self._show_cursor_status_near_cursor("翻译完成", duration_ms=1000)
                    if self.status_var is not None:
                        self.status_var.set(
                            self._status_enabled_text() if self._is_translate_enabled() else self._status_disabled_text()
                        )

                self.root.after(0, _done)
        except Exception as exc:
            if self.root is not None:
                self.root.after(
                    0,
                    lambda t=text, e=exc: self._ui_show_error(t, _format_translate_failure(e)),
                )
                self._set_cursor_status_safe("翻译失败", duration_ms=1500)
                self._set_status_safe(
                    self._status_enabled_text() if self._is_translate_enabled() else self._status_disabled_text()
                )

    @staticmethod
    def _is_likely_english(text: str) -> bool:
        letters = re.findall(r"[A-Za-z]", text)
        return len(letters) >= 2

    @staticmethod
    def _speak_english_text(text: str, volume: int = 100) -> None:
        # 使用 Windows 内置 System.Speech 朗读，避免新增 Python 依赖。
        # Volume 为合成器输出电平（0–100），与系统音量滑块独立设置。
        vol = max(0, min(100, int(volume)))
        escaped = text.replace("'", "''")
        script = (
            "Add-Type -AssemblyName System.Speech; "
            "$s=New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            f"$s.Volume={vol}; "
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
        self._set_cursor_status_safe("读取划词内容…")
        text = self.copy_selected_text()
        self._translate_text_job(text, no_text_hint=f"未检测到选中文本，请先划词再按 {self._hotkey_label('translate')}")

    def _do_save_last_translation_job(self) -> None:
        with self._last_lock:
            word = self._last_original or ""
            meaning = self._last_translated or ""
        word = self.clean_text(word)
        meaning = self.clean_text(meaning)
        if not word or not meaning:
            if self.root is not None:
                self.root.after(
                    0,
                    lambda: self._ui_vocab_feedback(
                        "生词本",
                        f"暂无可收录内容：请先按 {self._hotkey_label('translate')} 翻译一次",
                        floating=False,
                    ),
                )
            return
        self._do_save_vocab_job(word, meaning)

    def _do_screen_ocr_translate_job(self, bbox: tuple[int, int, int, int]) -> None:
        if not self._is_translate_enabled():
            return
        with self._ocr_lock:
            if self._ocr_running:
                self._set_cursor_status_safe("OCR 正在进行中，请稍候…", duration_ms=1300)
                self._set_status_safe("OCR 正在进行中，请稍候…")
                return
            self._ocr_running = True
        self._set_cursor_status_safe("OCR 准备中…")
        try:
            if ImageGrab is None or pytesseract is None:
                if self.root is not None:
                    self.root.after(
                        0,
                        lambda: self._ui_show_error(
                            "OCR 不可用",
                            "缺少 OCR 依赖：请安装 pillow 与 pytesseract，并确保系统已安装 Tesseract-OCR。",
                        ),
                    )
                    self._set_cursor_status_safe("OCR 不可用", duration_ms=1500)
                return
            tesseract_bin = self._pick_tesseract_binary()
            if tesseract_bin:
                pytesseract.pytesseract.tesseract_cmd = tesseract_bin
            tessdata_dir = ""
            lang_files: set[str] = set()
            if tesseract_bin:
                maybe_dir = os.path.join(os.path.dirname(tesseract_bin), "tessdata")
                if os.path.isdir(maybe_dir):
                    tessdata_dir = maybe_dir
                    for fn in os.listdir(maybe_dir):
                        if fn.endswith(".traineddata"):
                            lang_files.add(fn.replace(".traineddata", ""))
            langs = OCR_LANG
            if self._ocr_langs_cache is None:
                try:
                    self._ocr_langs_cache = set(pytesseract.get_languages())
                except Exception:
                    self._ocr_langs_cache = set()
            known = self._ocr_langs_cache or lang_files
            if known:
                if {"eng", "chi_sim"}.issubset(known):
                    langs = "eng+chi_sim"
                elif "eng" in known:
                    langs = "eng"
                elif "chi_sim" in known:
                    langs = "chi_sim"
                else:
                    langs = next(iter(known))

            self._set_status_safe("OCR：正在截取屏幕…")
            self._set_cursor_status_safe("OCR 截图中…")
            old_prefix = os.environ.get("TESSDATA_PREFIX")
            image = None
            try:
                if tessdata_dir:
                    # 强制覆盖错误的全局 TESSDATA_PREFIX（例如指向 E:\tes）
                    os.environ["TESSDATA_PREFIX"] = tessdata_dir
                image = ImageGrab.grab(bbox=bbox, all_screens=True)
                try:
                    w, h = image.size
                    if w > 0 and h > 0:
                        pixels = w * h
                        if pixels > OCR_MAX_PIXELS:
                            scale = (OCR_MAX_PIXELS / float(pixels)) ** 0.5
                            nw = max(1, int(w * scale))
                            nh = max(1, int(h * scale))
                            image = image.resize((nw, nh))
                except Exception:
                    pass
                self._set_status_safe("OCR：正在识别文字（区域越大可能越慢，请稍候）…")
                self._set_cursor_status_safe("OCR 识别中…")
                text = pytesseract.image_to_string(image, lang=langs)
            except Exception as exc:
                self._set_status_safe(
                    self._status_enabled_text() if self._is_translate_enabled() else self._status_disabled_text()
                )
                if self.root is not None:
                    msg = str(exc)
                    if "TESSDATA_PREFIX" in msg or "couldn't load any languages" in msg.lower():
                        msg = (
                            "OCR 失败：未找到可用语言包。请安装 Tesseract 的 eng/chi_sim 语言文件，"
                            "或修正 TESSDATA_PREFIX 到 tessdata 目录。"
                        )
                    detail = f"[tesseract={tesseract_bin or '未找到'} | tessdata={tessdata_dir or '未找到'} | lang={langs}]"
                    self.root.after(0, lambda e=f"{msg}\n{detail}": self._ui_show_error("OCR 失败", e))
                self._set_cursor_status_safe("OCR 失败", duration_ms=1500)
                return
            finally:
                if image is not None:
                    try:
                        image.close()
                    except Exception:
                        pass
                    image = None
                if old_prefix is None:
                    os.environ.pop("TESSDATA_PREFIX", None)
                else:
                    os.environ["TESSDATA_PREFIX"] = old_prefix
            self._set_status_safe("OCR：正在翻译…")
            self._set_cursor_status_safe("并发翻译中…")
            self._translate_text_job(text, no_text_hint="截图区域未识别到文字，请重试更清晰区域")
        finally:
            with self._ocr_lock:
                self._ocr_running = False

    def _pick_tesseract_binary(self) -> str | None:
        """
        选择最合适的 tesseract.exe：
        - 优先官方默认目录且具备语言包（避免被历史 PATH 中的 E:\\tes 劫持）
        - 其次才使用 PATH 中的 tesseract
        """
        if os.path.isfile(FORCE_TESSERACT_PATH):
            return FORCE_TESSERACT_PATH

        candidates: list[str] = []
        for base in TESSERACT_CANDIDATE_DIRS:
            cand = os.path.join(base, "tesseract.exe")
            if os.path.isfile(cand):
                candidates.append(cand)
        from_path = shutil.which("tesseract")
        if from_path and from_path not in candidates:
            candidates.append(from_path)
        if not candidates:
            return None

        def lang_score(exe_path: str) -> tuple[int, int]:
            td = os.path.join(os.path.dirname(exe_path), "tessdata")
            if not os.path.isdir(td):
                return (0, 0)
            has_eng = os.path.isfile(os.path.join(td, "eng.traineddata"))
            has_zh = os.path.isfile(os.path.join(td, "chi_sim.traineddata"))
            # 分值越高越好：优先中英齐全，其次仅英文
            return (1 if has_eng else 0) + (2 if has_zh else 0), 1

        best = max(candidates, key=lang_score)
        return best

    def _win32_force_foreground_hwnd(self, hwnd: int) -> None:
        """
        将本进程窗口推到 Windows 前台。
        在其它应用拥有焦点时，仅靠 Tk 的 focus_force 往往无效；需 AttachThreadInput 配合 SetForegroundWindow。
        """
        if hwnd <= 0:
            return
        user32 = self.user32
        kernel32 = self.kernel32
        SW_RESTORE = 9
        target = wintypes.HWND(hwnd)
        try:
            user32.ShowWindow(target, SW_RESTORE)
        except Exception:
            pass
        fg = user32.GetForegroundWindow()
        if not fg:
            user32.SetForegroundWindow(target)
            user32.BringWindowToTop(target)
            return
        cur_tid = kernel32.GetCurrentThreadId()
        proc = wintypes.DWORD(0)
        fg_tid = user32.GetWindowThreadProcessId(fg, ctypes.byref(proc))
        if fg_tid == 0 or fg_tid == cur_tid:
            user32.SetForegroundWindow(target)
            user32.BringWindowToTop(target)
            return
        user32.AttachThreadInput(fg_tid, cur_tid, True)
        try:
            user32.SetForegroundWindow(target)
            user32.BringWindowToTop(target)
        finally:
            user32.AttachThreadInput(fg_tid, cur_tid, False)

    def _win32_activate_snip_overlay(self, overlay: tk.Toplevel) -> None:
        """从后台热键打开截屏层时，把主窗与遮罩一并抢到前台，避免必须再点一次主界面。"""
        if self.root is None:
            return
        try:
            self.root.deiconify()
        except tk.TclError:
            pass
        self.root.update_idletasks()
        overlay.update_idletasks()
        try:
            self._win32_force_foreground_hwnd(int(self.root.winfo_id()))
        except Exception:
            pass
        try:
            self._win32_force_foreground_hwnd(int(overlay.winfo_id()))
        except Exception:
            pass
        try:
            overlay.focus_force()
        except tk.TclError:
            pass

    def _snip_cancel(self, msg: str | None = None) -> None:
        if self._snip_overlay is not None:
            try:
                self._snip_overlay.destroy()
            except Exception:
                pass
        self._snip_overlay = None
        self._snip_canvas = None
        self._snip_start = None
        self._snip_rect_id = None
        self._snip_busy = False
        if msg and self.status_var is not None:
            self.status_var.set(msg)

    def _begin_screen_snip(self) -> None:
        if self.root is None:
            return
        if self._snip_busy:
            return
        self._snip_busy = True
        self.status_var and self.status_var.set("截图模式：拖拽选择区域，ESC 取消")

        overlay: tk.Toplevel | None = None
        try:
            overlay = tk.Toplevel(self.root)
            self._snip_overlay = overlay
            # Win32：必须先全屏再 overrideredirect，否则会 TclError: can't set fullscreen ... override-redirect flag is set
            overlay.attributes("-fullscreen", True)
            overlay.overrideredirect(True)
            overlay.configure(bg=SNIP_OVERLAY_CANVAS_BG)

            canvas = tk.Canvas(
                overlay,
                bg=SNIP_OVERLAY_CANVAS_BG,
                highlightthickness=0,
                cursor="crosshair",
            )
            canvas.pack(fill="both", expand=True)
            hint = "拖拽框选要 OCR 翻译的区域（松开鼠标确认 / ESC 取消）"
            canvas.create_text(18, 20, text=hint, fill="#f8fafc", anchor="w", font=(FONT_FAMILY, 12, "bold"))

            def _snip_on_map(_event: tk.Event | None = None) -> None:
                try:
                    overlay.attributes("-alpha", SNIP_OVERLAY_ALPHA)
                except tk.TclError:
                    pass

            overlay.bind("<Map>", _snip_on_map)

            self._snip_canvas = canvas
            self._snip_start = None
            self._snip_rect_id = None
            self._snip_info_id = canvas.create_text(
                18,
                48,
                text="",
                fill="#93c5fd",
                anchor="w",
                font=(FONT_FAMILY, 10, "bold"),
            )

            def on_press(event: tk.Event) -> None:
                self._snip_start = (event.x_root, event.y_root)
                if self._snip_rect_id is not None:
                    canvas.delete(self._snip_rect_id)
                self._snip_rect_id = canvas.create_rectangle(
                    event.x,
                    event.y,
                    event.x,
                    event.y,
                    outline="#38bdf8",
                    width=3,
                    fill="#38bdf8",
                    stipple="gray50",
                )

            def on_drag(event: tk.Event) -> None:
                if self._snip_start is None or self._snip_rect_id is None:
                    return
                x0, y0 = self._snip_start
                x1, y1 = event.x_root, event.y_root
                canvas.coords(
                    self._snip_rect_id,
                    x0 - overlay.winfo_rootx(),
                    y0 - overlay.winfo_rooty(),
                    event.x,
                    event.y,
                )
                w = abs(x1 - x0)
                h = abs(y1 - y0)
                if self._snip_info_id is not None:
                    canvas.itemconfigure(self._snip_info_id, text=f"区域大小：{w} × {h}")

            def on_release(event: tk.Event) -> None:
                if self._snip_start is None:
                    return
                x0, y0 = self._snip_start
                x1, y1 = event.x_root, event.y_root
                left, top = min(x0, x1), min(y0, y1)
                right, bottom = max(x0, x1), max(y0, y1)
                if right - left < 6 or bottom - top < 6:
                    self._snip_cancel("截图区域太小，已取消")
                    return
                self._snip_cancel("OCR 识别中…")
                threading.Thread(target=self._do_screen_ocr_translate_job, args=((left, top, right, bottom),), daemon=True).start()

            overlay.bind("<Escape>", lambda e: self._snip_cancel("已取消截图"))
            canvas.bind("<ButtonPress-1>", on_press)
            canvas.bind("<B1-Motion>", on_drag)
            canvas.bind("<ButtonRelease-1>", on_release)
            self._finalize_snip_overlay(overlay)
            self._win32_activate_snip_overlay(overlay)
        except Exception as exc:
            if overlay is not None:
                try:
                    overlay.destroy()
                except Exception:
                    pass
            self._snip_overlay = None
            self._snip_canvas = None
            self._snip_start = None
            self._snip_rect_id = None
            self._snip_info_id = None
            self._snip_busy = False
            if self.status_var is not None:
                self.status_var.set(f"截图模式启动失败：{exc}")

    def _finalize_snip_overlay(self, overlay: tk.Toplevel) -> None:
        """统一遮罩视觉：置顶 + 透明度在窗口映射后重复设置，减轻 Win32 上忽明忽暗。"""
        overlay.update_idletasks()
        try:
            overlay.attributes("-topmost", True)
        except tk.TclError:
            pass
        for _ in range(2):
            try:
                overlay.attributes("-alpha", SNIP_OVERLAY_ALPHA)
            except tk.TclError:
                break
            overlay.update_idletasks()
        try:
            overlay.lift(self.root)
        except tk.TclError:
            overlay.lift()

    def _tab_combo_loop(self) -> None:
        prev_translate = False
        prev_snip = False
        prev_snip_tab = False
        prev_snip_key = False
        prev_save = False
        prev_save_tab = False
        prev_save_key = False
        while not self._closing:
            pressed_translate = self._is_hotkey_pressed(self.hotkeys.get("translate", ""))

            combo_snip = self.hotkeys.get("snip", "")
            ps = self._parse_hotkey(combo_snip)
            if ps and ps[0] == "tab":
                fire_snip, prev_snip_tab, prev_snip_key = self._tab_mod_key_fire_edge(
                    combo_snip, prev_snip_tab, prev_snip_key
                )
                pressed_snip = fire_snip
            else:
                pressed_snip = self._is_hotkey_pressed(combo_snip)
                prev_snip_tab = False
                prev_snip_key = False

            combo_save = self.hotkeys.get("save_last", "")
            pz = self._parse_hotkey(combo_save)
            if pz and pz[0] == "tab":
                fire_save, prev_save_tab, prev_save_key = self._tab_mod_key_fire_edge(
                    combo_save, prev_save_tab, prev_save_key
                )
                pressed_save = fire_save
            else:
                pressed_save = self._is_hotkey_pressed(combo_save)
                prev_save_tab = False
                prev_save_key = False

            if pressed_translate and not prev_translate:
                threading.Thread(target=self._do_translate_job, daemon=True).start()
            if pressed_snip and not prev_snip and self.root is not None:
                self.root.after(0, self._begin_screen_snip)
            if pressed_save and not prev_save:
                threading.Thread(target=self._do_save_last_translation_job, daemon=True).start()

            prev_translate = pressed_translate
            prev_snip = pressed_snip
            prev_save = pressed_save
            time.sleep(0.008)

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
                self.root.after(0, lambda: self._ui_vocab_feedback("生词本", "暂无可记录内容：请先翻译一次", floating=False))
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
                self.root.after(0, self._refresh_recent_saved_ui)
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
        self.translate_source_var = tk.StringVar(master=self.root, value="google")
        self.tts_volume_var = tk.IntVar(master=self.root, value=self._tts_volume_default)
        self.hotkey_translate_var = tk.StringVar(master=self.root, value=self.hotkeys.get("translate", DEFAULT_HOTKEYS["translate"]))
        self.hotkey_snip_var = tk.StringVar(master=self.root, value=self.hotkeys.get("snip", DEFAULT_HOTKEYS["snip"]))
        self.hotkey_save_var = tk.StringVar(master=self.root, value=self.hotkeys.get("save_last", DEFAULT_HOTKEYS["save_last"]))
        self.hotkey_hint_var = tk.StringVar(master=self.root, value="")
        self.tts_volume_var.trace_add("write", self._on_tts_volume_change)
        self.recent_vars = [tk.StringVar(master=self.root, value="（暂无）") for _ in range(3)]
        self.recent_saved_vars = [tk.StringVar(master=self.root, value="（暂无）") for _ in range(5)]

        self.root.title("SnapTranslate")
        self.root.geometry("760x740")
        self.root.minsize(620, 520)
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
            text="选中文字后按快捷键翻译 · 简体中文 · 可朗读英文",
            font=tkfont.Font(family=FONT_FAMILY, size=10),
            bg=UI_BG,
            fg=UI_TEXT_MUTED,
        ).pack(anchor="w", pady=(4, 0))
        tk.Label(
            hdr,
            textvariable=self.hotkey_hint_var,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
            bg=UI_BG,
            fg=UI_TEXT_MUTED,
        ).pack(anchor="w", pady=(2, 0))

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
            text="启用划词翻译",
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
            text="自动竞速（Google 双线路 + MyMemory + Lingva 镜像，谁先成功用谁）",
            variable=self.translate_source_var,
            value="google",
            **rb_kw,
        ).pack(side="left")
        tk.Label(
            src_wrap,
            text="竞速含 translate.googleapis.com、clients5.google.com、MyMemory、Lingva；MyMemory 用量大时可能提示配额用尽。",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=8),
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        hk_wrap = tk.Frame(ctrl_card, bg=UI_CARD)
        hk_wrap.pack(fill="x", pady=(12, 0))
        tk.Label(
            hk_wrap,
            text="快捷键设置（格式：ctrl+l / tab+q）",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w")
        hk_row = tk.Frame(hk_wrap, bg=UI_CARD)
        hk_row.pack(fill="x", pady=(5, 0))
        hk_ent_kw = dict(
            font=tkfont.Font(family="Consolas", size=9),
            bg=UI_LOG_BG,
            fg=UI_TEXT,
            insertbackground=UI_TEXT,
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            width=10,
        )
        tk.Label(hk_row, text="翻译", bg=UI_CARD, fg=UI_TEXT_MUTED, font=tkfont.Font(family=FONT_FAMILY, size=9)).pack(side="left")
        tk.Entry(hk_row, textvariable=self.hotkey_translate_var, **hk_ent_kw).pack(side="left", padx=(4, 10))
        tk.Label(hk_row, text="截图", bg=UI_CARD, fg=UI_TEXT_MUTED, font=tkfont.Font(family=FONT_FAMILY, size=9)).pack(side="left")
        tk.Entry(hk_row, textvariable=self.hotkey_snip_var, **hk_ent_kw).pack(side="left", padx=(4, 10))
        tk.Label(hk_row, text="收录", bg=UI_CARD, fg=UI_TEXT_MUTED, font=tkfont.Font(family=FONT_FAMILY, size=9)).pack(side="left")
        tk.Entry(hk_row, textvariable=self.hotkey_save_var, **hk_ent_kw).pack(side="left", padx=(4, 10))
        tk.Button(
            hk_row,
            text="应用并保存",
            command=self._on_apply_hotkeys,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
            bg=UI_CARD,
            fg=UI_ACCENT,
            activebackground=UI_CHIP,
            activeforeground=UI_ACCENT_HOVER,
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            padx=8,
            pady=3,
            cursor="hand2",
        ).pack(side="left")

        tts_wrap = tk.Frame(ctrl_card, bg=UI_CARD)
        tts_wrap.pack(fill="x", pady=(12, 0))
        tk.Label(
            tts_wrap,
            text="英文朗读音量（独立于系统音量）",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w")
        assert self.tts_volume_var is not None
        tk.Scale(
            tts_wrap,
            from_=0,
            to=100,
            orient="horizontal",
            variable=self.tts_volume_var,
            resolution=1,
            showvalue=True,
            bg=UI_CARD,
            fg=UI_TEXT,
            troughcolor=UI_LOG_BG,
            highlightthickness=0,
            length=260,
        ).pack(anchor="w", pady=(4, 0))

        recent_wrap = tk.Frame(ctrl_card, bg=UI_CARD)
        recent_wrap.pack(fill="x", pady=(12, 0))
        tk.Label(
            recent_wrap,
            text="最近 3 条翻译（可直接收录生词本）",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w")
        for i in range(3):
            r = tk.Frame(recent_wrap, bg=UI_CARD)
            r.pack(fill="x", pady=(6 if i == 0 else 4, 0))
            tk.Label(
                r,
                textvariable=self.recent_vars[i],
                bg=UI_LOG_BG,
                fg=UI_TEXT,
                anchor="w",
                justify="left",
                padx=10,
                pady=7,
                font=tkfont.Font(family=FONT_FAMILY, size=9),
                highlightbackground=UI_BORDER,
                highlightthickness=1,
            ).pack(side="left", fill="x", expand=True)
            tk.Button(
                r,
                text="收录",
                command=lambda n=i: self._on_recent_save_click(n),
                font=tkfont.Font(family=FONT_FAMILY, size=9, weight="bold"),
                bg=UI_ACCENT,
                fg="#ffffff",
                activebackground=UI_ACCENT_HOVER,
                activeforeground="#ffffff",
                relief="flat",
                padx=10,
                pady=6,
                cursor="hand2",
            ).pack(side="left", padx=(8, 0))

        recent_saved_wrap = tk.Frame(ctrl_card, bg=UI_CARD)
        recent_saved_wrap.pack(fill="x", pady=(12, 0))
        tk.Label(
            recent_saved_wrap,
            text="最近加入生词本（可一键删除）",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w")
        for i in range(5):
            r = tk.Frame(recent_saved_wrap, bg=UI_CARD)
            r.pack(fill="x", pady=(6 if i == 0 else 4, 0))
            tk.Label(
                r,
                textvariable=self.recent_saved_vars[i],
                bg=UI_LOG_BG,
                fg=UI_TEXT,
                anchor="w",
                justify="left",
                padx=10,
                pady=7,
                font=tkfont.Font(family=FONT_FAMILY, size=9),
                highlightbackground=UI_BORDER,
                highlightthickness=1,
            ).pack(side="left", fill="x", expand=True)
            tk.Button(
                r,
                text="删除",
                command=lambda n=i: self._delete_saved_word(n),
                font=tkfont.Font(family=FONT_FAMILY, size=9, weight="bold"),
                bg="#fee2e2",
                fg="#b91c1c",
                activebackground="#fecaca",
                activeforeground="#991b1b",
                relief="flat",
                padx=10,
                pady=6,
                cursor="hand2",
            ).pack(side="left", padx=(8, 0))

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
        self._refresh_hotkey_hint()
        self._refresh_recent_saved_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self) -> None:
        self._closing = True
        tid = self.hotkey_thread_id
        if tid:
            self.user32.PostThreadMessageW(tid, WM_QUIT, 0, 0)
        if self.tab_q_thread is not None:
            self.tab_q_thread.join(timeout=1.0)
        if self.root is not None:
            self.root.destroy()

    def run(self) -> None:
        self._build_ui()
        assert self.root is not None
        ok, detail = self._backup_vocab_on_startup()
        if ok:
            print(f"[{time.strftime('%H:%M:%S')}] 生词本已备份：{detail}")
        else:
            print(f"[{time.strftime('%H:%M:%S')}] {detail}")

        self.tab_q_thread = threading.Thread(target=self._tab_combo_loop, daemon=True)
        self.tab_q_thread.start()

        if self.status_var is not None:
            self.status_var.set(self._status_enabled_text())
        self.root.mainloop()


def main() -> None:
    app = TranslatorApp()
    app.run()


if __name__ == "__main__":
    main()
