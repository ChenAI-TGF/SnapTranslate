"""
本地生词本复习 + 用 DeepSeek API 批量生成「英文例句 + 例句中文翻译」，写入 vocab.json。
词条字段：word, meaning, example, example_zh, score（0–100，默认 50）。
复习时通过「认识 / 模糊 / 不认识」计分：「认识」加分，「模糊」「不认识」扣分；是否已打开「显示释义/例句」会影响幅度。
API Key 优先从 api_key.txt 读取，若无则可在界面输入并保存。
"""
from __future__ import annotations

import json
import os
import random
import re
import subprocess
import threading
import time
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, scrolledtext

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_VOCAB = os.path.join(SCRIPT_DIR, "vocab.json")
DEFAULT_KEY_FILE = os.path.join(SCRIPT_DIR, "api_key.txt")
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-chat"

# 界面主题（与 main.py 一致）
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
UI_EXAMPLE_PANEL = "#f8fafc"
UI_MEANING = "#047857"
UI_KEYWORD = "#4f46e5"
UI_ZH = "#0d9488"
FONT_FAMILY = "Microsoft YaHei UI"

DEFAULT_SCORE = 50.0
SCORE_MIN = 0.0
SCORE_MAX = 100.0
# (档位, 是否已在当前张上点过「显示释义/例句」): 分数变化
GRADE_DELTA: dict[tuple[str, bool], float] = {
    ("know", False): 10.0,
    ("vague", False): -4.0,
    ("unknown", False): -8.0,
    ("know", True): 5.0,
    ("vague", True): -7.0,
    ("unknown", True): -12.0,
}


def item_score(it: dict) -> float:
    s = it.get("score")
    if s is None:
        v = DEFAULT_SCORE
    else:
        try:
            v = float(s)
        except (TypeError, ValueError):
            v = DEFAULT_SCORE
    return max(SCORE_MIN, min(SCORE_MAX, v))


def normalize_vocab_scores(items: list[dict]) -> None:
    for it in items:
        it["score"] = item_score(it)


def load_vocab(path: str) -> list[dict]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            return []
        return [x for x in data if isinstance(x, dict)]
    except FileNotFoundError:
        return []
    except Exception:
        return []


def save_vocab(path: str, items: list[dict]) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def read_api_key_file(key_path: str) -> str | None:
    try:
        with open(key_path, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if s:
                    return s
    except FileNotFoundError:
        return None
    except Exception:
        return None
    return None


def write_api_key_file(key_path: str, key: str) -> None:
    tmp = key_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(key.strip() + "\n")
    os.replace(tmp, key_path)


def needs_bilingual_example(it: dict) -> bool:
    """英文例句或其中文译本任一为缺，则需要生成/补全（一次请求同时写入两者）。"""
    ex = it.get("example", "")
    if ex is None or not str(ex).strip():
        return True
    zh = it.get("example_zh", "")
    if zh is None or not str(zh).strip():
        return True
    return False


def count_pending_examples(items: list[dict]) -> int:
    return sum(1 for it in items if needs_bilingual_example(it))


def parse_bilingual_response(raw: str) -> tuple[str, str]:
    """从模型输出中解析 {"example":"...","example_zh":"..."}。"""
    s = (raw or "").strip()
    if s.startswith("```"):
        s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s*```\s*$", "", s)
    start = s.find("{")
    end = s.rfind("}")
    if start >= 0 and end > start:
        s = s[start : end + 1]
    data = json.loads(s)
    if not isinstance(data, dict):
        raise ValueError("模型返回不是 JSON 对象")
    en = str(data.get("example", "")).strip()
    zh = str(data.get("example_zh", "")).strip()
    return en, zh


def _is_insufficient_balance_error(exc: BaseException) -> bool:
    """DeepSeek/OpenAI SDK：余额不足通常返回 HTTP 402。"""
    code = getattr(exc, "status_code", None)
    if code == 402:
        return True
    msg = str(exc).lower()
    return "insufficient balance" in msg or ("402" in msg and "balance" in msg)


class VocabReviewApp:
    def __init__(self) -> None:
        self.vocab_path = DEFAULT_VOCAB
        self.key_path = DEFAULT_KEY_FILE
        self.vocab: list[dict] = load_vocab(self.vocab_path)
        normalize_vocab_scores(self.vocab)
        self.order: list[int] = []
        self.pos = 0
        self.reveal_meaning = False
        self._gen_running = False
        self._speak_generation = 0

        self.root = tk.Tk()
        self.root.title("SnapTranslate — 生词复习")
        self.root.geometry("680x640")
        self.root.minsize(560, 480)
        self.root.configure(bg=UI_BG)

        # Tk 变量必须在 root 创建之后绑定，否则会 RuntimeError: Too early to create variable
        # 顺序：random | score_asc | score_desc
        self.sort_mode_var = tk.StringVar(master=self.root, value="random")
        self.api_key_var = tk.StringVar(master=self.root, value=read_api_key_file(self.key_path) or "")
        self.status_var = tk.StringVar(master=self.root, value="就绪")
        self.word_var = tk.StringVar(master=self.root, value="")
        self.meaning_var = tk.StringVar(master=self.root, value="（点击「显示释义 / 例句」）")
        self.progress_var = tk.StringVar(master=self.root, value="0 / 0")
        self.score_var = tk.StringVar(master=self.root, value="熟练度 —")
        # 朗读：none | word | word_example
        self.read_mode_var = tk.StringVar(master=self.root, value="none")

        self._build_ui()
        self._rebuild_order()
        self._show_card()

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=UI_BG, padx=18, pady=14)
        outer.pack(fill="both", expand=True)

        hdr = tk.Frame(outer, bg=UI_BG)
        hdr.pack(fill="x", pady=(0, 12))
        tk.Label(
            hdr,
            text="生词复习",
            font=tkfont.Font(family=FONT_FAMILY, size=20, weight="bold"),
            bg=UI_BG,
            fg=UI_TEXT,
        ).pack(anchor="w")
        tk.Label(
            hdr,
            text="自评保存熟练度 · DeepSeek 可批量生成英例与中译",
            font=tkfont.Font(family=FONT_FAMILY, size=10),
            bg=UI_BG,
            fg=UI_TEXT_MUTED,
        ).pack(anchor="w", pady=(4, 0))

        settings = tk.Frame(
            outer,
            bg=UI_CARD,
            highlightbackground=UI_BORDER,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        settings.pack(fill="x", pady=(0, 10))

        path_row = tk.Frame(settings, bg=UI_CARD)
        path_row.pack(fill="x", pady=(0, 8))
        tk.Label(path_row, text="词表", bg=UI_CARD, fg=UI_TEXT_MUTED, font=tkfont.Font(family=FONT_FAMILY, size=9)).pack(
            anchor="w"
        )
        path_inner = tk.Frame(path_row, bg=UI_CARD)
        path_inner.pack(fill="x", pady=(4, 0))
        self.path_label = tk.Label(
            path_inner,
            text=self.vocab_path,
            bg=UI_LOG_BG,
            fg=UI_ACCENT,
            cursor="hand2",
            font=tkfont.Font(family="Consolas", size=9),
            anchor="w",
            padx=10,
            pady=8,
            highlightbackground=UI_BORDER,
            highlightthickness=1,
        )
        self.path_label.pack(side="left", fill="x", expand=True)
        self.path_label.bind("<Button-1>", lambda e: self._pick_vocab_file())
        tk.Button(
            path_inner,
            text="浏览…",
            command=self._pick_vocab_file,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
            bg=UI_ACCENT,
            fg="#ffffff",
            activebackground=UI_ACCENT_HOVER,
            activeforeground="#ffffff",
            relief="flat",
            padx=12,
            pady=6,
            cursor="hand2",
        ).pack(side="right", padx=(8, 0))

        key_row = tk.Frame(settings, bg=UI_CARD)
        key_row.pack(fill="x", pady=(0, 8))
        tk.Label(
            key_row, text="DeepSeek API Key", bg=UI_CARD, fg=UI_TEXT_MUTED, font=tkfont.Font(family=FONT_FAMILY, size=9)
        ).pack(anchor="w")
        key_inner = tk.Frame(key_row, bg=UI_CARD)
        key_inner.pack(fill="x", pady=(4, 0))
        ent = tk.Entry(
            key_inner,
            textvariable=self.api_key_var,
            show="*",
            font=tkfont.Font(family="Consolas", size=10),
            bg=UI_LOG_BG,
            fg=UI_TEXT,
            insertbackground=UI_TEXT,
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
        )
        ent.pack(side="left", fill="x", expand=True, ipady=6)
        tk.Button(
            key_inner,
            text="保存到文件",
            command=self._save_key_clicked,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
            bg=UI_CARD,
            fg=UI_ACCENT,
            activebackground=UI_CHIP,
            activeforeground=UI_ACCENT_HOVER,
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            padx=12,
            pady=6,
            cursor="hand2",
        ).pack(side="right", padx=(8, 0))

        sort_row = tk.Frame(settings, bg=UI_CARD)
        sort_row.pack(fill="x", pady=(4, 4))
        tk.Label(
            sort_row,
            text="复习顺序",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w")
        sort_btns = tk.Frame(settings, bg=UI_CARD)
        sort_btns.pack(fill="x", pady=(4, 0))
        rb_kw = dict(
            bg=UI_CARD,
            activebackground=UI_CARD,
            fg=UI_TEXT,
            selectcolor=UI_CHIP,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        )
        for label, val in (
            ("随机", "random"),
            ("得分低→高", "score_asc"),
            ("得分高→低", "score_desc"),
        ):
            tk.Radiobutton(
                sort_btns,
                text=label,
                variable=self.sort_mode_var,
                value=val,
                command=self._on_sort_mode_change,
                **rb_kw,
            ).pack(side="left", padx=(0, 14))

        read_row = tk.Frame(settings, bg=UI_CARD)
        read_row.pack(fill="x", pady=(8, 0))
        tk.Label(
            read_row,
            text="朗读（系统英文语音）",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w")
        read_btns = tk.Frame(settings, bg=UI_CARD)
        read_btns.pack(fill="x", pady=(4, 0))
        for label, val in (
            ("不朗读", "none"),
            ("单词", "word"),
            ("单词+例句", "word_example"),
        ):
            tk.Radiobutton(
                read_btns,
                text=label,
                variable=self.read_mode_var,
                value=val,
                **rb_kw,
            ).pack(side="left", padx=(0, 14))

        card = tk.Frame(
            outer,
            bg=UI_CARD,
            highlightbackground=UI_BORDER,
            highlightthickness=1,
            padx=18,
            pady=16,
        )
        card.pack(fill="both", expand=True, pady=(0, 10))

        card_hdr = tk.Frame(card, bg=UI_CARD)
        card_hdr.pack(fill="x")
        tk.Label(
            card_hdr,
            textvariable=self.progress_var,
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(side="left")
        tk.Label(
            card_hdr,
            textvariable=self.score_var,
            bg=UI_CARD,
            fg=UI_TEXT,
            font=tkfont.Font(family=FONT_FAMILY, size=9, weight="bold"),
        ).pack(side="right")

        tk.Label(
            card,
            textvariable=self.word_var,
            bg=UI_CARD,
            fg=UI_TEXT,
            font=tkfont.Font(family=FONT_FAMILY, size=20, weight="bold"),
            wraplength=520,
            justify="center",
        ).pack(pady=(8, 12))

        tk.Button(
            card,
            text="显示释义 / 例句",
            command=self._toggle_reveal,
            font=tkfont.Font(family=FONT_FAMILY, size=10),
            bg=UI_CHIP,
            fg=UI_ACCENT,
            activebackground=UI_ACCENT,
            activeforeground="#ffffff",
            relief="flat",
            padx=18,
            pady=8,
            cursor="hand2",
        ).pack(pady=(0, 8))

        tk.Label(
            card,
            textvariable=self.meaning_var,
            bg=UI_CARD,
            fg=UI_MEANING,
            font=tkfont.Font(family=FONT_FAMILY, size=12),
            wraplength=520,
            justify="center",
        ).pack(pady=(0, 8))

        self._font_example = tkfont.Font(family=FONT_FAMILY, size=11)
        self._font_example_bold = tkfont.Font(family=FONT_FAMILY, size=11, weight="bold")
        ex_wrap = tk.Frame(card, bg=UI_EXAMPLE_PANEL, highlightbackground=UI_BORDER, highlightthickness=1)
        ex_wrap.pack(fill="x", pady=(0, 4))
        self.example_text = tk.Text(
            ex_wrap,
            height=6,
            width=58,
            wrap="word",
            state="disabled",
            bg=UI_EXAMPLE_PANEL,
            fg=UI_KEYWORD,
            insertbackground=UI_KEYWORD,
            relief="flat",
            padx=10,
            pady=10,
            cursor="arrow",
            font=self._font_example,
            highlightthickness=0,
        )
        self.example_text.tag_configure("keyword", font=self._font_example_bold, foreground=UI_KEYWORD)
        self.example_text.tag_configure("zh_line", foreground=UI_ZH)
        self.example_text.tag_configure("muted", foreground=UI_TEXT_MUTED)
        self.example_text.pack(fill="x")

        sep = tk.Frame(card, bg=UI_BORDER, height=1)
        sep.pack(fill="x", pady=(14, 10))

        grade_fr = tk.Frame(card, bg=UI_CARD)
        grade_fr.pack(fill="x")
        tk.Label(
            grade_fr,
            text="自评（是否已看释义/例句会影响加减分）",
            bg=UI_CARD,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
        ).pack(anchor="w", pady=(0, 8))
        grade_btns = tk.Frame(grade_fr, bg=UI_CARD)
        grade_btns.pack(fill="x")
        btn_grade_kw = dict(relief="flat", font=tkfont.Font(family=FONT_FAMILY, size=10, weight="bold"), cursor="hand2", pady=10)
        tk.Button(
            grade_btns,
            text="认识",
            command=lambda: self._apply_grade("know"),
            bg="#059669",
            fg="#FFFFFF",
            activebackground="#047857",
            activeforeground="#FFFFFF",
            **btn_grade_kw,
        ).pack(side="left", expand=True, fill="x", padx=(0, 6))
        tk.Button(
            grade_btns,
            text="模糊",
            command=lambda: self._apply_grade("vague"),
            bg="#ea580c",
            fg="#FFFFFF",
            activebackground="#c2410c",
            activeforeground="#FFFFFF",
            **btn_grade_kw,
        ).pack(side="left", expand=True, fill="x", padx=(0, 6))
        tk.Button(
            grade_btns,
            text="不认识",
            command=lambda: self._apply_grade("unknown"),
            bg="#dc2626",
            fg="#FFFFFF",
            activebackground="#b91c1c",
            activeforeground="#FFFFFF",
            **btn_grade_kw,
        ).pack(side="left", expand=True, fill="x")

        actions = tk.Frame(outer, bg=UI_BG)
        actions.pack(fill="x", pady=(0, 8))

        empty_n = count_pending_examples(self.vocab)
        self.gen_btn = tk.Button(
            actions,
            text=f"用 DeepSeek 生成英例句 + 中译（约 {empty_n} 条待补全）",
            command=self._start_generate_examples,
            bg=UI_ACCENT,
            fg="#FFFFFF",
            activebackground=UI_ACCENT_HOVER,
            activeforeground="#FFFFFF",
            relief="flat",
            padx=14,
            pady=12,
            font=tkfont.Font(family=FONT_FAMILY, size=10, weight="bold"),
            cursor="hand2",
        )
        self.gen_btn.pack(fill="x")

        log_fr = tk.Frame(outer, bg=UI_BG)
        log_fr.pack(fill="both", expand=True, pady=(0, 6))
        stat_bar = tk.Frame(log_fr, bg=UI_STATUS_BG, padx=10, pady=8)
        stat_bar.pack(fill="x", pady=(0, 8))
        tk.Label(
            stat_bar,
            textvariable=self.status_var,
            bg=UI_STATUS_BG,
            fg=UI_TEXT_MUTED,
            font=tkfont.Font(family=FONT_FAMILY, size=9),
            anchor="w",
        ).pack(fill="x")

        self.log_text = scrolledtext.ScrolledText(
            log_fr,
            height=6,
            wrap="word",
            state="disabled",
            font=tkfont.Font(family="Consolas", size=9),
            bg=UI_LOG_BG,
            fg=UI_TEXT,
            insertbackground=UI_TEXT,
            relief="flat",
            bd=0,
            padx=10,
            pady=10,
            highlightthickness=0,
        )
        self.log_text.pack(fill="both", expand=True)

    def _log(self, line: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _pick_vocab_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 vocab.json",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
            initialdir=SCRIPT_DIR,
        )
        if not path:
            return
        self.vocab_path = path
        self.path_label.configure(text=self.vocab_path)
        self.vocab = load_vocab(self.vocab_path)
        normalize_vocab_scores(self.vocab)
        self._rebuild_order()
        self.pos = 0
        self.reveal_meaning = False
        self._refresh_gen_button_label()
        self._show_card()
        self._log(f"已加载词表：{self.vocab_path}（{len(self.vocab)} 条）")

    def _save_key_clicked(self) -> None:
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showwarning("提示", "请先填写 API Key。")
            return
        try:
            write_api_key_file(self.key_path, key)
            messagebox.showinfo("成功", f"已保存到 {self.key_path}")
        except Exception as e:
            messagebox.showerror("失败", str(e))

    def _on_sort_mode_change(self) -> None:
        self._rebuild_order()
        self.pos = 0
        self.reveal_meaning = False
        self._show_card()

    def _rebuild_order(self) -> None:
        n = len(self.vocab)
        self.order = list(range(n))
        if n <= 1:
            return
        mode = self.sort_mode_var.get()
        if mode == "random":
            random.shuffle(self.order)
        elif mode == "score_asc":
            self.order.sort(key=lambda i: (item_score(self.vocab[i]), i))
        elif mode == "score_desc":
            self.order.sort(key=lambda i: (-item_score(self.vocab[i]), i))

    def _current_item(self) -> dict | None:
        if not self.order or self.pos < 0 or self.pos >= len(self.order):
            return None
        idx = self.order[self.pos]
        if idx < 0 or idx >= len(self.vocab):
            return None
        return self.vocab[idx]

    @staticmethod
    def _speak_english_blocking(text: str, timeout_sec: float = 120.0) -> None:
        if not (text and str(text).strip()):
            return
        escaped = str(text).replace("'", "''")
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
                timeout=timeout_sec,
                check=False,
            )
        except Exception:
            pass

    def _maybe_speak_for_card(self, it: dict | None) -> None:
        if not it:
            return
        mode = self.read_mode_var.get()
        if mode == "none":
            return
        word = str(it.get("word", "")).strip()
        ex = str(it.get("example", "") or "").strip()
        self._speak_generation += 1
        gen = self._speak_generation

        def worker() -> None:
            if mode == "word":
                if gen != self._speak_generation:
                    return
                if word:
                    self._speak_english_blocking(word)
                return
            if mode == "word_example":
                if gen != self._speak_generation:
                    return
                if word:
                    self._speak_english_blocking(word, timeout_sec=60.0)
                if gen != self._speak_generation:
                    return
                if ex:
                    self._speak_english_blocking(ex, timeout_sec=120.0)

        threading.Thread(target=worker, daemon=True).start()

    def _apply_grade(self, grade: str) -> None:
        """grade: know | vague | unknown"""
        it = self._current_item()
        if not it:
            return
        revealed = bool(self.reveal_meaning)
        delta = GRADE_DELTA.get((grade, revealed))
        if delta is None:
            return
        old = item_score(it)
        new = max(SCORE_MIN, min(SCORE_MAX, old + delta))
        it["score"] = round(new, 1)
        it["reviews"] = int(it.get("reviews") or 0) + 1
        word = str(it.get("word", "")).strip()
        label = {"know": "认识", "vague": "模糊", "unknown": "不认识"}.get(grade, grade)
        hint = "已看释义/例句" if revealed else "未看释义/例句"
        try:
            save_vocab(self.vocab_path, self.vocab)
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return
        self._log(f"评分「{word}」{label}（{hint}）{old:.1f} → {new:.1f}（Δ{delta:+.1f}）")
        self._advance_after_grade()

    def _advance_after_grade(self) -> None:
        if not self.order:
            return
        n = len(self.order)
        cur_idx = self.order[self.pos]
        self._rebuild_order()
        try:
            new_pos = self.order.index(cur_idx)
        except ValueError:
            new_pos = 0
        self.pos = (new_pos + 1) % n
        self.reveal_meaning = False
        self._show_card()

    def _clear_example_display(self) -> None:
        self.example_text.configure(state="normal")
        self.example_text.delete("1.0", "end")
        self.example_text.configure(state="disabled")

    @staticmethod
    def _insert_text_with_keyword_bold(widget: tk.Text, body: str, keyword: str, tag: str = "keyword") -> None:
        keyword = (keyword or "").strip()
        body = body or ""
        if not keyword:
            widget.insert("end", body)
            return
        pattern = re.escape(keyword)
        try:
            parts = re.split(f"({pattern})", body, flags=re.IGNORECASE)
        except re.error:
            widget.insert("end", body)
            return
        kw_lower = keyword.lower()
        for part in parts:
            if not part:
                continue
            if part.lower() == kw_lower:
                widget.insert("end", part, (tag,))
            else:
                widget.insert("end", part)

    def _render_example_display(self, it: dict) -> None:
        word_kw = str(it.get("word", "")).strip()
        ex = it.get("example", "") or ""
        exs = str(ex).strip()
        zh = it.get("example_zh", "") or ""
        zhs = str(zh).strip()

        self.example_text.configure(state="normal")
        self.example_text.delete("1.0", "end")
        if not exs and not zhs:
            self.example_text.insert("end", "例句：（暂无）", ("muted",))
        else:
            self.example_text.insert("end", "例句（英）：")
            if exs:
                self._insert_text_with_keyword_bold(self.example_text, exs, word_kw)
            else:
                self.example_text.insert("end", "（暂无）", ("muted",))
            self.example_text.insert("end", "\n例句（译）：", ("zh_line",))
            self.example_text.insert("end", zhs if zhs else "（暂无）", ("zh_line",))
        self.example_text.configure(state="disabled")

    def _show_card(self) -> None:
        n = len(self.vocab)
        if n == 0:
            self.progress_var.set("0 / 0")
            self.score_var.set("熟练度 —")
            self.word_var.set("（词表为空或无法读取）")
            self.meaning_var.set("")
            self._clear_example_display()
            return
        self.progress_var.set(f"{self.pos + 1} / {len(self.order)}")
        it = self._current_item()
        if not it:
            return
        sc = item_score(it)
        rev = int(it.get("reviews") or 0)
        self.score_var.set(f"熟练度 {sc:.1f} / {SCORE_MAX:.0f}（已评 {rev} 次）")
        word = str(it.get("word", ""))
        self.word_var.set(word)
        self.reveal_meaning = False
        self.meaning_var.set("（点击「显示释义 / 例句」）")
        self._clear_example_display()
        self._maybe_speak_for_card(it)

    def _toggle_reveal(self) -> None:
        it = self._current_item()
        if not it:
            return
        self.reveal_meaning = not self.reveal_meaning
        if self.reveal_meaning:
            self.meaning_var.set(f"释义：{it.get('meaning', '')}")
            self._render_example_display(it)
        else:
            self.meaning_var.set("（已隐藏，再次点击显示）")
            self._clear_example_display()

    def _refresh_gen_button_label(self) -> None:
        empty_n = count_pending_examples(self.vocab)
        self.gen_btn.configure(text=f"用 DeepSeek 生成英例句 + 中译（约 {empty_n} 条待补全）")

    def _get_client(self):
        try:
            from openai import OpenAI
        except ImportError:
            messagebox.showerror(
                "缺少依赖",
                "请先安装：pip install openai\n然后重新运行本程序。",
            )
            return None
        key = self.api_key_var.get().strip() or (read_api_key_file(self.key_path) or "")
        if not key:
            messagebox.showwarning("提示", "请先填写 DeepSeek API Key，或写入 api_key.txt。")
            return None
        return OpenAI(api_key=key, base_url=DEEPSEEK_BASE_URL)

    def _call_example_bilingual(self, client, word: str, meaning: str) -> tuple[str, str]:
        user = (
            "为英语学习者写一句自然地道的英文例句，并给出这句英文的完整简体中文翻译（整句译文，不是只翻译词条）。\n"
            f"词条（可能是词或短语）：{word}\n"
            f"词条中文释义：{meaning}\n\n"
            "只输出一个 JSON 对象，不要 markdown 代码块，不要前缀或解释。\n"
            '格式严格为：{"example":"英文例句","example_zh":"例句的完整中文翻译"}\n'
            "自然、难度适合中高级学习者；若词条是短语，请在例句中自然使用该短语。\n"
        )
        resp = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": 'Reply with a single JSON object only, keys: "example" (English), "example_zh" (Chinese).',
                },
                {"role": "user", "content": user},
            ],
            stream=False,
        )
        raw = (resp.choices[0].message.content or "").strip()
        en, zh = parse_bilingual_response(raw)
        if not en or not zh:
            raise ValueError("模型未返回完整 example / example_zh")
        return en, zh

    def _start_generate_examples(self) -> None:
        if self._gen_running:
            return
        client = self._get_client()
        if client is None:
            return
        pending = [(i, it) for i, it in enumerate(self.vocab) if needs_bilingual_example(it)]
        if not pending:
            messagebox.showinfo("提示", "没有需要生成的例句（英文与中译均已齐全）。")
            return
        self._gen_running = True
        self.gen_btn.configure(state="disabled")
        self.status_var.set("正在请求 DeepSeek…")

        def worker() -> None:
            total = len(pending)
            ok = 0
            stop_reason: str | None = None
            for n, (idx, it) in enumerate(pending, start=1):
                if not self.root.winfo_exists():
                    break
                word = str(it.get("word", ""))
                meaning = str(it.get("meaning", ""))
                try:
                    ex, ex_zh = self._call_example_bilingual(client, word, meaning)
                    it["example"] = ex
                    it["example_zh"] = ex_zh
                    save_vocab(self.vocab_path, self.vocab)
                    ok += 1
                    self.root.after(0, lambda w=word, nn=n, t=total: self._log(f"[{nn}/{t}] OK：{w}"))
                except Exception as e:
                    if _is_insufficient_balance_error(e):
                        stop_reason = (
                            "DeepSeek 返回 402：账户余额不足（Insufficient Balance）。\n"
                            "请登录 DeepSeek 开放平台充值或更换有余额的 API Key；"
                            "这不是本程序 bug，未充值成功前批量生成会持续失败。"
                        )
                        hint = "已因余额不足中止，后续请求已跳过（避免无意义重试）。"
                        self.root.after(0, lambda h=hint: self._log(h))
                        break
                    self.root.after(0, lambda w=word, err=str(e), nn=n, t=total: self._log_fail(nn, t, w, err))
                time.sleep(0.35)
            self.root.after(
                0,
                lambda o=ok, tt=total, sr=stop_reason: self._gen_finished(o, tt, sr),
            )

        threading.Thread(target=worker, daemon=True).start()

    def _log_fail(self, n: int, t: int, w: str, err: str) -> None:
        self._log(f"[{n}/{t}] 失败：{w} — {err}")

    def _gen_finished(self, ok: int, total: int, stop_reason: str | None = None) -> None:
        self._gen_running = False
        self.gen_btn.configure(state="normal")
        self._refresh_gen_button_label()
        self.status_var.set(f"完成：成功 {ok} / 计划 {total}")
        self.reveal_meaning = False
        self._show_card()
        if stop_reason:
            messagebox.showwarning(
                "批量生成已中止",
                f"{stop_reason}\n\n已成功写入：{ok} / 本次计划：{total}",
            )
        else:
            messagebox.showinfo("完成", f"例句生成结束。\n成功写入：{ok} / 本次任务：{total}")

    def run(self) -> None:
        # 启动时若文件里有 key 而输入框空，同步到界面（界面已用 read 初始化，此处仅刷新按钮计数）
        self._refresh_gen_button_label()
        if not self.api_key_var.get().strip():
            file_key = read_api_key_file(self.key_path)
            if file_key:
                self.api_key_var.set(file_key)
        n = len(self.vocab)
        self._log(
            f"已加载 {self.vocab_path}，共 {n} 条；待生成/待补全（英或中译）约 {count_pending_examples(self.vocab)} 条。"
        )
        self.root.mainloop()


def main() -> None:
    app = VocabReviewApp()
    app.run()


if __name__ == "__main__":
    main()
