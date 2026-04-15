from __future__ import annotations

import json
import os
import random
import re
import shutil
import time
from typing import Any

import streamlit as st
import streamlit.components.v1 as components

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_VOCAB = os.path.join(SCRIPT_DIR, "vocab.json")
DEFAULT_KEY_FILE = os.path.join(SCRIPT_DIR, "api_key.txt")
BACKUP_DIR = os.path.join(SCRIPT_DIR, "backups")
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-chat"

DEFAULT_SCORE = 50.0
SCORE_MIN = 0.0
SCORE_MAX = 100.0
GRADE_DELTA: dict[tuple[str, bool], float] = {
    ("know", False): 10.0,
    ("vague", False): -4.0,
    ("unknown", False): -8.0,
    ("know", True): 5.0,
    ("vague", True): -7.0,
    ("unknown", True): -12.0,
}


def item_score(it: dict[str, Any]) -> float:
    s = it.get("score")
    if s is None:
        v = DEFAULT_SCORE
    else:
        try:
            v = float(s)
        except (TypeError, ValueError):
            v = DEFAULT_SCORE
    return max(SCORE_MIN, min(SCORE_MAX, v))


def normalize_vocab_scores(items: list[dict[str, Any]]) -> None:
    for it in items:
        it["score"] = item_score(it)


def load_vocab(path: str) -> list[dict[str, Any]]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            return []
        return [x for x in data if isinstance(x, dict)]
    except Exception:
        return []


def save_vocab(path: str, items: list[dict[str, Any]]) -> None:
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
    except Exception:
        return None
    return None


def write_api_key_file(key_path: str, key: str) -> None:
    tmp = key_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(key.strip() + "\n")
    os.replace(tmp, key_path)


def needs_bilingual_example(it: dict[str, Any]) -> bool:
    ex = str(it.get("example", "") or "").strip()
    zh = str(it.get("example_zh", "") or "").strip()
    return (not ex) or (not zh)


def count_pending_examples(items: list[dict[str, Any]]) -> int:
    return sum(1 for it in items if needs_bilingual_example(it))


def parse_bilingual_response(raw: str) -> tuple[str, str]:
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


def is_insufficient_balance_error(exc: BaseException) -> bool:
    code = getattr(exc, "status_code", None)
    if code == 402:
        return True
    msg = str(exc).lower()
    return "insufficient balance" in msg or ("402" in msg and "balance" in msg)


def call_example_bilingual(client: Any, word: str, meaning: str) -> tuple[str, str]:
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


def backup_vocab(path: str, count: int) -> tuple[bool, str]:
    if not os.path.isfile(path):
        return False, "未找到词表文件，跳过备份"
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = time.strftime("%Y-%m-%d_%H-%M-%S")
    name = f"vocab_backup_{stamp}_entries-{count}.json"
    out = os.path.join(BACKUP_DIR, name)
    try:
        shutil.copy2(path, out)
    except Exception as exc:
        return False, f"备份失败：{exc}"
    return True, out


def rebuild_order(items: list[dict[str, Any]], mode: str) -> list[int]:
    order = list(range(len(items)))
    if len(order) <= 1:
        return order
    if mode == "random":
        random.shuffle(order)
    elif mode == "score_asc":
        order.sort(key=lambda i: (item_score(items[i]), i))
    elif mode == "score_desc":
        order.sort(key=lambda i: (-item_score(items[i]), i))
    return order


def current_item() -> dict[str, Any] | None:
    items = st.session_state.vocab
    order = st.session_state.order
    pos = st.session_state.pos
    if not order or pos < 0 or pos >= len(order):
        return None
    idx = order[pos]
    if idx < 0 or idx >= len(items):
        return None
    return items[idx]


def reset_for_new_vocab(path: str) -> None:
    items = load_vocab(path)
    normalize_vocab_scores(items)
    st.session_state.vocab_path = path
    st.session_state.vocab = items
    st.session_state.pos = 0
    st.session_state.show_meaning = False
    st.session_state.show_example = False
    st.session_state.show_example_zh = False
    st.session_state.order = rebuild_order(items, st.session_state.sort_mode)
    st.session_state.boot_backup_done = False


def apply_grade(grade: str) -> None:
    it = current_item()
    if not it:
        return
    revealed = bool(
        st.session_state.show_meaning
        or st.session_state.show_example
        or st.session_state.show_example_zh
    )
    delta = GRADE_DELTA.get((grade, revealed))
    if delta is None:
        return
    old = item_score(it)
    new = max(SCORE_MIN, min(SCORE_MAX, old + delta))
    it["score"] = round(new, 1)
    it["reviews"] = int(it.get("reviews") or 0) + 1
    cur_idx = st.session_state.order[st.session_state.pos]
    save_vocab(st.session_state.vocab_path, st.session_state.vocab)
    st.session_state.order = rebuild_order(st.session_state.vocab, st.session_state.sort_mode)
    if st.session_state.order:
        try:
            new_pos = st.session_state.order.index(cur_idx)
        except ValueError:
            new_pos = 0
        st.session_state.pos = (new_pos + 1) % len(st.session_state.order)
    else:
        st.session_state.pos = 0
    st.session_state.show_meaning = False
    st.session_state.show_example = False
    st.session_state.show_example_zh = False
    st.session_state.last_msg = f"评分已保存：{old:.1f} -> {new:.1f}（Δ{delta:+.1f}）"


def init_state() -> None:
    if "vocab_path" not in st.session_state:
        st.session_state.vocab_path = DEFAULT_VOCAB
    if "key_path" not in st.session_state:
        st.session_state.key_path = DEFAULT_KEY_FILE
    if "sort_mode" not in st.session_state:
        st.session_state.sort_mode = "random"
    if "show_meaning" not in st.session_state:
        st.session_state.show_meaning = False
    if "show_example" not in st.session_state:
        st.session_state.show_example = False
    if "show_example_zh" not in st.session_state:
        st.session_state.show_example_zh = False
    if "pos" not in st.session_state:
        st.session_state.pos = 0
    if "vocab" not in st.session_state:
        st.session_state.vocab = load_vocab(st.session_state.vocab_path)
        normalize_vocab_scores(st.session_state.vocab)
    if "order" not in st.session_state:
        st.session_state.order = rebuild_order(st.session_state.vocab, st.session_state.sort_mode)
    if "last_msg" not in st.session_state:
        st.session_state.last_msg = ""
    if "boot_backup_done" not in st.session_state:
        st.session_state.boot_backup_done = False
    if "read_mode" not in st.session_state:
        st.session_state.read_mode = "none"
    if "last_auto_speak_token" not in st.session_state:
        st.session_state.last_auto_speak_token = ""
    if "pending_speak_example" not in st.session_state:
        st.session_state.pending_speak_example = False


def get_client():
    try:
        from openai import OpenAI
    except ImportError:
        st.error("缺少依赖：请先安装 openai。")
        return None
    key = read_api_key_file(st.session_state.key_path) or ""
    if not key:
        key = st.session_state.get("manual_api_key", "").strip()
        if key:
            try:
                write_api_key_file(st.session_state.key_path, key)
            except Exception as exc:
                st.error(f"保存 API Key 失败：{exc}")
                return None
        else:
            st.warning("未检测到 API Key，请在侧边栏填写后重试。")
            return None
    return OpenAI(api_key=key, base_url=DEEPSEEK_BASE_URL)


def generate_examples() -> None:
    items = st.session_state.vocab
    pending = [(i, it) for i, it in enumerate(items) if needs_bilingual_example(it)]
    if not pending:
        st.info("没有需要生成的条目。")
        return
    client = get_client()
    if client is None:
        return
    ok = 0
    total = len(pending)
    bar = st.progress(0)
    status = st.empty()
    for n, (idx, it) in enumerate(pending, start=1):
        word = str(it.get("word", "")).strip()
        meaning = str(it.get("meaning", "")).strip()
        status.write(f"[{n}/{total}] 生成中：{word}")
        try:
            en, zh = call_example_bilingual(client, word, meaning)
            st.session_state.vocab[idx]["example"] = en
            st.session_state.vocab[idx]["example_zh"] = zh
            ok += 1
        except Exception as exc:
            if is_insufficient_balance_error(exc):
                st.warning("检测到余额不足（402），已中止后续生成。")
                break
            st.error(f"[{n}/{total}] 失败：{word} — {exc}")
        bar.progress(n / total)
        time.sleep(0.2)
    save_vocab(st.session_state.vocab_path, st.session_state.vocab)
    status.write(f"完成：成功 {ok} / 计划 {total}")
    st.session_state.last_msg = f"批量生成完成：成功 {ok} / 计划 {total}"


def _speak_text_once_lang(text: str, lang: str = "en-US") -> None:
    safe = json.dumps(text or "")
    lang_safe = json.dumps(lang or "")
    components.html(
        f"""
        <script>
        const txt = {safe};
        const lang = {lang_safe};
        if (window.speechSynthesis) {{
            window.speechSynthesis.cancel();
            const u = new SpeechSynthesisUtterance(txt);
            if (lang) {{
                u.lang = lang;
            }}
            window.speechSynthesis.speak(u);
        }}
        </script>
        """,
        height=0,
    )


def _auto_speak_once_for_card(it: dict[str, Any], card_idx: int) -> None:
    mode = st.session_state.get("read_mode", "none")
    if mode == "none":
        return
    token = f"{card_idx}:{mode}"
    if st.session_state.get("last_auto_speak_token", "") == token:
        return
    word = str(it.get("word", "") or "").strip()
    if mode == "word":
        if word:
            _speak_text_once_lang(word, lang="en-US")
    elif mode == "word_example":
        if word:
            _speak_text_once_lang(word, lang="en-US")
    st.session_state.last_auto_speak_token = token


def main() -> None:
    st.set_page_config(page_title="SnapTranslate 生词复习", page_icon="📘", layout="centered")
    st.title("📘 SnapTranslate")

    init_state()

    with st.sidebar:
        st.subheader("设置")
        vocab_path_in = st.text_input("词表路径", value=st.session_state.vocab_path)
        if st.button("加载词表", use_container_width=True):
            reset_for_new_vocab(vocab_path_in.strip() or DEFAULT_VOCAB)
            st.rerun()

        sort_mode = st.selectbox(
            "复习顺序",
            options=["random", "score_asc", "score_desc"],
            format_func=lambda x: {"random": "随机", "score_asc": "得分低→高", "score_desc": "得分高→低"}[x],
            index=["random", "score_asc", "score_desc"].index(st.session_state.sort_mode),
        )
        if sort_mode != st.session_state.sort_mode:
            st.session_state.sort_mode = sort_mode
            st.session_state.order = rebuild_order(st.session_state.vocab, sort_mode)
            st.session_state.pos = 0
            st.session_state.show_meaning = False
            st.session_state.show_example = False
            st.session_state.show_example_zh = False
            st.session_state.last_auto_speak_token = ""
            st.rerun()

        st.divider()
        st.subheader("DeepSeek（可选）")
        key_path_in = st.text_input("API Key 文件路径", value=st.session_state.key_path)
        st.session_state.key_path = key_path_in.strip() or DEFAULT_KEY_FILE
        st.text_input("手动填写 API Key（仅缺少本地文件时）", key="manual_api_key", type="password")

    if not st.session_state.boot_backup_done:
        ok, msg = backup_vocab(st.session_state.vocab_path, len(st.session_state.vocab))
        st.session_state.boot_backup_done = True
        st.session_state.last_msg = f"启动备份：{msg}" if ok else msg

    total_n = len(st.session_state.vocab)
    pending_n = count_pending_examples(st.session_state.vocab)
    it = current_item()
    if it is None:
        st.warning("词表为空。")
        return
    word_text = str(it.get("word", "")).strip()
    st.markdown(f"<h1 style='margin-bottom:0.2rem;'>{word_text}</h1>", unsafe_allow_html=True)
    st.write(f"熟练度：**{item_score(it):.1f}**")
    _auto_speak_once_for_card(it, st.session_state.pos + 1)
    ex_text = str(it.get("example", "") or "").strip()
    if st.session_state.pending_speak_example:
        if ex_text:
            _speak_text_once_lang(ex_text, lang="en-US")
        st.session_state.pending_speak_example = False

    t1, t2 = st.columns(2)
    with t1:
        if st.button("显示 / 隐藏 释义", use_container_width=True):
            st.session_state.show_meaning = not st.session_state.show_meaning
            st.rerun()
    with t2:
        if st.button("显示 / 隐藏 例句", use_container_width=True):
            turning_on = not st.session_state.show_example
            st.session_state.show_example = not st.session_state.show_example
            if not st.session_state.show_example:
                st.session_state.show_example_zh = False
                st.session_state.pending_speak_example = False
            elif turning_on and st.session_state.read_mode == "word_example":
                st.session_state.pending_speak_example = True
            st.rerun()

    if st.session_state.show_example:
        if st.button("显示 / 隐藏 例句翻译", use_container_width=True):
            st.session_state.show_example_zh = not st.session_state.show_example_zh
            st.rerun()

    if st.session_state.show_meaning:
        st.success(f"释义：{it.get('meaning', '')}")
    if st.session_state.show_example:
        ex = str(it.get("example", "") or "").strip()
        st.write("**例句（英）**")
        st.write(ex if ex else "（暂无）")
        if st.session_state.show_example_zh:
            ex_zh = str(it.get("example_zh", "") or "").strip()
            st.write("**例句（译）**")
            st.write(ex_zh if ex_zh else "（暂无）")

    if (
        not st.session_state.show_meaning
        and not st.session_state.show_example
        and not st.session_state.show_example_zh
    ):
        st.caption("点击按钮后显示释义/例句。")

    st.divider()
    b1, b2, b3 = st.columns(3)
    with b1:
        if st.button("认识", use_container_width=True):
            apply_grade("know")
            st.rerun()
    with b2:
        if st.button("模糊", use_container_width=True):
            apply_grade("vague")
            st.rerun()
    with b3:
        if st.button("不认识", use_container_width=True):
            apply_grade("unknown")
            st.rerun()

    st.divider()
    read_mode_main = st.selectbox(
        "自动朗读模式",
        options=["none", "word", "word_example"],
        format_func=lambda x: {"none": "不朗读", "word": "只朗读单词", "word_example": "朗读单词+例句（显示例句后才读例句）"}[x],
        index=["none", "word", "word_example"].index(st.session_state.read_mode),
        key="read_mode_main",
    )
    if read_mode_main != st.session_state.read_mode:
        st.session_state.read_mode = read_mode_main
        st.session_state.last_auto_speak_token = ""
        st.rerun()

    s1, s2 = st.columns(2)
    with s1:
        if st.button("朗读单词", use_container_width=True):
            _speak_text_once_lang(str(it.get("word", "")).strip(), lang="en-US")
    with s2:
        if st.button("朗读例句", use_container_width=True, disabled=not bool(ex_text)):
            _speak_text_once_lang(ex_text, lang="en-US")

    st.divider()
    st.write(f"词表：`{st.session_state.vocab_path}`")
    st.write(f"条目总数：**{total_n}**，待补全例句：**{pending_n}**")
    if st.session_state.last_msg:
        st.info(st.session_state.last_msg)

    c1, c2 = st.columns(2)
    with c1:
        if st.button(f"用 DeepSeek 生成英例句+中译（约 {pending_n} 条）", use_container_width=True):
            generate_examples()
            st.rerun()
    with c2:
        if st.button("刷新读取磁盘词表", use_container_width=True):
            reset_for_new_vocab(st.session_state.vocab_path)
            st.rerun()


if __name__ == "__main__":
    main()
