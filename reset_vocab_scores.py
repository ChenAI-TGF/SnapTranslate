"""
无界面：重置生词本 vocab.json 中每条词条的熟练度与复习次数。

默认处理脚本同目录下的 vocab.json；也可传入路径：
  python reset_vocab_scores.py
  python reset_vocab_scores.py D:\\path\\to\\vocab.json
"""
from __future__ import annotations

import json
import os
import sys

DEFAULT_SCORE = 50.0
DEFAULT_REVIEWS = 0


def _default_vocab_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "vocab.json")


def _load(path: str) -> list[dict]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError("vocab.json 顶层必须是数组")
    return [x for x in data if isinstance(x, dict)]


def _save(path: str, items: list[dict]) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def main() -> int:
    path = sys.argv[1].strip() if len(sys.argv) > 1 else _default_vocab_path()
    if not os.path.isfile(path):
        print(f"错误：找不到文件 {path}")
        return 1
    try:
        items = _load(path)
    except Exception as e:
        print(f"错误：无法读取 JSON — {e}")
        return 1

    n = 0
    for it in items:
        it["score"] = float(DEFAULT_SCORE)
        it["reviews"] = int(DEFAULT_REVIEWS)
        n += 1

    try:
        _save(path, items)
    except Exception as e:
        print(f"错误：写入失败 — {e}")
        return 1

    print(f"已重置 {n} 条：score={DEFAULT_SCORE:g}，reviews={DEFAULT_REVIEWS}")
    print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
