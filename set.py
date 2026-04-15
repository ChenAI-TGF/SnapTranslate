from __future__ import annotations

import json
import os
import tkinter as tk
from tkinter import font as tkfont, messagebox
from typing import Any

DEFAULT_SCORE = 50.0
DEFAULT_REVIEWS = 0
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_VOCAB_PATH = os.path.join(SCRIPT_DIR, "vocab.json")
DEFAULT_BACKUP_DIR = os.path.join(SCRIPT_DIR, "backups")


def load_vocab(path: str) -> list[dict[str, Any]]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError("vocab.json 顶层必须是数组")
    return [x for x in data if isinstance(x, dict)]


def save_vocab(path: str, items: list[dict[str, Any]]) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def count_with_example(items: list[dict[str, Any]]) -> int:
    n = 0
    for it in items:
        ex = str(it.get("example", "") or "").strip()
        ex_zh = str(it.get("example_zh", "") or "").strip()
        if ex and ex_zh:
            n += 1
    return n


def list_backups(backup_dir: str) -> list[str]:
    if not os.path.isdir(backup_dir):
        return []
    out: list[tuple[float, str]] = []
    for name in os.listdir(backup_dir):
        p = os.path.join(backup_dir, name)
        if os.path.isfile(p) and name.lower().endswith(".json"):
            out.append((os.path.getmtime(p), p))
    out.sort(key=lambda x: x[0], reverse=True)
    return [p for _, p in out]


def reset_scores(path: str) -> int:
    items = load_vocab(path)
    n = 0
    for it in items:
        it["score"] = float(DEFAULT_SCORE)
        it["reviews"] = int(DEFAULT_REVIEWS)
        n += 1
    save_vocab(path, items)
    return n


def cleanup_backups_keep_latest(backup_dir: str) -> tuple[int, str | None]:
    backups = list_backups(backup_dir)
    if len(backups) <= 1:
        return 0, backups[0] if backups else None
    keep = backups[0]
    removed = 0
    for p in backups[1:]:
        try:
            os.remove(p)
            removed += 1
        except Exception:
            continue
    return removed, keep


class AdminApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("词表后台管理功能")
        self.root.geometry("760x520")
        self.root.minsize(680, 460)
        self.vocab_path_var = tk.StringVar(master=self.root, value=DEFAULT_VOCAB_PATH)
        self.backup_dir_var = tk.StringVar(master=self.root, value=DEFAULT_BACKUP_DIR)
        self.total_var = tk.StringVar(master=self.root, value="-")
        self.with_example_var = tk.StringVar(master=self.root, value="-")
        self.pending_var = tk.StringVar(master=self.root, value="-")
        self.backup_count_var = tk.StringVar(master=self.root, value="-")
        self.latest_backup_var = tk.StringVar(master=self.root, value="暂无")
        self.status_var = tk.StringVar(master=self.root, value="就绪")

        self._build_ui()
        self.refresh_status()

    def _build_ui(self) -> None:
        f_title = tkfont.Font(family="Microsoft YaHei UI", size=18, weight="bold")
        f_label = tkfont.Font(family="Microsoft YaHei UI", size=10)
        f_value = tkfont.Font(family="Consolas", size=10)

        outer = tk.Frame(self.root, padx=16, pady=14)
        outer.pack(fill="both", expand=True)

        tk.Label(outer, text="词表后台管理功能", font=f_title).pack(anchor="w", pady=(0, 10))

        p1 = tk.Frame(outer)
        p1.pack(fill="x", pady=(0, 6))
        tk.Label(p1, text="词表路径", font=f_label, width=10, anchor="w").pack(side="left")
        tk.Entry(p1, textvariable=self.vocab_path_var, font=f_value).pack(side="left", fill="x", expand=True)

        p2 = tk.Frame(outer)
        p2.pack(fill="x", pady=(0, 10))
        tk.Label(p2, text="备份目录", font=f_label, width=10, anchor="w").pack(side="left")
        tk.Entry(p2, textvariable=self.backup_dir_var, font=f_value).pack(side="left", fill="x", expand=True)

        status_box = tk.LabelFrame(outer, text="状态")
        status_box.pack(fill="x", pady=(0, 12))
        tk.Label(status_box, textvariable=self.total_var, anchor="w").pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(status_box, textvariable=self.with_example_var, anchor="w").pack(fill="x", padx=10)
        tk.Label(status_box, textvariable=self.pending_var, anchor="w").pack(fill="x", padx=10)
        tk.Label(status_box, textvariable=self.backup_count_var, anchor="w").pack(fill="x", padx=10)
        tk.Label(status_box, textvariable=self.latest_backup_var, anchor="w", wraplength=700, justify="left").pack(
            fill="x", padx=10, pady=(0, 8)
        )

        btn_row = tk.Frame(outer)
        btn_row.pack(fill="x", pady=(0, 8))
        tk.Button(btn_row, text="刷新状态", command=self.refresh_status).pack(side="left")
        tk.Button(btn_row, text="1) 词表评分重置系统", command=self.on_reset_scores).pack(side="left", padx=(8, 0))
        tk.Button(btn_row, text="2) 删除备份（仅保留最新1个）", command=self.on_cleanup_backups).pack(side="left", padx=(8, 0))

        tk.Label(outer, textvariable=self.status_var, anchor="w", fg="#334155").pack(fill="x")

    def refresh_status(self) -> None:
        vocab_path = self.vocab_path_var.get().strip()
        backup_dir = self.backup_dir_var.get().strip()

        try:
            items = load_vocab(vocab_path)
            total = len(items)
            with_example = count_with_example(items)
            pending = total - with_example
            self.total_var.set(f"词表总数：{total}")
            self.with_example_var.set(f"有例句+翻译：{with_example}")
            self.pending_var.set(f"待补全例句：{pending}")
        except Exception as exc:
            self.total_var.set("词表总数：读取失败")
            self.with_example_var.set("有例句+翻译：读取失败")
            self.pending_var.set("待补全例句：读取失败")
            self.status_var.set(f"读取词表失败：{exc}")

        backups = list_backups(backup_dir)
        self.backup_count_var.set(f"备份文件数：{len(backups)}")
        if backups:
            self.latest_backup_var.set(f"最新备份：{os.path.basename(backups[0])}")
        else:
            self.latest_backup_var.set("最新备份：暂无")

    def on_reset_scores(self) -> None:
        vocab_path = self.vocab_path_var.get().strip()
        if not os.path.isfile(vocab_path):
            messagebox.showerror("错误", f"找不到词表文件：{vocab_path}")
            return
        if not messagebox.askyesno("确认", "确定重置所有词条评分和复习次数吗？"):
            return
        try:
            n = reset_scores(vocab_path)
            self.status_var.set(f"已重置 {n} 条：score={DEFAULT_SCORE:g}，reviews={DEFAULT_REVIEWS}")
            self.refresh_status()
        except Exception as exc:
            messagebox.showerror("失败", str(exc))

    def on_cleanup_backups(self) -> None:
        backup_dir = self.backup_dir_var.get().strip()
        if not messagebox.askyesno("确认", "确定删除旧备份，仅保留最新 1 个吗？"):
            return
        try:
            removed, keep = cleanup_backups_keep_latest(backup_dir)
            if keep:
                self.status_var.set(f"已删除 {removed} 个备份，保留：{os.path.basename(keep)}")
            else:
                self.status_var.set("没有可删除的备份文件。")
            self.refresh_status()
        except Exception as exc:
            messagebox.showerror("失败", str(exc))

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    AdminApp().run()


if __name__ == "__main__":
    main()
