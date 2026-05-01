"""
rename_gui.py — 批量重命名工具 图形界面

基于 Tkinter 构建，支持：
- 多步规则链（添加/编辑/删除/排序）
- 10 种重命名方式（含正则替换、修改扩展名）
- 按扩展名过滤文件
- 预览、执行、多步撤销
- 进度条与取消操作
- 键盘快捷键
- 名称碰撞检测
"""

from __future__ import annotations

import os
import re
import threading
import tkinter as tk
from dataclasses import dataclass, field
from tkinter import ttk, filedialog, messagebox
from typing import Callable, Optional

# ═══════════════════════════════════════════════════════════════════
# 核心：规则引擎 & 文件操作
# ═══════════════════════════════════════════════════════════════════

SYSTEM_FILES = {
    ".DS_Store", ".localized", ".Spotlight-V100", ".Trashes",
    ".fseventsd", ".TemporaryItems", ".apdisk", "Thumbs.db",
}

METHOD_LABELS: dict[str, str] = {
    "delete_after": "删除指定字符串及其后内容",
    "delete_before": "删除指定字符串及其前内容",
    "replace":      "替换字符串",
    "regex_replace":"正则替换",
    "add_prefix":   "添加前缀",
    "add_suffix":   "添加后缀",
    "remove_spaces":"删除空格",
    "to_lower":     "转换为小写",
    "to_upper":     "转换为大写",
    "change_ext":   "修改扩展名",
}

LABEL_TO_METHOD: dict[str, str] = {v: k for k, v in METHOD_LABELS.items()}

METHODS_ORDERED = [
    "delete_after", "delete_before", "replace", "regex_replace",
    "add_prefix", "add_suffix", "remove_spaces", "to_lower", "to_upper", "change_ext",
]

METHOD_PARAMS: dict[str, list[str]] = {
    "delete_after":   ["sep"],
    "delete_before":  ["sep"],
    "replace":        ["old", "new"],
    "regex_replace":  ["pattern", "repl"],
    "add_prefix":     ["prefix"],
    "add_suffix":     ["suffix"],
    "remove_spaces":  [],
    "to_lower":       [],
    "to_upper":       [],
    "change_ext":     ["ext"],
}

PARAM_LABELS: dict[str, str] = {
    "sep": "分隔字符串", "old": "查找", "new": "替换为",
    "pattern": "正则模式", "repl": "替换为",
    "prefix": "前缀", "suffix": "后缀", "ext": "新扩展名",
}


def is_system_item(name: str) -> bool:
    _, ext = os.path.splitext(name)
    return ext in SYSTEM_FILES or name in SYSTEM_FILES


@dataclass
class RenameRule:
    """一条重命名规则"""
    method: str
    params: dict = field(default_factory=dict)

    @property
    def label(self) -> str:
        labels: dict[str, str] = {
            "delete_after": "删除「{sep}」及其后",
            "delete_before": "删除「{sep}」及其前",
            "replace": "替换「{old}」→「{new}」",
            "regex_replace": "正则替换「{pattern}」→「{repl}」",
            "add_prefix": "添加前缀「{prefix}」",
            "add_suffix": "添加后缀「{suffix}」",
            "remove_spaces": "删除空格",
            "to_lower": "转换为小写",
            "to_upper": "转换为大写",
            "change_ext": "修改扩展名 → .{ext}",
        }
        tmpl = labels.get(self.method, self.method)
        try:
            return tmpl.format(**self.params)
        except KeyError:
            return tmpl

    def to_dict(self) -> dict:
        return {"method": self.method, "params": dict(self.params)}

    @classmethod
    def from_dict(cls, d: dict) -> "RenameRule":
        return cls(method=d["method"], params=dict(d.get("params", {})))


def apply_rule(name: str, rule: RenameRule) -> str:
    """对单个文件名应用一条规则"""
    method, params = rule.method, rule.params

    if method == "delete_after":
        sep = params.get("sep", "")
        idx = name.find(sep)
        return name[:idx] if idx != -1 else name
    if method == "delete_before":
        sep = params.get("sep", "")
        idx = name.find(sep)
        return name[idx + len(sep):] if idx != -1 else name
    if method == "replace":
        return name.replace(params.get("old", ""), params.get("new", ""))
    if method == "regex_replace":
        try:
            return re.sub(params.get("pattern", ""), params.get("repl", ""), name)
        except re.error:
            return name
    if method == "add_prefix":
        return params.get("prefix", "") + name
    if method == "add_suffix":
        suffix = params.get("suffix", "")
        name_part, ext_part = os.path.splitext(name)
        return name_part + suffix + ext_part
    if method == "remove_spaces":
        return name.replace(" ", "")
    if method == "to_lower":
        return name.lower()
    if method == "to_upper":
        return name.upper()
    if method == "change_ext":
        new_ext = params.get("ext", "")
        if not new_ext.startswith("."):
            new_ext = "." + new_ext
        return os.path.splitext(name)[0] + new_ext
    return name


def apply_rules(name: str, rules: list[RenameRule]) -> str:
    for rule in rules:
        name = apply_rule(name, rule)
    return name


# ── 过滤 ──

@dataclass
class FilterConfig:
    extensions: Optional[list[str]] = None
    name_include: Optional[str] = None
    name_exclude: Optional[str] = None


def _normalise_ext(ext: str) -> str:
    ext = ext.strip().lower()
    return ext if ext.startswith(".") else "." + ext


def should_process(name: str, filter_cfg: Optional[FilterConfig] = None) -> bool:
    if is_system_item(name):
        return False
    if filter_cfg is None:
        return True
    if filter_cfg.extensions is not None:
        _, ext = os.path.splitext(name)
        allowed = {_normalise_ext(e) for e in filter_cfg.extensions}
        if ext.lower() not in allowed:
            return False
    if filter_cfg.name_include:
        try:
            if not re.search(filter_cfg.name_include, name):
                return False
        except re.error:
            pass
    if filter_cfg.name_exclude:
        try:
            if re.search(filter_cfg.name_exclude, name):
                return False
        except re.error:
            pass
    return True


# ── 变更 ──

@dataclass
class Change:
    item_type: str
    old_name: str
    new_name: str
    full_path: str
    new_full_path: str


# ── 遍历 & 计算 ──

def _iter_top(target_dir: str):
    try:
        root, dirs, files = next(os.walk(target_dir))
        yield root, dirs, files
    except StopIteration:
        pass


def _iter_items(
    target_dir: str,
    item_type: str,
    recursive: bool = True,
    filter_cfg: Optional[FilterConfig] = None,
) -> list[tuple[str, str, str]]:
    """收集所有需要处理的项目。返回 [(类型, 父目录, 名称)]"""
    results: list[tuple[str, str, str]] = []

    # 文件夹（深优先）
    if item_type in ("folders", "both"):
        folder_entries: list[tuple[str, str]] = []
        if recursive:
            for root, dirs, _ in os.walk(target_dir, topdown=False):
                for d in dirs:
                    folder_entries.append((root, d))
        else:
            try:
                for entry in os.scandir(target_dir):
                    if entry.is_dir():
                        folder_entries.append((target_dir, entry.name))
            except PermissionError:
                pass
        for root, d in folder_entries:
            if should_process(d, filter_cfg):
                results.append(("文件夹", root, d))

    # 文件
    if item_type in ("files", "both"):
        walk = os.walk(target_dir) if recursive else _iter_top(target_dir)
        for root, _, files in walk:
            for f in files:
                if should_process(f, filter_cfg):
                    results.append(("文件", root, f))
    return results


def compute_changes(
    target_dir: str,
    rules: list[RenameRule],
    item_type: str = "files",
    filter_cfg: Optional[FilterConfig] = None,
    recursive: bool = True,
) -> tuple[list[Change], int]:
    """计算所有变更，返回 (changes, 碰撞跳过数)"""
    changes: list[Change] = []
    items = _iter_items(target_dir, item_type, recursive, filter_cfg)
    items.sort(key=lambda x: (x[1], x[2]))

    planned: set[str] = set()
    skipped = 0

    for type_label, root, name in items:
        new_name = apply_rules(name, rules)
        if new_name == name or not new_name.strip():
            continue
        new_path = os.path.join(root, new_name)
        if os.path.exists(new_path) or new_path in planned:
            skipped += 1
            continue
        planned.add(new_path)
        changes.append(Change(type_label, name, new_name,
                              os.path.join(root, name), new_path))
    return changes, skipped


def execute_changes(
    changes: list[Change],
    undo_stack: list[tuple[str, str]],
    cancel_event: Optional[Callable[[], bool]] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> tuple[int, int]:
    """执行重命名，返回 (成功, 失败)"""
    success = fail = 0
    for i, ch in enumerate(changes):
        if cancel_event and cancel_event():
            break
        try:
            os.rename(ch.full_path, ch.new_full_path)
            undo_stack.append((ch.new_full_path, ch.full_path))
            success += 1
        except OSError:
            fail += 1
        if progress_callback:
            progress_callback(i + 1, len(changes))
    return success, fail


def undo_changes(
    undo_stack: list[tuple[str, str]],
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> int:
    success = 0
    total = len(undo_stack)
    for i in range(total - 1, -1, -1):
        new_path, old_path = undo_stack[i]
        if os.path.exists(new_path):
            try:
                os.rename(new_path, old_path)
                success += 1
            except OSError:
                pass
        if progress_callback:
            progress_callback(total - i, total)
    undo_stack.clear()
    return success


# ═══════════════════════════════════════════════════════════════════
# GUI：规则编辑对话框
# ═══════════════════════════════════════════════════════════════════

class RuleDialog(tk.Toplevel):
    """添加 / 编辑一条规则的弹窗"""

    def __init__(self, parent, rule: RenameRule | None = None):
        super().__init__(parent)
        self.result: RenameRule | None = None
        self._param_vars: dict[str, tk.StringVar] = {}

        title = "编辑规则" if rule else "添加规则"
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)
        ttk.Label(top, text="方式:").pack(side=tk.LEFT)
        self._method_var = tk.StringVar(value=METHOD_LABELS[METHODS_ORDERED[0]])
        cb = ttk.Combobox(top, textvariable=self._method_var,
                          values=[METHOD_LABELS[m] for m in METHODS_ORDERED],
                          state="readonly", width=28)
        cb.pack(side=tk.LEFT, padx=5)
        cb.bind("<<ComboboxSelected>>", self._on_method_changed)

        self._param_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        self._param_frame.pack(fill=tk.X)

        btn_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="确定", command=self._on_ok).pack(
            side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side=tk.RIGHT)

        if rule:
            self._method_var.set(METHOD_LABELS.get(rule.method, rule.method))
            self._build_params()
            for k, v in rule.params.items():
                if k in self._param_vars:
                    self._param_vars[k].set(v)
        else:
            self._build_params()

        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2 - w // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2 - h // 2
        self.geometry(f"+{pw}+{ph}")

    @property
    def _method(self) -> str:
        return LABEL_TO_METHOD.get(self._method_var.get(), "replace")

    def _on_method_changed(self, *_):
        self._build_params()

    def _build_params(self):
        for w in self._param_frame.winfo_children():
            w.destroy()
        self._param_vars.clear()

        needed = METHOD_PARAMS.get(self._method, [])
        if not needed:
            ttk.Label(self._param_frame, text="（此方式无需参数）").pack(
                anchor=tk.W, pady=5)
            return

        for key in needed:
            row = ttk.Frame(self._param_frame)
            row.pack(fill=tk.X, pady=2)
            lbl = PARAM_LABELS.get(key, key) + ":"
            ttk.Label(row, text=lbl, width=10).pack(side=tk.LEFT)
            var = tk.StringVar()
            ttk.Entry(row, textvariable=var, width=30).pack(side=tk.LEFT, padx=5)
            self._param_vars[key] = var

    def _on_ok(self):
        method = self._method
        params = {k: v.get() for k, v in self._param_vars.items()}
        for key in METHOD_PARAMS.get(method, []):
            if not params.get(key, "").strip():
                messagebox.showwarning(
                    "提示", f"请填写「{PARAM_LABELS.get(key, key)}」", parent=self)
                return
        self.result = RenameRule(method=method, params=params)
        self.destroy()


# ═══════════════════════════════════════════════════════════════════
# GUI：主应用
# ═══════════════════════════════════════════════════════════════════

class RenameApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("批量重命名工具")
        self.root.geometry("900x700")
        self.root.minsize(700, 500)

        self.rules: list[RenameRule] = []
        self._changes_cache: list[Change] | None = None
        self._undo_stack: list[tuple[str, str]] = []
        self._cancel_event = threading.Event()

        self._build_ui()
        self._bind_shortcuts()

    # ── UI 构建 ──

    def _build_ui(self):
        self._build_top()
        self._build_rules()
        self._build_actions()
        self._build_progress()
        self._build_results()
        self._build_log()

        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self._status_var,
                  relief=tk.SUNKEN, anchor=tk.W, padding=(5, 2)).pack(
            fill=tk.X, side=tk.BOTTOM)

    def _build_top(self):
        frm = ttk.LabelFrame(self.root, text="目标与选项", padding=10)
        frm.pack(fill=tk.X, padx=10, pady=(10, 5))

        r1 = ttk.Frame(frm)
        r1.pack(fill=tk.X)
        ttk.Label(r1, text="目标目录:").pack(side=tk.LEFT)
        self._dir_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self._dir_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(r1, text="浏览...", command=self._browse_dir).pack(side=tk.LEFT)

        r2 = ttk.Frame(frm)
        r2.pack(fill=tk.X, pady=(8, 0))

        self._item_type = tk.StringVar(value="files")
        ttk.Label(r2, text="对象:").pack(side=tk.LEFT)
        ttk.Radiobutton(r2, text="文件", variable=self._item_type, value="files").pack(
            side=tk.LEFT, padx=(2, 10))
        ttk.Radiobutton(r2, text="文件夹", variable=self._item_type, value="folders").pack(
            side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(r2, text="文件和文件夹", variable=self._item_type, value="both").pack(
            side=tk.LEFT, padx=(0, 20))

        self._recursive = tk.BooleanVar(value=True)
        ttk.Checkbutton(r2, text="递归子目录", variable=self._recursive).pack(side=tk.LEFT)

        ttk.Label(r2, text="扩展名过滤:").pack(side=tk.LEFT, padx=(20, 0))
        self._ext_filter_var = tk.StringVar()
        ttk.Entry(r2, textvariable=self._ext_filter_var, width=18).pack(
            side=tk.LEFT, padx=5)
        ttk.Label(r2, text="逗号分隔，如 .mp4,.avi ；留空=全部",
                  foreground="gray").pack(side=tk.LEFT)

    def _build_rules(self):
        frm = ttk.LabelFrame(self.root, text="重命名规则链 (按顺序执行)", padding=10)
        frm.pack(fill=tk.X, padx=10, pady=5)

        bar = ttk.Frame(frm)
        bar.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(bar, text="＋ 添加规则", command=self._add_rule).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(bar, text="✎ 编辑", command=self._edit_rule).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(bar, text="✕ 删除", command=self._remove_rule).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(bar, text="▲ 上移", command=self._move_rule_up).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(bar, text="▼ 下移", command=self._move_rule_down).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(bar, text="清空", command=self._clear_rules).pack(side=tk.RIGHT)

        lf = ttk.Frame(frm)
        lf.pack(fill=tk.X)
        self._rule_listbox = tk.Listbox(lf, height=4, exportselection=False)
        scroll = ttk.Scrollbar(lf, orient=tk.VERTICAL,
                               command=self._rule_listbox.yview)
        self._rule_listbox.configure(yscrollcommand=scroll.set)
        self._rule_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._rule_listbox.bind("<Double-Button-1>", lambda e: self._edit_rule())

    def _build_actions(self):
        frm = ttk.Frame(self.root)
        frm.pack(fill=tk.X, padx=10, pady=5)

        self._preview_btn = ttk.Button(frm, text="预览 (Ctrl+P)", command=self._preview)
        self._preview_btn.pack(side=tk.LEFT, padx=(0, 10))

        self._execute_btn = ttk.Button(frm, text="执行重命名 (Ctrl+E)",
                                       command=self._execute)
        self._execute_btn.pack(side=tk.LEFT, padx=(0, 10))

        self._cancel_btn = ttk.Button(frm, text="取消 (Esc)", command=self._cancel,
                                      state=tk.DISABLED)
        self._cancel_btn.pack(side=tk.LEFT, padx=(0, 10))

        self._undo_btn = ttk.Button(frm, text="撤销 (Ctrl+Z)", command=self._undo)
        self._undo_btn.pack(side=tk.LEFT)

    def _build_progress(self):
        frm = ttk.Frame(self.root)
        frm.pack(fill=tk.X, padx=10, pady=(0, 5))
        self._progress_var = tk.IntVar(value=0)
        self._progress_bar = ttk.Progressbar(
            frm, variable=self._progress_var, mode="determinate", length=400)
        self._progress_bar.pack(fill=tk.X)
        self._progress_bar.pack_forget()

    def _build_results(self):
        frm = ttk.LabelFrame(self.root, text="预览 / 结果", padding=10)
        frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("status", "type", "old_name", "new_name")
        self._tree = ttk.Treeview(frm, columns=cols, show="headings", height=10)
        self._tree.heading("status", text="", anchor=tk.CENTER)
        self._tree.heading("type", text="类型", anchor=tk.CENTER)
        self._tree.heading("old_name", text="原名")
        self._tree.heading("new_name", text="新名")
        self._tree.column("status", width=36, anchor=tk.CENTER)
        self._tree.column("type", width=56, anchor=tk.CENTER)
        self._tree.column("old_name", width=300)
        self._tree.column("new_name", width=300)

        scroll = ttk.Scrollbar(frm, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=scroll.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._tree_menu = tk.Menu(self.root, tearoff=0)
        self._tree_menu.add_command(label="复制原名", command=self._copy_old_name)
        self._tree_menu.add_command(label="复制新名", command=self._copy_new_name)
        self._tree_menu.add_command(label="复制所在路径", command=self._copy_path)
        self._tree.bind("<Button-2>" if os.name != "nt" else "<Button-3>",
                        self._tree_popup)

    def _build_log(self):
        frm = ttk.LabelFrame(self.root, text="日志", padding=5)
        frm.pack(fill=tk.X, padx=10, pady=(0, 5))
        self._log_text = tk.Text(frm, height=4, state=tk.DISABLED, wrap=tk.WORD)
        scroll = ttk.Scrollbar(frm, orient=tk.VERTICAL, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=scroll.set)
        self._log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ── 快捷键 ──

    def _bind_shortcuts(self):
        r = self.root
        r.bind("<Control-o>", lambda e: self._browse_dir())
        r.bind("<Control-p>", lambda e: self._preview())
        r.bind("<Control-e>", lambda e: self._execute())
        r.bind("<Control-r>", lambda e: self._execute())
        r.bind("<Control-z>", lambda e: self._undo())
        r.bind("<Control-n>", lambda e: self._add_rule())
        r.bind("<Escape>", lambda e: self._cancel())
        r.bind("<Delete>", lambda e: self._remove_rule())

    # ── 规则操作 ──

    def _refresh_rule_list(self):
        self._rule_listbox.delete(0, tk.END)
        for i, rule in enumerate(self.rules, 1):
            self._rule_listbox.insert(tk.END, f"{i}. {rule.label}")
        self._changes_cache = None

    def _add_rule(self):
        dlg = RuleDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            self.rules.append(dlg.result)
            self._refresh_rule_list()
            self._log(f"已添加规则: {dlg.result.label}")

    def _edit_rule(self):
        sel = self._rule_listbox.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先选择一条规则")
            return
        idx = sel[0]
        dlg = RuleDialog(self.root, self.rules[idx])
        self.root.wait_window(dlg)
        if dlg.result:
            self.rules[idx] = dlg.result
            self._refresh_rule_list()
            self._log(f"已更新规则: {dlg.result.label}")

    def _remove_rule(self):
        sel = self._rule_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        removed = self.rules.pop(idx)
        self._refresh_rule_list()
        self._log(f"已删除规则: {removed.label}")

    def _move_rule_up(self):
        sel = self._rule_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self.rules[idx], self.rules[idx - 1] = self.rules[idx - 1], self.rules[idx]
        self._refresh_rule_list()
        self._rule_listbox.selection_set(idx - 1)

    def _move_rule_down(self):
        sel = self._rule_listbox.curselection()
        if not sel or sel[0] >= len(self.rules) - 1:
            return
        idx = sel[0]
        self.rules[idx], self.rules[idx + 1] = self.rules[idx + 1], self.rules[idx]
        self._refresh_rule_list()
        self._rule_listbox.selection_set(idx + 1)

    def _clear_rules(self):
        if not self.rules:
            return
        if messagebox.askyesno("确认", "确定要清空所有规则吗？"):
            self.rules.clear()
            self._refresh_rule_list()
            self._log("已清空所有规则")

    # ── 过滤 ──

    def _get_filter(self) -> FilterConfig:
        cfg = FilterConfig()
        ext_str = self._ext_filter_var.get().strip()
        if ext_str:
            cfg.extensions = [e.strip() for e in ext_str.split(",") if e.strip()]
        return cfg

    # ── 预览 ──

    def _preview(self):
        self._tree.delete(*self._tree.get_children())

        target = self._dir_var.get().strip()
        if not target:
            messagebox.showwarning("提示", "请先选择目标目录")
            return
        if not os.path.isdir(target):
            messagebox.showwarning("提示", "目标目录不存在")
            return
        if not self.rules:
            messagebox.showwarning("提示", "请先添加至少一条规则")
            return

        changes, skipped = compute_changes(
            target_dir=target, rules=self.rules,
            item_type=self._item_type.get(),
            filter_cfg=self._get_filter(),
            recursive=self._recursive.get(),
        )
        self._changes_cache = changes

        if not changes:
            self._log("没有发现需要重命名的项目")
            if skipped:
                self._log(f"  (另有 {skipped} 个项目因名称冲突被跳过)")
            self._status_var.set("预览完成: 0 个变更")
            return

        for ch in changes:
            self._tree.insert("", tk.END, values=("→", ch.item_type, ch.old_name, ch.new_name))

        msg = f"预览: 发现 {len(changes)} 个可重命名项目"
        if skipped:
            msg += f", {skipped} 个因名称冲突被跳过"
        self._log(msg)
        self._status_var.set(f"预览完成: {len(changes)} 个变更")

    # ── 执行 ──

    def _execute(self):
        target = self._dir_var.get().strip()
        if not target or not os.path.isdir(target):
            messagebox.showwarning("提示", "请先选择有效目录")
            return

        if self._changes_cache is not None:
            changes = self._changes_cache
        else:
            changes, _ = compute_changes(
                target_dir=target, rules=self.rules,
                item_type=self._item_type.get(),
                filter_cfg=self._get_filter(),
                recursive=self._recursive.get(),
            )
        self._changes_cache = changes

        if not changes:
            self._log("没有需要重命名的项目")
            return

        if not messagebox.askyesno(
            "确认",
            f"即将重命名 {len(changes)} 个项目，此操作不可逆。\n\n"
            "可在执行后按 Ctrl+Z 撤销。\n\n确定要继续吗？",
        ):
            return

        self._log(f"开始重命名 {len(changes)} 个项目...")
        self._status_var.set("正在重命名...")

        self._progress_var.set(0)
        self._progress_bar.configure(maximum=len(changes))
        self._progress_bar.pack(fill=tk.X)

        self._set_busy(True)
        self._cancel_event.clear()

        def worker():
            def on_progress(cur, tot):
                self.root.after(0, lambda c=cur: self._progress_var.set(c))

            success, fail = execute_changes(
                changes=changes, undo_stack=self._undo_stack,
                cancel_event=lambda: self._cancel_event.is_set(),
                progress_callback=on_progress,
            )
            for i, ch in enumerate(changes):
                if i >= success:
                    st = "✗" if i < success + fail else "—"
                else:
                    st = "✓"
                self.root.after(0, lambda s=st, c=ch: self._update_tree(s, c))
            self.root.after(0, lambda: self._on_done(success, fail))

        threading.Thread(target=worker, daemon=True).start()

    def _update_tree(self, status: str, change: Change):
        for item in self._tree.get_children():
            vals = self._tree.item(item, "values")
            if vals[2] == change.old_name and vals[3] == change.new_name:
                self._tree.item(item, values=(status, vals[1], vals[2], vals[3]))
                break

    def _on_done(self, success: int, fail: int):
        self._set_busy(False)
        self._progress_bar.pack_forget()
        self._log(f"\n完成: 成功 {success} 个, 失败 {fail} 个")
        self._status_var.set(
            f"完成: 成功 {success}, 失败 {fail}   —   可按 Ctrl+Z 撤销")
        self._changes_cache = None

    def _cancel(self):
        if self._cancel_btn["state"] == tk.NORMAL:
            self._cancel_event.set()
            self._log("正在取消...")
            self._status_var.set("取消中...")

    def _set_busy(self, busy: bool):
        st = tk.DISABLED if busy else tk.NORMAL
        self._preview_btn.configure(state=st)
        self._execute_btn.configure(state=st)
        self._undo_btn.configure(state=st)
        self._cancel_btn.configure(state=tk.NORMAL if busy else tk.DISABLED)

    # ── 撤销 ──

    def _undo(self):
        if not self._undo_stack:
            messagebox.showinfo("提示", "没有可撤销的操作")
            return

        count = len(self._undo_stack)
        if not messagebox.askyesno("确认",
                                   f"即将撤销 {count} 个重命名操作。\n\n确定要继续吗？"):
            return

        self._log(f"开始撤销 {count} 个操作...")

        def on_progress(cur, tot):
            self.root.after(0, lambda c=cur: self._progress_var.set(c))

        self._progress_var.set(0)
        self._progress_bar.configure(maximum=count)
        self._progress_bar.pack(fill=tk.X)

        success = undo_changes(self._undo_stack, on_progress)

        self._progress_bar.pack_forget()
        self._tree.delete(*self._tree.get_children())
        self._log(f"已撤销 {success} 个操作")
        self._status_var.set(f"已撤销 {success} 个操作")

    # ── 辅助 ──

    def _browse_dir(self):
        d = filedialog.askdirectory(title="选择目标目录")
        if d:
            self._dir_var.set(d)

    def _log(self, msg: str):
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.insert(tk.END, msg + "\n")
        self._log_text.see(tk.END)
        self._log_text.configure(state=tk.DISABLED)

    def _tree_popup(self, event):
        sel = self._tree.selection()
        if sel:
            self._tree.selection_set(sel)
            self._tree_menu.tk_popup(event.x_root, event.y_root)

    def _copy_old_name(self):
        self._copy_tree_col(2)

    def _copy_new_name(self):
        self._copy_tree_col(3)

    def _copy_path(self):
        sel = self._tree.selection()
        if sel:
            vals = self._tree.item(sel[0], "values")
            target = self._dir_var.get().strip()
            full = os.path.join(target, vals[2])
            self.root.clipboard_clear()
            self.root.clipboard_append(full)

    def _copy_tree_col(self, col: int):
        sel = self._tree.selection()
        if sel:
            val = self._tree.item(sel[0], "values")[col]
            self.root.clipboard_clear()
            self.root.clipboard_append(val)


# ═══════════════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    RenameApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
