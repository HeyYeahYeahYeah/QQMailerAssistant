"""图形界面：选表格、填发件信息、预览名单、发送并查看结果。"""

from __future__ import annotations

import os
import queue
import re
import sys
import tempfile
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Callable
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from . import APP_TITLE, __version__
from .config import AppConfig, config_path
from .excel_io import (
    ExcelError,
    ParseResult,
    normalize_suffix,
    parse_workbook,
    read_sheet_names,
    write_report,
)
from .mailer import (
    QQ_HOST,
    QQ_PORT_SSL,
    MailerSettings,
    Sender,
    test_connection,
)

FONT_UI = ("Microsoft YaHei UI", 10)
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")
FONT_SMALL = ("Microsoft YaHei UI", 9)
COLOR_MUTED = "#666666"
COLOR_OK = "#1a7f37"
COLOR_WARN = "#b26a00"
COLOR_ERROR = "#c62828"

HELP_TEXT = """使用步骤

1. 准备 Excel：A 列放 QQ 号（纯数字），B 列放要发送的文字内容，第一行可以是标题。
2. 点「浏览…」选表格，程序会自动读取并在「名单预览」里列出每一封的收件邮箱。
3. 填「发件邮箱」和「SMTP 授权码」。
   授权码获取：QQ 邮箱网页版 → 设置 → 账号 → 开启「POP3/SMTP 服务」→ 按提示生成授权码。
   注意：授权码不是 QQ 登录密码。
4. 先点「测试登录」确认能连上，再点「开始发送」。
5. 发送完成后，表格同目录下会生成「发送结果_原文件名_时间.xlsx」，逐行记录成功/失败原因。

小提示

· 发送间隔建议 5–10 秒，太快容易被腾讯限流。
· 想先试一遍不真发，可以勾选「试运行」。
· 勾选「记住授权码」后，授权码用 Windows 本机加密保存，只有当前电脑账户能解开。
· 发送中可以「暂停」「停止」；停止后剩余行会标记为「未发送」。
· 主题支持占位符：{qq} {email} {序号} {总数} {日期} {内容}
· 改动了「收件邮箱后缀」或「重复 QQ」选项后，点「重新解析」刷新名单。

常见问题

· 提示登录失败 / 535：授权码不对、没开启 SMTP 服务，或邮箱地址写错。
· 提示发送过于频繁 / exceed：把间隔调大，隔几小时再发。
· 提示收件地址被拒绝：对方 QQ 邮箱可能未开通或地址不存在。
· 公司网络连不上 465 端口：程序会自动尝试 587 端口，仍失败请换网络。
"""


class App(tk.Tk):
    POLL_MS = 120

    def __init__(self) -> None:
        super().__init__()
        self.cfg = AppConfig.load()
        self.parse_result: ParseResult | None = None
        self._parse_key: tuple | None = None
        self.sender: Sender | None = None
        self.worker: threading.Thread | None = None
        self.events: queue.Queue[dict] = queue.Queue()
        self._last_result_dir: Path | None = None
        self._log_lines: list[str] = []
        self._run_started = 0.0
        self._config_path: Path | None = None

        self.title(f"{APP_TITLE} v{__version__}")
        self.minsize(900, 600)
        self._init_style()
        self._make_vars()
        self._build_ui()
        self._apply_config()
        self.after(self.POLL_MS, self._poll)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._log("欢迎使用。先选 Excel 表格，再填写发件邮箱与授权码。", "muted")
        if self.cfg.last_excel and Path(self.cfg.last_excel).is_file():
            self.after(300, lambda: self.on_parse(silent=True))

    # ------------------------------------------------------------------ #
    # 外观
    # ------------------------------------------------------------------ #
    def _init_style(self) -> None:
        style = ttk.Style(self)
        for theme in ("vista", "winnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        self.option_add("*Font", FONT_UI)
        style.configure(".", font=FONT_UI)
        style.configure("Treeview", font=FONT_UI, rowheight=26)
        style.configure("Treeview.Heading", font=FONT_BOLD)
        style.configure("TLabelframe.Label", font=FONT_BOLD)
        style.configure("Hint.TLabel", foreground=COLOR_MUTED, font=FONT_SMALL)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=(12, 8, 12, 8))
        outer.pack(fill="both", expand=True)
        # 先占住底部操作栏（side=bottom），这样窗口再矮也不会把按钮挤出可视区域
        self._build_bottom(outer)
        self._build_file_box(outer)
        self._build_settings_box(outer)
        self._build_notebook(outer)

    def _build_file_box(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(parent, text=" 第 1 步 · 选择表格 ", padding=(10, 8))
        box.pack(fill="x")
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="Excel 文件").grid(row=0, column=0, sticky="w")
        self.ent_excel = ttk.Entry(box, textvariable=self.var_excel, state="readonly")
        self.ent_excel.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        self.btn_browse = ttk.Button(box, text="浏览…", width=8, command=self.on_browse)
        self.btn_browse.grid(row=0, column=2)
        self.btn_reparse = ttk.Button(box, text="重新解析", width=10, command=self.on_parse)
        self.btn_reparse.grid(row=0, column=3, padx=(6, 0))

        ttk.Label(box, text="工作表").grid(row=0, column=4, padx=(16, 6))
        self.cmb_sheet = ttk.Combobox(box, textvariable=self.var_sheet, state="readonly", width=16)
        self.cmb_sheet.grid(row=0, column=5)
        self.cmb_sheet.bind("<<ComboboxSelected>>", lambda _event: self.on_parse(silent=True))

        ttk.Label(
            box,
            text="格式：A 列放 QQ 号（纯数字），B 列放文字内容；首行可为标题，会自动跳过。",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(6, 0))

    def _build_settings_box(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(parent, text=" 第 2 步 · 发件设置（QQ 邮箱 SMTP） ", padding=(10, 8))
        box.pack(fill="x", pady=(10, 0))
        box.columnconfigure(1, weight=1)
        box.columnconfigure(3, weight=1)

        ttk.Label(box, text="发件邮箱").grid(row=0, column=0, sticky="w", pady=3)
        self.ent_email = ttk.Entry(box, textvariable=self.var_email, width=14)
        self.ent_email.grid(row=0, column=1, sticky="ew", padx=(8, 18), pady=3)

        ttk.Label(box, text="SMTP 授权码").grid(row=0, column=2, sticky="w", pady=3)
        auth_wrap = ttk.Frame(box)
        auth_wrap.grid(row=0, column=3, sticky="ew", padx=(8, 0), pady=3)
        auth_wrap.columnconfigure(0, weight=1)
        self.ent_auth = ttk.Entry(auth_wrap, textvariable=self.var_auth, show="*", width=14)
        self.ent_auth.grid(row=0, column=0, sticky="ew")
        ttk.Checkbutton(
            auth_wrap, text="显示", variable=self.var_show_auth, command=self._toggle_auth
        ).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(box, text="显示名称").grid(row=1, column=0, sticky="w", pady=3)
        self.ent_name = ttk.Entry(box, textvariable=self.var_name, width=14)
        self.ent_name.grid(row=1, column=1, sticky="ew", padx=(8, 18), pady=3)

        ttk.Label(box, text="邮件主题").grid(row=1, column=2, sticky="w", pady=3)
        self.ent_subject = ttk.Entry(box, textvariable=self.var_subject, width=14)
        self.ent_subject.grid(row=1, column=3, sticky="ew", padx=(8, 0), pady=3)

        ttk.Label(box, text="发送间隔").grid(row=2, column=0, sticky="w", pady=3)
        interval_wrap = ttk.Frame(box)
        interval_wrap.grid(row=2, column=1, sticky="w", padx=(8, 18), pady=3)
        self.spn_interval = ttk.Spinbox(
            interval_wrap, from_=1, to=300, increment=1, width=6, textvariable=self.var_interval
        )
        self.spn_interval.pack(side="left")
        ttk.Label(interval_wrap, text="秒 / 封", style="Hint.TLabel").pack(side="left", padx=(4, 0))

        ttk.Label(box, text="失败重试").grid(row=2, column=2, sticky="w", pady=3)
        retry_wrap = ttk.Frame(box)
        retry_wrap.grid(row=2, column=3, sticky="w", padx=(8, 0), pady=3)
        self.spn_retry = ttk.Spinbox(
            retry_wrap, from_=0, to=5, increment=1, width=4, textvariable=self.var_retry
        )
        self.spn_retry.pack(side="left")
        ttk.Label(retry_wrap, text="次", style="Hint.TLabel").pack(side="left", padx=(4, 0))
        ttk.Label(retry_wrap, text="　本次最多").pack(side="left", padx=(16, 4))
        self.spn_limit = ttk.Spinbox(
            retry_wrap, from_=0, to=9999, increment=10, width=6, textvariable=self.var_limit
        )
        self.spn_limit.pack(side="left")
        ttk.Label(retry_wrap, text="封（0=不限）", style="Hint.TLabel").pack(side="left", padx=(4, 0))

        suffix_row = ttk.Frame(box)
        suffix_row.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        ttk.Label(suffix_row, text="收件邮箱后缀").pack(side="left")
        self.ent_suffix = ttk.Entry(suffix_row, textvariable=self.var_suffix, width=12)
        self.ent_suffix.pack(side="left", padx=(6, 0))
        self.btn_test = ttk.Button(
            suffix_row, text="测试登录", width=10, command=self.on_test_login
        )
        self.btn_test.pack(side="right")

        options_row = ttk.Frame(box)
        options_row.grid(row=4, column=0, columnspan=4, sticky="w", pady=(4, 0))
        ttk.Checkbutton(
            options_row, text="记住授权码（本机加密）", variable=self.var_remember
        ).pack(side="left")
        ttk.Checkbutton(
            options_row, text="重复 QQ 只发第一条", variable=self.var_dedupe
        ).pack(side="left", padx=(16, 0))
        ttk.Checkbutton(options_row, text="试运行（不真发）", variable=self.var_dry).pack(
            side="left", padx=(16, 0)
        )

        ttk.Label(
            box,
            text="授权码：QQ 邮箱网页版 → 设置 → 账号 → 开启 POP3/SMTP 服务后生成（不是 QQ 密码）。",
            style="Hint.TLabel",
        ).grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 0))

    def _build_notebook(self, parent: ttk.Frame) -> None:
        self.nb = ttk.Notebook(parent)
        self.nb.pack(fill="both", expand=True, pady=(10, 0))

        # -- 名单预览 --
        self.tab_preview = ttk.Frame(self.nb, padding=(6, 6))
        self.nb.add(self.tab_preview, text="名单预览（0）")
        self.tab_preview.rowconfigure(0, weight=1)
        self.tab_preview.columnconfigure(0, weight=1)

        columns = ("row", "qq", "email", "content")
        self.tree = ttk.Treeview(
            self.tab_preview, columns=columns, show="headings", selectmode="browse", height=8
        )
        self.tree.heading("row", text="行号")
        self.tree.heading("qq", text="QQ 号")
        self.tree.heading("email", text="收件邮箱")
        self.tree.heading("content", text="内容摘要")
        self.tree.column("row", width=60, anchor="center", stretch=False)
        self.tree.column("qq", width=130, anchor="center", stretch=False)
        self.tree.column("email", width=250, anchor="w", stretch=False)
        self.tree.column("content", width=520, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(self.tab_preview, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.bind("<Double-1>", self.show_row_detail)

        ttk.Label(
            self.tab_preview,
            text="双击任意一行可查看完整内容；发送前请核对收件邮箱。",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # -- 问题行 --
        self.tab_issues = ttk.Frame(self.nb, padding=(6, 6))
        self.nb.add(self.tab_issues, text="问题行（0）")
        self.tab_issues.rowconfigure(0, weight=1)
        self.tab_issues.columnconfigure(0, weight=1)
        self.issue_tree = ttk.Treeview(
            self.tab_issues,
            columns=("row", "raw", "reason"),
            show="headings",
            selectmode="browse",
            height=8,
        )
        self.issue_tree.heading("row", text="行号")
        self.issue_tree.heading("raw", text="第一列内容")
        self.issue_tree.heading("reason", text="跳过原因")
        self.issue_tree.column("row", width=60, anchor="center", stretch=False)
        self.issue_tree.column("raw", width=200, stretch=False)
        self.issue_tree.column("reason", width=600)
        self.issue_tree.grid(row=0, column=0, sticky="nsew")
        issue_scroll = ttk.Scrollbar(
            self.tab_issues, orient="vertical", command=self.issue_tree.yview
        )
        issue_scroll.grid(row=0, column=1, sticky="ns")
        self.issue_tree.configure(yscrollcommand=issue_scroll.set)
        ttk.Label(
            self.tab_issues,
            text="这些行不会被发送，可在 Excel 里修正后点「重新解析」。",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # -- 日志 --
        self.tab_log = ttk.Frame(self.nb, padding=(6, 6))
        self.nb.add(self.tab_log, text="发送日志")
        self.tab_log.rowconfigure(0, weight=1)
        self.tab_log.columnconfigure(0, weight=1)
        self.txt_log = scrolledtext.ScrolledText(
            self.tab_log, wrap="word", height=12, font=("Consolas", 9), state="disabled"
        )
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        self.txt_log.tag_configure("ok", foreground=COLOR_OK)
        self.txt_log.tag_configure("warn", foreground=COLOR_WARN)
        self.txt_log.tag_configure("error", foreground=COLOR_ERROR)
        self.txt_log.tag_configure("info", foreground="#222222")
        self.txt_log.tag_configure("muted", foreground=COLOR_MUTED)

    def _build_bottom(self, parent: ttk.Frame) -> None:
        bottom = ttk.Frame(parent)
        bottom.pack(side="bottom", fill="x", pady=(8, 0))
        self.bottom_bar = bottom

        self.pb = ttk.Progressbar(bottom, mode="determinate", maximum=100)
        self.pb.pack(fill="x")

        info = ttk.Frame(bottom)
        info.pack(fill="x", pady=(8, 0))

        self.btn_open = ttk.Button(
            info, text="打开结果文件夹", command=self.on_open_folder, state="disabled"
        )
        self.btn_open.pack(side="right")
        self.btn_stop = ttk.Button(info, text="停止", width=8, command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="right", padx=(8, 0))
        self.btn_pause = ttk.Button(
            info, text="暂停", width=8, command=self.on_pause, state="disabled"
        )
        self.btn_pause.pack(side="right", padx=(8, 0))
        self.btn_start = ttk.Button(info, text="开始发送", width=12, command=self.on_start)
        self.btn_start.pack(side="right", padx=(8, 0))

        self.lbl_status = ttk.Label(info, textvariable=self.var_status)
        self.lbl_status.pack(side="left")
        ttk.Label(info, textvariable=self.var_counts, style="Hint.TLabel").pack(
            side="left", padx=(14, 0)
        )

        menubar = tk.Menu(self)
        action_menu = tk.Menu(menubar, tearoff=0)
        action_menu.add_command(label="开始发送", accelerator="Ctrl+Enter", command=self.on_start)
        action_menu.add_command(label="暂停 / 继续", command=self.on_pause)
        action_menu.add_command(label="停止", command=self.on_stop)
        menubar.add_cascade(label="操作", menu=action_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)
        self.configure(menu=menubar)
        self.bind("<Control-Return>", lambda _event: self.on_start())

    # ------------------------------------------------------------------ #
    # 变量与配置
    # ------------------------------------------------------------------ #
    def _make_vars(self) -> None:
        self.var_excel = tk.StringVar()
        self.var_sheet = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_auth = tk.StringVar()
        self.var_show_auth = tk.BooleanVar(value=False)
        self.var_name = tk.StringVar()
        self.var_subject = tk.StringVar(value="消息通知")
        self.var_interval = tk.DoubleVar(value=5.0)
        self.var_retry = tk.IntVar(value=2)
        self.var_limit = tk.IntVar(value=0)
        self.var_suffix = tk.StringVar(value="@qq.com")
        self.var_remember = tk.BooleanVar(value=True)
        self.var_dedupe = tk.BooleanVar(value=False)
        self.var_dry = tk.BooleanVar(value=False)
        self.var_status = tk.StringVar(value="就绪")
        self.var_counts = tk.StringVar(value="待发送 0 封")

    def _apply_config(self) -> None:
        cfg = self.cfg
        self.var_email.set(cfg.sender_email)
        self.var_name.set(cfg.display_name)
        self.var_auth.set(cfg.auth_code)
        self.var_subject.set(cfg.subject_template or "消息通知")
        self.var_interval.set(max(1.0, float(cfg.interval)))
        self.var_retry.set(max(0, int(cfg.retries)))
        self.var_limit.set(max(0, int(cfg.limit)))
        self.var_suffix.set(cfg.addr_suffix or "@qq.com")
        self.var_remember.set(bool(cfg.remember_auth))
        self.var_dedupe.set(bool(cfg.dedupe))
        self.var_dry.set(bool(cfg.dry_run))
        self._toggle_auth()

        geometry = self._clamp_geometry(cfg.window_geometry or "")
        width, height = self._default_size()
        try:
            self.geometry(geometry or f"{width}x{height}")
        except tk.TclError:
            self.geometry(f"{width}x{height}")

        if cfg.last_excel and Path(cfg.last_excel).is_file():
            self.var_excel.set(cfg.last_excel)
            try:
                names = read_sheet_names(cfg.last_excel)
            except ExcelError:
                names = []
            if names:
                self.cmb_sheet.configure(values=names)
                self.var_sheet.set(names[0])

    def _default_size(self) -> tuple[int, int]:
        """默认窗口尺寸：不超过屏幕可用范围，保证底部按钮始终可见。"""
        try:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
        except tk.TclError:
            return 1100, 800
        width = min(1160, max(940, screen_w - 100))
        height = min(820, max(620, screen_h - 140))
        return width, height

    def _clamp_geometry(self, geometry: str) -> str:
        """把上次保存的窗口尺寸限制在当前屏幕内，避免底部按钮跑到屏幕外。"""
        match = re.match(r"^(\d+)x(\d+)([+-]\d+[+-]\d+)?$", geometry or "")
        if not match:
            return ""
        try:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
        except tk.TclError:
            return geometry
        width = min(int(match.group(1)), max(800, screen_w - 40))
        height = min(int(match.group(2)), max(560, screen_h - 80))
        return f"{width}x{height}{match.group(3) or ''}"

    def _collect_config(self) -> AppConfig:
        cfg = self.cfg
        cfg.sender_email = self.var_email.get().strip()
        cfg.display_name = self.var_name.get().strip()
        cfg.auth_code = self.var_auth.get().strip()
        cfg.remember_auth = bool(self.var_remember.get())
        cfg.subject_template = self.var_subject.get().strip() or "消息通知"
        cfg.interval = self._num(self.var_interval, 5.0)
        cfg.retries = int(self._num(self.var_retry, 2, int))
        cfg.limit = int(self._num(self.var_limit, 0, int))
        cfg.addr_suffix = self.var_suffix.get().strip() or "@qq.com"
        cfg.dedupe = bool(self.var_dedupe.get())
        cfg.dry_run = bool(self.var_dry.get())
        cfg.last_excel = self.var_excel.get().strip()
        try:
            cfg.window_geometry = self.geometry()
        except tk.TclError:
            pass
        return cfg

    @staticmethod
    def _num(var: tk.Variable, default: float, cast: type = float):
        try:
            return cast(var.get())
        except Exception:
            return default

    def _toggle_auth(self) -> None:
        self.ent_auth.configure(show="" if self.var_show_auth.get() else "*")

    # ------------------------------------------------------------------ #
    # 解析
    # ------------------------------------------------------------------ #
    def on_browse(self) -> None:
        initial = self.var_excel.get() or ""
        initial_dir = str(Path(initial).parent) if initial else str(Path.cwd())
        path = filedialog.askopenfilename(
            title="选择 Excel 名单",
            initialdir=initial_dir,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.var_excel.set(path)
        try:
            names = read_sheet_names(path)
        except ExcelError as exc:
            messagebox.showerror("无法读取文件", str(exc), parent=self)
            return
        if not names:
            messagebox.showerror("无法读取文件", "这个工作簿里没有任何工作表。", parent=self)
            return
        self.cmb_sheet.configure(values=names)
        self.var_sheet.set(names[0])
        self.parse_result = None
        self._parse_key = None
        self.on_parse(silent=True)

    def _suffix(self) -> str:
        try:
            return normalize_suffix(self.var_suffix.get())
        except ExcelError as exc:
            raise ValueError(str(exc)) from exc

    def on_parse(self, silent: bool = False) -> None:
        path = self.var_excel.get().strip()
        if not path:
            if not silent:
                messagebox.showinfo(APP_TITLE, "请先选择 Excel 文件。", parent=self)
            return
        try:
            suffix = self._suffix()
        except ValueError as exc:
            if not silent:
                messagebox.showwarning(APP_TITLE, str(exc), parent=self)
            return
        try:
            result = parse_workbook(
                path,
                sheet=self.var_sheet.get() or None,
                suffix=suffix,
                dedupe=bool(self.var_dedupe.get()),
            )
        except ExcelError as exc:
            self.parse_result = None
            self._parse_key = None
            self._fill_preview(None)
            self._log(f"解析失败：{exc}", "error")
            if not silent:
                messagebox.showerror("表格读取失败", str(exc), parent=self)
            return

        self.parse_result = result
        self._parse_key = (path, self.var_sheet.get() or "", suffix, bool(self.var_dedupe.get()))
        self._fill_preview(result)

        self._log(
            f"已读取 {Path(path).name}（工作表：{result.sheet}）：可发送 {len(result.rows)} 行，"
            f"跳过 {len(result.skipped)} 行，空行 {result.blank} 行。",
            "ok",
        )
        if result.duplicates:
            self._log(f"其中有 {result.duplicates} 行是重复的 QQ 号。", "warn")
        for item in result.skipped[:200]:
            self._log(f"  跳过第 {item.row_no} 行：{item.reason}", "warn")
        if len(result.skipped) > 200:
            self._log(f"  （其余 {len(result.skipped) - 200} 条见「问题行」标签）", "warn")

        self.nb.select(self.tab_preview)
        self.var_status.set("名单已就绪")
        self.var_counts.set(
            f"待发送 {len(result.rows)} 封 · 跳过 {len(result.skipped)} 行"
            + ("（试运行）" if self.var_dry.get() else "")
        )
        if self.var_dry.get():
            self._log("当前为试运行模式：不会真的发出邮件。", "warn")

    def _fill_preview(self, result: ParseResult | None) -> None:
        self.tree.delete(*self.tree.get_children())
        self.issue_tree.delete(*self.issue_tree.get_children())
        if result is None:
            self.nb.tab(self.tab_preview, text="名单预览（0）")
            self.nb.tab(self.tab_issues, text="问题行（0）")
            return

        for item in result.rows[:3000]:
            self.tree.insert(
                "", "end", values=(item.row_no, item.qq, item.email, summarize(item.content))
            )
        for item in result.skipped:
            self.issue_tree.insert("", "end", values=(item.row_no, item.raw_qq, item.reason))

        self.nb.tab(self.tab_preview, text=f"名单预览（{len(result.rows)}）")
        self.nb.tab(self.tab_issues, text=f"问题行（{len(result.skipped)}）")

    # ------------------------------------------------------------------ #
    # 设置校验
    # ------------------------------------------------------------------ #
    def _settings(self) -> MailerSettings:
        email = self.var_email.get().strip().replace(" ", "")
        if not email:
            raise ValueError("请填写发件邮箱，例如 123456789@qq.com。")
        if email.count("@") != 1:
            raise ValueError(f"发件邮箱「{email}」格式不对，应该像 123456789@qq.com。")
        local, domain = email.split("@")
        if not local or "." not in domain:
            raise ValueError(f"发件邮箱「{email}」格式不对，应该像 123456789@qq.com。")
        code = self.var_auth.get().replace(" ", "").strip()
        dry_run = bool(self.var_dry.get())
        if not code and not dry_run:
            raise ValueError(
                "请填写 SMTP 授权码。\n"
                "获取方式：QQ 邮箱网页版 → 设置 → 账号 → 开启 POP3/SMTP 服务 → 生成授权码。"
            )
        if code and len(code) < 6 and not dry_run:
            raise ValueError("授权码看起来太短了，QQ 邮箱的授权码一般是 16 位字母。")
        interval = self._num(self.var_interval, 5.0)
        if interval < 1:
            raise ValueError("发送间隔至少 1 秒。")
        retries = int(self._num(self.var_retry, 2, int))
        limit = int(self._num(self.var_limit, 0, int))
        return MailerSettings(
            host=QQ_HOST,
            port=QQ_PORT_SSL,
            use_ssl=True,
            username=email,
            password=code,
            display_name=self.var_name.get().strip(),
            subject_template=self.var_subject.get().strip() or "消息通知",
            interval=float(interval),
            retries=max(0, retries),
            limit=max(0, limit),
            dry_run=bool(self.var_dry.get()),
        )

    def _ensure_result(self) -> ParseResult:
        path = self.var_excel.get().strip()
        if not path:
            raise ValueError("请先选择 Excel 文件。")
        suffix = self._suffix()
        key = (path, self.var_sheet.get() or "", suffix, bool(self.var_dedupe.get()))
        if self.parse_result is None or self._parse_key != key:
            result = parse_workbook(
                path,
                sheet=key[1] or None,
                suffix=suffix,
                dedupe=key[3],
            )
            self.parse_result = result
            self._parse_key = key
            self._fill_preview(result)
        return self.parse_result

    # ------------------------------------------------------------------ #
    # 测试登录
    # ------------------------------------------------------------------ #
    def on_test_login(self) -> None:
        try:
            settings = self._settings()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc), parent=self)
            return
        self.btn_test.configure(state="disabled", text="测试中…")
        self.var_status.set("正在连接邮件服务器…")
        self.nb.select(self.tab_log)
        self._log("开始测试登录…", "info")

        def work() -> None:
            try:
                message = test_connection(
                    settings,
                    log=lambda text, level: self.events.put(
                        {"type": "log", "level": level, "text": text}
                    ),
                )
                self.events.put({"type": "test_ok", "message": message})
            except Exception as exc:  # noqa: BLE001
                self.events.put({"type": "test_fail", "message": str(exc)})

        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------------ #
    # 发送
    # ------------------------------------------------------------------ #
    def on_start(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo(APP_TITLE, "正在发送中，请先等待或点「停止」。", parent=self)
            return
        try:
            result = self._ensure_result()
        except ExcelError as exc:
            messagebox.showerror("表格读取失败", str(exc), parent=self)
            return
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc), parent=self)
            return

        try:
            settings = self._settings()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc), parent=self)
            return

        rows = result.rows
        if not rows:
            messagebox.showwarning(APP_TITLE, "没有可发送的行。", parent=self)
            return

        count = len(rows)
        if settings.limit:
            count = min(count, settings.limit)
        minutes = count * (settings.interval + 0.6) / 60
        mode = "试运行（不会真的发出）" if settings.dry_run else "真实发送"
        lines = [
            f"模式：{mode}",
            f"收件人：{count} 个 QQ 邮箱",
            f"发送间隔：{settings.interval:.0f} 秒/封，预计约 {format_duration(minutes * 60)}",
            f"发件人：{settings.username}",
        ]
        if not settings.dry_run and count > 50:
            lines += [
                "",
                "注意：一次发送过多可能被腾讯判定为异常并临时限制，建议分批发送或把间隔调大。",
            ]
        lines += ["", "确认开始吗？"]
        if not messagebox.askokcancel("确认发送", "\n".join(lines), parent=self):
            return

        self._save_config()
        self._run_started = time.monotonic()
        self.sender = Sender(settings, rows, on_event=self.events.put)
        self.worker = threading.Thread(target=self._run_sender, daemon=True)
        self._set_busy(True)
        self.pb.configure(value=0, maximum=100)
        self.var_status.set("正在发送…")
        self.var_counts.set(f"0 / {count}")
        self.nb.select(self.tab_log)
        self._log(
            f"开始发送：共 {count} 封，间隔 {settings.interval:.0f} 秒，"
            f"重试 {settings.retries} 次，模式：{mode}",
            "info",
        )
        self.worker.start()

    def _run_sender(self) -> None:
        assert self.sender is not None
        try:
            results = self.sender.run()
            self.events.put(
                {"type": "finished", "results": results, "aborted": self.sender.last_aborted}
            )
        except Exception:  # noqa: BLE001
            self.events.put({"type": "crashed", "error": traceback.format_exc()})

    def on_pause(self) -> None:
        if self.sender is None:
            return
        if self.sender.paused:
            self.sender.resume()
            self.btn_pause.configure(text="暂停")
            self.var_status.set("继续发送…")
            self._log("已继续发送。", "info")
        else:
            self.sender.pause()
            self.btn_pause.configure(text="继续")
            self.var_status.set("已暂停（当前这封发完后停下）")
            self._log("已暂停。", "warn")

    def on_stop(self) -> None:
        if self.sender is None:
            return
        if not messagebox.askokcancel(
            "停止发送", "确定要停止吗？剩余未发送的行会标记为「未发送」。", parent=self
        ):
            return
        self.sender.stop()
        self.var_status.set("正在停止…")
        self.btn_stop.configure(state="disabled")
        self.btn_pause.configure(state="disabled")
        self._log("已请求停止，正在收尾…", "warn")

    # ------------------------------------------------------------------ #
    # 事件循环
    # ------------------------------------------------------------------ #
    def _poll(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        except tk.TclError:
            return
        try:
            if self.winfo_exists():
                self.after(self.POLL_MS, self._poll)
        except tk.TclError:
            pass

    def _handle_event(self, event: dict) -> None:
        kind = event.get("type")
        if kind == "log":
            self._log(event.get("text", ""), event.get("level", "info"))
        elif kind == "item":
            if event.get("status") == "成功":
                self.var_status.set(f"已发送 {event.get('index')} / {event.get('total')}")
        elif kind == "progress":
            total = max(1, int(event.get("total") or 1))
            done = int(event.get("done") or 0)
            self.pb.configure(maximum=100, value=min(100, done * 100 / total))
            self.var_counts.set(
                f"{done} / {total}　成功 {int(event.get('ok') or 0)} · "
                f"失败 {int(event.get('fail') or 0)} · 未发 {int(event.get('skipped') or 0)}"
            )
        elif kind == "test_ok":
            self.btn_test.configure(state="normal", text="测试登录")
            self.var_status.set("登录测试成功")
            self._log(event.get("message", "登录成功"), "ok")
            messagebox.showinfo("登录成功", event.get("message", "登录成功"), parent=self)
        elif kind == "test_fail":
            self.btn_test.configure(state="normal", text="测试登录")
            self.var_status.set("登录测试失败")
            self._log(f"登录测试失败：{event.get('message')}", "error")
            messagebox.showerror("登录失败", event.get("message", ""), parent=self)
        elif kind == "finished":
            self._finish(event.get("results") or [], event.get("aborted", ""))
        elif kind == "crashed":
            self._set_busy(False)
            self.var_status.set("发送异常中止")
            self._log(event.get("error", "未知错误"), "error")
            messagebox.showerror(
                "程序异常",
                "发送过程中出现异常，详情见日志。\n\n" + event.get("error", ""),
                parent=self,
            )

    def _finish(self, results: list[dict], aborted: str = "") -> None:
        self._set_busy(False)
        self.pb.configure(maximum=100, value=100)
        ok = sum(1 for row in results if row.get("status") == "成功")
        fail = sum(1 for row in results if row.get("status") == "失败")
        skipped = sum(1 for row in results if row.get("status") == "未发送")
        trial = sum(1 for row in results if row.get("status") == "试运行")
        elapsed = format_duration(time.monotonic() - self._run_started)

        report_path: Path | None = None
        write_error = ""
        if results:
            try:
                report_path = self._write_outputs(results)
            except OSError as exc:
                write_error = str(exc)
                self._log(f"结果文件写入失败：{write_error}", "error")
            except Exception as exc:  # noqa: BLE001
                write_error = str(exc)
                self._log(f"结果文件写入失败：{write_error}", "error")

        summary = [f"用时 {elapsed}", f"成功 {ok} 封"]
        if fail:
            summary.append(f"失败 {fail} 封")
        if skipped:
            summary.append(f"未发送 {skipped} 封")
        if trial:
            summary.append(f"试运行 {trial} 封")
        summary_text = " · ".join(summary)
        self.var_status.set("已停止" if aborted else "发送完成")
        self.var_counts.set(summary_text)
        self._log(f"本次结束：{summary_text}", "ok" if not fail else "warn")

        if report_path:
            self._last_result_dir = report_path.parent
            self.btn_open.configure(state="normal")
            self._log(f"结果已保存：{report_path}", "ok")
            text = summary_text + f"\n\n结果文件：\n{report_path}"
            if fail:
                text += "\n\n失败原因见结果文件或日志。"
            messagebox.showinfo("发送结束", text, parent=self)
        elif write_error:
            messagebox.showwarning(
                "邮件已发出，但结果文件没写成",
                summary_text
                + "\n\n邮件已经按照名单发出，「发送日志」里保留了每一封的结果，"
                "可以直接截图或全选复制。\n\n"
                "写文件失败的原因：\n"
                + write_error
                + "\n\n提示：如果程序是从别的工具里启动的，直接到文件夹里双击 exe 再运行一次即可正常保存。",
                parent=self,
            )
        else:
            messagebox.showinfo("发送结束", summary_text, parent=self)

    def _output_candidates(self) -> list[Path]:
        """结果文件可以写的位置，按优先级：表格同目录 → 桌面 → 文档 → 临时目录 → 当前目录。"""
        source = Path(self.var_excel.get().strip() or ".")
        home = Path.home()
        candidates = [
            source.parent,
            home / "Desktop",
            home / "Documents",
            Path(tempfile.gettempdir()),
            Path.cwd(),
        ]
        if getattr(sys, "frozen", False):
            candidates.append(Path(sys.executable).resolve().parent)
        candidates.append(Path(__file__).resolve().parents[1])

        unique: list[Path] = []
        seen: set[str] = set()
        for path in candidates:
            key = str(path).lower()
            if key not in seen:
                seen.add(key)
                unique.append(path)
        return unique

    def _write_outputs(self, results: list[dict]) -> Path:
        source = Path(self.var_excel.get().strip())
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = "试运行结果" if self.var_dry.get() else "发送结果"
        report_name = f"{prefix}_{source.stem}_{stamp}.xlsx"
        log_name = f"发送记录_{source.stem}_{stamp}.txt"

        def writer(folder: Path) -> Path:
            report_path = write_report(folder / report_name, results)
            self._write_text_log(folder / log_name, source)
            return report_path

        report_path = write_with_fallback(self._output_candidates(), writer)
        try:
            moved = str(report_path.parent.resolve()).lower() != str(source.parent.resolve()).lower()
        except OSError:
            moved = True
        if moved:
            self._log(f"表格所在目录不可写，结果改存到：{report_path.parent}", "warn")
        return report_path

    def _write_text_log(self, log_path: Path, source: Path) -> None:
        header = [
            f"{APP_TITLE} 发送记录",
            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"表格：{source}",
            f"发件人：{self.var_email.get().strip()}",
            f"模式：{'试运行' if self.var_dry.get() else '真实发送'}",
            "-" * 60,
        ]
        try:
            log_path.write_text("\n".join(header + self._log_lines), encoding="utf-8")
            self._log(f"文本记录：{log_path}", "muted")
        except OSError as exc:
            self._log(f"文本记录写入失败：{exc}", "warn")

    def _set_busy(self, busy: bool) -> None:
        widgets = {
            self.btn_browse: "normal",
            self.btn_reparse: "normal",
            self.ent_email: "normal",
            self.ent_auth: "normal",
            self.ent_name: "normal",
            self.ent_subject: "normal",
            self.ent_suffix: "normal",
            self.spn_interval: "normal",
            self.spn_retry: "normal",
            self.spn_limit: "normal",
            self.btn_test: "normal",
            self.cmb_sheet: "readonly",
        }
        for widget, state in widgets.items():
            try:
                widget.configure(state="disabled" if busy else state)
            except tk.TclError:
                pass
        if busy:
            self.btn_start.configure(state="disabled")
            self.btn_pause.configure(state="normal", text="暂停")
            self.btn_stop.configure(state="normal")
        else:
            self.btn_start.configure(state="normal")
            self.btn_pause.configure(state="disabled", text="暂停")
            self.btn_stop.configure(state="disabled")

    # ------------------------------------------------------------------ #
    # 杂项
    # ------------------------------------------------------------------ #
    def _log(self, text: str, level: str = "info") -> None:
        if not text:
            return
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {text}"
        self._log_lines.append(line)
        try:
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", line + "\n", level)
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        except tk.TclError:
            pass

    def show_row_detail(self, _event: object = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        values = self.tree.item(selection[0], "values")
        if not values:
            return
        row_no = values[0]
        item = None
        if self.parse_result is not None:
            item = next(
                (row for row in self.parse_result.rows if str(row.row_no) == str(row_no)), None
            )
        window = tk.Toplevel(self)
        window.title(f"第 {row_no} 行内容")
        window.geometry("660x480")
        window.transient(self)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"QQ：{values[1]}　收件邮箱：{values[2]}", font=FONT_BOLD).pack(
            anchor="w"
        )
        ttk.Label(frame, text="正文内容：", style="Hint.TLabel").pack(anchor="w", pady=(8, 2))
        text = tk.Text(frame, wrap="word", font=FONT_UI)
        text.pack(fill="both", expand=True)
        text.insert("1.0", item.content if item else "（内容不可用）")
        text.configure(state="disabled")
        ttk.Button(frame, text="关闭", command=window.destroy).pack(anchor="e", pady=(8, 0))

    def show_help(self) -> None:
        window = tk.Toplevel(self)
        window.title("使用说明")
        window.geometry("720x600")
        window.transient(self)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word", font=FONT_UI)
        text.pack(fill="both", expand=True)
        text.insert("1.0", HELP_TEXT)
        text.configure(state="disabled")
        ttk.Button(frame, text="关闭", command=window.destroy).pack(anchor="e", pady=(8, 0))

    def show_about(self) -> None:
        try:
            location = str(self._config_path or config_path())
        except Exception:
            location = "（无法确定配置目录）"
        messagebox.showinfo(
            "关于",
            f"{APP_TITLE} v{__version__}\n\n"
            "按 Excel 名单把文字内容逐封发送到对应的 QQ 邮箱。\n"
            f"配置文件：{location}\n\n"
            "请在获得对方同意的前提下发送通知类邮件，并遵守腾讯邮箱的使用规则。",
            parent=self,
        )

    def on_open_folder(self) -> None:
        target = self._last_result_dir
        if target is None:
            source = self.var_excel.get().strip()
            target = Path(source).parent if source else Path.cwd()
        try:
            os.startfile(str(target))  # type: ignore[attr-defined]
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning(APP_TITLE, f"无法打开文件夹：{exc}", parent=self)

    def _save_config(self) -> None:
        try:
            saved = self._collect_config().save()
            self._config_path = saved
            if saved.parent != config_path().parent:
                self._log(f"配置保存到备用位置：{saved}", "muted")
        except Exception as exc:  # noqa: BLE001
            self._log(f"配置没能保存（不影响发送）：{exc}", "warn")

    def _on_close(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            if not messagebox.askokcancel(
                "正在发送", "还在发送中，确定要退出吗？未发送的部分会标记为「未发送」。", parent=self
            ):
                return
            if self.sender is not None:
                self.sender.stop()
            self.worker.join(timeout=3)
        self._save_config()
        self.destroy()


def write_with_fallback(candidates: list[Path], writer: Callable[[Path], Path]) -> Path:
    """依次尝试候选目录，直到 writer 成功写出文件；全部失败时汇总原因抛出。"""
    errors: list[str] = []
    for folder in candidates:
        try:
            folder.mkdir(parents=True, exist_ok=True)
            return writer(folder)
        except OSError as exc:
            errors.append(f"{folder}：{exc.strerror or exc}")
    raise OSError("以下位置都无法写入结果文件：\n" + "\n".join(errors))


def summarize(content: str, limit: int = 60) -> str:
    text = " ".join(content.split())
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


def format_duration(seconds: float) -> str:
    seconds = int(max(0, seconds))
    if seconds < 60:
        return f"{seconds} 秒"
    minutes, rest = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes} 分 {rest:02d} 秒"
    hours, minutes = divmod(minutes, 60)
    return f"{hours} 小时 {minutes:02d} 分"


def main() -> None:
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
