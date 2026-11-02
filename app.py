"""程序入口：双击本文件或用 python app.py 启动图形界面。"""

from __future__ import annotations

import os
import sys
import tempfile
import traceback
from pathlib import Path

FROZEN = bool(getattr(sys, "frozen", False))
BASE_DIR = Path(sys.executable).parent if FROZEN else Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))


def _write_crash_log(text: str) -> Path | None:
    candidates = [
        BASE_DIR / "启动错误.log",
        Path(os.environ.get("APPDATA") or Path.home()) / "QQMailerAssistant" / "启动错误.log",
        Path(tempfile.gettempdir()) / "QQMailerAssistant_启动错误.log",
    ]
    for path in candidates:
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
            return path
        except OSError:
            continue
    return None


def _self_test() -> int:
    """--self-test：只构建界面然后立刻退出，用来验证打包结果是否可用。"""
    try:
        from qq_mailer.ui import App

        app = App()
        app.withdraw()
        app.update_idletasks()
        app.update()
        tabs = app.nb.tabs()
        app.destroy()
        message = f"OK 自检通过（标签页 {len(tabs)} 个，版本 {__import__('qq_mailer').__version__}）"
        code = 0
    except Exception:
        message = "FAIL 自检失败\n" + traceback.format_exc()
        code = 2
    try:
        (BASE_DIR / "自检结果.txt").write_text(message, encoding="utf-8")
    except OSError:
        pass
    print(message)
    return code


def run() -> None:
    if "--self-test" in sys.argv:
        raise SystemExit(_self_test())
    try:
        from qq_mailer.ui import main

        main()
    except Exception:
        text = traceback.format_exc()
        log_path = _write_crash_log(text)
        detail = text.strip().splitlines()[-1] if text.strip() else "未知错误"
        message = f"程序启动时出错：\n\n{detail}"
        if log_path is not None:
            message += f"\n\n详细信息已写入：\n{log_path}"
        try:
            import tkinter.messagebox as messagebox

            messagebox.showerror("启动失败", message)
        except Exception:
            pass
        raise SystemExit(1)


if __name__ == "__main__":
    run()
