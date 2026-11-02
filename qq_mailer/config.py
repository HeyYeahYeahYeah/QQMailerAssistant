"""配置读写。

授权码默认用 Windows DPAPI 加密后保存在本机（只有当前 Windows 账户能解密），
在没有 DPAPI 的环境下退化为 base64 编码（仅防误读，不防有心人）。
"""

from __future__ import annotations

import base64
import json
import os
import sys
import tempfile
from dataclasses import dataclass, fields
from pathlib import Path
from typing import Any

APP_DIR_NAME = "QQMailerAssistant"
CONFIG_FILENAME = "config.json"


def candidate_dirs() -> list[Path]:
    """配置可能存放的位置，按优先级排列（某个位置没权限就换下一个）。"""
    dirs: list[Path] = []
    for env_name in ("APPDATA", "LOCALAPPDATA"):
        base = os.environ.get(env_name) or ""
        if base:
            dirs.append(Path(base) / APP_DIR_NAME)
    if getattr(sys, "frozen", False):
        dirs.append(Path(sys.executable).resolve().parent / "data")
    dirs.append(Path(__file__).resolve().parents[1] / "data")
    dirs.append(Path.home() / APP_DIR_NAME)

    unique: list[Path] = []
    seen: set[str] = set()
    for directory in dirs:
        key = str(directory).lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(directory)
    return unique


def config_dir() -> Path:
    """首选配置目录（不创建，避免无权限时启动失败）。"""
    return candidate_dirs()[0]


def config_path() -> Path:
    return config_dir() / CONFIG_FILENAME


# --------------------------------------------------------------------------- #
# Windows DPAPI
# --------------------------------------------------------------------------- #
def _dpapi(data: bytes, protect: bool) -> bytes:
    """调用 Windows 的 CryptProtectData / CryptUnprotectData。"""
    import ctypes
    from ctypes import wintypes

    class DataBlob(ctypes.Structure):
        _fields_ = [
            ("cbData", wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_char)),
        ]

    buffer = ctypes.create_string_buffer(data, len(data))
    blob_in = DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char)))
    blob_out = DataBlob()

    crypt32 = ctypes.windll.crypt32
    if protect:
        ok = crypt32.CryptProtectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
        )
    else:
        ok = crypt32.CryptUnprotectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
        )
    if not ok:
        raise OSError("CryptProtectData/CryptUnprotectData 调用失败")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


def protect_secret(text: str) -> str:
    """加密授权码，返回可直接写进 JSON 的字符串。"""
    if not text:
        return ""
    if os.name == "nt":
        try:
            return "dpapi:" + base64.b64encode(_dpapi(text.encode("utf-8"), True)).decode("ascii")
        except Exception:
            pass
    return "b64:" + base64.b64encode(text.encode("utf-8")).decode("ascii")


def unprotect_secret(stored: str) -> str:
    """解密配置里的授权码；解不开就返回空字符串。"""
    if not stored:
        return ""
    try:
        if stored.startswith("dpapi:"):
            return _dpapi(base64.b64decode(stored[6:]), False).decode("utf-8")
        if stored.startswith("b64:"):
            return base64.b64decode(stored[4:]).decode("utf-8")
    except Exception:
        return ""
    return stored


# --------------------------------------------------------------------------- #
# 配置对象
# --------------------------------------------------------------------------- #
@dataclass
class AppConfig:
    sender_email: str = ""
    display_name: str = ""
    auth_code: str = ""
    remember_auth: bool = True
    subject_template: str = "消息通知"
    interval: float = 5.0
    retries: int = 2
    addr_suffix: str = "@qq.com"
    dedupe: bool = False
    dry_run: bool = False
    limit: int = 0
    last_excel: str = ""
    window_geometry: str = ""

    # -- 序列化 -------------------------------------------------------------
    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for field in fields(self):
            if field.name == "auth_code":
                continue
            data[field.name] = getattr(self, field.name)
        data["auth_code_enc"] = protect_secret(self.auth_code) if self.remember_auth else ""
        data["_version"] = 1
        return data

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "AppConfig":
        cfg = cls()
        for field in fields(cls):
            if field.name not in raw:
                continue
            value = raw[field.name]
            setattr(cfg, field.name, _coerce(field.name, value))
        enc = raw.get("auth_code_enc") or raw.get("auth_code") or ""
        cfg.auth_code = unprotect_secret(enc) if isinstance(enc, str) else ""
        return cfg

    # -- 磁盘 ---------------------------------------------------------------
    @classmethod
    def load(cls, path: Path | None = None) -> "AppConfig":
        targets = [path] if path else [d / CONFIG_FILENAME for d in candidate_dirs()]
        for target in targets:
            try:
                if not target.is_file():
                    continue
                raw = json.loads(target.read_text(encoding="utf-8"))
            except Exception:
                continue
            if isinstance(raw, dict):
                return cls.from_dict(raw)
        return cls()

    def save(self, path: Path | None = None) -> Path:
        """保存配置，返回实际写入的文件；一个位置不行就换下一个。"""
        payload = json.dumps(self.to_dict(), ensure_ascii=False, indent=2)
        targets = [path] if path else [d / CONFIG_FILENAME for d in candidate_dirs()]
        last_error: OSError | None = None
        for target in targets:
            tmp_name = ""
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                fd, tmp_name = tempfile.mkstemp(
                    prefix="cfg-", suffix=".json", dir=str(target.parent)
                )
                with os.fdopen(fd, "w", encoding="utf-8") as handle:
                    handle.write(payload)
                os.replace(tmp_name, target)
                return target
            except OSError as exc:
                last_error = exc
                if tmp_name and os.path.exists(tmp_name):
                    try:
                        os.unlink(tmp_name)
                    except OSError:
                        pass
        raise last_error or OSError("配置无法保存")


def _coerce(name: str, value: Any) -> Any:
    """把 JSON 里的值转成 dataclass 期望的类型，坏值回退默认。"""
    default = AppConfig.__dataclass_fields__[name].default
    try:
        if isinstance(default, bool):
            return bool(value)
        if isinstance(default, int):
            return max(0, int(float(value)))
        if isinstance(default, float):
            return max(0.0, float(value))
        if isinstance(default, str):
            return str(value)
    except (TypeError, ValueError):
        return default
    return value
