"""SMTP 发送核心：构造邮件、连接登录、限速重试、逐条发送。"""

from __future__ import annotations

import random
import smtplib
import socket
import ssl
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from typing import Callable, Sequence

from .excel_io import RowItem

QQ_HOST = "smtp.qq.com"
QQ_PORT_SSL = 465
QQ_PORT_TLS = 587

AUTH_HINT = (
    "登录失败。请检查：\n"
    "1) 发件邮箱是否填写完整（例如 123456789@qq.com）；\n"
    "2) 密码栏填的是「SMTP 授权码」，不是 QQ 登录密码；\n"
    "3) 是否已在 QQ 邮箱 → 设置 → 账号 里开启「POP3/SMTP 服务」。\n"
    "授权码在开启服务时由腾讯给出，通常为 16 位字母。"
)


class MailerError(Exception):
    """可直接展示给用户的发送错误。"""


class AuthFailed(MailerError):
    """登录 / 认证失败：重试没有意义，必须终止。"""


class RateLimited(MailerError):
    """被服务器判定为发送过于频繁。"""


class RecipientRejected(MailerError):
    """单个收件地址被拒绝，不影响后续地址。"""


@dataclass
class MailerSettings:
    host: str = QQ_HOST
    port: int = QQ_PORT_SSL
    use_ssl: bool = True
    starttls: bool = True
    username: str = ""
    password: str = ""
    display_name: str = ""
    subject_template: str = "消息通知"
    interval: float = 5.0
    retries: int = 2
    limit: int = 0
    dry_run: bool = False
    timeout: float = 30.0


# --------------------------------------------------------------------------- #
# 邮件内容
# --------------------------------------------------------------------------- #
def _decode(value: object) -> str:
    if isinstance(value, bytes):
        return value.decode("utf-8", "replace")
    return str(value)


def sanitize_header(text: str) -> str:
    """去掉换行与控制字符，避免邮件头被正文内容污染。"""
    cleaned = "".join(ch for ch in text.replace("\r", " ").replace("\n", " ") if ch >= " ")
    return " ".join(cleaned.split())


def content_preview(content: str, limit: int = 40) -> str:
    text = " ".join(content.split())
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


def render_subject(
    template: str, item: RowItem, index: int, total: int, now: datetime | None = None
) -> str:
    now = now or datetime.now()
    preview = content_preview(item.content)
    text = template or "消息通知"
    mapping = {
        "{qq}": item.qq,
        "{QQ}": item.qq,
        "{email}": item.email,
        "{邮箱}": item.email,
        "{index}": str(index),
        "{序号}": str(index),
        "{total}": str(total),
        "{总数}": str(total),
        "{date}": now.strftime("%Y-%m-%d"),
        "{日期}": now.strftime("%Y-%m-%d"),
        "{content}": preview,
        "{内容}": preview,
    }
    for key, value in mapping.items():
        text = text.replace(key, value)
    text = sanitize_header(text)
    return (text or "消息通知")[:200]


def build_message(
    settings: MailerSettings, item: RowItem, index: int, total: int, now: datetime | None = None
) -> EmailMessage:
    now = now or datetime.now()
    sender = sanitize_header(settings.username)
    display = sanitize_header(settings.display_name) or sender.split("@")[0]
    subject = render_subject(settings.subject_template, item, index, total, now)

    message = EmailMessage()
    message["From"] = formataddr((display, sender)) if display else sender
    message["To"] = sanitize_header(item.email)
    message["Subject"] = subject
    message["Date"] = formatdate(localtime=True)
    domain = sender.split("@")[-1] if "@" in sender else "localhost"
    message["Message-ID"] = make_msgid(domain=domain)
    message["X-Mailer"] = "QQMailerAssistant"
    body = item.content if item.content.strip() else "（无内容）"
    message.set_content(body, subtype="plain", charset="utf-8")
    return message


# --------------------------------------------------------------------------- #
# 连接
# --------------------------------------------------------------------------- #
def _safe_close(server: smtplib.SMTP | None) -> None:
    if server is None:
        return
    try:
        server.quit()
    except Exception:
        try:
            server.close()
        except Exception:
            pass


def _connect(
    host: str, port: int, use_ssl: bool, timeout: float, starttls: bool = True
) -> smtplib.SMTP:
    context = ssl.create_default_context()
    if use_ssl:
        return smtplib.SMTP_SSL(host=host, port=port, timeout=timeout, context=context)
    server = smtplib.SMTP(host=host, port=port, timeout=timeout)
    server.ehlo()
    if starttls:
        server.starttls(context=context)
        server.ehlo()
    return server


def _login(server: smtplib.SMTP, settings: MailerSettings) -> None:
    try:
        server.login(settings.username, settings.password)
    except smtplib.SMTPAuthenticationError as exc:
        raise AuthFailed(f"{AUTH_HINT}\n\n服务器返回：{exc.smtp_code} {_decode(exc.smtp_error)}") from exc
    except smtplib.SMTPNotSupportedError as exc:
        raise AuthFailed(f"服务器不支持登录认证：{exc}\n请确认服务器地址是 SMTP 发信服务器。") from exc
    except smtplib.SMTPException as exc:
        raise AuthFailed(f"{AUTH_HINT}\n\n错误信息：{exc}") from exc


def open_connection(
    settings: MailerSettings, log: Callable[[str, str], None] | None = None
) -> smtplib.SMTP:
    """连接并登录；SSL 端口失败时自动尝试 STARTTLS 端口。"""
    if not settings.username or not settings.password:
        raise AuthFailed("请先填写发件邮箱和 SMTP 授权码。")

    candidates: list[tuple[int, bool, bool]] = [
        (settings.port, settings.use_ssl, settings.starttls)
    ]
    if settings.port == QQ_PORT_SSL:
        alternate = (QQ_PORT_TLS, False, True)
    else:
        alternate = (QQ_PORT_SSL, True, True)
    candidates.append(alternate)

    errors: list[str] = []
    for port, use_ssl, starttls in candidates:
        label = f"{settings.host}:{port}（{'SSL' if use_ssl else 'STARTTLS' if starttls else '明文'}）"
        try:
            if log:
                log(f"正在连接 {label} …", "info")
            server = _connect(settings.host, port, use_ssl, settings.timeout, starttls)
        except ssl.SSLError as exc:
            errors.append(f"{label} TLS 握手失败：{exc}")
            continue
        except socket.gaierror as exc:
            errors.append(f"{label} 域名解析失败：{exc}")
            continue
        except (socket.timeout, TimeoutError):
            errors.append(f"{label} 连接超时")
            continue
        except OSError as exc:
            errors.append(f"{label} 连接失败：{exc}")
            continue
        except smtplib.SMTPException as exc:
            errors.append(f"{label} SMTP 错误：{exc}")
            continue

        try:
            if log:
                log("已连接，正在登录…", "info")
            _login(server, settings)
        except MailerError:
            _safe_close(server)
            raise
        return server

    raise MailerError(
        "连接邮件服务器失败：\n" + "\n".join(errors) + "\n\n请检查网络、公司防火墙/代理，或稍后重试。"
    )


def test_connection(settings: MailerSettings, log: Callable[[str, str], None] | None = None) -> str:
    server = open_connection(settings, log=log)
    try:
        peer = ""
        try:
            if server.sock is not None:
                peer = f"{server.sock.getpeername()[0]}:{server.sock.getpeername()[1]}"
        except Exception:
            peer = ""
        return f"登录成功：{settings.username}" + (f"（服务器 {peer}）" if peer else "")
    finally:
        _safe_close(server)


# --------------------------------------------------------------------------- #
# 错误描述
# --------------------------------------------------------------------------- #
def _describe(exc: BaseException | None) -> str:
    if exc is None:
        return "未知错误"
    if isinstance(exc, smtplib.SMTPServerDisconnected):
        return "与服务器的连接被断开（网络不稳定或服务器限流）"
    if isinstance(exc, smtplib.SMTPConnectError):
        return "无法建立 SMTP 连接"
    if isinstance(exc, (socket.timeout, TimeoutError)):
        return "连接或发送超时"
    if isinstance(exc, socket.gaierror):
        return "服务器地址解析失败，请检查网络"
    if isinstance(exc, OSError):
        code = getattr(exc, "winerror", None)
        if code == 10054:
            return "连接被服务器重置（10054）"
        if code == 10060:
            return "连接超时（10060）"
        if code == 10061:
            return "目标端口拒绝连接（10061），可能被防火墙拦截"
    text = str(exc).strip()
    return text or exc.__class__.__name__


def _is_rate_limited(text: str) -> bool:
    lowered = text.lower()
    return any(
        key in lowered
        for key in ("exceed", "too many", "frequency", "rate limit", "421", "频率", "限制", "过多")
    )


def _recipient_detail(exc: smtplib.SMTPRecipientsRefused) -> str:
    parts = []
    for address, (code, message) in exc.recipients.items():
        parts.append(f"{address} → {code} {_decode(message)}")
    return "；".join(parts) or str(exc)


# --------------------------------------------------------------------------- #
# 发送器
# --------------------------------------------------------------------------- #
class Sender:
    """在后台线程里逐条发送，通过 on_event 回调把状态推给界面。"""

    def __init__(
        self,
        settings: MailerSettings,
        items: Sequence[RowItem],
        on_event: Callable[[dict], None] | None = None,
    ) -> None:
        self.settings = settings
        self.items = list(items)
        self.on_event = on_event or (lambda event: None)
        self.results: list[dict] = []
        self.last_aborted = ""
        self._ok = 0
        self._fail = 0
        self._skipped = 0
        self._server: smtplib.SMTP | None = None
        self._stop = threading.Event()
        self._pause = threading.Event()

    # -- 外部控制 -----------------------------------------------------------
    def stop(self) -> None:
        self._stop.set()
        self._pause.clear()

    def pause(self) -> None:
        self._pause.set()

    def resume(self) -> None:
        self._pause.clear()

    @property
    def paused(self) -> bool:
        return self._pause.is_set()

    # -- 事件 ---------------------------------------------------------------
    def _emit(self, event_type: str, **payload) -> None:
        try:
            self.on_event({"type": event_type, **payload})
        except Exception:
            pass

    def _log(self, text: str, level: str = "info") -> None:
        self._emit("log", level=level, text=text)

    def _record(self, item: RowItem, status: str, detail: str, subject: str = "") -> None:
        if status == "成功":
            self._ok += 1
        elif status == "失败":
            self._fail += 1
        else:
            self._skipped += 1
        self.results.append(
            {
                "row_no": item.row_no,
                "qq": item.qq,
                "email": item.email,
                "status": status,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "detail": detail,
                "subject": subject,
            }
        )

    # -- 停 / 暂停 ----------------------------------------------------------
    def _wait_pause(self) -> bool:
        """暂停期间阻塞；返回 True 表示应当停止。"""
        while self._pause.is_set():
            if self._stop.wait(0.2):
                return True
        return self._stop.is_set()

    def _sleep_between(self) -> None:
        base = max(0.0, float(self.settings.interval))
        if base <= 0:
            return
        jitter = random.uniform(-0.25, 0.25) * base
        self._stop.wait(max(0.2, base + jitter))

    # -- 连接 ---------------------------------------------------------------
    def _ensure_server(self) -> smtplib.SMTP:
        if self._server is not None:
            return self._server
        self._server = open_connection(self.settings, log=self._log)
        return self._server

    def _close(self) -> None:
        _safe_close(self._server)
        self._server = None

    # -- 单封发送 -----------------------------------------------------------
    def _send_one(self, message: EmailMessage) -> None:
        attempts = max(1, int(self.settings.retries) + 1)
        last_error: BaseException | None = None

        for attempt in range(1, attempts + 1):
            try:
                server = self._ensure_server()
                server.send_message(message)
                return
            except (AuthFailed, RateLimited):
                raise
            except smtplib.SMTPRecipientsRefused as exc:
                self._server = None
                raise RecipientRejected(f"收件地址被拒绝：{_recipient_detail(exc)}") from exc
            except smtplib.SMTPSenderRefused as exc:
                raise MailerError(
                    f"发件地址被拒绝（{exc.smtp_code} {_decode(exc.smtp_error)}），可能当天发送量已超限。"
                ) from exc
            except smtplib.SMTPDataError as exc:
                text = _decode(exc.smtp_error)
                self._close()
                if _is_rate_limited(text):
                    raise RateLimited(
                        f"服务器提示发送过于频繁或已达上限：{exc.smtp_code} {text}\n"
                        "建议把发送间隔调大（例如 10–20 秒），隔一段时间再继续。"
                    ) from exc
                if 400 <= int(exc.smtp_code) < 500 and attempt < attempts:
                    last_error = exc
                    self._log(f"服务器暂时拒收（{exc.smtp_code}），{2 * attempt} 秒后重试…", "warn")
                    if self._stop.wait(2 * attempt):
                        raise MailerError("已手动停止") from exc
                    continue
                raise MailerError(f"服务器拒收邮件内容（{exc.smtp_code}）：{text}") from exc
            except MailerError as exc:
                last_error = exc
                self._close()
                if attempt < attempts and not self._stop.is_set():
                    self._log(f"连接失败：{_describe(exc)}，{2 * attempt} 秒后重试…", "warn")
                    if self._stop.wait(2 * attempt):
                        raise MailerError("已手动停止") from exc
                    continue
                raise
            except (
                smtplib.SMTPServerDisconnected,
                smtplib.SMTPConnectError,
                smtplib.SMTPHeloError,
                smtplib.SMTPResponseException,
                socket.timeout,
                TimeoutError,
                OSError,
            ) as exc:
                if isinstance(exc, smtplib.SMTPResponseException) and _is_rate_limited(
                    _decode(exc.smtp_error)
                ):
                    self._close()
                    raise RateLimited(
                        f"服务器提示发送过于频繁：{exc.smtp_code} {_decode(exc.smtp_error)}\n"
                        "建议把发送间隔调大，过一段时间再继续。"
                    ) from exc
                last_error = exc
                self._close()
                if attempt < attempts and not self._stop.is_set():
                    self._log(f"第 {attempt} 次发送失败：{_describe(exc)}，{2 * attempt} 秒后重试…", "warn")
                    if self._stop.wait(2 * attempt):
                        raise MailerError("已手动停止") from exc
                    continue
                break
            except smtplib.SMTPException as exc:
                last_error = exc
                self._close()
                if attempt < attempts and not self._stop.is_set():
                    self._log(f"第 {attempt} 次发送失败：{exc}，{2 * attempt} 秒后重试…", "warn")
                    if self._stop.wait(2 * attempt):
                        raise MailerError("已手动停止") from exc
                    continue
                break

        raise MailerError(f"发送失败：{_describe(last_error)}")

    # -- 主流程 -------------------------------------------------------------
    def run(self) -> list[dict]:
        total = len(self.items)
        started = time.monotonic()
        self._emit("start", total=total)
        aborted = ""

        for pos, item in enumerate(self.items):
            if self._stop.is_set():
                self._mark_rest(pos, "已手动停止")
                aborted = "stopped"
                break

            if self._pause.is_set():
                self._log("已暂停，点「继续」恢复发送。", "warn")
                if self._wait_pause():
                    self._mark_rest(pos, "已手动停止")
                    aborted = "stopped"
                    break

            if self.settings.limit and pos >= self.settings.limit:
                self._mark_rest(pos, f"超出本次上限（{self.settings.limit} 封）")
                aborted = "limit"
                break

            index = pos + 1
            message = build_message(self.settings, item, index, total)
            subject = str(message["Subject"] or "")

            if self.settings.dry_run:
                self._record(item, "试运行", "试运行模式：未真实发送", subject)
                self._emit("item", index=index, total=total, status="试运行", detail="未真实发送", item=item)
            else:
                try:
                    self._send_one(message)
                except AuthFailed as exc:
                    self._record(item, "失败", str(exc), subject)
                    self._emit("item", index=index, total=total, status="失败", detail=str(exc), item=item)
                    self._log("登录失败，已终止本次发送。", "error")
                    self._mark_rest(pos + 1, "登录失败，已终止")
                    aborted = "auth"
                    break
                except RateLimited as exc:
                    self._record(item, "失败", str(exc), subject)
                    self._emit("item", index=index, total=total, status="失败", detail=str(exc), item=item)
                    self._log(f"已被服务器限流，提前结束：{exc}", "error")
                    self._mark_rest(pos + 1, "触发服务器限流，已终止")
                    aborted = "rate"
                    break
                except (MailerError, RecipientRejected) as exc:
                    self._record(item, "失败", str(exc), subject)
                    self._emit("item", index=index, total=total, status="失败", detail=str(exc), item=item)
                    self._log(f"第 {index} 封失败：{exc}", "error")
                else:
                    self._record(item, "成功", "", subject)
                    self._emit(
                        "item",
                        index=index,
                        total=total,
                        status="成功",
                        detail=f"→ {item.email}",
                        item=item,
                    )
                    self._log(f"第 {index}/{total} 封已发送：{item.qq}（{item.email}）", "ok")

            self._emit(
                "progress",
                done=len(self.results),
                total=total,
                ok=self._ok,
                fail=self._fail,
                skipped=self._skipped,
            )

            if index < total and not self._stop.is_set():
                self._sleep_between()

        self._close()
        elapsed = time.monotonic() - started
        self.last_aborted = aborted
        self._emit(
            "progress",
            done=len(self.results),
            total=total,
            ok=self._ok,
            fail=self._fail,
            skipped=self._skipped,
        )
        self._emit(
            "done",
            total=total,
            ok=self._ok,
            fail=self._fail,
            skipped=self._skipped,
            elapsed=elapsed,
            aborted=aborted,
        )
        return self.results

    def _mark_rest(self, start: int, reason: str) -> None:
        rest = self.items[start:]
        for item in rest:
            self._record(item, "未发送", reason)
        if rest:
            self._log(f"剩余 {len(rest)} 个未发送（{reason}）。", "warn")
