"""离线端到端测试：自建一个最小 SMTP 服务器，验证解析与发送逻辑。

运行：python -m unittest discover -s tests -t . -v
"""

from __future__ import annotations

import base64
import shutil
import socket
import tempfile
import threading
import time
import unittest
import uuid
from email import policy
from email.parser import BytesParser
from pathlib import Path

from openpyxl import Workbook, load_workbook

from qq_mailer.excel_io import ExcelError, parse_workbook, write_report
from qq_mailer.mailer import (
    AuthFailed,
    MailerSettings,
    Sender,
    build_message,
    content_preview,
    render_subject,
)
from qq_mailer.excel_io import RowItem


# --------------------------------------------------------------------------- #
# 迷你 SMTP 服务器
# --------------------------------------------------------------------------- #
class FakeSMTPServer:
    """只实现 smtplib 会用到的那几条命令，用来离线验证发送流程。"""

    def __init__(self, *, accept_auth: bool = True, drop_first_data: bool = False) -> None:
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(8)
        self._sock.settimeout(0.2)
        self.port = self._sock.getsockname()[1]
        self.accept_auth = accept_auth
        self.drop_first_data = drop_first_data
        self.messages: list[dict] = []
        self.auth_attempts = 0
        self.connections = 0
        self._data_failures = 0
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._serve, daemon=True)

    def start(self) -> None:
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        try:
            self._sock.close()
        except OSError:
            pass
        self._thread.join(timeout=2)

    def _serve(self) -> None:
        while not self._stop.is_set():
            try:
                conn, _ = self._sock.accept()
            except (socket.timeout, OSError):
                continue
            self.connections += 1
            try:
                self._handle(conn)
            except Exception:
                pass
            finally:
                try:
                    conn.close()
                except OSError:
                    pass

    def _handle(self, conn: socket.socket) -> None:
        reader = conn.makefile("rb")
        conn.sendall(b"220 fake.local ESMTP\r\n")
        envelope: dict[str, str] = {}
        while True:
            raw = reader.readline()
            if not raw:
                return
            command = raw.decode("utf-8", "replace").strip()
            upper = command.upper()
            if upper.startswith(("EHLO", "HELO")):
                conn.sendall(
                    b"250-fake.local\r\n250-SIZE 33554432\r\n250-AUTH PLAIN LOGIN\r\n250 HELP\r\n"
                )
            elif upper.startswith("AUTH PLAIN"):
                parts = command.split()
                if len(parts) >= 3:
                    self._check_auth(parts[2], conn)
                else:
                    conn.sendall(b"334 \r\n")
                    reply = reader.readline().strip()
                    self._check_auth(reply.decode("ascii", "replace"), conn)
            elif upper.startswith("AUTH LOGIN"):
                conn.sendall(b"334 VXNlcm5hbWU6\r\n")
                user = reader.readline().strip()
                conn.sendall(b"334 UGFzc3dvcmQ6\r\n")
                password = reader.readline().strip()
                self._check_auth_login(user, password, conn)
            elif upper.startswith("MAIL FROM"):
                envelope["from"] = command
                conn.sendall(b"250 OK\r\n")
            elif upper.startswith("RCPT TO"):
                envelope["to"] = command
                conn.sendall(b"250 OK\r\n")
            elif upper.startswith("DATA"):
                if self.drop_first_data and self._data_failures == 0:
                    self._data_failures += 1
                    try:
                        conn.shutdown(socket.SHUT_RDWR)
                    except OSError:
                        pass
                    return
                conn.sendall(b"354 End data with <CR><LF>.<CR><LF>\r\n")
                chunks: list[bytes] = []
                while True:
                    line = reader.readline()
                    if not line:
                        return
                    if line == b".\r\n":
                        break
                    chunks.append(line)
                payload = b"".join(chunks)
                message = BytesParser(policy=policy.default).parsebytes(payload)
                self.messages.append(
                    {
                        "envelope": dict(envelope),
                        "to": message["To"],
                        "subject": message["Subject"],
                        "body": message.get_content().strip(),
                        "from": message["From"],
                    }
                )
                conn.sendall(b"250 OK queued\r\n")
            elif upper.startswith("RSET") or upper.startswith("NOOP"):
                conn.sendall(b"250 OK\r\n")
            elif upper.startswith("QUIT"):
                conn.sendall(b"221 Bye\r\n")
                return
            else:
                conn.sendall(b"250 OK\r\n")

    def _check_auth(self, payload: str, conn: socket.socket) -> None:
        self.auth_attempts += 1
        try:
            decoded = base64.b64decode(payload).decode("utf-8", "replace")
        except Exception:
            decoded = ""
        parts = decoded.split("\x00")
        ok = self.accept_auth and len(parts) >= 3 and parts[1] and parts[2]
        self._reply_auth(ok, conn)

    def _check_auth_login(self, user: bytes, password: bytes, conn: socket.socket) -> None:
        self.auth_attempts += 1
        ok = self.accept_auth and bool(user) and bool(password)
        self._reply_auth(ok, conn)

    @staticmethod
    def _reply_auth(ok: bool, conn: socket.socket) -> None:
        if ok:
            conn.sendall(b"235 2.7.0 Authentication successful\r\n")
        else:
            conn.sendall(b"535 5.7.8 Error: authentication failed\r\n")


# --------------------------------------------------------------------------- #
# 测试工具
# --------------------------------------------------------------------------- #
def make_items(count: int, start: int = 1) -> list[RowItem]:
    return [
        RowItem(
            row_no=index + 1,
            qq=f"1000{index:04d}",
            email=f"1000{index:04d}@qq.com",
            content=f"第 {index} 条内容\n第二行",
        )
        for index in range(start, start + count)
    ]


def make_settings(port: int, **overrides) -> MailerSettings:
    settings = MailerSettings(
        host="127.0.0.1",
        port=port,
        use_ssl=False,
        starttls=False,
        username="sender@qq.com",
        password="abcdefghijklmnop",
        display_name="测试发件人",
        subject_template="通知 #{序号}",
        interval=0.0,
        retries=1,
        timeout=5.0,
    )
    for key, value in overrides.items():
        setattr(settings, key, value)
    return settings


class ParseTests(unittest.TestCase):
    def setUp(self) -> None:
        workspace = Path(__file__).resolve().parents[1]
        scratch = workspace / "build"
        scratch.mkdir(parents=True, exist_ok=True)
        # 不用 tempfile.mkdtemp：它在 Windows 沙箱下会生成受限 ACL 的目录
        self.dir = scratch / f"parse-{uuid.uuid4().hex[:8]}"
        self.dir.mkdir(parents=True, exist_ok=True)
        self.addCleanup(lambda: shutil.rmtree(self.dir, ignore_errors=True))

    def _workbook(self, rows, name: str = "名单.xlsx") -> Path:
        path = self.dir / name
        book = Workbook()
        sheet = book.active
        sheet.title = "名单"
        for row in rows:
            sheet.append(list(row))
        book.save(path)
        return path

    def test_header_and_rows(self) -> None:
        path = self._workbook(
            [
                ("QQ号", "内容"),
                (123456789, "你好，这是第一条\n带换行"),
                ("987654321", "第二条"),
                (None, None),
            ]
        )
        result = parse_workbook(path)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[0].qq, "123456789")
        self.assertEqual(result.rows[0].email, "123456789@qq.com")
        self.assertIn("带换行", result.rows[0].content)
        self.assertEqual(result.rows[1].qq, "987654321")
        self.assertEqual(result.blank, 1)
        self.assertEqual(result.sheet, "名单")

    def test_no_header(self) -> None:
        path = self._workbook([(123456789, "内容一"), (234567890, "内容二")])
        result = parse_workbook(path)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[0].row_no, 1)

    def test_cleanup_and_skips(self) -> None:
        path = self._workbook(
            [
                ("QQ号", "内容"),
                (" １２３４５６７８９ ", "全角数字"),
                ("'234567890", "带引号"),
                ("abc", "不是号码"),
                ("123456789", ""),
                ("someone@qq.com", "邮箱直接写在第一列"),
                ("1000", "位数太短"),
            ]
        )
        result = parse_workbook(path)
        self.assertEqual([row.qq for row in result.rows], ["123456789", "234567890"])
        reasons = " ".join(item.reason for item in result.skipped)
        self.assertIn("不是纯数字", reasons)
        self.assertIn("为空", reasons)
        self.assertIn("邮箱地址", reasons)
        self.assertIn("位数", reasons)

    def test_duplicates_kept_by_default(self) -> None:
        path = self._workbook([("QQ号", "内容"), (123456789, "一"), (123456789, "二")])
        result = parse_workbook(path)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.duplicates, 1)

    def test_duplicates_deduped(self) -> None:
        path = self._workbook([("QQ号", "内容"), (123456789, "一"), (123456789, "二")])
        result = parse_workbook(path, dedupe=True)
        self.assertEqual(len(result.rows), 1)
        self.assertEqual(len(result.skipped), 1)
        self.assertIn("重复", result.skipped[0].reason)

    def test_suffix(self) -> None:
        path = self._workbook([("QQ号", "内容"), (123456789, "一")])
        result = parse_workbook(path, suffix="corp.example.com")
        self.assertEqual(result.rows[0].email, "123456789@corp.example.com")

    def test_header_only_raises(self) -> None:
        path = self._workbook([("QQ号", "内容")])
        with self.assertRaises(ExcelError):
            parse_workbook(path)

    def test_missing_file(self) -> None:
        with self.assertRaises(ExcelError):
            parse_workbook(self.dir / "不存在.xlsx")

    def test_write_report(self) -> None:
        path = self.dir / "结果.xlsx"
        results = [
            {"row_no": 2, "qq": "123456789", "email": "123456789@qq.com", "status": "成功", "time": "2026-01-01 10:00:00", "detail": ""},
            {"row_no": 3, "qq": "234567890", "email": "234567890@qq.com", "status": "失败", "time": "2026-01-01 10:00:01", "detail": "收件地址被拒绝"},
        ]
        write_report(path, results)
        book = load_workbook(path)
        self.assertIn("发送结果", book.sheetnames)
        sheet = book["发送结果"]
        self.assertEqual(sheet["A1"].value, "行号")
        self.assertEqual(sheet["D2"].value, "成功")
        self.assertEqual(sheet["F3"].value, "收件地址被拒绝")
        self.assertIn("统计", book.sheetnames)


class MessageTests(unittest.TestCase):
    def test_subject_placeholders(self) -> None:
        item = RowItem(row_no=2, qq="123456789", email="123456789@qq.com", content="第一行\n第二行很长很长")
        subject = render_subject("{qq} 的通知（{序号}/{总数}）", item, 3, 10)
        self.assertEqual(subject, "123456789 的通知（3/10）")

    def test_subject_sanitizes_newlines(self) -> None:
        item = RowItem(row_no=2, qq="123456789", email="123456789@qq.com", content="x")
        subject = render_subject("标题\r\nBcc: evil@example.com", item, 1, 1)
        self.assertNotIn("\n", subject)
        self.assertNotIn("\r", subject)

    def test_content_preview(self) -> None:
        self.assertEqual(content_preview("a\nb   c"), "a b c")
        self.assertTrue(content_preview("x" * 100).endswith("…"))

    def test_message_headers(self) -> None:
        settings = make_settings(465)
        item = RowItem(row_no=2, qq="123456789", email="123456789@qq.com", content="正文内容")
        message = build_message(settings, item, 1, 1)
        self.assertEqual(message["To"], "123456789@qq.com")
        self.assertIn("sender@qq.com", message["From"])
        self.assertEqual(message.get_content().strip(), "正文内容")


class SendTests(unittest.TestCase):
    def setUp(self) -> None:
        self.server = FakeSMTPServer()
        self.server.start()
        self.addCleanup(self.server.stop)

    def test_send_all(self) -> None:
        events: list[dict] = []
        sender = Sender(make_settings(self.server.port), make_items(3), on_event=events.append)
        results = sender.run()

        self.assertEqual([row["status"] for row in results], ["成功", "成功", "成功"])
        self.assertEqual(len(self.server.messages), 3)
        self.assertEqual(self.server.messages[0]["to"], "10000001@qq.com")
        self.assertEqual(self.server.messages[0]["body"].replace("\r\n", "\n"), "第 1 条内容\n第二行")
        self.assertEqual(self.server.messages[0]["subject"], "通知 #1")
        self.assertIn("10000003@qq.com", self.server.messages[2]["envelope"]["to"])
        done = [event for event in events if event["type"] == "done"]
        self.assertEqual(done[0]["ok"], 3)
        self.assertEqual(done[0]["aborted"], "")

    def test_auth_failed_aborts(self) -> None:
        self.server.accept_auth = False
        sender = Sender(make_settings(self.server.port), make_items(3))
        results = sender.run()
        statuses = [row["status"] for row in results]
        self.assertEqual(statuses, ["失败", "未发送", "未发送"])
        self.assertIn("授权码", results[0]["detail"])
        self.assertEqual(sender.last_aborted, "auth")
        self.assertEqual(len(self.server.messages), 0)

    def test_retry_after_disconnect(self) -> None:
        self.server.drop_first_data = True
        sender = Sender(make_settings(self.server.port, retries=2), make_items(1))
        results = sender.run()
        self.assertEqual(results[0]["status"], "成功", results[0]["detail"])
        self.assertEqual(len(self.server.messages), 1)
        self.assertGreaterEqual(self.server.connections, 2)

    def test_stop_before_start(self) -> None:
        sender = Sender(make_settings(self.server.port), make_items(3))
        sender.stop()
        results = sender.run()
        self.assertEqual([row["status"] for row in results], ["未发送"] * 3)
        self.assertEqual(sender.last_aborted, "stopped")
        self.assertEqual(len(self.server.messages), 0)

    def test_stop_during_run(self) -> None:
        sender = Sender(
            make_settings(self.server.port, interval=0.6, retries=0), make_items(6)
        )
        thread = threading.Thread(target=sender.run)
        thread.start()
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline and len(self.server.messages) < 2:
            time.sleep(0.02)
        sender.stop()
        thread.join(timeout=10)
        sent = len(self.server.messages)
        self.assertLess(sent, 6)
        statuses = [row["status"] for row in sender.results]
        self.assertEqual(len(statuses), 6)
        self.assertIn("未发送", statuses)

    def test_limit(self) -> None:
        sender = Sender(make_settings(self.server.port, limit=2), make_items(4))
        results = sender.run()
        self.assertEqual([row["status"] for row in results], ["成功", "成功", "未发送", "未发送"])
        self.assertEqual(len(self.server.messages), 2)
        self.assertEqual(sender.last_aborted, "limit")

    def test_dry_run_does_not_connect(self) -> None:
        settings = make_settings(1, dry_run=True)  # 端口 1 不可用，dry run 不应连接
        sender = Sender(settings, make_items(2))
        results = sender.run()
        self.assertEqual([row["status"] for row in results], ["试运行", "试运行"])
        self.assertEqual(self.server.messages, [])

    def test_single_failure_continues(self) -> None:
        from qq_mailer.mailer import MailerError

        sender = Sender(make_settings(self.server.port, retries=0), make_items(2))

        def boom(message):  # noqa: ANN001
            raise MailerError("模拟的单封失败")

        sender._send_one = boom  # type: ignore[method-assign]
        results = sender.run()
        self.assertEqual([row["status"] for row in results], ["失败", "失败"])
        self.assertIn("模拟的单封失败", results[0]["detail"])
        self.assertEqual(sender.last_aborted, "")


class OutputFallbackTests(unittest.TestCase):
    """结果文件写不进去时，应该自动换下一个目录，而不是直接失败。"""

    def setUp(self) -> None:
        workspace = Path(__file__).resolve().parents[1]
        scratch = workspace / "build"
        scratch.mkdir(parents=True, exist_ok=True)
        self.dir = scratch / f"fallback-{uuid.uuid4().hex[:8]}"
        self.dir.mkdir(parents=True, exist_ok=True)
        self.addCleanup(lambda: shutil.rmtree(self.dir, ignore_errors=True))

    def test_falls_back_to_next_writable_dir(self) -> None:
        from qq_mailer.ui import write_with_fallback

        folders = [self.dir / "a", self.dir / "b", self.dir / "c"]

        def writer(folder: Path) -> Path:
            if folder.name in {"a", "b"}:
                raise PermissionError(13, "拒绝访问")
            return folder / "发送结果.xlsx"

        result = write_with_fallback(folders, writer)
        self.assertEqual(result.parent.name, "c")
        self.assertEqual(result.name, "发送结果.xlsx")

    def test_first_dir_wins_when_writable(self) -> None:
        from qq_mailer.ui import write_with_fallback

        folders = [self.dir / "a", self.dir / "b"]
        result = write_with_fallback(folders, lambda folder: folder / "r.xlsx")
        self.assertEqual(result.parent.name, "a")

    def test_all_fail_raises_with_reasons(self) -> None:
        from qq_mailer.ui import write_with_fallback

        folders = [self.dir / "a", self.dir / "b"]

        def writer(folder: Path) -> Path:
            raise PermissionError(13, "拒绝访问")

        with self.assertRaises(OSError) as ctx:
            write_with_fallback(folders, writer)
        message = str(ctx.exception)
        self.assertIn("都无法写入", message)
        self.assertIn("a", message)
        self.assertIn("b", message)


if __name__ == "__main__":
    unittest.main()
