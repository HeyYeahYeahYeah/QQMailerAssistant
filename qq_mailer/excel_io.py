"""Excel 名单读取、校验与结果回写。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

QQ_PATTERN = re.compile(r"^\d{5,12}$")
HEADER_WORDS = {
    "qq",
    "qq号",
    "qq号码",
    "qq帐号",
    "qq账号",
    "qq邮箱",
    "号码",
    "账号",
    "帐号",
    "qqnumber",
    "qq_number",
    "qq number",
    "收件人",
    "收件qq",
}

_FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
_ZERO_WIDTH = dict.fromkeys(map(ord, "\ufeff\u200b\u200c\u200d\xa0"), None)


class ExcelError(Exception):
    """表格读取 / 校验失败，消息可以直接展示给用户。"""


@dataclass
class RowItem:
    row_no: int
    qq: str
    email: str
    content: str


@dataclass
class SkippedRow:
    row_no: int
    raw_qq: str
    reason: str


@dataclass
class ParseResult:
    rows: list[RowItem] = field(default_factory=list)
    skipped: list[SkippedRow] = field(default_factory=list)
    sheet: str = ""
    sheet_names: list[str] = field(default_factory=list)
    duplicates: int = 0
    blank: int = 0

    @property
    def total_source_rows(self) -> int:
        return len(self.rows) + len(self.skipped)


# --------------------------------------------------------------------------- #
# 读取
# --------------------------------------------------------------------------- #
def read_sheet_names(path: str | Path) -> list[str]:
    try:
        workbook = load_workbook(filename=str(path), read_only=True, data_only=True)
    except FileNotFoundError as exc:
        raise ExcelError(f"找不到文件：{path}") from exc
    except PermissionError as exc:
        raise ExcelError("文件被占用或没有读取权限，请先在 Excel 里关闭它再试。") from exc
    except InvalidFileException as exc:
        raise ExcelError(
            "这个文件不是 .xlsx 格式（可能是旧版 .xls 或改了后缀）。\n"
            "请用 Excel 打开后「另存为 → Excel 工作簿(*.xlsx)」再试。"
        ) from exc
    except Exception as exc:  # 损坏、加密、非 xlsx 等
        raise ExcelError(f"无法打开这个 Excel 文件：{exc}") from exc
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value)
    return str(value)


def normalize_qq(raw: str) -> str:
    """把单元格里的 QQ 号清洗成纯数字串（去空格、全角、引号、小数点尾巴）。"""
    text = raw.translate(_FULLWIDTH_DIGITS).translate(_ZERO_WIDTH)
    text = text.strip().strip("'\"‘’“”").strip()
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"\.0+$", "", text)
    return text


def _looks_like_header(text: str) -> bool:
    return text.strip().lower().replace(" ", "") in HEADER_WORDS


def parse_workbook(
    path: str | Path,
    sheet: str | None = None,
    suffix: str = "@qq.com",
    dedupe: bool = False,
) -> ParseResult:
    """读取表格：A 列 QQ 号，B 列 文字内容。首行可以是标题行。"""
    path = Path(path)
    if not path.is_file():
        raise ExcelError(f"找不到文件：{path}")

    suffix = normalize_suffix(suffix)
    try:
        workbook = load_workbook(filename=str(path), read_only=True, data_only=True)
    except PermissionError as exc:
        raise ExcelError("文件被占用或没有读取权限，请先在 Excel 里关闭它再试。") from exc
    except InvalidFileException as exc:
        raise ExcelError(
            "这个文件不是 .xlsx 格式（可能是旧版 .xls 或改了后缀）。\n"
            "请用 Excel 打开后「另存为 → Excel 工作簿(*.xlsx)」再试。"
        ) from exc
    except Exception as exc:
        raise ExcelError(f"无法打开这个 Excel 文件：{exc}") from exc

    try:
        sheet_names = list(workbook.sheetnames)
        if not sheet_names:
            raise ExcelError("这个工作簿里没有任何工作表。")
        if sheet:
            if sheet not in sheet_names:
                raise ExcelError(f"找不到工作表「{sheet}」，现有的工作表：{'、'.join(sheet_names)}")
            worksheet = workbook[sheet]
        else:
            worksheet = workbook[sheet_names[0]]

        raw_rows: list[tuple[int, str, str]] = []
        for index, row in enumerate(
            worksheet.iter_rows(min_row=1, max_col=2, values_only=True), start=1
        ):
            if len(raw_rows) > 20000:
                raise ExcelError("表格行数超过 2 万行，请先拆分文件。")
            first = _cell_text(row[0]) if len(row) > 0 else ""
            second = _cell_text(row[1]) if len(row) > 1 else ""
            raw_rows.append((index, first, second))
        sheet_title = getattr(worksheet, "title", "") or (sheet or "")
    finally:
        workbook.close()

    result = ParseResult(sheet=sheet_title, sheet_names=sheet_names)
    if not raw_rows:
        raise ExcelError("表格是空的，没有任何内容。")

    start_index = _detect_header_offset(raw_rows)
    seen: dict[str, int] = {}

    for row_no, first, second in raw_rows[start_index:]:
        if not first.strip() and not second.strip():
            result.blank += 1
            continue
        qq = normalize_qq(first)
        if not qq:
            result.skipped.append(SkippedRow(row_no, first, "第一列（QQ 号）为空"))
            continue
        if "@" in qq:
            result.skipped.append(
                SkippedRow(row_no, first, f"第一列是邮箱地址「{first}」，这里需要纯 QQ 号")
            )
            continue
        if not qq.isdigit():
            result.skipped.append(
                SkippedRow(row_no, first, f"第一列「{first}」不是纯数字 QQ 号")
            )
            continue
        if not QQ_PATTERN.match(qq):
            result.skipped.append(
                SkippedRow(row_no, first, f"QQ 号「{qq}」位数不对（正常为 5–12 位数字）")
            )
            continue
        if not second.strip():
            result.skipped.append(SkippedRow(row_no, first, "第二列（文字内容）为空，未发送"))
            continue

        if qq in seen:
            result.duplicates += 1
            if dedupe:
                result.skipped.append(
                    SkippedRow(row_no, first, f"重复的 QQ 号（第 {seen[qq]} 行已出现过），按设置跳过")
                )
                continue
        else:
            seen[qq] = row_no

        result.rows.append(RowItem(row_no=row_no, qq=qq, email=f"{qq}{suffix}", content=second))

    if not result.rows:
        hint = ""
        if result.skipped:
            swapped = [
                item
                for item in result.skipped
                if not item.raw_qq.strip().isdigit()
            ]
            if len(swapped) >= max(1, len(result.skipped) * 0.6):
                hint = "\n提示：看起来第一列不是 QQ 号，可能两列顺序放反了（要求 A 列 QQ 号、B 列 内容）。"
        raise ExcelError(
            "没有读到任何可发送的行。\n"
            "请确认：A 列是 QQ 号（纯数字），B 列是要发送的文字内容，第一行可以是标题。"
            + hint
        )
    return result


def _detect_header_offset(raw_rows: Sequence[tuple[int, str, str]]) -> int:
    """返回第一条数据行的下标（0 基）。首行是标题时返回 1。"""
    populated = [i for i, (_, a, b) in enumerate(raw_rows) if a.strip() or b.strip()]
    if not populated:
        return 0
    first = populated[0]
    first_qq = normalize_qq(raw_rows[first][1])
    if QQ_PATTERN.match(first_qq):
        return 0
    if _looks_like_header(raw_rows[first][1]):
        return first + 1
    if len(populated) > 1:
        second_qq = normalize_qq(raw_rows[populated[1]][1])
        if QQ_PATTERN.match(second_qq):
            return first + 1
    return 0


def normalize_suffix(suffix: str) -> str:
    text = (suffix or "").strip().replace(" ", "")
    if not text:
        return "@qq.com"
    if not text.startswith("@"):
        text = "@" + text.lstrip("@")
    if "." not in text[1:]:
        raise ExcelError(f"邮箱后缀「{suffix}」看起来不对，应该像 @qq.com 这样。")
    return text


# --------------------------------------------------------------------------- #
# 回写结果
# --------------------------------------------------------------------------- #
REPORT_HEADERS = ["行号", "QQ 号", "收件邮箱", "状态", "发送时间", "说明"]
HEADER_FILL = PatternFill("solid", fgColor="DCE8F7")
STATUS_FILL = {
    "成功": PatternFill("solid", fgColor="D6F0D6"),
    "失败": PatternFill("solid", fgColor="FAD9D9"),
    "未发送": PatternFill("solid", fgColor="EFEFEF"),
    "试运行": PatternFill("solid", fgColor="FFF3CD"),
}
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_report(path: str | Path, results: Iterable[dict[str, Any]]) -> Path:
    """把每行发送结果写成 xlsx，返回文件路径。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "发送结果"
    sheet.append(REPORT_HEADERS)

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    rows = list(results)
    for item in rows:
        sheet.append(
            [
                item.get("row_no", ""),
                item.get("qq", ""),
                item.get("email", ""),
                item.get("status", ""),
                item.get("time", ""),
                item.get("detail") or "",
            ]
        )
        status = str(item.get("status", ""))
        fill = STATUS_FILL.get(status)
        for cell in sheet[sheet.max_row]:
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=False)
        if fill is not None:
            sheet.cell(row=sheet.max_row, column=4).fill = fill

    widths = [8, 14, 28, 10, 20, 60]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    if rows:
        sheet.auto_filter.ref = f"A1:F{sheet.max_row}"

    summary = workbook.create_sheet("统计")
    counts: dict[str, int] = {}
    for item in rows:
        key = str(item.get("status", "") or "未知")
        counts[key] = counts.get(key, 0) + 1
    summary.append(["项目", "数量"])
    summary.append(["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    for key in ("成功", "失败", "未发送", "试运行"):
        summary.append([key, counts.get(key, 0)])
    summary.append(["合计", len(rows)])
    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 28
    for cell in summary[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL

    workbook.save(str(path))
    return path
