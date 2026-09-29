"""
条形码生成模块：读取表格C的指定列，把每个值生成条形码图片导出到指定目录。
依赖：python-barcode, Pillow
"""
import os
import re
import sys

from openpyxl import load_workbook

from barcode import get_barcode_class
from barcode.writer import ImageWriter

# 仅支持纯数字的条形码类型（需要 12/13 位等），其余类型（code128 等）可接受任意字符
_DIGIT_ONLY = {"ean13", "ean8", "upca", "isbn13", "issn", "ean", "upc", "isbn"}


def _find_font() -> str | None:
    """定位一个可用的 TTF 字体，用于条形码下方的人眼可读文字。
    优先级：打包内置字体 > Windows 系统字体。找不到返回 None（调用方将关闭文字以保证能生成）。
    """
    candidates = []
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, "fonts", "DejaVuSans.ttf"))
        candidates.append(os.path.join(sys._MEIPASS, "arial.ttf"))
    candidates += [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\calibri.ttf",
        r"C:\Windows\Fonts\msyh.ttc",
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


def _safe_name(value: str, max_len: int = 80) -> str:
    s = str(value).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    return s[:max_len]


def get_headers(path: str, header_row: int = 1) -> list:
    """读取工作簿第一个工作表的指定行作为表头，返回 list[str]。
    header_row 为 1 基行号（默认 1；若列名在第 2 行则传 2）。
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    idx = max(0, header_row - 1)
    if idx >= len(rows):
        return []
    first = rows[idx]
    return [str(c) if c is not None else f"列{i + 1}" for i, c in enumerate(first)]


def analyze_column(
    input_path: str,
    col_index: int,
    start_row: int = 2,
) -> dict:
    """扫描指定列（从 start_row 行起）的内容，返回统计：
        {
            "non_empty": 非空单元格数量,
            "non_ascii": 含非ASCII(中文等)字符的数量,
        }
    用于在生成前判断该列是否适合做一维条形码。
    """
    wb = load_workbook(input_path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    non_empty = 0
    non_ascii = 0
    start = max(0, start_row - 1)
    for i, r in enumerate(rows):
        if i < start:
            continue
        if col_index >= len(r):
            continue
        v = r[col_index]
        if v is None or str(v).strip() == "":
            continue
        non_empty += 1
        try:
            str(v).encode("ascii")
        except UnicodeEncodeError:
            non_ascii += 1
    return {"non_empty": non_empty, "non_ascii": non_ascii}


def generate_barcodes(
    input_path: str,
    col_index: int,
    output_dir: str,
    barcode_type: str = "code128",
    with_text: bool = True,
    start_row: int = 2,
) -> dict:
    """
    把表格指定列（从 start_row 行开始）的每个值生成条形码图片并保存到 output_dir。
    start_row 为 1 基的行号：1 表示从首行开始（不跳过），2 表示跳过首行（默认，首行为表头）。
    返回统计字典：
        {
            "count": 成功生成数量,
            "skipped": [被跳过的 (行号, 值, 原因), ...],
            "output_dir": 输出目录,
        }
    """
    wb = load_workbook(input_path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    os.makedirs(output_dir, exist_ok=True)

    writer = ImageWriter()
    font = _find_font()
    # 先尝试加载字体；若打包环境下字体加载失败，回退到关闭文字以避免乱码/重叠。
    if with_text and font:
        try:
            from PIL import ImageFont
            test_size = int(300 * 11 / 72)  # 约等于 11pt @ 300dpi
            ImageFont.truetype(font, test_size)
        except Exception:  # noqa: BLE001
            font = None

    # text_distance：条码底部到文字基线的距离(毫米)。需明显大于字体高度。
    # 实测 arial 11pt 字高约 2.7mm，这里给 10.0mm 足够任何常见字体都不重叠。
    options = {
        "write_text": with_text,
        "font_size": 10,
        "text_distance": 10.0,
        "module_height": 12.0,  # 条码条高(毫米)
        "quiet_zone": 6.5,      # 左右静区(毫米)
    }
    if font:
        options["font_path"] = font
    elif with_text:
        # 无可用字体时关闭文字，保证条形码本体仍能生成
        options["write_text"] = False
    cls = get_barcode_class(barcode_type)

    count = 0
    skipped = []
    used = {}

    start = max(0, start_row - 1)  # 转换为 0 基跳过行数
    for i, r in enumerate(rows):
        if i < start:
            continue
        if col_index >= len(r):
            continue
        val = r[col_index]
        if val is None or str(val).strip() == "":
            continue
        val = str(val).strip()

        # 数字型条码校验
        if barcode_type in _DIGIT_ONLY:
            digits = re.sub(r"\D", "", val)
            valid = False
            if barcode_type in ("ean8",):
                valid = len(digits) in (7, 8)
            elif barcode_type in ("upca",):
                valid = len(digits) in (11, 12)
            else:  # ean13 / isbn13 / issn 等需要 12 或 13 位
                valid = len(digits) in (12, 13)
            if not valid:
                skipped.append((i + 1, val, "数字长度不符合该条码类型要求"))
                continue
            val = digits

        try:
            rv = cls(val, writer=writer)
        except Exception as e:  # noqa: BLE001
            skipped.append((i + 1, val, f"生成失败: {e}"))
            continue

        base = _safe_name(val) or f"row{i + 1}"
        name = base
        if name in used:
            used[name] += 1
            name = f"{base}_{used[name]}"
        else:
            used[name] = 0

        fpath = os.path.join(output_dir, name)  # save 会自动补 .png
        try:
            rv.save(fpath, options=options)
            count += 1
        except Exception as e:  # noqa: BLE001
            skipped.append((i + 1, val, f"保存失败: {e}"))

    return {"count": count, "skipped": skipped, "output_dir": output_dir}
