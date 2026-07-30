import re
from io import BytesIO
from openpyxl import load_workbook


def escape_markdown(value):
    if value is None:
        return ""

    text = str(value)
    text = text.replace("|", "\\|")
    text = text.replace("\r\n", "<br>")
    text = text.replace("\n", "<br>")

    return text


def format_cell(cell):
    value = cell.value

    if value is None:
        return ""

    number_format = str(cell.number_format) if cell.number_format else "General"

    # Built-in Excel format IDs 9 = "0%", 10 = "0.00%"
    # openpyxl sometimes returns "General" for these instead of resolving the %
    if "%" not in number_format:
        try:
            from openpyxl.styles.numbers import BUILTIN_FORMATS
            fmt_id = cell._style.numFmtId
            builtin_fmt = BUILTIN_FORMATS.get(fmt_id, "")
            if "%" in str(builtin_fmt):
                number_format = builtin_fmt
        except Exception:
            pass

    # Percentage: robust decimal extraction via regex (handles "#,##0.0%" edge case)
    if "%" in number_format and isinstance(value, (int, float)):
        match = re.search(r'\.(\d+)%', number_format)
        decimals = len(match.group(1)) if match else 0
        value = f"{value * 100:.{decimals}f}%"

    return escape_markdown(value)


def worksheet_to_markdown(ws):

    rows = list(ws.iter_rows())

    if not rows:
        return "_Empty Sheet_"

    headers = [format_cell(c) for c in rows[0]]

    lines = []

    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

    for row in rows[1:]:
        values = [format_cell(c) for c in row]
        lines.append("| " + " | ".join(values) + " |")

    return "\n".join(lines)


def extract_excel_to_markdown(excel_bytes: BytesIO, workbook_name: str):

    excel_bytes.seek(0)

    wb = load_workbook(
        excel_bytes,
        data_only=True
    )

    md = []

    md.append(f"# {workbook_name}")
    md.append("")

    for ws in wb.worksheets:

        md.append(f"## {ws.title}")
        md.append("")

        md.append(worksheet_to_markdown(ws))
        md.append("")
        md.append("---")
        md.append("")

    return "\n".join(md)
