from io import BytesIO
import pandas as pd


def escape_markdown(value):

    if pd.isna(value):
        return ""

    text = str(value)
    text = text.replace("|", "\\|")
    text = text.replace("\r\n", "<br>")
    text = text.replace("\n", "<br>")

    return text


def dataframe_to_markdown(df):

    headers = [escape_markdown(col) for col in df.columns]

    lines = []

    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

    for _, row in df.iterrows():
        values = [escape_markdown(v) for v in row]
        lines.append("| " + " | ".join(values) + " |")

    return "\n".join(lines)


def extract_excel_to_markdown(excel_bytes: BytesIO, workbook_name: str) -> str:

    xls = pd.ExcelFile(excel_bytes)

    md = []

    md.append(f"# {workbook_name}")
    md.append("")

    for sheet in xls.sheet_names:

        md.append(f"## {sheet}")
        md.append("")

        df = pd.read_excel(
            excel_bytes,
            sheet_name=sheet,
            dtype=object
        )

        if df.empty and len(df.columns) == 0:

            md.append("_Empty Sheet_")
            md.append("")
            md.append("---")
            md.append("")
            continue

        md.append(dataframe_to_markdown(df))
        md.append("")
        md.append("---")
        md.append("")

        excel_bytes.seek(0)

    return "\n".join(md)
