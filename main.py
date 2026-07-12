import streamlit as st
from io import BytesIO
from pathlib import Path

from sexcel_to_md import extract_excel_to_markdown

st.set_page_config(
    page_title="Excel → Markdown",
    layout="wide"
)

st.title("📊 Excel → Markdown Converter")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    st.write(f"Total Excel files uploaded: {len(uploaded_files)}")

    for idx, uploaded_file in enumerate(uploaded_files):

        st.divider()

        st.subheader(f"📄 {uploaded_file.name}")

        excel_bytes = BytesIO(uploaded_file.read())

        md_text = extract_excel_to_markdown(
            excel_bytes,
            Path(uploaded_file.name).stem
        )

        output_filename = (
            Path(uploaded_file.name).stem + ".md"
        )

        st.text_area(
            "Preview",
            md_text,
            height=300,
            key=f"preview_{idx}"
        )

        st.download_button(
            label=f"Download {output_filename}",
            data=md_text,
            file_name=output_filename,
            mime="text/markdown",
            key=f"download_{idx}"
        )