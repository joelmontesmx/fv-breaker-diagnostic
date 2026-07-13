import streamlit as st
import pandas as pd
from pypdf import PdfReader
from io import BytesIO

st.set_page_config(
    page_title="FV Excel Diagnostic",
    layout="wide"
)

st.title("FV Excel Diagnostic")
st.write("This test checks PDF text extraction, pandas tables, and Excel generation.")

uploaded_files = st.file_uploader(
    "Upload one or more PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    rows = []

    for file in uploaded_files:
        data = file.getvalue()
        reader = PdfReader(BytesIO(data))

        full_text = ""
        for page in reader.pages:
            full_text += (page.extract_text() or "") + "\n"

        rows.append({
            "File Name": file.name,
            "File Size": file.size,
            "Pages": len(reader.pages),
            "Characters Extracted": len(full_text),
            "Has BOX BOM": "BOX BOM" in full_text,
            "Has INTERIOR BOM": "INTERIOR BOM" in full_text,
            "Has Breaker Catalog Number": "Breaker Catalog Number" in full_text,
        })

    df = pd.DataFrame(rows)

    st.success("PDFs processed successfully.")
    st.dataframe(df, width="stretch")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Diagnostic")

    st.download_button(
        "Download diagnostic Excel",
        data=output.getvalue(),
        file_name="diagnostic.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
