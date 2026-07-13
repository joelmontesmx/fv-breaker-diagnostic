import streamlit as st
from pypdf import PdfReader
from io import BytesIO

st.set_page_config(
    page_title="FV PDF Read Diagnostic",
    layout="wide"
)

st.title("FV PDF Read Diagnostic")
st.write("This test checks whether Streamlit can read text from a PDF using pypdf.")

uploaded_files = st.file_uploader(
    "Upload one or more PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} file(s) uploaded successfully.")

    for file in uploaded_files:
        st.write({
            "file_name": file.name,
            "file_type": file.type,
            "file_size_bytes": file.size,
        })

    if st.button("Read PDF text"):
        for file in uploaded_files:
            try:
                data = file.getvalue()
                reader = PdfReader(BytesIO(data))

                st.subheader(file.name)
                st.write(f"Pages detected: {len(reader.pages)}")

                first_page_text = reader.pages[0].extract_text() or ""
                st.write(f"Characters extracted from page 1: {len(first_page_text)}")

                with st.expander("Preview extracted text"):
                    st.text(first_page_text[:3000])

            except Exception as error:
                st.error(f"Error reading {file.name}")
                st.exception(error)
