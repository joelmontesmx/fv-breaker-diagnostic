import streamlit as st

st.set_page_config(
    page_title="FV Upload Diagnostic",
    layout="wide"
)

st.title("FV Upload Diagnostic")
st.write("This test only checks whether Streamlit can receive PDF files without crashing.")

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

    if st.button("Read uploaded bytes"):
        for file in uploaded_files:
            data = file.getvalue()
            st.write(f"{file.name}: {len(data)} bytes read successfully.")
