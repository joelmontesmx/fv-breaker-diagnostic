from __future__ import annotations

import base64
from html import escape
from pathlib import Path

import streamlit as st

APP_FOLDER = Path(__file__).resolve().parent
PRIMARY_MAPPING_FILE = APP_FOLDER / "part_number_mapping.xlsx"
LEGACY_MAPPING_FILE = APP_FOLDER / "np_equivalencias.xlsx"
MAPPING_FILE = PRIMARY_MAPPING_FILE if PRIMARY_MAPPING_FILE.exists() else LEGACY_MAPPING_FILE

MAX_PDF_FILES = 50
MAX_SINGLE_FILE_MB = 15
MAX_TOTAL_UPLOAD_MB = 100

st.set_page_config(page_title="Front View BOM Extractor", page_icon="⚡", layout="wide")


def human_file_size(size_bytes: int) -> str:
    if size_bytes >= 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{size_bytes / 1024:.0f} KB"


def make_excel_download_link(data: bytes, file_name: str, label: str) -> str:
    encoded = base64.b64encode(data).decode("utf-8")
    return f'''
    <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{encoded}"
       download="{escape(file_name, quote=True)}"
       style="display:inline-block;padding:0.75rem 1.1rem;border:1px solid #ff000f;border-radius:8px;background:#ff000f;color:white;text-decoration:none;font-weight:700;">
       {escape(label)}
    </a>
    '''


def clear_data() -> None:
    for key in ["fv_result", "prepared_excel_bytes", "last_files"]:
        st.session_state.pop(key, None)


def file_batch_errors(uploaded_files) -> list[str]:
    """Validate only metadata at upload time. Do not read file bytes here."""
    errors: list[str] = []
    if not uploaded_files:
        return errors

    total_bytes = sum(int(file.size) for file in uploaded_files)
    if len(uploaded_files) > MAX_PDF_FILES:
        errors.append(f"Maximum batch size is {MAX_PDF_FILES} PDF files.")
    if total_bytes > MAX_TOTAL_UPLOAD_MB * 1024 * 1024:
        errors.append(f"Total batch size must be {MAX_TOTAL_UPLOAD_MB} MB or less.")

    oversized = [file.name for file in uploaded_files if int(file.size) > MAX_SINGLE_FILE_MB * 1024 * 1024]
    if oversized:
        errors.append(f"Each PDF must be {MAX_SINGLE_FILE_MB} MB or smaller. Oversized: {', '.join(oversized)}")

    not_pdf = [file.name for file in uploaded_files if not file.name.lower().endswith(".pdf")]
    if not_pdf:
        errors.append("Only PDF files are allowed: " + ", ".join(not_pdf))
    return errors


def show_file_list(uploaded_files) -> None:
    if not uploaded_files:
        return
    st.caption(f"Selected files: {len(uploaded_files)}")
    visible = uploaded_files[:8]
    for file in visible:
        st.write(f"• {file.name} — {human_file_size(int(file.size))}")
    if len(uploaded_files) > len(visible):
        with st.expander(f"Show all {len(uploaded_files)} selected files"):
            for file in uploaded_files:
                st.write(f"• {file.name} — {human_file_size(int(file.size))}")


def show_summary_text(metrics: dict) -> None:
    st.subheader("2. Processing Summary")
    lines = [
        ("PDFs", metrics.get("pdfs", 0)),
        ("Sales Order-Items", metrics.get("so_items", 0)),
        ("Breaker Rows", metrics.get("breaker_rows", 0)),
        ("Spacer Breakers", metrics.get("spacer_rows", 0)),
        ("Box BOM Rows", metrics.get("box_bom_rows", 0)),
        ("Interior BOM Rows", metrics.get("interior_bom_rows", 0)),
        ("Missing Mappings", metrics.get("unmapped", 0)),
        ("Errors", metrics.get("errors", 0)),
    ]
    st.write(" | ".join(f"**{label}:** {value}" for label, value in lines))


def dataframe_csv_preview(df, max_rows: int = 25) -> str:
    if df is None or df.empty:
        return "No rows."
    return df.head(max_rows).to_csv(index=False)


left, right = st.columns([6, 1])
with left:
    st.title("Front View BOM Extractor")
    st.caption("ABB Electrification · Stability build")
with right:
    st.write("")
    if st.button("Clear Data", use_container_width=True):
        clear_data()
        st.success("Data cleared.")

if not MAPPING_FILE.exists():
    st.error("The part-number mapping database was not found. Place part_number_mapping.xlsx in the same folder as app.py.")
    st.stop()

st.subheader("1. Upload Front Views")
st.caption(f"Batch limits: up to {MAX_PDF_FILES} PDFs, {MAX_SINGLE_FILE_MB} MB per file, and {MAX_TOTAL_UPLOAD_MB} MB total.")
uploaded_files = st.file_uploader("Upload one or more Front View PDFs", type=["pdf"], accept_multiple_files=True)

errors = file_batch_errors(uploaded_files)
if uploaded_files:
    show_file_list(uploaded_files)
    for error in errors:
        st.error(error)

process_clicked = st.button(
    "Process Front Views",
    type="primary",
    use_container_width=True,
    disabled=(not uploaded_files or bool(errors)),
)

if process_clicked:
    try:
        # Import lazily so upload-only rendering does not initialize the engine.
        from breaker_engine import ALL_SECTIONS, UploadedPDF, process_uploaded_pdfs

        pdfs = []
        for file in uploaded_files:
            data = file.getvalue()
            if not data.startswith(b"%PDF-"):
                raise ValueError(f"{file.name} does not appear to be a valid PDF.")
            pdfs.append(UploadedPDF(name=file.name, data=data))

        with st.spinner("Processing Front Views..."):
            dataframes, _, metrics = process_uploaded_pdfs(
                pdfs=pdfs,
                mapping_data=MAPPING_FILE.read_bytes(),
                sections_to_extract=ALL_SECTIONS,
                generate_excel=False,
            )
        st.session_state["fv_result"] = {"dataframes": dataframes, "metrics": metrics}
        st.session_state["last_files"] = [(file.name, int(file.size)) for file in uploaded_files]
        st.session_state.pop("prepared_excel_bytes", None)
        st.success("Front Views processed successfully.")
    except Exception as error:
        st.session_state.pop("fv_result", None)
        st.error("The files could not be processed.")
        with st.expander("Error details"):
            st.exception(error)

result = st.session_state.get("fv_result")
if result:
    dataframes = result["dataframes"]
    metrics = result["metrics"]

    show_summary_text(metrics)

    st.subheader("3. Results Preview")
    st.caption("This stability build avoids heavy interactive grids. Previews are shown as text while we isolate the Streamlit Cloud crash.")
    section_names = [name for name, df in dataframes.items() if getattr(df, "empty", True) is False]
    selected_section = st.selectbox("Preview section", section_names if section_names else ["No data"])
    if selected_section in dataframes:
        st.code(dataframe_csv_preview(dataframes[selected_section], max_rows=30), language="csv")

    st.subheader("4. Excel Report")
    st.caption("Excel generation is separated from processing.")
    if st.button("Prepare Excel Report", type="primary"):
        try:
            from breaker_engine import dataframes_to_excel
            with st.spinner("Preparing Excel report..."):
                st.session_state["prepared_excel_bytes"] = dataframes_to_excel(dataframes)
            st.success("Excel report prepared.")
        except Exception as error:
            st.error("Excel report could not be prepared.")
            with st.expander("Error details"):
                st.exception(error)

    if st.session_state.get("prepared_excel_bytes"):
        st.markdown(
            make_excel_download_link(
                st.session_state["prepared_excel_bytes"],
                "front_view_bom_report.xlsx",
                "Download Excel Report",
            ),
            unsafe_allow_html=True,
        )
