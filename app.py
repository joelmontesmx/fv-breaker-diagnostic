from __future__ import annotations

import base64
from html import escape
from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st

APP_FOLDER = Path(__file__).resolve().parent
PRIMARY_MAPPING_FILE = APP_FOLDER / "part_number_mapping.xlsx"
LEGACY_MAPPING_FILE = APP_FOLDER / "np_equivalencias.xlsx"
MAPPING_FILE = PRIMARY_MAPPING_FILE if PRIMARY_MAPPING_FILE.exists() else LEGACY_MAPPING_FILE

MAX_PDF_FILES = 50
MAX_SINGLE_FILE_MB = 15
MAX_TOTAL_UPLOAD_MB = 100

ABB_RED = "#FF000F"
ABB_LILAC = "#6764f6"
ABB_RED_LIGHT = "#ff957e"
ABB_RED_PALE = "#ffdccd"
ABB_LILAC_LIGHT = "#93a1ff"
ABB_LILAC_PALE = "#e4e7ff"
DARK = "#1f1f2e"
GREY_BG = "#f5f6f8"
GREY_BORDER = "#e3e6ea"

st.set_page_config(page_title="Front View BOM Extractor", page_icon="⚡", layout="wide")


# -----------------------------
# Stable ABB-inspired styling
# -----------------------------
st.markdown(
    f"""
    <style>
        :root {{
            --abb-red: {ABB_RED};
            --abb-lilac: {ABB_LILAC};
            --abb-red-light: {ABB_RED_LIGHT};
            --abb-red-pale: {ABB_RED_PALE};
            --abb-lilac-pale: {ABB_LILAC_PALE};
            --abb-dark: {DARK};
            --abb-bg: {GREY_BG};
            --abb-border: {GREY_BORDER};
        }}

        .stApp {{
            background: linear-gradient(180deg, #ffffff 0%, #fafafa 52%, #f7f8fb 100%);
        }}

        .block-container {{
            padding-top: 1.6rem;
            padding-bottom: 3rem;
            max-width: 1320px;
        }}

        h1, h2, h3 {{
            color: var(--abb-dark);
            letter-spacing: -0.02em;
        }}

        .abb-hero {{
            background: #ffffff;
            border: 1px solid var(--abb-border);
            border-radius: 22px;
            padding: 1.35rem 1.55rem 1.2rem 1.55rem;
            box-shadow: 0 12px 34px rgba(31, 31, 46, 0.07);
            margin-bottom: 1.2rem;
            position: relative;
            overflow: hidden;
        }}
        .abb-hero:before {{
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            height: 6px;
            width: 100%;
            background: linear-gradient(90deg, var(--abb-red) 0%, var(--abb-lilac) 100%);
        }}
        .abb-logo {{
            color: var(--abb-red);
            font-weight: 900;
            font-size: 2.45rem;
            line-height: 0.95;
            letter-spacing: -0.06em;
        }}
        .abb-business {{
            color: #4d5161;
            font-weight: 700;
            font-size: 0.92rem;
            margin-top: 0.2rem;
        }}
        .abb-title {{
            margin-top: 0.55rem;
            font-size: clamp(2.05rem, 4vw, 3.2rem);
            font-weight: 850;
            color: var(--abb-dark);
            letter-spacing: -0.045em;
        }}
        .abb-subtitle {{
            margin-top: 0.15rem;
            color: #5b6070;
            font-size: 1.02rem;
            max-width: 880px;
        }}
        .abb-slogan {{
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            margin-top: 0.9rem;
            padding: 0.46rem 0.72rem;
            border-radius: 999px;
            background: linear-gradient(90deg, var(--abb-red-pale), var(--abb-lilac-pale));
            color: var(--abb-dark);
            font-weight: 800;
            font-size: 0.9rem;
        }}
        .abb-slogan-dot {{
            width: 8px;
            height: 8px;
            border-radius: 99px;
            background: var(--abb-red);
            box-shadow: 0 0 0 4px rgba(255, 0, 15, 0.12);
        }}

        .abb-section-card {{
            background: #ffffff;
            border: 1px solid var(--abb-border);
            border-radius: 18px;
            padding: 1rem 1.05rem;
            box-shadow: 0 10px 24px rgba(31, 31, 46, 0.055);
            margin: 0.85rem 0 1rem 0;
        }}
        .abb-section-title {{
            display: flex;
            align-items: center;
            gap: 0.65rem;
            font-size: 1.25rem;
            color: var(--abb-dark);
            font-weight: 850;
            margin-bottom: 0.22rem;
        }}
        .abb-step {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 2rem;
            height: 2rem;
            border-radius: 11px;
            background: var(--abb-red);
            color: white;
            font-weight: 850;
            font-size: 0.95rem;
        }}
        .abb-help {{
            color: #777d8d;
            font-size: 0.93rem;
            margin-bottom: 0.45rem;
        }}

        .file-chip {{
            display: inline-flex;
            align-items: center;
            gap: 0.48rem;
            padding: 0.48rem 0.7rem;
            margin: 0.18rem 0.18rem 0.18rem 0;
            border-radius: 999px;
            border: 1px solid #dfe3ea;
            background: #f7f8fb;
            color: #303445;
            font-size: 0.88rem;
        }}
        .file-chip span {{ color: #7a8090; }}

        .metric-grid {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 0.78rem;
            margin-top: 0.85rem;
        }}
        .metric-card {{
            background: #ffffff;
            border: 1px solid var(--abb-border);
            border-radius: 16px;
            padding: 0.9rem 1rem;
            box-shadow: 0 8px 20px rgba(31, 31, 46, 0.045);
            min-height: 90px;
        }}
        .metric-label {{
            color: #6a7080;
            font-size: 0.82rem;
            font-weight: 750;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }}
        .metric-value {{
            color: var(--abb-dark);
            font-size: 1.8rem;
            font-weight: 900;
            margin-top: 0.2rem;
        }}
        .metric-accent {{
            width: 32px;
            height: 4px;
            border-radius: 99px;
            margin-top: 0.55rem;
            background: linear-gradient(90deg, var(--abb-red), var(--abb-lilac));
        }}

        .status-pill {{
            display: inline-block;
            padding: 0.35rem 0.58rem;
            border-radius: 999px;
            background: var(--abb-lilac-pale);
            color: #34338a;
            font-weight: 800;
            font-size: 0.78rem;
            margin-right: 0.35rem;
            margin-top: 0.35rem;
        }}
        .status-pill.red {{
            background: var(--abb-red-pale);
            color: #9b0008;
        }}
        .status-pill.dark {{
            background: #eff1f5;
            color: #303445;
        }}

        div.stButton > button[kind="primary"] {{
            background: var(--abb-red) !important;
            border: 1px solid var(--abb-red) !important;
            color: white !important;
            font-weight: 800 !important;
            border-radius: 12px !important;
            min-height: 2.8rem;
        }}
        div.stButton > button[kind="primary"]:hover {{
            background: #d9000d !important;
            border-color: #d9000d !important;
        }}
        div.stButton > button:not([kind="primary"]) {{
            border-radius: 12px !important;
            font-weight: 750 !important;
            border-color: #2d3140 !important;
            color: #2d3140 !important;
        }}

        div[data-testid="stFileUploader"] section {{
            border-radius: 16px !important;
            border: 1px dashed #cfd5df !important;
            background: #f6f7fa !important;
        }}

        .download-link {{
            display: inline-block;
            padding: 0.8rem 1.12rem;
            border-radius: 12px;
            background: var(--abb-red);
            color: white !important;
            text-decoration: none !important;
            font-weight: 850;
            box-shadow: 0 10px 22px rgba(255, 0, 15, 0.22);
        }}
        .download-link:hover {{
            background: #d9000d;
            color: white !important;
            text-decoration: none !important;
        }}

        @media (max-width: 900px) {{
            .metric-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
            .abb-title {{ font-size: 2.1rem; }}
            .block-container {{ padding-left: 1rem; padding-right: 1rem; }}
        }}
        @media (max-width: 560px) {{
            .metric-grid {{ grid-template-columns: 1fr; }}
            .abb-logo {{ font-size: 2rem; }}
            .abb-section-title {{ font-size: 1.08rem; }}
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Utilities
# -----------------------------
def human_file_size(size_bytes: int) -> str:
    if size_bytes >= 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{size_bytes / 1024:.0f} KB"


def make_excel_download_link(data: bytes, file_name: str, label: str) -> str:
    encoded = base64.b64encode(data).decode("utf-8")
    return f'''
    <a class="download-link"
       href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{encoded}"
       download="{escape(file_name, quote=True)}">
       {escape(label)}
    </a>
    '''


def clear_data() -> None:
    for key in [
        "fv_result",
        "prepared_excel_bytes",
        "last_files",
        "show_Breaker BOM",
        "show_Interior BOM",
        "show_Box BOM",
    ]:
        st.session_state.pop(key, None)


def file_batch_errors(uploaded_files) -> list[str]:
    """Validate metadata only. Do not read PDF bytes until processing starts."""
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


def render_section_open(title: str, step: str | None = None, help_text: str | None = None) -> None:
    step_html = f'<span class="abb-step">{escape(step)}</span>' if step else ""
    help_html = f'<div class="abb-help">{escape(help_text)}</div>' if help_text else ""
    st.markdown(
        f'''
        <div class="abb-section-card">
            <div class="abb-section-title">{step_html}<span>{escape(title)}</span></div>
            {help_html}
        </div>
        ''',
        unsafe_allow_html=True,
    )


def render_file_chips(uploaded_files) -> None:
    if not uploaded_files:
        return

    visible = uploaded_files[:8]
    chips = []
    for file in visible:
        chips.append(
            f'<span class="file-chip">📄 {escape(file.name)} <span>{escape(human_file_size(int(file.size)))}</span></span>'
        )
    st.markdown("".join(chips), unsafe_allow_html=True)

    if len(uploaded_files) > len(visible):
        with st.expander(f"Show all {len(uploaded_files)} selected files"):
            for file in uploaded_files:
                st.write(f"• {file.name} — {human_file_size(int(file.size))}")


def render_metric_grid(metrics: dict) -> None:
    items = [
        ("PDFs", metrics.get("pdfs", 0)),
        ("Sales Order-Items", metrics.get("so_items", 0)),
        ("Breaker Rows", metrics.get("breaker_rows", 0)),
        ("Spacer Breakers", metrics.get("spacer_rows", 0)),
        ("Box BOM Rows", metrics.get("box_bom_rows", 0)),
        ("Interior BOM Rows", metrics.get("interior_bom_rows", 0)),
        ("Missing Mappings", metrics.get("unmapped", 0)),
        ("Errors", metrics.get("errors", 0)),
    ]
    cards = []
    for label, value in items:
        cards.append(
            f'''
            <div class="metric-card">
                <div class="metric-label">{escape(str(label))}</div>
                <div class="metric-value">{escape(str(value))}</div>
                <div class="metric-accent"></div>
            </div>
            '''
        )
    st.markdown(f'<div class="metric-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def dataframe_csv_preview(df: pd.DataFrame, columns: Iterable[str] | None = None, max_rows: int = 40) -> str:
    if df is None or df.empty:
        return "No rows."
    preview = df.copy()
    if columns:
        keep = [column for column in columns if column in preview.columns]
        if keep:
            preview = preview[keep]
    return preview.head(max_rows).to_csv(index=False)


def filter_dataframe(df: pd.DataFrame, query: str, visible_sections: list[str]) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    filtered = df.copy()

    if "BOM Section" in filtered.columns and visible_sections:
        filtered = filtered[filtered["BOM Section"].isin(visible_sections)].copy()

    q = query.strip().upper()
    if q:
        searchable_columns = [
            column
            for column in [
                "Sales Order",
                "Item",
                "Sales Order-Item",
                "Front View",
                "BOM Section",
                "Record Type",
                "Catalog Number",
                "Alternate SKU",
                "ABB Part Number",
                "Description",
            ]
            if column in filtered.columns
        ]
        if searchable_columns:
            mask = pd.Series(False, index=filtered.index)
            for column in searchable_columns:
                mask = mask | filtered[column].astype(str).str.upper().str.contains(q, na=False, regex=False)
            filtered = filtered[mask].copy()
    return filtered


def get_visible_sections() -> list[str]:
    available_sections = ["Breaker BOM", "Interior BOM", "Box BOM"]
    visible: list[str] = []
    st.markdown("**BOM sections to show**")
    cols = st.columns(3)
    for index, section in enumerate(available_sections):
        key = f"show_{section}"
        if key not in st.session_state:
            st.session_state[key] = True
        with cols[index]:
            is_on = st.toggle(section, value=bool(st.session_state[key]), key=key)
            if is_on:
                visible.append(section)
    return visible


def get_breaker_columns() -> list[str]:
    base = [
        "Sales Order",
        "Item",
        "Sales Order-Item",
        "Front View",
        "BOM Section",
        "Record Type",
        "Catalog Number",
        "Alternate SKU",
        "ABB Part Number",
        "Quantity",
        "Description",
    ]
    optional_fields = [
        "Trip Amps",
        "Poles",
        "Lug Cable Size / Neut Sensor / Grnd Fault Cable",
        "Sensor Amps",
        "Trip Unit",
        "Phases Used",
        "Notes",
    ]
    st.markdown("**Optional breaker fields**")
    cols = st.columns(4)
    selected = []
    default_on = {"Trip Amps", "Poles", "Lug Cable Size / Neut Sensor / Grnd Fault Cable"}
    for index, field in enumerate(optional_fields):
        with cols[index % 4]:
            if st.toggle(field, value=field in default_on, key=f"field_{field}"):
                selected.append(field)
    return base + selected


# -----------------------------
# Header
# -----------------------------
st.markdown(
    '''
    <div class="abb-hero">
        <div class="abb-logo">ABB</div>
        <div class="abb-business">Electrical Control Systems</div>
        <div class="abb-title">Front View BOM Extractor</div>
        <div class="abb-subtitle">
            Extract Breaker BOM, Interior BOM, and Box BOM information from Front View PDFs,
            cross-reference catalog numbers to ABB part numbers, and generate a structured Excel report.
        </div>
        <div class="abb-slogan"><span class="abb-slogan-dot"></span>Engineered to Outrun</div>
    </div>
    ''',
    unsafe_allow_html=True,
)

control_left, control_right = st.columns([5, 1])
with control_right:
    if st.button("Clear Data", use_container_width=True):
        clear_data()
        st.success("Data cleared.")

if not MAPPING_FILE.exists():
    st.error("The part-number mapping database was not found. Place part_number_mapping.xlsx in the same folder as app.py.")
    st.stop()

# -----------------------------
# Upload
# -----------------------------
render_section_open(
    "Upload Front Views",
    step="1",
    help_text=f"Batch limits: up to {MAX_PDF_FILES} PDFs, {MAX_SINGLE_FILE_MB} MB per file, and {MAX_TOTAL_UPLOAD_MB} MB total.",
)

uploaded_files = st.file_uploader(
    "Upload one or more Front View PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="visible",
)

errors = file_batch_errors(uploaded_files)
if uploaded_files:
    render_file_chips(uploaded_files)
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

# -----------------------------
# Results
# -----------------------------
result = st.session_state.get("fv_result")
if result:
    dataframes: dict[str, pd.DataFrame] = result["dataframes"]
    metrics: dict = result["metrics"]

    render_section_open("Processing Summary", step="2", help_text="Processed output by section. Use the controls below to customize what is visible on screen.")
    render_metric_grid(metrics)

    render_section_open("Filter and Customize View", step="3", help_text="The extraction is complete. These controls only change the on-screen preview, not the extracted data.")
    visible_sections = get_visible_sections()
    selected_columns = get_breaker_columns()
    search_query = st.text_input(
        "Search by Sales Order, Sales Order-Item, catalog number, ABB part number, or description",
        placeholder="Example: 154258789 or XT5HU340ABYN000XXX",
    )

    render_section_open("Results Preview", step="4", help_text="Stable preview mode uses CSV text to avoid Streamlit Cloud download/grid crashes. The Excel report contains the structured sheets.")
    available_previews = [name for name, df in dataframes.items() if getattr(df, "empty", True) is False]
    if not available_previews:
        st.info("No rows were extracted.")
    else:
        selected_preview = st.selectbox("Preview section", available_previews)
        preview_df = filter_dataframe(dataframes[selected_preview], search_query, visible_sections)
        row_count = 0 if preview_df is None or preview_df.empty else len(preview_df)
        st.markdown(
            f'<span class="status-pill">{escape(selected_preview)}</span>'
            f'<span class="status-pill dark">Rows shown: {row_count}</span>'
            f'<span class="status-pill red">Full detail available in Excel</span>',
            unsafe_allow_html=True,
        )
        columns = selected_columns if selected_preview in {"Results", "Breaker BOM", "Spacer Breakers"} else None
        st.code(dataframe_csv_preview(preview_df, columns=columns, max_rows=50), language="csv")

    render_section_open("Excel Report", step="5", help_text="Prepare the complete Excel workbook only after confirming the preview looks correct.")
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
