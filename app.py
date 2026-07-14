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
.stApp {{ background: linear-gradient(180deg, #ffffff 0%, #fbfbfd 55%, #f5f6f8 100%); }}
.block-container {{ padding-top: 1.15rem; padding-bottom: 3rem; max-width: 1280px; }}
h1, h2, h3 {{ color: var(--abb-dark); letter-spacing: -0.02em; }}
.abb-hero {{
    background: #ffffff;
    border: 1px solid var(--abb-border);
    border-radius: 24px;
    padding: 1.45rem 1.55rem 1.35rem 1.55rem;
    box-shadow: 0 16px 44px rgba(31, 31, 46, 0.075);
    margin-bottom: 1.15rem;
    position: relative;
    overflow: hidden;
}}
.abb-hero:before {{ content: ""; position: absolute; inset: 0 auto auto 0; height: 6px; width: 100%; background: linear-gradient(90deg, var(--abb-red) 0%, var(--abb-lilac) 100%); }}
.abb-brand-row {{ display: flex; align-items: center; justify-content: space-between; gap: 1.2rem; flex-wrap: wrap; margin-top: 0.15rem; }}
.abb-brand-left {{ display: flex; align-items: center; gap: 1rem; flex-wrap: wrap; }}
.abb-logo {{ color: var(--abb-red); font-weight: 950; font-size: clamp(2.2rem, 5vw, 3.9rem); line-height: 0.88; letter-spacing: -0.085em; }}
.abb-slogan-wordmark {{ color: var(--abb-dark); font-weight: 950; font-size: clamp(1.25rem, 2.7vw, 2.25rem); line-height: 0.92; letter-spacing: -0.035em; text-transform: uppercase; }}
.abb-slogan-wordmark span {{ display: block; }}
.abb-business {{ color: #4d5161; font-weight: 780; font-size: 0.95rem; padding: 0.45rem 0.8rem; border: 1px solid #eef0f4; border-radius: 999px; background: #fafbfc; }}
.abb-title {{ margin-top: 1.15rem; font-size: clamp(1.9rem, 4vw, 3rem); font-weight: 900; color: var(--abb-dark); letter-spacing: -0.045em; }}
.abb-subtitle {{ margin-top: 0.25rem; color: #5b6070; font-size: 1rem; max-width: 900px; line-height: 1.55; }}
.abb-rail {{ display: flex; gap: 0.4rem; margin-top: 1.05rem; }}
.abb-rail div {{ height: 4px; border-radius: 99px; }}
.abb-rail-red {{ width: 90px; background: var(--abb-red); }}
.abb-rail-lilac {{ width: 54px; background: var(--abb-lilac); }}
.abb-section-card {{ background: #ffffff; border: 1px solid var(--abb-border); border-radius: 18px; padding: 1rem 1.05rem; box-shadow: 0 10px 26px rgba(31, 31, 46, 0.055); margin: 0.85rem 0 1rem 0; }}
.abb-section-title {{ display: flex; align-items: center; gap: 0.65rem; font-size: 1.22rem; color: var(--abb-dark); font-weight: 880; margin-bottom: 0.22rem; }}
.abb-step {{ display: inline-flex; align-items: center; justify-content: center; min-width: 2rem; height: 2rem; border-radius: 11px; background: var(--abb-red); color: white; font-weight: 900; font-size: 0.95rem; }}
.abb-help {{ color: #777d8d; font-size: 0.93rem; margin-bottom: 0.45rem; }}
.file-chip {{ display: inline-flex; align-items: center; gap: 0.48rem; padding: 0.48rem 0.7rem; margin: 0.18rem 0.18rem 0.18rem 0; border-radius: 999px; border: 1px solid #dfe3ea; background: #f7f8fb; color: #303445; font-size: 0.88rem; }}
.file-chip span {{ color: #7a8090; }}
.metric-grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 0.78rem; margin-top: 0.85rem; }}
.metric-card {{ background: #ffffff; border: 1px solid var(--abb-border); border-radius: 16px; padding: 0.9rem 1rem; box-shadow: 0 8px 20px rgba(31, 31, 46, 0.045); min-height: 92px; }}
.metric-label {{ color: #6a7080; font-size: 0.78rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.04em; }}
.metric-value {{ color: var(--abb-dark); font-size: 1.75rem; font-weight: 950; margin-top: 0.2rem; }}
.metric-accent {{ width: 34px; height: 4px; border-radius: 99px; margin-top: 0.55rem; background: linear-gradient(90deg, var(--abb-red), var(--abb-lilac)); }}
.status-pill {{ display: inline-block; padding: 0.35rem 0.58rem; border-radius: 999px; background: var(--abb-lilac-pale); color: #34338a; font-weight: 820; font-size: 0.78rem; margin-right: 0.35rem; margin-top: 0.35rem; }}
.status-pill.red {{ background: var(--abb-red-pale); color: #9b0008; }}
.status-pill.dark {{ background: #eff1f5; color: #303445; }}
.preview-table-wrap {{ overflow-x: auto; border: 1px solid var(--abb-border); border-radius: 16px; background: #ffffff; box-shadow: 0 8px 20px rgba(31,31,46,0.04); margin-top: 0.85rem; }}
.preview-table {{ width: 100%; border-collapse: collapse; min-width: 920px; font-size: 0.86rem; }}
.preview-table th {{ text-align: left; background: #f2f3f6; color: #4f5565; padding: 0.72rem 0.8rem; border-bottom: 1px solid var(--abb-border); font-weight: 850; white-space: nowrap; }}
.preview-table td {{ padding: 0.68rem 0.8rem; border-bottom: 1px solid #eef0f4; color: #272b38; vertical-align: top; }}
.preview-table tr:nth-child(even) td {{ background: #fbfcfd; }}
.preview-empty {{ padding: 1rem; color: #777d8d; background: #ffffff; border: 1px solid var(--abb-border); border-radius: 14px; }}
div.stButton > button[kind="primary"] {{ background: var(--abb-red) !important; border: 1px solid var(--abb-red) !important; color: white !important; font-weight: 850 !important; border-radius: 12px !important; min-height: 2.8rem; }}
div.stButton > button[kind="primary"]:hover {{ background: #d9000d !important; border-color: #d9000d !important; }}
div.stButton > button:not([kind="primary"]) {{ border-radius: 12px !important; font-weight: 780 !important; border-color: #2d3140 !important; color: #2d3140 !important; }}
div[data-testid="stFileUploader"] section {{ border-radius: 16px !important; border: 1px dashed #cfd5df !important; background: #f6f7fa !important; }}
.download-link {{ display: inline-block; padding: 0.8rem 1.12rem; border-radius: 12px; background: var(--abb-red); color: white !important; text-decoration: none !important; font-weight: 880; box-shadow: 0 10px 22px rgba(255, 0, 15, 0.22); }}
.download-link:hover {{ background: #d9000d; color: white !important; text-decoration: none !important; }}
@media (max-width: 900px) {{ .metric-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} .abb-title {{ font-size: 2rem; }} .block-container {{ padding-left: 1rem; padding-right: 1rem; }} .abb-brand-row {{ align-items: flex-start; }} }}
@media (max-width: 560px) {{ .metric-grid {{ grid-template-columns: 1fr; }} .abb-logo {{ font-size: 2rem; }} .abb-slogan-wordmark {{ font-size: 1.25rem; }} .abb-section-title {{ font-size: 1.08rem; }} }}
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
            '<div class="metric-card">'
            f'<div class="metric-label">{escape(str(label))}</div>'
            f'<div class="metric-value">{escape(str(value))}</div>'
            '<div class="metric-accent"></div>'
            '</div>'
        )
    st.markdown('<div class="metric-grid">' + ''.join(cards) + '</div>', unsafe_allow_html=True)


def dataframe_html_preview(df: pd.DataFrame, columns: Iterable[str] | None = None, max_rows: int = 40) -> str:
    if df is None or df.empty:
        return '<div class="preview-empty">No rows.</div>'
    preview = df.copy()
    if columns:
        keep = [column for column in columns if column in preview.columns]
        if keep:
            preview = preview[keep]
    preview = preview.head(max_rows).fillna("").astype(str)
    headers = ''.join(f'<th>{escape(col)}</th>' for col in preview.columns)
    body_rows = []
    for _, row in preview.iterrows():
        cells = ''.join(f'<td>{escape(str(value))}</td>' for value in row.tolist())
        body_rows.append(f'<tr>{cells}</tr>')
    return '<div class="preview-table-wrap"><table class="preview-table"><thead><tr>' + headers + '</tr></thead><tbody>' + ''.join(body_rows) + '</tbody></table></div>'


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



def get_breaker_columns() -> list[str]:
    base = [
        "Sales Order",
        "Item",
        "Sales Order-Item",
        "Front View",
        "Record Type",
        "Catalog Number",
        "Alternate SKU",
        "ABB Part Number",
        "Quantity",
    ]
    optional_fields = [
        "Trip Amps",
        "Poles",
        "Lug Cable Size / Neut Sensor / Grnd Fault Cable",
        "Sensor Amps",
        "Trip Unit",
        "Phases Used",
        "Notes",
        "X Space",
    ]
    optional_fields = list(dict.fromkeys(optional_fields))
    default_fields = [
        "Trip Amps",
        "Poles",
        "Lug Cable Size / Neut Sensor / Grnd Fault Cable",
    ]
    selected = st.multiselect(
        "Breaker fields to display",
        options=optional_fields,
        default=[field for field in default_fields if field in optional_fields],
        help="These fields only apply to Breaker BOM rows. Interior BOM and Box BOM use their own relevant columns.",
    )
    return list(dict.fromkeys(base + selected + ["Description"]))


def columns_available(df: pd.DataFrame, desired: list[str]) -> list[str]:
    if df is None or df.empty:
        return desired
    return [column for column in desired if column in df.columns]


def section_detail_columns(section: str, breaker_columns: list[str] | None = None) -> list[str]:
    if section in {"Breaker BOM", "Spacer Breakers"}:
        return breaker_columns or get_breaker_columns()
    if section in {"Interior BOM", "Box BOM"}:
        return [
            "Sales Order",
            "Item",
            "Sales Order-Item",
            "Front View",
            "Catalog Number",
            "Alternate SKU",
            "ABB Part Number",
            "Quantity",
            "Description",
            "Source File",
            "Page",
        ]
    if section == "Unmapped Codes":
        return ["Detected Code", "BOM Section", "Record Type", "Sales Order-Item", "Front View", "Quantity", "Source File"]
    if section == "Processing Log":
        return []
    return []


def filter_by_query(df: pd.DataFrame, query: str) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    q = query.strip().upper()
    if not q:
        return df.copy()
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
            "Source File",
        ]
        if column in df.columns
    ]
    if not searchable_columns:
        return df.copy()
    mask = pd.Series(False, index=df.index)
    for column in searchable_columns:
        mask = mask | df[column].astype(str).str.upper().str.contains(q, na=False, regex=False)
    return df[mask].copy()


def apply_section_filter(df: pd.DataFrame, sections: list[str]) -> pd.DataFrame:
    if df is None or df.empty or not sections or "BOM Section" not in df.columns:
        return df
    return df[df["BOM Section"].isin(sections)].copy()


def render_html_table(df: pd.DataFrame, columns: list[str] | None = None, max_rows: int = 60) -> None:
    st.markdown(dataframe_html_preview(df, columns=columns, max_rows=max_rows), unsafe_allow_html=True)


def result_subset_for_section(dataframes: dict[str, pd.DataFrame], section: str) -> pd.DataFrame:
    if section in dataframes:
        return dataframes[section].copy()
    results = dataframes.get("Results", pd.DataFrame()).copy()
    if results.empty:
        return results
    if "BOM Section" in results.columns:
        return results[results["BOM Section"] == section].copy()
    return results


def render_grouped_order_view(dataframes: dict[str, pd.DataFrame], query: str, breaker_columns: list[str]) -> None:
    results = dataframes.get("Results", pd.DataFrame()).copy()
    if results.empty:
        st.info("No extracted rows available for grouped view.")
        return
    results = filter_by_query(results, query)
    if results.empty:
        st.info("No rows match the current search.")
        return

    group_col = "Sales Order-Item" if "Sales Order-Item" in results.columns else None
    if not group_col:
        render_html_table(results, max_rows=80)
        return

    open_by_default = st.checkbox("Open grouped results by default", value=False)
    so_items = sorted(results[group_col].astype(str).dropna().unique().tolist())
    st.caption(f"Showing {len(so_items)} Sales Order-Item group(s). Use the search box to narrow the list.")

    for so_item in so_items:
        group = results[results[group_col].astype(str) == so_item].copy()
        sales_order = group["Sales Order"].iloc[0] if "Sales Order" in group.columns and not group.empty else ""
        item = group["Item"].iloc[0] if "Item" in group.columns and not group.empty else ""
        breaker_qty = int(pd.to_numeric(group.loc[group.get("BOM Section", "") == "Breaker BOM", "Quantity"], errors="coerce").fillna(0).sum()) if "BOM Section" in group.columns and "Quantity" in group.columns else 0
        interior_rows = int((group.get("BOM Section", pd.Series(dtype=str)) == "Interior BOM").sum()) if "BOM Section" in group.columns else 0
        box_rows = int((group.get("BOM Section", pd.Series(dtype=str)) == "Box BOM").sum()) if "BOM Section" in group.columns else 0
        title = f"{so_item}  •  SO {sales_order}  •  Item {item}  •  {breaker_qty} breaker qty  •  {interior_rows} interior rows  •  {box_rows} box rows"

        with st.expander(title, expanded=open_by_default):
            panels = ["Single Panel"]
            if "Front View" in group.columns:
                panels = sorted(group["Front View"].astype(str).dropna().unique().tolist())
            for panel in panels:
                panel_df = group[group["Front View"].astype(str) == panel].copy() if "Front View" in group.columns else group
                st.markdown(f"**{escape(panel)}**")
                for section in ["Breaker BOM", "Interior BOM", "Box BOM"]:
                    if "BOM Section" not in panel_df.columns:
                        continue
                    section_df = panel_df[panel_df["BOM Section"] == section].copy()
                    if section_df.empty:
                        continue
                    if section == "Breaker BOM":
                        columns = columns_available(section_df, breaker_columns)
                    else:
                        columns = columns_available(section_df, section_detail_columns(section, breaker_columns))
                    st.markdown(f'<span class="status-pill">{escape(section)}</span><span class="status-pill dark">Rows: {len(section_df)}</span>', unsafe_allow_html=True)
                    render_html_table(section_df, columns=columns, max_rows=35)


def render_recommended_view(dataframes: dict[str, pd.DataFrame], view_name: str, query: str, breaker_columns: list[str]) -> None:
    if view_name == "Grouped by Sales Order-Item":
        render_grouped_order_view(dataframes, query, breaker_columns)
        return

    view_to_section = {
        "Breaker BOM detail": "Breaker BOM",
        "Interior BOM detail": "Interior BOM",
        "Box BOM detail": "Box BOM",
        "Spacer breakers excluded": "Spacer Breakers",
        "Missing mappings": "Unmapped Codes",
        "Processing log": "Processing Log",
    }
    section = view_to_section.get(view_name, "Results")
    df = result_subset_for_section(dataframes, section)
    df = filter_by_query(df, query)
    columns = section_detail_columns(section, breaker_columns)
    st.markdown(
        f'<span class="status-pill">{escape(view_name)}</span>'
        f'<span class="status-pill dark">Rows shown: {0 if df is None or df.empty else len(df)}</span>'
        f'<span class="status-pill red">Complete workbook available in Excel</span>',
        unsafe_allow_html=True,
    )
    render_html_table(df, columns=columns, max_rows=80)


def render_custom_table(dataframes: dict[str, pd.DataFrame], query: str) -> None:
    results = dataframes.get("Results", pd.DataFrame()).copy()
    if results.empty:
        st.info("No results available.")
        return
    section_options = [section for section in ["Breaker BOM", "Interior BOM", "Box BOM"] if section in set(results.get("BOM Section", []))]
    selected_sections = st.multiselect("BOM sections", options=section_options, default=section_options)
    filtered = apply_section_filter(results, selected_sections)
    filtered = filter_by_query(filtered, query)
    available_columns = list(filtered.columns) if not filtered.empty else list(results.columns)
    default_columns = [
        column
        for column in ["Sales Order-Item", "Front View", "BOM Section", "Record Type", "Catalog Number", "ABB Part Number", "Quantity", "Description"]
        if column in available_columns
    ]
    selected_columns = st.multiselect("Columns to show", options=available_columns, default=default_columns)
    render_html_table(filtered, columns=selected_columns, max_rows=100)


def render_raw_table(dataframes: dict[str, pd.DataFrame], query: str) -> None:
    available = [name for name, df in dataframes.items() if getattr(df, "empty", True) is False]
    if not available:
        st.info("No raw tables are available.")
        return
    selected = st.selectbox("Raw sheet", available)
    df = filter_by_query(dataframes[selected], query)
    st.caption("Raw view shows the extracted sheet columns with minimal formatting.")
    render_html_table(df, columns=list(df.columns), max_rows=100)

# -----------------------------
# Header
# -----------------------------
st.markdown(
    '''<div class="abb-hero"><div class="abb-brand-row"><div class="abb-brand-left"><div class="abb-logo">ABB</div><div class="abb-slogan-wordmark"><span>Engineered</span><span>to Outrun</span></div></div><div class="abb-business">Electrical Control Systems</div></div><div class="abb-title">Front View BOM Extractor</div><div class="abb-subtitle">Extract Breaker BOM, Interior BOM, and Box BOM information from Front View PDFs, cross-reference catalog numbers to ABB part numbers, and generate a structured Excel report.</div><div class="abb-rail"><div class="abb-rail-red"></div><div class="abb-rail-lilac"></div></div></div>''',
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

    render_section_open("Filter and Customize View", step="3", help_text="Choose a recommended view, search specific orders, or build a custom table. These controls only change the on-screen preview; the Excel report keeps the full extraction.")
    search_query = st.text_input(
        "Search by Sales Order, Sales Order-Item, catalog number, ABB part number, source file, or description",
        placeholder="Example: 154258789, 154258789-10, XT5HU340ABYN000XXX, or ER8440A",
    )

    view_mode = st.radio(
        "Preview mode",
        options=["Recommended views", "Custom table", "Raw extraction"],
        horizontal=True,
        help="Recommended views separate Breaker, Interior, and Box information. Custom table lets you choose sections and columns. Raw extraction shows the workbook sheets as extracted.",
    )

    breaker_columns = get_breaker_columns()

    render_section_open("Results Preview", step="4", help_text="Grouped and section-specific views keep breaker-only fields away from Interior BOM and Box BOM rows.")

    if view_mode == "Recommended views":
        recommended_view = st.selectbox(
            "Recommended view",
            options=[
                "Grouped by Sales Order-Item",
                "Breaker BOM detail",
                "Interior BOM detail",
                "Box BOM detail",
                "Spacer breakers excluded",
                "Missing mappings",
                "Processing log",
            ],
        )
        render_recommended_view(dataframes, recommended_view, search_query, breaker_columns)
    elif view_mode == "Custom table":
        render_custom_table(dataframes, search_query)
    else:
        render_raw_table(dataframes, search_query)

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
