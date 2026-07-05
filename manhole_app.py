from __future__ import annotations

from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from manhole_report import (
    DEFAULT_LOGO,
    DEFAULT_SOURCE,
    _build_report,
    export_dxf,
    export_pdf,
    export_report,
    get_design_manhole_names,
)


st.set_page_config(page_title="Manhole Setting-Out Report", layout="wide")

logo_path = Path(DEFAULT_LOGO)
header_cols = st.columns([1, 5])
if logo_path.exists():
    header_cols[0].image(str(logo_path), width=140)
header_cols[1].title("Kesheshiwe Engineering Surveyors Cc")
header_cols[1].subheader("Manhole Setting-Out Report Generator")

with st.sidebar:
    st.header("Source")
    uploaded = st.file_uploader("Upload fieldbook CSV or Excel", type=["csv", "xlsx", "xlsm", "xls"])
    use_sample = st.checkbox("Use Fransfontein fieldbook", value=uploaded is None)

    st.header("Report Settings")
    reverse_flow = st.toggle("Reverse flow direction", value=False)
    level_tolerance = st.number_input("Level tolerance (mm)", min_value=0.0, value=20.0, step=5.0)
    max_match_distance = st.number_input("Maximum match distance (m)", min_value=1.0, value=25.0, step=1.0)


def load_input() -> pd.DataFrame | None:
    if uploaded is not None:
        name = uploaded.name.lower()
        if name.endswith((".xlsx", ".xlsm", ".xls")):
            return pd.read_excel(uploaded)
        return pd.read_csv(uploaded)
    if use_sample:
        return pd.read_csv(DEFAULT_SOURCE)
    return None


df = load_input()
if df is None:
    st.info("Upload a fieldbook or enable the Fransfontein sample to begin.")
    st.stop()

try:
    all_points = get_design_manhole_names(df)
except Exception as exc:
    st.error(f"Could not detect design manholes in this fieldbook: {exc}")
    st.stop()

with st.sidebar:
    st.header("Manholes")
    included_points = st.multiselect(
        "Include manholes",
        all_points,
        default=all_points,
        help="Remove manholes by unticking them here.",
    )
    default_flow_text = "\n".join(included_points)
    flow_text = st.text_area(
        "Flow order",
        value=default_flow_text,
        height=220,
        help="One manhole per line. Reorder these names before generating the layout and longsection.",
    )
    typed_flow = [line.strip() for line in flow_text.splitlines() if line.strip()]
    valid_points = set(included_points)
    selected_flow = [point for point in typed_flow if point in valid_points]
    missing_from_flow = [point for point in included_points if point not in selected_flow]
    selected_flow.extend(missing_from_flow)
    excluded_points = [point for point in all_points if point not in included_points]
    unknown_flow = [point for point in typed_flow if point not in valid_points]
    if unknown_flow:
        st.warning("Ignored unknown or excluded manholes: " + ", ".join(unknown_flow))

sheets = _build_report(
    df,
    level_tolerance_mm=level_tolerance,
    max_match_distance_m=max_match_distance,
    excluded_points=excluded_points,
    reverse_flow=reverse_flow,
    flow_order=selected_flow,
)

summary = sheets["Summary"]
comparison = sheets["Manhole Comparison"]
surveyed_comparison = sheets["Surveyed Comparison"]
slope = sheets["Trench Slopes"]

metric_values = dict(summary.itertuples(index=False, name=None))
metric_cols = st.columns(5)
metric_cols[0].metric("Design points", metric_values.get("Design points", 0))
metric_cols[1].metric("Measured observations", metric_values.get("Valid measured observations", 0))
metric_cols[2].metric("Trench points", metric_values.get("Design points with trench/invert", 0))
metric_cols[3].metric("Cut points", metric_values.get("Design points needing cut", 0))
metric_cols[4].metric("Fill points", metric_values.get("Design points needing fill", 0))

tabs = st.tabs(["Manholes", "All Surveyed Points", "Longsection", "Measured Assignments", "QA"])
tabs[0].dataframe(comparison, use_container_width=True, hide_index=True)
tabs[1].dataframe(surveyed_comparison, use_container_width=True, hide_index=True)
tabs[2].dataframe(slope, use_container_width=True, hide_index=True)
if not slope.empty:
    st.caption("Layout/CAD plotting uses inverted coordinates: plotted East = -East, plotted North = -North.")
    tabs[2].line_chart(
        slope,
        x="Chainage m",
        y=["Design Level", "Measured Trench Level"],
        color=["#d9480f", "#2563eb"],
    )
tabs[3].dataframe(sheets["Measured Assignments"], use_container_width=True, hide_index=True)
tabs[4].dataframe(sheets["QA Excluded Rows"], use_container_width=True, hide_index=True)

st.divider()
st.subheader("Exports")
export_cols = st.columns(3)

with tempfile.TemporaryDirectory() as tmpdir:
    tmp = Path(tmpdir)
    xlsx_path = tmp / "manhole_settingout_report.xlsx"
    pdf_path = tmp / "manhole_layout_longsection.pdf"
    dxf_path = tmp / "manhole_layout_longsection.dxf"
    export_report(sheets, xlsx_path, logo_path=logo_path)
    export_pdf(sheets, pdf_path, logo_path=logo_path)
    export_dxf(sheets, dxf_path)

    export_cols[0].download_button(
        "Download Excel A3 Report",
        data=xlsx_path.read_bytes(),
        file_name="manhole_settingout_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    export_cols[1].download_button(
        "Download PDF",
        data=pdf_path.read_bytes(),
        file_name="manhole_layout_longsection.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
    export_cols[2].download_button(
        "Download DXF",
        data=dxf_path.read_bytes(),
        file_name="manhole_layout_longsection.dxf",
        mime="application/dxf",
        use_container_width=True,
    )
