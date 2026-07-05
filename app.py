from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

from checker import (
    ColumnMap,
    analyze_fieldbook,
    detect_columns,
    detect_date_column,
    export_excel,
    load_fieldbook,
)


st.set_page_config(page_title="Control Point Placement Checker", layout="wide")

st.title("Control Point Placement Checker")

uploaded = st.file_uploader("Upload GPS/CSV fieldbook", type=["csv", "xlsx", "xlsm", "xls"])

with st.sidebar:
    st.header("Comparison")
    comparison_mode = st.selectbox(
        "Mode",
        ["Position and height", "Positional settingout only"],
    )
    position_only = comparison_mode == "Positional settingout only"
    latest_day_only = st.checkbox("Use latest surveyed day only", value=True)

    st.header("Tolerances")
    position_tolerance = st.number_input(
        "Position tolerance (mm)",
        min_value=0.0,
        value=100.0,
        step=5.0,
    )
    if position_only:
        height_tolerance = 20.0
    else:
        height_tolerance = st.number_input(
            "Height tolerance (mm)",
            min_value=0.0,
            value=20.0,
            step=1.0,
        )
    invert_for_autocad = st.checkbox("Invert for AutoCAD Plotting", value=False)

if uploaded is None:
    st.info("Upload a CSV or Excel fieldbook to begin.")
    st.stop()

try:
    df = load_fieldbook(uploaded)
except Exception as exc:
    st.error(f"Could not read the uploaded file: {exc}")
    st.stop()

if df.empty:
    st.warning("The uploaded file has no rows.")
    st.stop()

detected, missing = detect_columns(df)
detected_date_column = detect_date_column(df)
columns = list(df.columns)

st.subheader("Column Mapping")
st.caption("Detected columns can be changed before running the report.")

mapping_cols = st.columns(5)
field_labels = {
    "point": "Point name",
    "easting": "Easting",
    "northing": "Northing",
    "height": "Height / elevation",
    "solution": "Solution / fix status",
}
selected: dict[str, str] = {}
for idx, field in enumerate(["point", "easting", "northing", "height", "solution"]):
    detected_value = detected.get(field)
    default_index = columns.index(detected_value) if detected_value in columns else 0
    selected[field] = mapping_cols[idx].selectbox(
        field_labels[field],
        columns,
        index=default_index,
        key=f"column-{field}",
    )

if missing:
    st.warning("Some columns were not detected automatically. Confirm the mapping above before export.")

column_map = ColumnMap(**selected)

latest_date_column = None
if latest_day_only:
    st.subheader("Survey Date Filter")
    date_options = ["Do not filter"] + columns
    default_date_index = (
        date_options.index(detected_date_column)
        if detected_date_column in date_options
        else 0
    )
    selected_date_column = st.selectbox(
        "Date / timestamp column",
        date_options,
        index=default_date_index,
    )
    if selected_date_column == "Do not filter":
        st.warning("Latest surveyed day filter is on, but no date column is selected.")
    else:
        latest_date_column = selected_date_column

results = analyze_fieldbook(
    df,
    column_map,
    position_tolerance_mm=position_tolerance,
    height_tolerance_mm=height_tolerance,
    invert_for_autocad=invert_for_autocad,
    latest_date_column=latest_date_column,
    position_only=position_only,
)

summary = results["Summary"]
report = results["Report"]
unmatched = results["Unmatched Measurements"]
errors = results["Errors"]

metric_cols = st.columns(5)
summary_values = dict(zip(summary["Metric"], summary["Value"]))
metric_cols[0].metric("Matched controls", int(summary_values["Total matched control points"]))
metric_cols[1].metric("Unmatched measurements", int(summary_values["Total unmatched measured points"]))
metric_cols[2].metric("Position pass", f"{summary_values['Position pass percentage']}%")
metric_cols[3].metric("Height pass", "N/A" if position_only else f"{summary_values['Height pass percentage']}%")
metric_cols[4].metric("Overall pass", f"{summary_values['Overall pass percentage']}%")

st.subheader("Preview")
tabs = st.tabs(["Report", "Failed Points", "Unmatched Measurements", "Raw Classified Data", "Errors"])
tabs[0].dataframe(report, use_container_width=True, hide_index=True)
tabs[1].dataframe(results["Failed Points"], use_container_width=True, hide_index=True)
tabs[2].dataframe(unmatched, use_container_width=True, hide_index=True)
tabs[3].dataframe(results["Raw Classified Data"], use_container_width=True, hide_index=True)
tabs[4].dataframe(errors, use_container_width=True, hide_index=True)

excel_bytes = export_excel(results)
st.download_button(
    "Export Excel Report",
    data=excel_bytes,
    file_name="control_point_placement_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

cleaned_csv = results["Raw Classified Data"].to_csv(index=False).encode("utf-8")
st.download_button(
    "Export Cleaned CSV",
    data=cleaned_csv,
    file_name="control_point_classified_data.csv",
    mime="text/csv",
)
