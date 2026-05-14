"""
Streamlit version of the LL Sizing Report dashboard.

This app avoids local folder browsing and server-side report files so it can run
on Streamlit Cloud. Users upload Calypso `.xls` files by batch, review tables,
and download the generated HTML report directly from the browser.
"""

from __future__ import annotations

import html
import re
from datetime import datetime
from typing import Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import xlrd


SPEC_LIMITS = {
    "32mm": {"lower": 1.2610, "upper": 1.2635},
    "34mm": {"lower": 1.3396, "upper": 1.3421},
    "36mm": {"lower": 1.4180, "upper": 1.4205},
    "38mm": {"lower": 1.4965, "upper": 1.4990},
    "40mm": {"lower": 1.5760, "upper": 1.5785},
}
for spec in SPEC_LIMITS.values():
    spec["mid"] = round((spec["lower"] + spec["upper"]) / 2, 7)


PATTERNS = {
    "L_Out1": re.compile(r"#\d+_ID[\d.]+_L\s*Out-1$", re.IGNORECASE),
    "L_Out2": re.compile(r"#\d+_ID[\d.]+_L\s*Out-2$", re.IGNORECASE),
    "L_In1": re.compile(r"#\d+_ID[\d.]+_L\s*In-1$", re.IGNORECASE),
    "L_In2": re.compile(r"#\d+_ID[\d.]+_L\s*In-2$", re.IGNORECASE),
    "R_Out1": re.compile(r"#\d+_ID[\d.]+_R\s*Out-1$", re.IGNORECASE),
    "R_Out2": re.compile(r"#\d+_ID[\d.]+_R\s*Out-2$", re.IGNORECASE),
    "R_In1": re.compile(r"#\d+_ID[\d.]+_R\s*In-1$", re.IGNORECASE),
    "R_In2": re.compile(r"#\d+_ID[\d.]+_R\s*In-2$", re.IGNORECASE),
}
REQUIRED_MEASUREMENTS = tuple(PATTERNS.keys())
POS_KEYS = ("AS_U", "AS_L", "DS_U", "DS_L")
POSITIONS = ("AS U", "AS L", "DS U", "DS L")
MM_TO_IN = 1 / 25.4
IN_TO_MM = 25.4
COLORS = ("#2563eb", "#f97316", "#64748b", "#16a34a", "#db2777", "#7c3aed", "#0891b2", "#ca8a04")

st.set_page_config(page_title="LL Sizing Report", layout="wide")


def parse_xls_bytes(content: bytes, filename: str) -> dict[str, Any]:
    workbook = xlrd.open_workbook(file_contents=content)
    sheet = workbook.sheet_by_index(0)
    measurements: dict[str, float] = {}
    all_chars: dict[str, dict[str, float]] = {}
    warnings: list[str] = []

    try:
        raw_part_no = sheet.cell_value(7, 5)
        part_no = str(int(raw_part_no)) if isinstance(raw_part_no, float) else str(raw_part_no)
    except (IndexError, TypeError, ValueError) as exc:
        part_no = ""
        warnings.append(f"Could not read part number: {exc}")

    skipped_rows = 0
    for row in range(12, sheet.nrows):
        char_name = str(sheet.cell_value(row, 0)).strip()
        if not char_name or char_name == "Characteristic":
            continue
        try:
            all_chars[char_name] = {
                "actual": float(sheet.cell_value(row, 1)),
                "nominal": float(sheet.cell_value(row, 2)),
                "upper_tol": float(sheet.cell_value(row, 3)),
                "lower_tol": float(sheet.cell_value(row, 4)),
                "deviation": float(sheet.cell_value(row, 5)),
            }
        except (IndexError, TypeError, ValueError):
            skipped_rows += 1

    for char_name, char_data in all_chars.items():
        for key, pattern in PATTERNS.items():
            if pattern.match(char_name):
                measurements[key] = char_data["actual"]
                break

    missing = [key for key in REQUIRED_MEASUREMENTS if key not in measurements]
    if skipped_rows:
        warnings.append(f"Skipped {skipped_rows} non-numeric or malformed characteristic rows")
    if missing:
        warnings.append(f"Missing required measurements: {', '.join(missing)}")

    return {
        "filename": filename,
        "part_no": part_no,
        "measurements": measurements,
        "positions": compute_positions(measurements),
        "warnings": warnings,
    }


def compute_positions(measurements: dict[str, float]) -> dict[str, float]:
    pairs = {
        "AS_U": ("L_Out1", "L_Out2"),
        "AS_L": ("L_In1", "L_In2"),
        "DS_U": ("R_Out1", "R_Out2"),
        "DS_L": ("R_In1", "R_In2"),
    }
    positions: dict[str, float] = {}
    for position, (left_key, right_key) in pairs.items():
        left = measurements.get(left_key)
        right = measurements.get(right_key)
        if left is not None and right is not None:
            positions[position] = round((left + right) / 2, 7)
    return positions


def convert_value(value: float | None, from_unit: str, to_unit: str) -> float | None:
    if value is None or from_unit == to_unit:
        return value
    return value * MM_TO_IN if from_unit == "mm" else value * IN_TO_MM


def batch_source_unit(batch_type: str) -> str:
    return "mm" if batch_type == "raw" else "inch"


def batch_label(batch_type: str, sizer_size: float | None) -> str:
    if batch_type == "raw":
        return "Raw Material"
    if batch_type == "bushing_unsized":
        return "Bushing ID (unsized)"
    if sizer_size is None:
        return "Bushing ID (sized)"
    return f'Sized by {sizer_size:.5f}" sizer'


def positions_dataframe(parts: list[dict[str, Any]], source_unit: str, display_unit: str) -> pd.DataFrame:
    rows = []
    for index, part in enumerate(parts, 1):
        row = {"Part": str(index)}
        for key, label in zip(POS_KEYS, POSITIONS):
            row[label] = convert_value(part["positions"].get(key), source_unit, display_unit)
        rows.append(row)
    return pd.DataFrame(rows)


def part_status(row: pd.Series, spec: dict[str, float], batch_type: str, display_unit: str) -> str:
    if any(pd.isna(row[position]) for position in POSITIONS):
        return "Missing"
    if batch_type == "raw":
        return "Data"
    spec_lower = convert_value(spec["lower"], "inch", display_unit)
    spec_upper = convert_value(spec["upper"], "inch", display_unit)
    if any(row[position] < spec_lower or row[position] > spec_upper for position in POSITIONS):
        return "OOS"
    return "OK"


def build_plotly_chart(table: pd.DataFrame, spec: dict[str, float], batch_type: str, display_unit: str) -> go.Figure:
    fig = go.Figure()
    decimals = 5 if display_unit == "inch" else 4

    for index, row in table.iterrows():
        values = [row[position] for position in POSITIONS]
        fig.add_trace(
            go.Scatter(
                x=list(POSITIONS),
                y=values,
                mode="lines+markers",
                name=f"Part {row['Part']}",
                line={"width": 2.5, "color": COLORS[index % len(COLORS)]},
                marker={"size": 8},
                hovertemplate="Part %{customdata}<br>%{x}: %{y:." + str(decimals) + "f}<extra></extra>",
                customdata=[row["Part"]] * len(POSITIONS),
            )
        )

    if batch_type != "raw":
        for label, raw_value in (("Max", spec["upper"]), ("Mid", spec["mid"]), ("Min", spec["lower"])):
            value = convert_value(raw_value, "inch", display_unit)
            fig.add_trace(
                go.Scatter(
                    x=list(POSITIONS),
                    y=[value] * len(POSITIONS),
                    mode="lines",
                    name=label,
                    line={"width": 2, "color": "#dc2626", "dash": "dash" if label != "Mid" else "dot"},
                    hovertemplate=f"{label}: %{{y:.{decimals}f}}<extra></extra>",
                )
            )

    fig.update_layout(
        height=360,
        margin={"l": 20, "r": 20, "t": 18, "b": 20},
        paper_bgcolor="white",
        plot_bgcolor="white",
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "left", "x": 0},
        xaxis_title="Position",
        yaxis_title=f"Value ({display_unit})",
        font={"color": "#344054", "size": 12},
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eaecf0", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="#eaecf0", zeroline=False)
    return fig


def render_batch(batch: dict[str, Any], spec: dict[str, float], display_unit: str) -> None:
    source_unit = batch_source_unit(batch["type"])
    table = positions_dataframe(batch["parts"], source_unit, display_unit)
    table["Status"] = table.apply(lambda row: part_status(row, spec, batch["type"], display_unit), axis=1)

    st.markdown(
        f"""
        <div class="batch-header">
          <div>
            <div class="batch-label">{html.escape(batch["label"])}</div>
            <div class="batch-name">{html.escape(batch["name"])}</div>
          </div>
          <div class="batch-meta">{len(batch["parts"])} parts | Source: {source_unit} | Display: {display_unit}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if batch["warnings"]:
        with st.expander(f"Parsing warnings: {len(batch['warnings'])}", expanded=False):
            for warning in batch["warnings"]:
                st.warning(warning)

    st.plotly_chart(
        build_plotly_chart(table, spec, batch["type"], display_unit),
        use_container_width=True,
        config={"displayModeBar": False},
    )

    decimals = 5 if display_unit == "inch" else 4
    display_table = table[["Part", "AS U", "AS L", "DS U", "DS L", "Status"]].copy()

    def highlight_status(row: pd.Series) -> list[str]:
        if row["Status"] == "OOS":
            return ["background-color: #fff0f0; color: #b42318; font-weight: 700"] * len(row)
        if row["Status"] == "Missing":
            return ["background-color: #fff8e6; color: #915930; font-weight: 700"] * len(row)
        if row["Status"] == "OK":
            return ["color: #067647; font-weight: 600" if col == "Status" else "" for col in row.index]
        return ["" for _ in row.index]

    styler = display_table.style.format({position: f"{{:.{decimals}f}}" for position in POSITIONS}, na_rep="-").apply(highlight_status, axis=1)
    st.dataframe(styler, use_container_width=True, hide_index=True)

    if batch["type"] != "raw":
        spec_lower = convert_value(spec["lower"], "inch", display_unit)
        spec_mid = convert_value(spec["mid"], "inch", display_unit)
        spec_upper = convert_value(spec["upper"], "inch", display_unit)
        st.markdown(
            f"""
            <div class="spec-strip">
              <span>Spec reference ({display_unit})</span>
              <strong>Min {spec_lower:.{decimals}f}</strong>
              <strong>Mid {spec_mid:.{decimals}f}</strong>
              <strong>Max {spec_upper:.{decimals}f}</strong>
            </div>
            """,
            unsafe_allow_html=True,
        )


def build_html_report(batches: list[dict[str, Any]], spec_size: str, display_unit: str) -> str:
    spec = SPEC_LIMITS[spec_size]
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    html_parts = [
        "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>LL Sizing Report</title>",
        "<style>body{font-family:Arial,sans-serif;max-width:1200px;margin:0 auto;padding:24px;color:#333}"
        "h1,h2{color:#FF6B00}h1{border-bottom:3px solid #FF6B00;padding-bottom:10px}"
        "table{width:100%;border-collapse:collapse;margin:16px 0;font-size:13px}"
        "th{background:#FF6B00;color:white;padding:8px 12px;text-align:center}"
        "td{padding:6px 12px;text-align:center;border-bottom:1px solid #ddd;font-family:'Courier New',monospace}"
        ".meta{color:#666;font-size:12px}.warn{background:#FFF6D8;border:1px solid #E2B33C;padding:8px;margin:8px 0}"
        ".oos{color:#C22;font-weight:700;background:#FFF0F0}.ok{color:#067647;font-weight:700}</style></head><body>",
        "<h1>FOX LL Sizing Report</h1>",
        f"<p class='meta'>Generated: {html.escape(generated)} | Spec: {html.escape(spec_size)} "
        f"({spec['lower']:.4f}&quot; ~ {spec['upper']:.4f}&quot;) | Display unit: {html.escape(display_unit)}</p>",
    ]

    decimals = 5 if display_unit == "inch" else 4
    spec_lower = convert_value(spec["lower"], "inch", display_unit)
    spec_mid = convert_value(spec["mid"], "inch", display_unit)
    spec_upper = convert_value(spec["upper"], "inch", display_unit)

    for batch in batches:
        source_unit = batch_source_unit(batch["type"])
        table = positions_dataframe(batch["parts"], source_unit, display_unit)
        table["Status"] = table.apply(lambda row: part_status(row, spec, batch["type"], display_unit), axis=1)
        html_parts.append(f"<h2>{html.escape(batch['label'])} - {html.escape(batch['name'])}</h2>")
        html_parts.append(f"<p class='meta'>{len(batch['parts'])} parts | Source: {source_unit}</p>")
        if batch["warnings"]:
            html_parts.append(f"<div class='warn'><strong>Parsing warnings: {len(batch['warnings'])}</strong>")
            for warning in batch["warnings"][:10]:
                html_parts.append(f"<div>{html.escape(warning)}</div>")
            html_parts.append("</div>")

        html_parts.append("<table><thead><tr><th>Part</th><th>AS U</th><th>AS L</th><th>DS U</th><th>DS L</th><th>Status</th></tr></thead><tbody>")
        if batch["type"] != "raw":
            for label, value in (("Max", spec_upper), ("Mid", spec_mid), ("Min", spec_lower)):
                html_parts.append(
                    f"<tr><td>{label}</td>"
                    + "".join(f"<td>{value:.{decimals}f}</td>" for _ in POSITIONS)
                    + "<td></td></tr>"
                )
        for _, row in table.iterrows():
            row_class = " class='oos'" if row["Status"] == "OOS" else ""
            html_parts.append(f"<tr{row_class}><td>{html.escape(row['Part'])}</td>")
            for position in POSITIONS:
                value = row[position]
                out_of_spec = batch["type"] != "raw" and pd.notna(value) and (value < spec_lower or value > spec_upper)
                css_class = " class='oos'" if out_of_spec else ""
                text = "" if pd.isna(value) else f"{value:.{decimals}f}"
                html_parts.append(f"<td{css_class}>{text}</td>")
            status_class = " class='ok'" if row["Status"] == "OK" else ""
            html_parts.append(f"<td{status_class}>{html.escape(row['Status'])}</td></tr>")
        html_parts.append("</tbody></table>")

    html_parts.append("</body></html>")
    return "".join(html_parts)


def parse_uploaded_batch(name: str, batch_type: str, sizer_size: float | None, uploaded_files: list[Any]) -> dict[str, Any]:
    parts = []
    warnings = []
    for uploaded_file in uploaded_files:
        try:
            part = parse_xls_bytes(uploaded_file.getvalue(), uploaded_file.name)
            missing_positions = [key for key in POS_KEYS if key not in part["positions"]]
            if missing_positions:
                part["warnings"].append(f"Missing computed positions: {', '.join(missing_positions)}")
            for warning in part["warnings"]:
                warnings.append(f"{uploaded_file.name}: {warning}")
            parts.append(part)
        except Exception as exc:
            warnings.append(f"{uploaded_file.name}: {exc}")

    return {
        "name": name or "Uploaded Batch",
        "type": batch_type,
        "label": batch_label(batch_type, sizer_size),
        "sizer_size": sizer_size,
        "parts": parts,
        "warnings": warnings,
    }


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1.5rem;
            max-width: 1500px;
        }
        .app-title {
            font-size: 2rem;
            font-weight: 800;
            color: #101828;
            margin-bottom: 0.15rem;
        }
        .app-subtitle {
            color: #667085;
            font-size: 0.95rem;
            margin-bottom: 1rem;
        }
        .import-title {
            font-size: 1.05rem;
            font-weight: 800;
            color: #f05a22;
            margin-bottom: 0.25rem;
        }
        .import-note {
            color: #667085;
            font-size: 0.85rem;
            margin-bottom: 0.5rem;
        }
        .batch-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            padding: 0.9rem 1rem;
            margin-top: 1rem;
            background: #fff7ed;
            border: 1px solid #fed7aa;
            border-radius: 8px;
        }
        .batch-label {
            display: inline-block;
            color: #c2410c;
            font-size: 0.78rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }
        .batch-name {
            color: #101828;
            font-size: 1.2rem;
            font-weight: 800;
            line-height: 1.2;
        }
        .batch-meta {
            color: #475467;
            font-size: 0.86rem;
            white-space: nowrap;
        }
        .spec-strip {
            display: flex;
            flex-wrap: wrap;
            gap: 0.75rem;
            align-items: center;
            padding: 0.65rem 0.8rem;
            margin: 0.4rem 0 1rem;
            border: 1px solid #e4e7ec;
            border-radius: 8px;
            background: #f9fafb;
            color: #475467;
            font-size: 0.85rem;
        }
        .spec-strip strong {
            color: #b42318;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid #e4e7ec;
            border-radius: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_import_panel() -> None:
    with st.container(border=True):
        st.markdown('<div class="import-title">Import Batch Folder</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="import-note">For Streamlit Cloud, select all .xls files from one batch folder, then assign the batch type below.</div>',
            unsafe_allow_html=True,
        )
        name_col, type_col, sizer_col = st.columns([1.25, 1.4, 0.8])
        with name_col:
            batch_name = st.text_input("Batch name", value=f"Batch {len(st.session_state.batches) + 1}")
        with type_col:
            batch_type = st.radio(
                "Batch Type",
                options=("raw", "bushing_unsized", "bushing_sized"),
                horizontal=True,
                format_func=lambda value: {
                    "raw": "Raw Material",
                    "bushing_unsized": "Bushing ID unsized",
                    "bushing_sized": "Bushing ID sized",
                }[value],
            )
        with sizer_col:
            sizer_size = None
            if batch_type == "bushing_sized":
                sizer_size = st.number_input("Sizer size", min_value=0.5, max_value=3.0, value=1.34050, step=0.00001, format="%.5f")
            else:
                st.text_input("Sizer size", value="N/A", disabled=True)

        upload_col, action_col = st.columns([3, 1])
        with upload_col:
            uploader_key = f"batch_files_{st.session_state.uploader_nonce}"
            uploaded_files = st.file_uploader("Batch .xls files", type=["xls"], accept_multiple_files=True, key=uploader_key)
        with action_col:
            st.write("")
            st.write("")
            add_clicked = st.button("Add Batch", type="primary", use_container_width=True, disabled=not uploaded_files)
            clear_clicked = st.button("Clear All", use_container_width=True, disabled=not st.session_state.batches)

        if add_clicked:
            batch = parse_uploaded_batch(batch_name, batch_type, sizer_size, uploaded_files)
            if batch["parts"]:
                st.session_state.batches.append(batch)
                st.session_state.uploader_nonce += 1
                st.success(f"Added {batch['name']} with {len(batch['parts'])} parsed file(s).")
                st.rerun()
            else:
                st.error("No files could be parsed.")
        if clear_clicked:
            st.session_state.batches = []
            st.session_state.uploader_nonce += 1
            st.rerun()


def render_batch_registry() -> None:
    if not st.session_state.batches:
        return
    rows = []
    for index, batch in enumerate(st.session_state.batches, 1):
        rows.append({
            "#": index,
            "Batch": batch["name"],
            "Type": batch["label"],
            "Files": len(batch["parts"]),
            "Warnings": len(batch["warnings"]),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def main() -> None:
    if "batches" not in st.session_state:
        st.session_state.batches = []
    if "uploader_nonce" not in st.session_state:
        st.session_state.uploader_nonce = 0

    inject_css()
    st.markdown('<div class="app-title">LL Sizing Report</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subtitle">Upload one batch at a time, review AS/DS position tables, and download an HTML report.</div>',
        unsafe_allow_html=True,
    )

    render_import_panel()
    render_batch_registry()

    spec_col, unit_col, download_col = st.columns([1.2, 0.8, 1])
    with spec_col:
        spec_size = st.selectbox("Bore Size Spec", options=list(SPEC_LIMITS.keys()), index=1)
    with unit_col:
        display_unit = st.segmented_control("Display Unit", options=("inch", "mm"), default="inch")

    if not st.session_state.batches:
        st.info("Select .xls files in the Import Batch panel to begin.")
        return

    spec = SPEC_LIMITS[spec_size]
    with download_col:
        report_html = build_html_report(st.session_state.batches, spec_size, display_unit)
        filename = f"Sizing_Report_{datetime.now().strftime('%Y%m%d')}.html"
        st.download_button(
            "Download HTML Report",
            data=report_html,
            file_name=filename,
            mime="text/html",
            type="primary",
            use_container_width=True,
        )

    st.metric("Selected Spec", f'{spec["lower"]:.4f}" ~ {spec["upper"]:.4f}"')

    for batch in st.session_state.batches:
        render_batch(batch, spec, display_unit)
        st.divider()


if __name__ == "__main__":
    main()
