import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="BPI Dashboard - Table 1", layout="wide")

TABLE_STATUSES: List[str] = [
    "ALL NEGATIVE",
    "CALL OUTS",
    "EMAIL",
    "FLD VST",
    "SELF CURED",
]


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


# Optional alias map so small variations in the source file still land in the expected bucket.
STATUS_ALIASES = {
    "CALL OUT": "CALL OUTS",
    "CALL-OUTS": "CALL OUTS",
    "FIELD VISIT": "FLD VST",
    "FLD VISIT": "FLD VST",
    "SELF-CURED": "SELF CURED",
}


def normalize_status(value: object) -> str:
    text = normalize_text(value)
    return STATUS_ALIASES.get(text, text)



def resolve_column(df: pd.DataFrame, *, by_name: str | None = None, by_excel_index: int | None = None) -> str:
    if by_name:
        normalized_target = normalize_text(by_name)
        for col in df.columns:
            if normalize_text(col) == normalized_target:
                return col

    if by_excel_index is not None:
        zero_based = by_excel_index - 1
        if 0 <= zero_based < len(df.columns):
            return df.columns[zero_based]

    missing = by_name or f"Excel column #{by_excel_index}"
    raise KeyError(f"Could not find required column: {missing}")


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    return {
        sheet_name: pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
        for sheet_name in excel_file.sheet_names
    }



def build_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    cycle_col = resolve_column(df, by_name="CYCLE", by_excel_index=2)
    ob_col = resolve_column(df, by_name="OB", by_excel_index=5)
    contact_col = resolve_column(df, by_name="CONTACT SOURCE (OVERALL)", by_excel_index=55)

    working = df.copy()
    working[cycle_col] = working[cycle_col].fillna("").astype(str).str.strip()
    working[contact_col] = working[contact_col].map(normalize_status)
    working[ob_col] = pd.to_numeric(working[ob_col], errors="coerce").fillna(0)
    working = working[working[cycle_col].ne("")].copy()

    cycle_order = (
        working[[cycle_col]]
        .drop_duplicates()
        .assign(_cycle_num=lambda x: x[cycle_col].str.extract(r"(\d+)").astype(float))
        .sort_values(["_cycle_num", cycle_col])
        [cycle_col]
        .tolist()
    )

    rows: List[Dict[Tuple[str, str], object]] = []
    for cycle in cycle_order:
        cycle_df = working[working[cycle_col] == cycle]
        row: Dict[Tuple[str, str], object] = {("", "Row Labels"): cycle}

        for status in TABLE_STATUSES:
            status_df = cycle_df[cycle_df[contact_col] == status]
            row[(status, "COUNT OF ACCOUNT CYCLE")] = int(len(status_df))
            row[(status, "OB PER CYCLE")] = float(status_df[ob_col].sum())

        row[("TOTAL", "TOTAL COUNT OF CASES")] = int(len(cycle_df))
        row[("TOTAL", "TOTAL OB")] = float(cycle_df[ob_col].sum())
        rows.append(row)

    summary = pd.DataFrame(rows)
    summary.columns = pd.MultiIndex.from_tuples(summary.columns)

    total_row: Dict[Tuple[str, str], object] = {("", "Row Labels"): "Grand Total"}
    for col in summary.columns:
        if col != ("", "Row Labels"):
            total_row[col] = summary[col].sum()

    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)
    return summary



def to_flat_preview(summary: pd.DataFrame) -> pd.DataFrame:
    preview = summary.copy()
    preview.columns = [
        sub if top in ("", "TOTAL") else f"{top} - {sub}"
        for top, sub in preview.columns
    ]
    return preview



def build_formatted_excel(summary: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Table 1"

    # Palette close to the screenshot.
    fill_title = PatternFill("solid", fgColor="D9EAD3")
    fill_red = PatternFill("solid", fgColor="C00000")
    fill_white = PatternFill("solid", fgColor="FFFFFF")
    thin_black = Side(style="thin", color="000000")
    border = Border(left=thin_black, right=thin_black, top=thin_black, bottom=thin_black)

    font_title = Font(name="Arial", size=12, bold=True, color="000000")
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=10, bold=False, color="000000")
    font_total = Font(name="Arial", size=10, bold=True, color="000000")

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    right = Alignment(horizontal="right", vertical="center")

    # Column layout A:M to mirror the screenshot.
    columns = [
        ("A", 18),
        ("B", 14), ("C", 14),
        ("D", 14), ("E", 14),
        ("F", 14), ("G", 14),
        ("H", 14), ("I", 14),
        ("J", 14), ("K", 14),
        ("L", 16), ("M", 14),
    ]
    for col, width in columns:
        ws.column_dimensions[col].width = width

    # Title row.
    ws.merge_cells("A1:M1")
    ws["A1"] = "OVERALL RESPONSE BY CASES AND BALANCE (JANUARY TO MARCH)"
    ws["A1"].fill = fill_title
    ws["A1"].font = font_title
    ws["A1"].alignment = left
    ws["A1"].border = border

    # Group headers.
    ws.merge_cells("B2:C2")
    ws.merge_cells("D2:E2")
    ws.merge_cells("F2:G2")
    ws.merge_cells("H2:I2")
    ws.merge_cells("J2:K2")
    ws.merge_cells("L2:L3")
    ws.merge_cells("M2:M3")

    ws["A2"] = "Row Labels"
    ws["B2"] = "ALL NEGATIVE"
    ws["D2"] = "CALL OUTS"
    ws["F2"] = "EMAIL"
    ws["H2"] = "FLD VST"
    ws["J2"] = "SELF CURED"
    ws["L2"] = "TOTAL COUNT OF CASES"
    ws["M2"] = "TOTAL OB"

    for cell in ["A2", "B2", "D2", "F2", "H2", "J2", "L2", "M2"]:
        ws[cell].fill = fill_red
        ws[cell].font = font_header
        ws[cell].alignment = center
        ws[cell].border = border

    # Need border/fill across merged header ranges.
    for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=13):
        for cell in row:
            cell.fill = fill_red
            cell.font = font_header
            cell.alignment = center
            cell.border = border

    sub_headers = [
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
    ]
    ws["A3"] = ""
    for idx, text in enumerate(sub_headers, start=2):
        ws.cell(row=3, column=idx, value=text)

    for row in ws.iter_rows(min_row=3, max_row=3, min_col=1, max_col=11):
        for cell in row:
            cell.fill = fill_red
            cell.font = font_header
            cell.alignment = center
            cell.border = border

    # Body rows.
    start_row = 4
    numeric_count_cols = [2, 4, 6, 8, 10, 12]
    numeric_ob_cols = [3, 5, 7, 9, 11, 13]

    for i, (_, row) in enumerate(summary.iterrows(), start=start_row):
        row_label = str(row[("", "Row Labels")])
        is_total = normalize_text(row_label) == "GRAND TOTAL"
        body_font = font_total if is_total else font_body

        ws.cell(row=i, column=1, value=row_label)
        ws.cell(row=i, column=2, value=int(row[("ALL NEGATIVE", "COUNT OF ACCOUNT CYCLE")]))
        ws.cell(row=i, column=3, value=float(row[("ALL NEGATIVE", "OB PER CYCLE")]))
        ws.cell(row=i, column=4, value=int(row[("CALL OUTS", "COUNT OF ACCOUNT CYCLE")]))
        ws.cell(row=i, column=5, value=float(row[("CALL OUTS", "OB PER CYCLE")]))
        ws.cell(row=i, column=6, value=int(row[("EMAIL", "COUNT OF ACCOUNT CYCLE")]))
        ws.cell(row=i, column=7, value=float(row[("EMAIL", "OB PER CYCLE")]))
        ws.cell(row=i, column=8, value=int(row[("FLD VST", "COUNT OF ACCOUNT CYCLE")]))
        ws.cell(row=i, column=9, value=float(row[("FLD VST", "OB PER CYCLE")]))
        ws.cell(row=i, column=10, value=int(row[("SELF CURED", "COUNT OF ACCOUNT CYCLE")]))
        ws.cell(row=i, column=11, value=float(row[("SELF CURED", "OB PER CYCLE")]))
        ws.cell(row=i, column=12, value=int(row[("TOTAL", "TOTAL COUNT OF CASES")]))
        ws.cell(row=i, column=13, value=float(row[("TOTAL", "TOTAL OB")]))

        for col_idx in range(1, 14):
            cell = ws.cell(row=i, column=col_idx)
            cell.fill = fill_white
            cell.font = body_font
            cell.border = border
            cell.alignment = left if col_idx == 1 else right
            if col_idx in numeric_count_cols:
                cell.number_format = '#,##0'
            elif col_idx in numeric_ob_cols:
                cell.number_format = '#,##0'

    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 32

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.title("BPI Project Dashboard Automation")
st.caption("Table 1: generates the Excel output file with the same table-style layout")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("Upload the workbook first to generate the formatted Table 1 export.")
    st.stop()

file_bytes = uploaded_file.getvalue()

try:
    workbook = load_workbook(file_bytes)
except Exception as exc:
    st.error(f"Unable to read the workbook: {exc}")
    st.stop()

sheet_name = st.selectbox("Select source sheet", options=list(workbook.keys()), index=0)
source_df = workbook[sheet_name]

with st.expander("Preview source data", expanded=False):
    st.dataframe(source_df.head(20), use_container_width=True)

try:
    summary = build_summary_table(source_df)
except Exception as exc:
    st.error(f"Unable to build Table 1: {exc}")
    st.stop()

st.subheader("Table 1 Preview")
st.dataframe(to_flat_preview(summary), use_container_width=True, hide_index=True)

try:
    excel_bytes = build_formatted_excel(summary)
except Exception as exc:
    st.error(f"Unable to format the Excel output file: {exc}")
    st.stop()

st.download_button(
    label="Download formatted Table 1 Excel",
    data=excel_bytes,
    file_name="table1_formatted_output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
