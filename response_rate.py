import io
import re
import calendar
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(page_title="BPI Dashboard - Table 1 and Table 2", layout="wide")

TABLE_STATUSES: List[str] = [
    "ALL NEGATIVE",
    "CALL OUTS",
    "EMAIL",
    "FLD VST",
    "SELF CURED",
]

MONTH_NAME_TO_NUM = {
    "JAN": 1, "JANUARY": 1,
    "FEB": 2, "FEBRUARY": 2,
    "MAR": 3, "MARCH": 3,
    "APR": 4, "APRIL": 4,
    "MAY": 5,
    "JUN": 6, "JUNE": 6,
    "JUL": 7, "JULY": 7,
    "AUG": 8, "AUGUST": 8,
    "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
    "OCT": 10, "OCTOBER": 10,
    "NOV": 11, "NOVEMBER": 11,
    "DEC": 12, "DECEMBER": 12,
}


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


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


def resolve_column(
    df: pd.DataFrame,
    *,
    by_name: Optional[str] = None,
    by_excel_index: Optional[int] = None,
) -> str:
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


def extract_month_number(value: object) -> Optional[int]:
    """
    Tries to convert CUT OFF MONTH values into a month number.

    Supported examples:
    - 1, 2, 3
    - Jan, January, FEBRUARY
    - 2024-01-31
    - 01/15/2024
    """
    if pd.isna(value):
        return None

    # Excel / pandas date-like values
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.notna(dt):
            return int(dt.month)
    except Exception:
        pass

    text = normalize_text(value)
    if not text:
        return None

    if text in MONTH_NAME_TO_NUM:
        return MONTH_NAME_TO_NUM[text]

    # Look for month names inside longer strings
    for name, month_num in MONTH_NAME_TO_NUM.items():
        if re.search(rf"\b{name}\b", text):
            return month_num

    # If it is numeric text like "1", "01", "3.0"
    numeric_match = re.fullmatch(r"(\d{1,2})(?:\.0+)?", text)
    if numeric_match:
        month_num = int(numeric_match.group(1))
        if 1 <= month_num <= 12:
            return month_num

    return None


def month_num_to_name(month_num: int) -> str:
    return calendar.month_name[month_num].upper()


def get_month_span_label(df: pd.DataFrame) -> str:
    """
    Builds a title suffix from all months present in CUT OFF MONTH.
    Examples:
    - JANUARY
    - JANUARY TO MARCH
    - JANUARY TO DECEMBER
    - ALL MONTHS
    """
    cutoff_col = resolve_column(df, by_name="CUT OFF MONTH", by_excel_index=8)

    months = (
        df[cutoff_col]
        .map(extract_month_number)
        .dropna()
        .astype(int)
        .tolist()
    )

    if not months:
        return "ALL MONTHS"

    unique_months = sorted(set(months))
    if len(unique_months) == 1:
        return month_num_to_name(unique_months[0])

    return f"{month_num_to_name(min(unique_months))} TO {month_num_to_name(max(unique_months))}"


def filter_by_cutoff_month(df: pd.DataFrame, target_month: int) -> pd.DataFrame:
    """
    Helper for the next tables.
    This filters the dataset to one month based on CUT OFF MONTH.
    """
    cutoff_col = resolve_column(df, by_name="CUT OFF MONTH", by_excel_index=8)
    working = df.copy()
    working["_cutoff_month_num"] = working[cutoff_col].map(extract_month_number)
    return working[working["_cutoff_month_num"] == target_month].copy()


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
        .assign(_cycle_num=lambda x: x[cycle_col].astype(str).str.extract(r"(\d+)").astype(float))
        .sort_values(["_cycle_num", cycle_col])[cycle_col]
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


def build_percentage_table(summary: pd.DataFrame) -> pd.DataFrame:
    detail = summary[summary[("", "Row Labels")].astype(str).str.upper() != "GRAND TOTAL"].copy()

    rows: List[Dict[Tuple[str, str], object]] = []
    for _, row in detail.iterrows():
        total_count = row[("TOTAL", "TOTAL COUNT OF CASES")]
        total_ob = row[("TOTAL", "TOTAL OB")]

        pct_row: Dict[Tuple[str, str], object] = {
            ("", "Row Labels"): row[("", "Row Labels")]
        }

        for status in TABLE_STATUSES:
            count_val = row[(status, "COUNT OF ACCOUNT CYCLE")]
            ob_val = row[(status, "OB PER CYCLE")]

            pct_row[(status, "COUNT %")] = (count_val / total_count) if total_count else 0
            pct_row[(status, "OB %")] = (ob_val / total_ob) if total_ob else 0

        pct_row[("TOTAL", "TOTAL COUNT %")] = 1 if total_count else 0
        pct_row[("TOTAL", "TOTAL OB %")] = 1 if total_ob else 0
        rows.append(pct_row)

    pct_df = pd.DataFrame(rows)
    pct_df.columns = pd.MultiIndex.from_tuples(pct_df.columns)

    grand_total_row = summary[summary[("", "Row Labels")].astype(str).str.upper() == "GRAND TOTAL"].iloc[0]
    grand_total_count = grand_total_row[("TOTAL", "TOTAL COUNT OF CASES")]
    grand_total_ob = grand_total_row[("TOTAL", "TOTAL OB")]

    total_pct_row: Dict[Tuple[str, str], object] = {("", "Row Labels"): "Grand Total"}
    for status in TABLE_STATUSES:
        status_count_total = grand_total_row[(status, "COUNT OF ACCOUNT CYCLE")]
        status_ob_total = grand_total_row[(status, "OB PER CYCLE")]

        total_pct_row[(status, "COUNT %")] = (
            status_count_total / grand_total_count if grand_total_count else 0
        )
        total_pct_row[(status, "OB %")] = (
            status_ob_total / grand_total_ob if grand_total_ob else 0
        )

    total_pct_row[("TOTAL", "TOTAL COUNT %")] = 1 if grand_total_count else 0
    total_pct_row[("TOTAL", "TOTAL OB %")] = 1 if grand_total_ob else 0

    pct_df = pd.concat([pct_df, pd.DataFrame([total_pct_row])], ignore_index=True)
    return pct_df


def to_flat_preview(df: pd.DataFrame) -> pd.DataFrame:
    preview = df.copy()
    preview.columns = [
        sub if top in ("", "TOTAL") else f"{top} - {sub}"
        for top, sub in preview.columns
    ]
    return preview


def to_flat_preview_pct(pct_df: pd.DataFrame) -> pd.DataFrame:
    preview = pct_df.copy()
    preview.columns = [
        sub if top in ("", "TOTAL") else f"{top} - {sub}"
        for top, sub in preview.columns
    ]
    return preview


def apply_cell_style(cell, fill, font, alignment, border, number_format=None):
    cell.fill = fill
    cell.font = font
    cell.alignment = alignment
    cell.border = border
    if number_format:
        cell.number_format = number_format


def build_formatted_excel(summary: pd.DataFrame, pct_df: pd.DataFrame, month_span_label: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Table 1"

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

    columns = [
        ("A", 18),
        ("B", 16), ("C", 14),
        ("D", 16), ("E", 14),
        ("F", 16), ("G", 14),
        ("H", 16), ("I", 14),
        ("J", 16), ("K", 14),
        ("L", 16), ("M", 14),
    ]
    for col, width in columns:
        ws.column_dimensions[col].width = width

    table1_title = f"OVERALL RESPONSE BY CASES AND BALANCE ({month_span_label})"
    table2_title = "OVERALL RESPONSE RATE"

    # =========================
    # TABLE 1
    # =========================
    ws.merge_cells("A1:M1")
    ws["A1"] = table1_title
    apply_cell_style(ws["A1"], fill_title, font_title, left, border)

    ws.merge_cells("A2:A3")
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

    for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=13):
        for cell in row:
            apply_cell_style(cell, fill_red, font_header, center, border)

    sub_headers = [
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
        "COUNT OF ACCOUNT CYCLE", "OB PER CYCLE",
    ]

    for idx, text in enumerate(sub_headers, start=2):
        ws.cell(row=3, column=idx, value=text)

    for row in ws.iter_rows(min_row=3, max_row=3, min_col=2, max_col=11):
        for cell in row:
            apply_cell_style(cell, fill_red, font_header, center, border)

    apply_cell_style(ws["A2"], fill_red, font_header, center, border)
    ws["A3"].border = border
    apply_cell_style(ws["L2"], fill_red, font_header, center, border)
    ws["L3"].border = border
    apply_cell_style(ws["M2"], fill_red, font_header, center, border)
    ws["M3"].border = border

    start_row = 4
    numeric_count_cols = [2, 4, 6, 8, 10, 12]
    numeric_ob_cols = [3, 5, 7, 9, 11, 13]

    for i, (_, row) in enumerate(summary.iterrows(), start=start_row):
        row_label = str(row[("", "Row Labels")])
        is_total = normalize_text(row_label) == "GRAND TOTAL"
        body_font = font_total if is_total else font_body

        values = [
            row_label,
            int(row[("ALL NEGATIVE", "COUNT OF ACCOUNT CYCLE")]),
            float(row[("ALL NEGATIVE", "OB PER CYCLE")]),
            int(row[("CALL OUTS", "COUNT OF ACCOUNT CYCLE")]),
            float(row[("CALL OUTS", "OB PER CYCLE")]),
            int(row[("EMAIL", "COUNT OF ACCOUNT CYCLE")]),
            float(row[("EMAIL", "OB PER CYCLE")]),
            int(row[("FLD VST", "COUNT OF ACCOUNT CYCLE")]),
            float(row[("FLD VST", "OB PER CYCLE")]),
            int(row[("SELF CURED", "COUNT OF ACCOUNT CYCLE")]),
            float(row[("SELF CURED", "OB PER CYCLE")]),
            int(row[("TOTAL", "TOTAL COUNT OF CASES")]),
            float(row[("TOTAL", "TOTAL OB")]),
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=i, column=col_idx, value=value)
            apply_cell_style(cell, fill_white, body_font, left if col_idx == 1 else right, border)

            if col_idx in numeric_count_cols:
                cell.number_format = "#,##0"
            elif col_idx in numeric_ob_cols:
                cell.number_format = "#,##0"

    # =========================
    # TABLE 2
    # =========================
    gap_row = start_row + len(summary) + 3
    title_row = gap_row - 1
    header_top = gap_row
    header_sub = gap_row + 1
    data_start = gap_row + 2

    ws.merge_cells(f"A{title_row}:M{title_row}")
    ws[f"A{title_row}"] = table2_title
    apply_cell_style(ws[f"A{title_row}"], fill_title, font_title, left, border)

    ws.merge_cells(f"A{header_top}:A{header_sub}")
    ws.merge_cells(f"B{header_top}:C{header_top}")
    ws.merge_cells(f"D{header_top}:E{header_top}")
    ws.merge_cells(f"F{header_top}:G{header_top}")
    ws.merge_cells(f"H{header_top}:I{header_top}")
    ws.merge_cells(f"J{header_top}:K{header_top}")
    ws.merge_cells(f"L{header_top}:L{header_sub}")
    ws.merge_cells(f"M{header_top}:M{header_sub}")

    ws[f"A{header_top}"] = "Row Labels"
    ws[f"B{header_top}"] = "ALL NEGATIVE"
    ws[f"D{header_top}"] = "CALL OUTS"
    ws[f"F{header_top}"] = "EMAIL"
    ws[f"H{header_top}"] = "FLD VST"
    ws[f"J{header_top}"] = "SELF CURED"
    ws[f"L{header_top}"] = "TOTAL COUNT %"
    ws[f"M{header_top}"] = "TOTAL OB %"

    for row in ws.iter_rows(min_row=header_top, max_row=header_top, min_col=1, max_col=13):
        for cell in row:
            apply_cell_style(cell, fill_red, font_header, center, border)

    pct_sub_headers = [
        "COUNT %", "OB %",
        "COUNT %", "OB %",
        "COUNT %", "OB %",
        "COUNT %", "OB %",
        "COUNT %", "OB %",
    ]

    for idx, text in enumerate(pct_sub_headers, start=2):
        ws.cell(row=header_sub, column=idx, value=text)

    for row in ws.iter_rows(min_row=header_sub, max_row=header_sub, min_col=2, max_col=11):
        for cell in row:
            apply_cell_style(cell, fill_red, font_header, center, border)

    apply_cell_style(ws[f"A{header_top}"], fill_red, font_header, center, border)
    ws[f"A{header_sub}"].border = border
    apply_cell_style(ws[f"L{header_top}"], fill_red, font_header, center, border)
    ws[f"L{header_sub}"].border = border
    apply_cell_style(ws[f"M{header_top}"], fill_red, font_header, center, border)
    ws[f"M{header_sub}"].border = border

    for i, (_, row) in enumerate(pct_df.iterrows(), start=data_start):
        row_label = str(row[("", "Row Labels")])
        is_total = normalize_text(row_label) == "GRAND TOTAL"
        body_font = font_total if is_total else font_body

        label_cell = ws.cell(row=i, column=1, value=row_label)
        apply_cell_style(label_cell, fill_white, body_font, left, border)

        values = [
            row[("ALL NEGATIVE", "COUNT %")],
            row[("ALL NEGATIVE", "OB %")],
            row[("CALL OUTS", "COUNT %")],
            row[("CALL OUTS", "OB %")],
            row[("EMAIL", "COUNT %")],
            row[("EMAIL", "OB %")],
            row[("FLD VST", "COUNT %")],
            row[("FLD VST", "OB %")],
            row[("SELF CURED", "COUNT %")],
            row[("SELF CURED", "OB %")],
            row[("TOTAL", "TOTAL COUNT %")],
            row[("TOTAL", "TOTAL OB %")],
        ]

        for col_idx, value in enumerate(values, start=2):
            cell = ws.cell(row=i, column=col_idx, value=float(value))
            apply_cell_style(cell, fill_white, body_font, right, border, "0%")

    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 32
    ws.row_dimensions[title_row].height = 22
    ws.row_dimensions[header_top].height = 20
    ws.row_dimensions[header_sub].height = 20

    # No freeze panes, as requested.

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.title("BPI Project Dashboard Automation")
st.caption("Table 1 and Table 2 use all months found in CUT OFF MONTH")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("Upload the workbook first.")
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
    month_span_label = get_month_span_label(source_df)
    summary = build_summary_table(source_df)
    pct_df = build_percentage_table(summary)
except Exception as exc:
    st.error(f"Unable to build the tables: {exc}")
    st.stop()

st.subheader("Table 1 Preview")
st.dataframe(to_flat_preview(summary), use_container_width=True, hide_index=True)

st.subheader("Table 2 Preview")
st.dataframe(to_flat_preview(pct_df), use_container_width=True, hide_index=True)

st.info(
    f"Detected CUT OFF MONTH range for overall tables: {month_span_label}. "
    f"These first two tables use all rows in the uploaded file."
)

# Generate Table 3 and Table 4 for each month (January, February, March, etc.)
for month_num in range(1, 13):
    filtered_data = filter_by_cutoff_month(source_df, month_num)
    if not filtered_data.empty:
        month_name = month_num_to_name(month_num)
        st.subheader(f"Table 3 - {month_name} - Summary")
        month_summary = build_summary_table(filtered_data)
        st.dataframe(to_flat_preview(month_summary), use_container_width=True, hide_index=True)

        st.subheader(f"Table 4 - {month_name} - Percentage")
        month_pct_df = build_percentage_table(month_summary)
        st.dataframe(to_flat_preview_pct(month_pct_df), use_container_width=True, hide_index=True)

        try:
            excel_bytes = build_formatted_excel(month_summary, month_pct_df, month_name)
        except Exception as exc:
            st.error(f"Unable to format the Excel output file for {month_name}: {exc}")
            continue

        st.download_button(
            label=f"Download {month_name} formatted Excel",
            data=excel_bytes,
            file_name=f"{month_name}_table3_table4_formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
