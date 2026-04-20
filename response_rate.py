import io
import re
import calendar
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(page_title="BPI Dashboard - Tables", layout="wide")

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
    if pd.isna(value):
        return None
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

    for name, month_num in MONTH_NAME_TO_NUM.items():
        if re.search(rf"\b{name}\b", text):
            return month_num

    numeric_match = re.fullmatch(r"(\d{1,2})(?:\.0+)?", text)
    if numeric_match:
        month_num = int(numeric_match.group(1))
        if 1 <= month_num <= 12:
            return month_num

    return None


def month_num_to_name(month_num: int) -> str:
    return calendar.month_name[month_num].upper()


def get_month_span_label(df: pd.DataFrame) -> str:
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


def get_detected_months(df: pd.DataFrame) -> List[int]:
    cutoff_col = resolve_column(df, by_name="CUT OFF MONTH", by_excel_index=8)
    months = (
        df[cutoff_col]
        .map(extract_month_number)
        .dropna()
        .astype(int)
        .tolist()
    )
    return sorted(set(months))


def filter_by_cutoff_month(df: pd.DataFrame, target_month: int) -> pd.DataFrame:
    cutoff_col = resolve_column(df, by_name="CUT OFF MONTH", by_excel_index=8)
    working = df.copy()
    working["_cutoff_month_num"] = working[cutoff_col].map(extract_month_number)
    return working[working["_cutoff_month_num"] == target_month].copy()


def filter_ptp(df: pd.DataFrame) -> pd.DataFrame:
    """Filter to PTP rows: REMARKS (PTP/NO PTP) == 'PTP'"""
    remarks_col = resolve_column(df, by_name="REMARKS (PTP/NO PTP)", by_excel_index=61)
    return df[df[remarks_col].apply(normalize_text) == "PTP"].copy()


def filter_cured(df: pd.DataFrame) -> pd.DataFrame:
    """Filter to CURED rows: FINAL STATUS == 'CURED'"""
    final_col = resolve_column(df, by_name="FINAL STATUS", by_excel_index=11)
    return df[df[final_col].apply(normalize_text) == "CURED"].copy()


def build_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    """Overall summary table: COUNT OF ACCOUNT CYCLE + OB PER CYCLE"""
    cycle_col   = resolve_column(df, by_name="CYCLE",                    by_excel_index=2)
    ob_col      = resolve_column(df, by_name="OB",                       by_excel_index=5)
    contact_col = resolve_column(df, by_name="CONTACT SOURCE (OVERALL)", by_excel_index=55)

    working = df.copy()
    working[cycle_col]   = working[cycle_col].fillna("").astype(str).str.strip()
    working[contact_col] = working[contact_col].map(normalize_status)
    working[ob_col]      = pd.to_numeric(working[ob_col], errors="coerce").fillna(0)
    working = working[working[cycle_col].ne("")].copy()

    cycle_order = (
        working[[cycle_col]]
        .drop_duplicates()
        .assign(_n=lambda x: x[cycle_col].astype(str).str.extract(r"(\d+)").astype(float))
        .sort_values(["_n", cycle_col])[cycle_col]
        .tolist()
    )

    rows: List[Dict] = []
    for cycle in cycle_order:
        cycle_df = working[working[cycle_col] == cycle]
        row: Dict = {("", "Row Labels"): cycle}
        for status in TABLE_STATUSES:
            s_df = cycle_df[cycle_df[contact_col] == status]
            row[(status, "COUNT OF ACCOUNT CYCLE")] = int(len(s_df))
            row[(status, "OB PER CYCLE")]            = float(s_df[ob_col].sum())
        row[("TOTAL", "TOTAL COUNT OF CASES")] = int(len(cycle_df))
        row[("TOTAL", "TOTAL OB")]             = float(cycle_df[ob_col].sum())
        rows.append(row)

    summary = pd.DataFrame(rows)
    summary.columns = pd.MultiIndex.from_tuples(summary.columns)

    total_row: Dict = {("", "Row Labels"): "Grand Total"}
    for col in summary.columns:
        if col != ("", "Row Labels"):
            total_row[col] = summary[col].sum()
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)
    return summary


def build_sub_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    """PTP / CURED summary table: NO. OF CASES + OB"""
    cycle_col   = resolve_column(df, by_name="CYCLE",                    by_excel_index=2)
    ob_col      = resolve_column(df, by_name="OB",                       by_excel_index=5)
    contact_col = resolve_column(df, by_name="CONTACT SOURCE (OVERALL)", by_excel_index=55)

    working = df.copy()
    working[cycle_col]   = working[cycle_col].fillna("").astype(str).str.strip()
    working[contact_col] = working[contact_col].map(normalize_status)
    working[ob_col]      = pd.to_numeric(working[ob_col], errors="coerce").fillna(0)
    working = working[working[cycle_col].ne("")].copy()

    cycle_order = (
        working[[cycle_col]]
        .drop_duplicates()
        .assign(_n=lambda x: x[cycle_col].astype(str).str.extract(r"(\d+)").astype(float))
        .sort_values(["_n", cycle_col])[cycle_col]
        .tolist()
    )

    rows: List[Dict] = []
    for cycle in cycle_order:
        cycle_df = working[working[cycle_col] == cycle]
        row: Dict = {("", "Row Labels"): cycle}
        for status in TABLE_STATUSES:
            s_df = cycle_df[cycle_df[contact_col] == status]
            row[(status, "NO. OF CASES")] = int(len(s_df))
            row[(status, "OB")]           = float(s_df[ob_col].sum())
        row[("TOTAL", "TOTAL COUNT OF CASES")] = int(len(cycle_df))
        row[("TOTAL", "TOTAL OB")]             = float(cycle_df[ob_col].sum())
        rows.append(row)

    summary = pd.DataFrame(rows)
    summary.columns = pd.MultiIndex.from_tuples(summary.columns)

    total_row: Dict = {("", "Row Labels"): "Grand Total"}
    for col in summary.columns:
        if col != ("", "Row Labels"):
            total_row[col] = summary[col].sum()
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)
    return summary


def build_percentage_table(summary: pd.DataFrame, is_sub: bool = False) -> pd.DataFrame:
    """Build rate table. is_sub=True uses NO. OF CASES/OB keys."""
    count_key = "NO. OF CASES"        if is_sub else "COUNT OF ACCOUNT CYCLE"
    ob_key    = "OB"                   if is_sub else "OB PER CYCLE"

    detail = summary[summary[("", "Row Labels")].astype(str).str.upper() != "GRAND TOTAL"].copy()
    rows: List[Dict] = []

    for _, row in detail.iterrows():
        total_count = row[("TOTAL", "TOTAL COUNT OF CASES")]
        total_ob    = row[("TOTAL", "TOTAL OB")]
        pct_row: Dict = {("", "Row Labels"): row[("", "Row Labels")]}
        for status in TABLE_STATUSES:
            pct_row[(status, "COUNT %")] = (row[(status, count_key)] / total_count) if total_count else 0
            pct_row[(status, "OB %")]    = (row[(status, ob_key)]    / total_ob)    if total_ob    else 0
        pct_row[("TOTAL", "TOTAL COUNT %")] = 1 if total_count else 0
        pct_row[("TOTAL", "TOTAL OB %")]    = 1 if total_ob    else 0
        rows.append(pct_row)

    pct_df = pd.DataFrame(rows)
    pct_df.columns = pd.MultiIndex.from_tuples(pct_df.columns)

    gt = summary[summary[("", "Row Labels")].astype(str).str.upper() == "GRAND TOTAL"].iloc[0]
    gt_count = gt[("TOTAL", "TOTAL COUNT OF CASES")]
    gt_ob    = gt[("TOTAL", "TOTAL OB")]

    gt_row: Dict = {("", "Row Labels"): "Grand Total"}
    for status in TABLE_STATUSES:
        gt_row[(status, "COUNT %")] = (gt[(status, count_key)] / gt_count) if gt_count else 0
        gt_row[(status, "OB %")]    = (gt[(status, ob_key)]    / gt_ob)    if gt_ob    else 0
    gt_row[("TOTAL", "TOTAL COUNT %")] = 1 if gt_count else 0
    gt_row[("TOTAL", "TOTAL OB %")]    = 1 if gt_ob    else 0

    pct_df = pd.concat([pct_df, pd.DataFrame([gt_row])], ignore_index=True)
    return pct_df


def to_flat_preview(df: pd.DataFrame) -> pd.DataFrame:
    preview = df.copy()
    preview.columns = [
        sub if top in ("", "TOTAL") else f"{top} - {sub}"
        for top, sub in preview.columns
    ]
    return preview


# ─────────────────────────────────────────────────────────────────────────────
#  Excel helpers
# ─────────────────────────────────────────────────────────────────────────────

def _make_styles() -> dict:
    fill_title = PatternFill("solid", fgColor="D9EAD3")
    fill_red   = PatternFill("solid", fgColor="C00000")
    fill_white = PatternFill("solid", fgColor="FFFFFF")
    thin       = Side(style="thin", color="000000")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    return dict(
        fill_title=fill_title, fill_red=fill_red, fill_white=fill_white,
        border=border,
        font_title=Font(name="Arial", size=12, bold=True,  color="000000"),
        font_hdr  =Font(name="Arial", size=10, bold=True,  color="FFFFFF"),
        font_body =Font(name="Arial", size=10, bold=False, color="000000"),
        font_total=Font(name="Arial", size=10, bold=True,  color="000000"),
        center=Alignment(horizontal="center", vertical="center"),
        left  =Alignment(horizontal="left",   vertical="center"),
        right =Alignment(horizontal="right",  vertical="center"),
    )


def _apply(cell, fill, font, alignment, border, number_format=None):
    cell.fill      = fill
    cell.font      = font
    cell.alignment = alignment
    cell.border    = border
    if number_format:
        cell.number_format = number_format


def _write_summary_block(ws, summary: pd.DataFrame, title: str,
                          start_row: int, s: dict, is_sub: bool = False) -> int:
    """Write a count/OB summary block. Returns next available row."""
    count_lbl = "NO. OF CASES"        if is_sub else "COUNT OF ACCOUNT CYCLE"
    ob_lbl    = "OB"                   if is_sub else "OB PER CYCLE"

    # title
    ws.merge_cells(f"A{start_row}:M{start_row}")
    ws[f"A{start_row}"] = title
    _apply(ws[f"A{start_row}"], s["fill_title"], s["font_title"], s["left"], s["border"])

    hr1 = start_row + 1
    hr2 = start_row + 2

    # header row 1
    for tag, col in [("A", "A"), ("B", "B"), ("D", "D"), ("F", "F"),
                     ("H", "H"), ("J", "J"), ("L", "L"), ("M", "M")]:
        pass  # handled below with merge

    ws.merge_cells(f"A{hr1}:A{hr2}")
    ws.merge_cells(f"B{hr1}:C{hr1}")
    ws.merge_cells(f"D{hr1}:E{hr1}")
    ws.merge_cells(f"F{hr1}:G{hr1}")
    ws.merge_cells(f"H{hr1}:I{hr1}")
    ws.merge_cells(f"J{hr1}:K{hr1}")
    ws.merge_cells(f"L{hr1}:L{hr2}")
    ws.merge_cells(f"M{hr1}:M{hr2}")

    ws[f"A{hr1}"] = "Row Labels"
    ws[f"B{hr1}"] = "ALL NEGATIVE"
    ws[f"D{hr1}"] = "CALL OUTS"
    ws[f"F{hr1}"] = "EMAIL"
    ws[f"H{hr1}"] = "FLD VST"
    ws[f"J{hr1}"] = "SELF CURED"
    ws[f"L{hr1}"] = "TOTAL COUNT OF CASES"
    ws[f"M{hr1}"] = "TOTAL OB"

    for row in ws.iter_rows(min_row=hr1, max_row=hr1, min_col=1, max_col=13):
        for cell in row:
            _apply(cell, s["fill_red"], s["font_hdr"], s["center"], s["border"])

    # header row 2
    for idx, text in enumerate([count_lbl, ob_lbl] * 5, start=2):
        ws.cell(row=hr2, column=idx, value=text)
    for row in ws.iter_rows(min_row=hr2, max_row=hr2, min_col=2, max_col=11):
        for cell in row:
            _apply(cell, s["fill_red"], s["font_hdr"], s["center"], s["border"])
    ws[f"A{hr2}"].border = s["border"]
    ws[f"L{hr2}"].border = s["border"]
    ws[f"M{hr2}"].border = s["border"]

    # data rows
    count_cols = [2, 4, 6, 8, 10, 12]
    ob_cols    = [3, 5, 7, 9, 11, 13]
    data_start = hr2 + 1

    for i, (_, row) in enumerate(summary.iterrows(), start=data_start):
        label    = str(row[("", "Row Labels")])
        is_total = normalize_text(label) == "GRAND TOTAL"
        bfont    = s["font_total"] if is_total else s["font_body"]

        values = [
            label,
            int(row[("ALL NEGATIVE", count_lbl)]), float(row[("ALL NEGATIVE", ob_lbl)]),
            int(row[("CALL OUTS",    count_lbl)]), float(row[("CALL OUTS",    ob_lbl)]),
            int(row[("EMAIL",        count_lbl)]), float(row[("EMAIL",        ob_lbl)]),
            int(row[("FLD VST",      count_lbl)]), float(row[("FLD VST",      ob_lbl)]),
            int(row[("SELF CURED",   count_lbl)]), float(row[("SELF CURED",   ob_lbl)]),
            int(row[("TOTAL", "TOTAL COUNT OF CASES")]),
            float(row[("TOTAL", "TOTAL OB")]),
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=i, column=col_idx, value=value)
            _apply(cell, s["fill_white"], bfont,
                   s["left"] if col_idx == 1 else s["right"], s["border"])
            if col_idx in count_cols:
                cell.number_format = "#,##0"
            elif col_idx in ob_cols:
                cell.number_format = "#,##0"

    return data_start + len(summary)


def _write_rate_block(ws, pct_df: pd.DataFrame, title: str,
                      start_row: int, s: dict) -> int:
    """Write a percentage/rate block. Returns next available row."""
    ws.merge_cells(f"A{start_row}:M{start_row}")
    ws[f"A{start_row}"] = title
    _apply(ws[f"A{start_row}"], s["fill_title"], s["font_title"], s["left"], s["border"])

    hr1 = start_row + 1
    hr2 = start_row + 2

    ws.merge_cells(f"A{hr1}:A{hr2}")
    ws.merge_cells(f"B{hr1}:C{hr1}")
    ws.merge_cells(f"D{hr1}:E{hr1}")
    ws.merge_cells(f"F{hr1}:G{hr1}")
    ws.merge_cells(f"H{hr1}:I{hr1}")
    ws.merge_cells(f"J{hr1}:K{hr1}")
    ws.merge_cells(f"L{hr1}:L{hr2}")
    ws.merge_cells(f"M{hr1}:M{hr2}")

    ws[f"A{hr1}"] = "Row Labels"
    ws[f"B{hr1}"] = "ALL NEGATIVE"
    ws[f"D{hr1}"] = "CALL OUTS"
    ws[f"F{hr1}"] = "EMAIL"
    ws[f"H{hr1}"] = "FLD VST"
    ws[f"J{hr1}"] = "SELF CURED"
    ws[f"L{hr1}"] = "TOTAL COUNT %"
    ws[f"M{hr1}"] = "TOTAL OB %"

    for row in ws.iter_rows(min_row=hr1, max_row=hr1, min_col=1, max_col=13):
        for cell in row:
            _apply(cell, s["fill_red"], s["font_hdr"], s["center"], s["border"])

    for idx, text in enumerate(["COUNT %", "OB %"] * 5, start=2):
        ws.cell(row=hr2, column=idx, value=text)
    for row in ws.iter_rows(min_row=hr2, max_row=hr2, min_col=2, max_col=11):
        for cell in row:
            _apply(cell, s["fill_red"], s["font_hdr"], s["center"], s["border"])
    ws[f"A{hr2}"].border = s["border"]
    ws[f"L{hr2}"].border = s["border"]
    ws[f"M{hr2}"].border = s["border"]

    data_start = hr2 + 1
    for i, (_, row) in enumerate(pct_df.iterrows(), start=data_start):
        label    = str(row[("", "Row Labels")])
        is_total = normalize_text(label) == "GRAND TOTAL"
        bfont    = s["font_total"] if is_total else s["font_body"]

        ws.cell(row=i, column=1, value=label)
        _apply(ws.cell(row=i, column=1), s["fill_white"], bfont, s["left"], s["border"])

        values = [
            row[("ALL NEGATIVE", "COUNT %")], row[("ALL NEGATIVE", "OB %")],
            row[("CALL OUTS",    "COUNT %")], row[("CALL OUTS",    "OB %")],
            row[("EMAIL",        "COUNT %")], row[("EMAIL",        "OB %")],
            row[("FLD VST",      "COUNT %")], row[("FLD VST",      "OB %")],
            row[("SELF CURED",   "COUNT %")], row[("SELF CURED",   "OB %")],
            row[("TOTAL", "TOTAL COUNT %")],  row[("TOTAL", "TOTAL OB %")],
        ]
        for col_idx, value in enumerate(values, start=2):
            cell = ws.cell(row=i, column=col_idx, value=float(value))
            _apply(cell, s["fill_white"], bfont, s["right"], s["border"], "0%")

    return data_start + len(pct_df)


def build_formatted_excel(source_df: pd.DataFrame, month_span_label: str) -> bytes:
    """
    One workbook, one sheet ("Dashboard"), stacked vertically:
      Block 1-2  : Overall summary + rate  (all months)
      Per month  : Month summary + rate
                   Month PTP summary + rate
                   Month CURED summary + rate
    3-row gap between every block.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    for col, width in [
        ("A", 22), ("B", 16), ("C", 16), ("D", 16), ("E", 16),
        ("F", 16), ("G", 16), ("H", 16), ("I", 16), ("J", 16),
        ("K", 16), ("L", 18), ("M", 18),
    ]:
        ws.column_dimensions[col].width = width

    s   = _make_styles()
    cur = 1  # current row pointer

    # ── OVERALL ─────────────────────────────────────────────────────────────
    overall_summary = build_summary_table(source_df)
    overall_pct     = build_percentage_table(overall_summary, is_sub=False)

    cur = _write_summary_block(ws, overall_summary,
        f"OVERALL RESPONSE BY CASES AND BALANCE ({month_span_label})",
        cur, s, is_sub=False)
    cur += 3
    cur = _write_rate_block(ws, overall_pct,
        "OVERALL RESPONSE RATE", cur, s)
    cur += 3

    # ── PER MONTH ────────────────────────────────────────────────────────────
    for month_num in get_detected_months(source_df):
        month_name   = month_num_to_name(month_num)
        month_df     = filter_by_cutoff_month(source_df, month_num)
        ptp_df_raw   = filter_ptp(month_df)
        cured_df_raw = filter_cured(month_df)

        # Month Overall
        m_summary = build_summary_table(month_df)
        m_pct     = build_percentage_table(m_summary, is_sub=False)
        cur = _write_summary_block(ws, m_summary,
            f"{month_name} RESPONSE BY CASES AND BALANCE",
            cur, s, is_sub=False)
        cur += 3
        cur = _write_rate_block(ws, m_pct,
            f"{month_name} OVERALL RESPONSE RATE", cur, s)
        cur += 3

        # Month PTP
        if not ptp_df_raw.empty:
            ptp_summary = build_sub_summary_table(ptp_df_raw)
            ptp_pct     = build_percentage_table(ptp_summary, is_sub=True)
            cur = _write_summary_block(ws, ptp_summary,
                f"{month_name} RESPONSE BY CASES AND BALANCE - PTP",
                cur, s, is_sub=True)
            cur += 3
            cur = _write_rate_block(ws, ptp_pct,
                f"{month_name} OVERALL RESPONSE RATE - PTP", cur, s)
            cur += 3

        # Month CURED
        if not cured_df_raw.empty:
            cured_summary = build_sub_summary_table(cured_df_raw)
            cured_pct     = build_percentage_table(cured_summary, is_sub=True)
            cur = _write_summary_block(ws, cured_summary,
                f"{month_name} RESPONSE BY CASES AND BALANCE - CURED",
                cur, s, is_sub=True)
            cur += 3
            cur = _write_rate_block(ws, cured_pct,
                f"{month_name} OVERALL RESPONSE RATE - CURED", cur, s)
            cur += 3

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
#  Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

st.title("BPI Project Dashboard Automation")
st.caption(
    "Overall tables (all months) + per-month Summary & Rate tables "
    "for Overall, PTP, and CURED — all in one Excel sheet."
)

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
source_df  = workbook[sheet_name]

with st.expander("Preview source data", expanded=False):
    st.dataframe(source_df.head(20), use_container_width=True)

try:
    month_span_label = get_month_span_label(source_df)
    detected_months  = get_detected_months(source_df)
except Exception as exc:
    st.error(f"Unable to parse month data: {exc}")
    st.stop()

st.info(
    f"Detected months: **{', '.join(month_num_to_name(m) for m in detected_months)}** "
    f"({month_span_label})"
)

# ── Overall previews ──────────────────────────────────────────────────────────
try:
    overall_summary = build_summary_table(source_df)
    overall_pct     = build_percentage_table(overall_summary, is_sub=False)
except Exception as exc:
    st.error(f"Unable to build overall tables: {exc}")
    st.stop()

st.subheader(f"Table 1 — Overall Response by Cases and Balance ({month_span_label})")
st.dataframe(to_flat_preview(overall_summary), use_container_width=True, hide_index=True)

st.subheader("Table 2 — Overall Response Rate")
st.dataframe(to_flat_preview(overall_pct), use_container_width=True, hide_index=True)

# ── Per-month previews ────────────────────────────────────────────────────────
for month_num in detected_months:
    month_name   = month_num_to_name(month_num)
    month_df     = filter_by_cutoff_month(source_df, month_num)
    ptp_df_raw   = filter_ptp(month_df)
    cured_df_raw = filter_cured(month_df)

    st.markdown(f"---\n### {month_name}")

    # Overall
    try:
        m_summary = build_summary_table(month_df)
        m_pct     = build_percentage_table(m_summary, is_sub=False)
        st.subheader(f"Table 3 — {month_name} Response by Cases and Balance")
        st.dataframe(to_flat_preview(m_summary), use_container_width=True, hide_index=True)
        st.subheader(f"Table 4 — {month_name} Overall Response Rate")
        st.dataframe(to_flat_preview(m_pct), use_container_width=True, hide_index=True)
    except Exception as exc:
        st.error(f"Error building {month_name} overall tables: {exc}")

    # PTP
    if not ptp_df_raw.empty:
        try:
            ptp_summary = build_sub_summary_table(ptp_df_raw)
            ptp_pct     = build_percentage_table(ptp_summary, is_sub=True)
            st.subheader(f"Table 5 — {month_name} Response by Cases and Balance - PTP")
            st.dataframe(to_flat_preview(ptp_summary), use_container_width=True, hide_index=True)
            st.subheader(f"Table 6 — {month_name} Overall Response Rate - PTP")
            st.dataframe(to_flat_preview(ptp_pct), use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(f"Error building {month_name} PTP tables: {exc}")
    else:
        st.caption(f"No PTP data found for {month_name}.")

    # CURED
    if not cured_df_raw.empty:
        try:
            cured_summary = build_sub_summary_table(cured_df_raw)
            cured_pct     = build_percentage_table(cured_summary, is_sub=True)
            st.subheader(f"Table 7 — {month_name} Response by Cases and Balance - CURED")
            st.dataframe(to_flat_preview(cured_summary), use_container_width=True, hide_index=True)
            st.subheader(f"Table 8 — {month_name} Overall Response Rate - CURED")
            st.dataframe(to_flat_preview(cured_pct), use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(f"Error building {month_name} CURED tables: {exc}")
    else:
        st.caption(f"No CURED data found for {month_name}.")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Download All Tables")

try:
    excel_bytes = build_formatted_excel(source_df, month_span_label)
    st.download_button(
        label="⬇️ Download Full Formatted Excel (All Tables — One Sheet)",
        data=excel_bytes,
        file_name="BPI_Dashboard_All_Tables.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception as exc:
    st.error(f"Unable to generate Excel file: {exc}")
