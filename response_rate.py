import io
import re
import calendar
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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

STATUS_ALIASES = {
    "CALL OUT": "CALL OUTS",
    "CALL-OUTS": "CALL OUTS",
    "FIELD VISIT": "FLD VST",
    "FLD VISIT": "FLD VST",
    "SELF-CURED": "SELF CURED",
}

# ── Column layout constants (1-based Excel column numbers) ────────────────────
#  Overall  : cols  1–13  (A–M)
#  gap      : col  14     (N)
#  PTP      : cols 15–27  (O–AA)
#  gap      : col  28     (AB)
#  CURED    : cols 29–41  (AC–AO)

OVERALL_START = 1   # col A
PTP_START     = 15  # col O
CURED_START   = 29  # col AC
BLOCK_WIDTH   = 13  # 13 columns per table (label + total×2 + 5×2)


# ─────────────────────────────────────────────────────────────────────────────
#  Data helpers
# ─────────────────────────────────────────────────────────────────────────────

def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip().upper())


def normalize_status(value: object) -> str:
    text = normalize_text(value)
    return STATUS_ALIASES.get(text, text)


def resolve_column(
    df: pd.DataFrame,
    *,
    by_name: Optional[str] = None,
    by_excel_index: Optional[int] = None
) -> str:
    if by_name:
        target = normalize_text(by_name)
        for col in df.columns:
            if normalize_text(col) == target:
                return col
    if by_excel_index is not None:
        idx = by_excel_index - 1
        if 0 <= idx < len(df.columns):
            return df.columns[idx]
    raise KeyError(f"Column not found: {by_name or f'Excel #{by_excel_index}'}")


@st.cache_data(show_spinner=False)
def load_workbook_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    return {
        s: pd.read_excel(io.BytesIO(file_bytes), sheet_name=s)
        for s in excel_file.sheet_names
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
    for name, num in MONTH_NAME_TO_NUM.items():
        if re.search(rf"\b{name}\b", text):
            return num
    m = re.fullmatch(r"(\d{1,2})(?:\.0+)?", text)
    if m:
        n = int(m.group(1))
        if 1 <= n <= 12:
            return n
    return None


def month_num_to_name(n: int) -> str:
    return calendar.month_name[n].upper()


def get_month_span_label(df: pd.DataFrame) -> str:
    cutoff_col = resolve_column(df, by_name="CUT OFF MONTH", by_excel_index=8)
    months = sorted(set(
        df[cutoff_col].map(extract_month_number).dropna().astype(int).tolist()
    ))
    if not months:
        return "ALL MONTHS"
    if len(months) == 1:
        return month_num_to_name(months[0])
    return f"{month_num_to_name(min(months))} TO {month_num_to_name(max(months))}"


def get_detected_months(df: pd.DataFrame) -> List[int]:
    cutoff_col = resolve_column(df, by_name="CUT OFF MONTH", by_excel_index=8)
    return sorted(set(
        df[cutoff_col].map(extract_month_number).dropna().astype(int).tolist()
    ))


def filter_by_cutoff_month(df: pd.DataFrame, target_month: int) -> pd.DataFrame:
    cutoff_col = resolve_column(df, by_name="CUT OFF MONTH", by_excel_index=8)
    working = df.copy()
    working["_m"] = working[cutoff_col].map(extract_month_number)
    return working[working["_m"] == target_month].copy()


def filter_ptp(df: pd.DataFrame) -> pd.DataFrame:
    col = resolve_column(df, by_name="REMARKS (PTP/NO PTP)", by_excel_index=61)
    return df[df[col].apply(normalize_text) == "PTP"].copy()


def filter_cured(df: pd.DataFrame) -> pd.DataFrame:
    col = resolve_column(df, by_name="FINAL STATUS", by_excel_index=11)
    return df[df[col].apply(normalize_text) == "CURED"].copy()


def _build_table(df: pd.DataFrame, is_sub: bool) -> pd.DataFrame:
    """
    Core aggregation. is_sub=True uses NO. OF CASES / OB labels (PTP/CURED).
    """
    cycle_col   = resolve_column(df, by_name="CYCLE",                    by_excel_index=2)
    ob_col      = resolve_column(df, by_name="OB",                       by_excel_index=5)
    contact_col = resolve_column(df, by_name="CONTACT SOURCE (OVERALL)", by_excel_index=55)

    count_lbl = "NO. OF CASES" if is_sub else "COUNT OF ACCOUNT CYCLE"
    ob_lbl    = "OB"           if is_sub else "OB PER CYCLE"

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
        cdf = working[working[cycle_col] == cycle]
        row: Dict = {
            ("", "Cycle"): cycle,
            ("TOTAL", "TOTAL COUNT OF CASES"): int(len(cdf)),
            ("TOTAL", "TOTAL OB"): float(cdf[ob_col].sum()),
        }
        for status in TABLE_STATUSES:
            sdf = cdf[cdf[contact_col] == status]
            row[(status, count_lbl)] = int(len(sdf))
            row[(status, ob_lbl)]    = float(sdf[ob_col].sum())
        rows.append(row)

    tbl = pd.DataFrame(rows)
    tbl.columns = pd.MultiIndex.from_tuples(tbl.columns)

    gt: Dict = {
        ("", "Cycle"): "Grand Total",
        ("TOTAL", "TOTAL COUNT OF CASES"): tbl[("TOTAL", "TOTAL COUNT OF CASES")].sum(),
        ("TOTAL", "TOTAL OB"): tbl[("TOTAL", "TOTAL OB")].sum(),
    }
    for col in tbl.columns:
        if col not in [("", "Cycle"), ("TOTAL", "TOTAL COUNT OF CASES"), ("TOTAL", "TOTAL OB")]:
            gt[col] = tbl[col].sum()

    return pd.concat([tbl, pd.DataFrame([gt])], ignore_index=True)


def build_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    return _build_table(df, is_sub=False)


def build_sub_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    return _build_table(df, is_sub=True)


def build_percentage_table(summary: pd.DataFrame, is_sub: bool = False) -> pd.DataFrame:
    count_key = "NO. OF CASES" if is_sub else "COUNT OF ACCOUNT CYCLE"
    ob_key    = "OB"           if is_sub else "OB PER CYCLE"

    detail = summary[summary[("", "Cycle")].astype(str).str.upper() != "GRAND TOTAL"].copy()
    rows: List[Dict] = []

    for _, row in detail.iterrows():
        tc = row[("TOTAL", "TOTAL COUNT OF CASES")]
        to = row[("TOTAL", "TOTAL OB")]

        pr: Dict = {
            ("", "Cycle"): row[("", "Cycle")],
            ("TOTAL", "TOTAL COUNT %"): 1 if tc else 0,
            ("TOTAL", "TOTAL OB %"): 1 if to else 0,
        }

        for s in TABLE_STATUSES:
            pr[(s, "COUNT %")] = (row[(s, count_key)] / tc) if tc else 0
            pr[(s, "OB %")]    = (row[(s, ob_key)]    / to) if to else 0

        rows.append(pr)

    pct = pd.DataFrame(rows)
    pct.columns = pd.MultiIndex.from_tuples(pct.columns)

    gt_row = summary[summary[("", "Cycle")].astype(str).str.upper() == "GRAND TOTAL"].iloc[0]
    gtc = gt_row[("TOTAL", "TOTAL COUNT OF CASES")]
    gto = gt_row[("TOTAL", "TOTAL OB")]

    gtr: Dict = {
        ("", "Cycle"): "Grand Total",
        ("TOTAL", "TOTAL COUNT %"): 1 if gtc else 0,
        ("TOTAL", "TOTAL OB %"): 1 if gto else 0,
    }

    for s in TABLE_STATUSES:
        gtr[(s, "COUNT %")] = (gt_row[(s, count_key)] / gtc) if gtc else 0
        gtr[(s, "OB %")]    = (gt_row[(s, ob_key)]    / gto) if gto else 0

    return pd.concat([pct, pd.DataFrame([gtr])], ignore_index=True)


def to_flat_preview(df: pd.DataFrame) -> pd.DataFrame:
    preview = df.copy()
    preview.columns = [
        sub if top in ("", "TOTAL") else f"{top} - {sub}"
        for top, sub in preview.columns
    ]
    return preview


# ─────────────────────────────────────────────────────────────────────────────
#  Excel formatting
# ─────────────────────────────────────────────────────────────────────────────

def _make_styles() -> dict:
    thin = Side(style="thin", color="000000")
    return dict(
        fill_title=PatternFill("solid", fgColor="D9EAD3"),
        fill_red  =PatternFill("solid", fgColor="C00000"),
        fill_white=PatternFill("solid", fgColor="FFFFFF"),
        border    =Border(left=thin, right=thin, top=thin, bottom=thin),
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


def _col(n: int) -> str:
    """1-based column number → Excel letter (e.g. 1→A, 15→O, 29→AC)."""
    return get_column_letter(n)


def _write_summary_side_by_side(
    ws,
    overall_summary: Optional[pd.DataFrame],
    ptp_summary: Optional[pd.DataFrame],
    cured_summary: Optional[pd.DataFrame],
    overall_title: str,
    ptp_title: str,
    cured_title: str,
    start_row: int,
    s: dict,
) -> int:
    """
    Write three summary tables side-by-side on the same rows.
    Overall → cols 1-13, PTP → cols 15-27, CURED → cols 29-41.
    Returns next available row after all three tables.
    """
    tables = [
        (overall_summary, overall_title, OVERALL_START, False),
        (ptp_summary,     ptp_title,     PTP_START,     True),
        (cured_summary,   cured_title,   CURED_START,   True),
    ]

    max_data_rows = 0
    for tbl, title, col_start, is_sub in tables:
        if tbl is None:
            continue

        count_lbl = "NO. OF CASES" if is_sub else "COUNT OF ACCOUNT CYCLE"
        ob_lbl    = "OB"           if is_sub else "OB PER CYCLE"
        c         = col_start

        # title
        title_start = _col(c)
        title_end   = _col(c + BLOCK_WIDTH - 1)
        ws.merge_cells(f"{title_start}{start_row}:{title_end}{start_row}")
        ws[f"{title_start}{start_row}"] = title
        _apply(
            ws[f"{title_start}{start_row}"],
            s["fill_title"], s["font_title"], s["left"], s["border"]
        )

        hr1 = start_row + 1
        hr2 = start_row + 2

        # header row 1 merges
        ws.merge_cells(f"{_col(c)}{hr1}:{_col(c)}{hr2}")
        ws.merge_cells(f"{_col(c+1)}{hr1}:{_col(c+1)}{hr2}")
        ws.merge_cells(f"{_col(c+2)}{hr1}:{_col(c+2)}{hr2}")
        ws.merge_cells(f"{_col(c+3)}{hr1}:{_col(c+4)}{hr1}")
        ws.merge_cells(f"{_col(c+5)}{hr1}:{_col(c+6)}{hr1}")
        ws.merge_cells(f"{_col(c+7)}{hr1}:{_col(c+8)}{hr1}")
        ws.merge_cells(f"{_col(c+9)}{hr1}:{_col(c+10)}{hr1}")
        ws.merge_cells(f"{_col(c+11)}{hr1}:{_col(c+12)}{hr1}")

        ws.cell(hr1, c,    "Cycle")
        ws.cell(hr1, c+1,  "TOTAL COUNT OF CASES")
        ws.cell(hr1, c+2,  "TOTAL OB")
        ws.cell(hr1, c+3,  "ALL NEGATIVE")
        ws.cell(hr1, c+5,  "CALL OUTS")
        ws.cell(hr1, c+7,  "EMAIL")
        ws.cell(hr1, c+9,  "FLD VST")
        ws.cell(hr1, c+11, "SELF CURED")

        for cell in ws.iter_rows(
            min_row=hr1, max_row=hr1,
            min_col=c, max_col=c+12, values_only=False
        ):
            for cl in cell:
                _apply(cl, s["fill_red"], s["font_hdr"], s["center"], s["border"])

        # header row 2 sub-headers (statuses only)
        for offset, text in enumerate([count_lbl, ob_lbl] * 5, start=3):
            ws.cell(hr2, c + offset, text)

        for cell in ws.iter_rows(
            min_row=hr2, max_row=hr2,
            min_col=c+3, max_col=c+12, values_only=False
        ):
            for cl in cell:
                _apply(cl, s["fill_red"], s["font_hdr"], s["center"], s["border"])

        ws.cell(hr2, c).border    = s["border"]
        ws.cell(hr2, c+1).border  = s["border"]
        ws.cell(hr2, c+2).border  = s["border"]

        # data rows
        count_offsets = [1, 3, 5, 7, 9, 11]
        ob_offsets    = [2, 4, 6, 8, 10, 12]
        data_start    = hr2 + 1

        for i, (_, row) in enumerate(tbl.iterrows(), start=data_start):
            label    = str(row[("", "Cycle")])
            is_total = normalize_text(label) == "GRAND TOTAL"
            bfont    = s["font_total"] if is_total else s["font_body"]

            values = [
                label,
                int(row[("TOTAL", "TOTAL COUNT OF CASES")]),
                float(row[("TOTAL", "TOTAL OB")]),
                int(row[("ALL NEGATIVE", count_lbl)]), float(row[("ALL NEGATIVE", ob_lbl)]),
                int(row[("CALL OUTS",    count_lbl)]), float(row[("CALL OUTS",    ob_lbl)]),
                int(row[("EMAIL",        count_lbl)]), float(row[("EMAIL",        ob_lbl)]),
                int(row[("FLD VST",      count_lbl)]), float(row[("FLD VST",      ob_lbl)]),
                int(row[("SELF CURED",   count_lbl)]), float(row[("SELF CURED",   ob_lbl)]),
            ]

            for offset, value in enumerate(values):
                cl = ws.cell(i, c + offset, value)
                _apply(
                    cl, s["fill_white"], bfont,
                    s["left"] if offset == 0 else s["right"],
                    s["border"]
                )
                if offset in count_offsets:
                    cl.number_format = "#,##0"
                elif offset in ob_offsets:
                    cl.number_format = "#,##0"

        max_data_rows = max(max_data_rows, len(tbl))

    return start_row + 3 + max_data_rows


def _write_rate_side_by_side(
    ws,
    overall_pct: Optional[pd.DataFrame],
    ptp_pct: Optional[pd.DataFrame],
    cured_pct: Optional[pd.DataFrame],
    overall_title: str,
    ptp_title: str,
    cured_title: str,
    start_row: int,
    s: dict,
) -> int:
    """
    Write three rate tables side-by-side on the same rows.
    Returns next available row.
    """
    tables = [
        (overall_pct, overall_title, OVERALL_START),
        (ptp_pct,     ptp_title,     PTP_START),
        (cured_pct,   cured_title,   CURED_START),
    ]

    max_data_rows = 0
    for tbl, title, col_start in tables:
        if tbl is None:
            continue

        c = col_start

        # title
        ws.merge_cells(f"{_col(c)}{start_row}:{_col(c+12)}{start_row}")
        ws[f"{_col(c)}{start_row}"] = title
        _apply(
            ws[f"{_col(c)}{start_row}"],
            s["fill_title"], s["font_title"], s["left"], s["border"]
        )

        hr1 = start_row + 1
        hr2 = start_row + 2

        ws.merge_cells(f"{_col(c)}{hr1}:{_col(c)}{hr2}")
        ws.merge_cells(f"{_col(c+1)}{hr1}:{_col(c+1)}{hr2}")
        ws.merge_cells(f"{_col(c+2)}{hr1}:{_col(c+2)}{hr2}")
        ws.merge_cells(f"{_col(c+3)}{hr1}:{_col(c+4)}{hr1}")
        ws.merge_cells(f"{_col(c+5)}{hr1}:{_col(c+6)}{hr1}")
        ws.merge_cells(f"{_col(c+7)}{hr1}:{_col(c+8)}{hr1}")
        ws.merge_cells(f"{_col(c+9)}{hr1}:{_col(c+10)}{hr1}")
        ws.merge_cells(f"{_col(c+11)}{hr1}:{_col(c+12)}{hr1}")

        ws.cell(hr1, c,    "Cycle")
        ws.cell(hr1, c+1,  "TOTAL COUNT %")
        ws.cell(hr1, c+2,  "TOTAL OB %")
        ws.cell(hr1, c+3,  "ALL NEGATIVE")
        ws.cell(hr1, c+5,  "CALL OUTS")
        ws.cell(hr1, c+7,  "EMAIL")
        ws.cell(hr1, c+9,  "FLD VST")
        ws.cell(hr1, c+11, "SELF CURED")

        for cell in ws.iter_rows(
            min_row=hr1, max_row=hr1,
            min_col=c, max_col=c+12, values_only=False
        ):
            for cl in cell:
                _apply(cl, s["fill_red"], s["font_hdr"], s["center"], s["border"])

        for offset, text in enumerate(["COUNT %", "OB %"] * 5, start=3):
            ws.cell(hr2, c + offset, text)

        for cell in ws.iter_rows(
            min_row=hr2, max_row=hr2,
            min_col=c+3, max_col=c+12, values_only=False
        ):
            for cl in cell:
                _apply(cl, s["fill_red"], s["font_hdr"], s["center"], s["border"])

        ws.cell(hr2, c).border    = s["border"]
        ws.cell(hr2, c+1).border  = s["border"]
        ws.cell(hr2, c+2).border  = s["border"]

        data_start = hr2 + 1
        for i, (_, row) in enumerate(tbl.iterrows(), start=data_start):
            label    = str(row[("", "Cycle")])
            is_total = normalize_text(label) == "GRAND TOTAL"
            bfont    = s["font_total"] if is_total else s["font_body"]

            ws.cell(i, c, label)
            _apply(ws.cell(i, c), s["fill_white"], bfont, s["left"], s["border"])

            values = [
                row[("TOTAL", "TOTAL COUNT %")],  row[("TOTAL", "TOTAL OB %")],
                row[("ALL NEGATIVE", "COUNT %")], row[("ALL NEGATIVE", "OB %")],
                row[("CALL OUTS",    "COUNT %")], row[("CALL OUTS",    "OB %")],
                row[("EMAIL",        "COUNT %")], row[("EMAIL",        "OB %")],
                row[("FLD VST",      "COUNT %")], row[("FLD VST",      "OB %")],
                row[("SELF CURED",   "COUNT %")], row[("SELF CURED",   "OB %")],
            ]

            for offset, value in enumerate(values, start=1):
                cl = ws.cell(i, c + offset, float(value))
                _apply(cl, s["fill_white"], bfont, s["right"], s["border"], "0%")

        max_data_rows = max(max_data_rows, len(tbl))

    return start_row + 3 + max_data_rows


def build_formatted_excel(source_df: pd.DataFrame, month_span_label: str) -> bytes:
    """
    One workbook, one sheet ("Dashboard"):
      Row group 1: Overall summary (3 side-by-side) + Overall rate (3 side-by-side)
      Per month  : Month summary (3 side-by-side) + Month rate (3 side-by-side)

    Layout per row group:
      Col  1-13  = Overall / Month Overall
      Col 15-27  = PTP
      Col 29-41  = CURED
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    # Set column widths for all 41 used columns
    for c in range(1, 42):
        letter = _col(c)
        # gap columns (N=14, AB=28)
        if c in (14, 28):
            ws.column_dimensions[letter].width = 2
        # label columns (A=1, O=15, AC=29)
        elif c in (1, 15, 29):
            ws.column_dimensions[letter].width = 22
        # total columns now placed right after Cycle
        elif c in (2, 3, 16, 17, 30, 31):
            ws.column_dimensions[letter].width = 18
        else:
            ws.column_dimensions[letter].width = 16

    s = _make_styles()
    cur = 1

    # ── OVERALL (all months) ──────────────────────────────────────────────────
    overall_summary = build_summary_table(source_df)
    overall_pct     = build_percentage_table(overall_summary, is_sub=False)
    ptp_all         = filter_ptp(source_df)
    cured_all       = filter_cured(source_df)
    ptp_all_summary = build_sub_summary_table(ptp_all)   if not ptp_all.empty   else None
    cured_all_sum   = build_sub_summary_table(cured_all) if not cured_all.empty else None
    ptp_all_pct     = build_percentage_table(ptp_all_summary, is_sub=True) if ptp_all_summary is not None else None
    cured_all_pct   = build_percentage_table(cured_all_sum, is_sub=True) if cured_all_sum is not None else None

    cur = _write_summary_side_by_side(
        ws,
        overall_summary, ptp_all_summary, cured_all_sum,
        f"OVERALL RESPONSE BY CASES AND BALANCE ({month_span_label})",
        f"OVERALL RESPONSE BY CASES AND BALANCE - PTP ({month_span_label})",
        f"OVERALL RESPONSE BY CASES AND BALANCE - CURED ({month_span_label})",
        cur, s,
    )
    cur += 2

    cur = _write_rate_side_by_side(
        ws,
        overall_pct, ptp_all_pct, cured_all_pct,
        "OVERALL RESPONSE RATE",
        "OVERALL RESPONSE RATE - PTP",
        "OVERALL RESPONSE RATE - CURED",
        cur, s,
    )
    cur += 3

    # ── PER MONTH ─────────────────────────────────────────────────────────────
    for month_num in get_detected_months(source_df):
        mname         = month_num_to_name(month_num)
        month_df      = filter_by_cutoff_month(source_df, month_num)
        ptp_df        = filter_ptp(month_df)
        cured_df      = filter_cured(month_df)

        m_summary     = build_summary_table(month_df)
        m_pct         = build_percentage_table(m_summary, is_sub=False)
        ptp_summary   = build_sub_summary_table(ptp_df)   if not ptp_df.empty   else None
        cured_summary = build_sub_summary_table(cured_df) if not cured_df.empty else None
        ptp_pct       = build_percentage_table(ptp_summary, is_sub=True) if ptp_summary is not None else None
        cured_pct     = build_percentage_table(cured_summary, is_sub=True) if cured_summary is not None else None

        cur = _write_summary_side_by_side(
            ws,
            m_summary, ptp_summary, cured_summary,
            f"{mname} RESPONSE BY CASES AND BALANCE",
            f"{mname} RESPONSE BY CASES AND BALANCE - PTP",
            f"{mname} RESPONSE BY CASES AND BALANCE - CURED",
            cur, s,
        )
        cur += 2

        cur = _write_rate_side_by_side(
            ws,
            m_pct, ptp_pct, cured_pct,
            f"{mname} OVERALL RESPONSE RATE",
            f"{mname} OVERALL RESPONSE RATE - PTP",
            f"{mname} OVERALL RESPONSE RATE - CURED",
            cur, s,
        )
        cur += 3

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
#  Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

def response_rate():
    st.title("BPI Reponse Rate Dashboard")

    uploaded_file = st.file_uploader(
        "PLEASE UPLOAD YOUR ENDORSEMENT WORKLIST(.xlsx)",
        type=["xlsx"]
    )

    if uploaded_file is None:
        st.info("Upload the workbook first.")
        st.stop()

    file_bytes = uploaded_file.getvalue()

    try:
        workbook = load_workbook_sheets(file_bytes)
    except Exception as exc:
        st.error(f"Unable to read the workbook: {exc}")
        st.stop()

    sheet_name = st.selectbox("Select source sheet", options=list(workbook.keys()), index=0)
    source_df = workbook[sheet_name]

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

    # ── Overall previews ──────────────────────────────────────────────────────
    try:
        overall_summary = build_summary_table(source_df)
        overall_pct     = build_percentage_table(overall_summary, is_sub=False)
        ptp_all         = filter_ptp(source_df)
        cured_all       = filter_cured(source_df)
    except Exception as exc:
        st.error(f"Unable to build overall tables: {exc}")
        st.stop()

    st.subheader(f"Overall — Response by Cases and Balance ({month_span_label})")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.caption("Overall")
        st.dataframe(to_flat_preview(overall_summary), use_container_width=True, hide_index=True)
    with col2:
        st.caption("PTP")
        if not ptp_all.empty:
            st.dataframe(
                to_flat_preview(build_sub_summary_table(ptp_all)),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.caption("No PTP data.")
    with col3:
        st.caption("CURED")
        if not cured_all.empty:
            st.dataframe(
                to_flat_preview(build_sub_summary_table(cured_all)),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.caption("No CURED data.")

    st.subheader("Overall — Response Rate")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.caption("Overall")
        st.dataframe(to_flat_preview(overall_pct), use_container_width=True, hide_index=True)
    with col2:
        st.caption("PTP")
        if not ptp_all.empty:
            p = build_sub_summary_table(ptp_all)
            st.dataframe(
                to_flat_preview(build_percentage_table(p, is_sub=True)),
                use_container_width=True,
                hide_index=True
            )
    with col3:
        st.caption("CURED")
        if not cured_all.empty:
            p = build_sub_summary_table(cured_all)
            st.dataframe(
                to_flat_preview(build_percentage_table(p, is_sub=True)),
                use_container_width=True,
                hide_index=True
            )

    # ── Per-month previews ────────────────────────────────────────────────────
    for month_num in detected_months:
        mname    = month_num_to_name(month_num)
        month_df = filter_by_cutoff_month(source_df, month_num)
        ptp_df   = filter_ptp(month_df)
        cured_df = filter_cured(month_df)

        st.markdown(f"---\n### {mname}")

        try:
            m_summary = build_summary_table(month_df)
            m_pct     = build_percentage_table(m_summary, is_sub=False)

            st.subheader(f"{mname} — Response by Cases and Balance")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.caption("Overall")
                st.dataframe(to_flat_preview(m_summary), use_container_width=True, hide_index=True)
            with c2:
                st.caption("PTP")
                if not ptp_df.empty:
                    st.dataframe(
                        to_flat_preview(build_sub_summary_table(ptp_df)),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.caption("No PTP data.")
            with c3:
                st.caption("CURED")
                if not cured_df.empty:
                    st.dataframe(
                        to_flat_preview(build_sub_summary_table(cured_df)),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.caption("No CURED data.")

            st.subheader(f"{mname} — Response Rate")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.caption("Overall")
                st.dataframe(to_flat_preview(m_pct), use_container_width=True, hide_index=True)
            with c2:
                st.caption("PTP")
                if not ptp_df.empty:
                    ps = build_sub_summary_table(ptp_df)
                    st.dataframe(
                        to_flat_preview(build_percentage_table(ps, is_sub=True)),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.caption("No PTP data.")
            with c3:
                st.caption("CURED")
                if not cured_df.empty:
                    cs = build_sub_summary_table(cured_df)
                    st.dataframe(
                        to_flat_preview(build_percentage_table(cs, is_sub=True)),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.caption("No CURED data.")

        except Exception as exc:
            st.error(f"Error building {mname} tables: {exc}")

    # ── Download ──────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Download All Tables")

    try:
        excel_bytes = build_formatted_excel(source_df, month_span_label)
        st.download_button(
            label="⬇️ Download Response Rate",
            data=excel_bytes,
            file_name="BPI_Response_Rate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Unable to generate Excel file: {exc}")
