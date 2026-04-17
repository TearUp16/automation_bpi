import io
import re
from typing import Dict, Optional, Set, Tuple

import pandas as pd
import streamlit as st


# -----------------------------------------------------------------------------
# CONFIG: LOCAL REFERENCE FILE PATH
# -----------------------------------------------------------------------------

REFERENCE_FILE_PATH = "JARGONS FOR FINAL RFD.xlsx"
REFERENCE_SHEET_NAME = 0
CUSTOMER_NUMBER_LENGTH = 16


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"[^A-Z0-9]", "", text)
    return text


@st.cache_data(show_spinner=False)
def load_reference_file() -> pd.DataFrame:
    return pd.read_excel(REFERENCE_FILE_PATH, sheet_name=REFERENCE_SHEET_NAME)


@st.cache_data(show_spinner=False)
def load_target_file(uploaded_file) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(uploaded_file, sheet_name=None)


def find_column(df: pd.DataFrame, candidates: Tuple[str, ...]) -> Optional[str]:
    normalized_headers = {normalize_code(col): col for col in df.columns}

    for candidate in candidates:
        key = normalize_code(candidate)
        if key in normalized_headers:
            return normalized_headers[key]

    return None


def build_jargon_map(jargon_df: pd.DataFrame) -> Dict[str, str]:
    jargon_col = find_column(jargon_df, ("JARGONS",))
    final_col = find_column(jargon_df, ("FINAL : RFD", "FINAL RFD"))

    if not jargon_col or not final_col:
        raise ValueError("Reference file must contain JARGONS and FINAL : RFD columns.")

    cleaned = jargon_df[[jargon_col, final_col]].copy()
    cleaned["_code"] = cleaned[jargon_col].apply(normalize_code)
    cleaned["_final"] = cleaned[final_col].fillna("").astype(str).str.strip()

    cleaned = cleaned[(cleaned["_code"] != "") & (cleaned["_final"] != "")]
    cleaned = cleaned.drop_duplicates(subset=["_code"])

    return dict(zip(cleaned["_code"], cleaned["_final"]))


def extract_rfd_code(remarks: object, valid_codes: Set[str]) -> str:
    if pd.isna(remarks):
        return ""

    text = str(remarks).upper()

    direct_match = re.search(r"RFD\s*:\s*\*?\s*([A-Z0-9]+)", text)
    if direct_match:
        code = normalize_code(direct_match.group(1))
        if code in valid_codes:
            return code

    for code in sorted(valid_codes, key=len, reverse=True):
        pattern = r"(?<![A-Z0-9])\*?" + re.escape(code) + r"(?![A-Z0-9])"
        if re.search(pattern, text):
            return code

    return ""


def format_customer_number(value: object, default_length: int = CUSTOMER_NUMBER_LENGTH) -> str:
    """
    Preserve customer numbers as text and remove scientific notation.
    """
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text == "":
        return ""

    try:
        if re.search(r"[Ee]", text) or re.fullmatch(r"[+-]?\d+(\.0+)?", text):
            digits = str(int(float(text)))
        else:
            digits = re.sub(r"\D", "", text)
    except Exception:
        digits = re.sub(r"\D", "", text)

    if digits == "":
        return ""

    return digits.zfill(default_length)


def apply_rfd_mapping(target_df: pd.DataFrame, jargon_map: Dict[str, str]) -> pd.DataFrame:
    ptp_col = find_column(target_df, ("PTP REMARKS",))

    if not ptp_col:
        raise ValueError("Target file must contain PTP REMARKS column.")

    result = target_df.copy()
    valid_codes = set(jargon_map.keys())

    customer_number_col = find_column(result, ("CUSTOMER NUMBER", "CUSTOMERNUMBER"))
    if customer_number_col:
        result[customer_number_col] = result[customer_number_col].apply(format_customer_number)

    result["DETECTED_RFD_CODE"] = result[ptp_col].apply(
        lambda x: extract_rfd_code(x, valid_codes)
    )
    result["FINAL : RFD"] = result["DETECTED_RFD_CODE"].map(jargon_map).fillna("")

    return result


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

        worksheet = writer.sheets["Sheet1"]
        customer_number_col = find_column(df, ("CUSTOMER NUMBER", "CUSTOMERNUMBER"))
        if customer_number_col:
            col_idx = df.columns.get_loc(customer_number_col) + 1
            for row in range(2, len(df) + 2):
                worksheet.cell(row=row, column=col_idx).number_format = "@"

    output.seek(0)
    return output.getvalue()


def render_rfd_mapper() -> None:
    st.markdown('<div class="app-title">RFD Mapper</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subtitle">Upload your Excel file and automatically map PTP REMARKS into FINAL : RFD.</div>',
        unsafe_allow_html=True,
    )

    try:
        reference_df = load_reference_file()
        jargon_map = build_jargon_map(reference_df)
        st.success("Reference file loaded successfully.")
    except Exception as e:
        st.error("Failed to load reference file. Make sure it's in the same folder.")
        st.exception(e)
        return

    st.subheader("Upload file")

    uploaded_file = st.file_uploader(
        "CUSTOMER NUMER | CUSTOMER NAME | PTP REMARKS",
        type=["xlsx", "xls"],
        key="rfd_upload_file",
    )

    if uploaded_file:
        try:
            sheets = load_target_file(uploaded_file)
            sheet_name = st.selectbox(
                "Select sheet",
                list(sheets.keys()),
                key="rfd_sheet_name",
            )
            df = sheets[sheet_name]

            st.caption("Preview")
            st.dataframe(df.head(10), use_container_width=True)

            if st.button("Process", use_container_width=True, key="rfd_process_button"):
                processed_df = apply_rfd_mapping(df, jargon_map)

                st.success("Processing complete")
                st.dataframe(processed_df.head(50), use_container_width=True)

                output_bytes = to_excel_bytes(processed_df)
                unmatched_df = processed_df[
                    processed_df["DETECTED_RFD_CODE"].astype(str).str.strip() == ""
                ].copy()

                st.download_button(
                    "Download processed file",
                    data=output_bytes,
                    file_name="processed_with_final_rfd.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="rfd_download_processed",
                )

                if not unmatched_df.empty:
                    unmatched_bytes = to_excel_bytes(unmatched_df)
                    st.download_button(
                        "Download accounts with no detected RFD",
                        data=unmatched_bytes,
                        file_name="accounts_with_no_detected_rfd.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="rfd_download_unmatched",
                    )
                else:
                    st.info("All rows have a detected RFD code.")

        except Exception as e:
            st.error("Error processing file")
            st.exception(e)