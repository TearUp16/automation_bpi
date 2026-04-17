import io
import os
import re
from io import BytesIO

import pandas as pd
import streamlit as st


def render_page_header(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="app-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="app-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def clean_sheet_name(name: str, used_names: set[str]) -> str:
    cleaned = re.sub(r'[\\/*?:\[\]]', "_", str(name).strip())

    if not cleaned:
        cleaned = "BLANK_STATUS"

    cleaned = cleaned[:31]
    original = cleaned
    counter = 1

    while cleaned in used_names:
        suffix = f"_{counter}"
        cleaned = original[:31 - len(suffix)] + suffix
        counter += 1

    used_names.add(cleaned)
    return cleaned


def process_drr_file(file):
    df = pd.read_csv(file, dtype=str)
    df.columns = df.columns.str.strip()

    if "Account No." in df.columns:
        def convert_value(x):
            try:
                value = str(x).strip()
                if "E+" in value.upper():
                    return "{:.0f}".format(float(value))
                return x
            except Exception:
                return x

        df["Account No."] = df["Account No."].apply(convert_value)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    reference_path = os.path.join(base_dir, "Reference.xlsx")

    if not os.path.exists(reference_path):
        st.error("❌ Reference.xlsx not found")
        return None

    ref_df = pd.read_excel(reference_path, dtype=str)
    ref_df.columns = ref_df.columns.str.strip()

    required_ref_cols = ["Status", "Final Status", "Cycle", "Date", "Cut off"]
    missing_ref_cols = [col for col in required_ref_cols if col not in ref_df.columns]
    if missing_ref_cols:
        st.error(f"❌ Missing columns in Reference.xlsx: {', '.join(missing_ref_cols)}")
        return None

    required_csv_cols = ["Status", "Card No.", "Date"]
    missing_csv_cols = [col for col in required_csv_cols if col not in df.columns]
    if missing_csv_cols:
        st.error(f"❌ Missing columns in uploaded CSV: {', '.join(missing_csv_cols)}")
        return None

    status_map = dict(zip(
        ref_df["Status"].fillna("").str.strip().str.upper(),
        ref_df["Final Status"].fillna("0")
    ))

    df["STATUS"] = df["Status"].fillna("").str.strip().str.upper().map(status_map)
    df["STATUS"] = df["STATUS"].replace("UNKNOWN", "").fillna("0")

    ref_df["lookup_key"] = (
        ref_df["Cycle"].fillna("").str.strip().str.upper() + "|" +
        pd.to_datetime(ref_df["Date"], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("")
    )

    cutoff_map = dict(zip(ref_df["lookup_key"], ref_df["Cut off"].fillna("")))

    df["CYCLE"] = "CYCLE " + df["Card No."].fillna("").str[:2]

    df["Month Extracted"] = pd.to_datetime(
        df["Date"], errors="coerce", dayfirst=True
    ).dt.strftime("%b").fillna("")

    df["lookup_key"] = (
        df["CYCLE"].fillna("").str.strip().str.upper() + "|" +
        pd.to_datetime(df["Date"], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("")
    )

    df["Month Cut Off"] = df["lookup_key"].map(cutoff_map).fillna("")
    df.drop(columns=["lookup_key"], inplace=True)

    desired_columns = [
        "STATUS",
        "CYCLE",
        "Month Cut Off",
        "Month Extracted",
        "S.No",
        "Date",
        "Time",
        "Debtor",
        "Account No.",
        "Card No.",
        "Status",
        "Remark",
        "Remark By",
        "Client",
        "Product Type",
        "PTP Amount",
        "Next Call",
        "PTP Date",
        "Claim Paid Amount",
        "Claim Paid Date",
        "Dialed Number",
        "Balance",
        "Call Duration"
    ]

    df = df[[col for col in desired_columns if col in df.columns]]
    return df


def process_positive_status(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str)
    df.columns = [col.strip() for col in df.columns]

    if "STATUS" not in df.columns:
        raise ValueError("Missing STATUS column.")

    required_defaults = {
        "Account No.": "",
        "Dialed Number": "",
        "Month Extracted": "",
    }

    for col_name, default_value in required_defaults.items():
        if col_name not in df.columns:
            df[col_name] = default_value

    for col_name in ["STATUS", "Account No.", "Dialed Number", "Month Extracted"]:
        df[col_name] = df[col_name].fillna("").astype(str).str.strip()

    df["Account No."] = df["Account No."].str.replace(r"\.0$", "", regex=True)
    df["Dialed Number"] = df["Dialed Number"].str.replace(r"\.0$", "", regex=True)

    df["Dialed Number"] = df["Dialed Number"].apply(
        lambda x: f"+{x}" if x != "" and not x.startswith("+") else x
    )

    df = df[
        (df["STATUS"].str.upper() != "NEGATIVE") &
        (df["STATUS"].str.upper() != "0") &
        (df["STATUS"] != "")
    ].copy()

    df["Account No. + Month Extracted"] = df["Account No."] + " | " + df["Month Extracted"]

    new_order = ["Account No. + Month Extracted"] + [
        col for col in df.columns if col != "Account No. + Month Extracted"
    ]

    return df[new_order]


def to_excel_bytes_by_status(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    used_sheet_names = set()

    status_values = sorted(df["STATUS"].dropna().astype(str).unique().tolist())

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for status_value in status_values:
            sheet_name = clean_sheet_name(status_value, used_sheet_names)
            sheet_df = df[df["STATUS"] == status_value]
            sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

    output.seek(0)
    return output.getvalue()


def filter_negative_status(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str)
    df.columns = [col.strip() for col in df.columns]

    if "STATUS" not in df.columns:
        raise ValueError("Missing STATUS column.")

    required_defaults = {
        "Account No.": "",
        "Dialed Number": "",
        "Month Extracted": "",
    }

    for col_name, default_value in required_defaults.items():
        if col_name not in df.columns:
            df[col_name] = default_value

    for col_name in ["STATUS", "Account No.", "Dialed Number", "Month Extracted"]:
        df[col_name] = df[col_name].fillna("").astype(str).str.strip()

    df["Account No."] = df["Account No."].str.replace(r"\.0$", "", regex=True)
    df["Dialed Number"] = df["Dialed Number"].str.replace(r"\.0$", "", regex=True)

    df = df[df["STATUS"].str.upper() == "NEGATIVE"].copy()

    df["Dialed Number"] = df["Dialed Number"].apply(
        lambda x: f"+{x}" if x != "" and not x.startswith("+") else x
    )

    df["Account No. + Month Extracted"] = df["Account No."] + " | " + df["Month Extracted"]

    new_order = ["Account No. + Month Extracted"] + [
        col for col in df.columns if col != "Account No. + Month Extracted"
    ]

    return df[new_order]


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output.getvalue()


@st.cache_data
def convert_df_to_excel(df: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="FilteredData")
    output.seek(0)
    return output


def render_report_generator(report_mode: str) -> None:
    if report_mode == "📂 DRR CSV Processor":
        render_page_header("📂 DRR CSV Processing Tool", "Report Generator")

        uploaded_file = st.file_uploader("Upload DRR CSV File", type=["csv"])

        if uploaded_file is not None:
            st.success("✅ File uploaded")

            if st.button("🚀 Process File"):
                with st.spinner("Processing..."):
                    processed_df = process_drr_file(uploaded_file)

                if processed_df is not None:
                    st.success("✅ Processing complete")
                    st.dataframe(processed_df.head(50), use_container_width=True)

                    output = BytesIO()
                    processed_df.to_excel(output, index=False, engine="openpyxl")
                    output.seek(0)

                    st.download_button(
                        label="📥 Download Excel File",
                        data=output.getvalue(),
                        file_name="processed_drr.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    elif report_mode == "✅ POSITIVE Status":
        render_page_header("✅ Positive Status", "Report Generator")

        uploaded_file = st.file_uploader("Upload CMS EXTRACTION file", type=["xlsx"])

        if uploaded_file:
            try:
                df = process_positive_status(uploaded_file)
                st.dataframe(df.head(100), use_container_width=True)

                excel_data = to_excel_bytes_by_status(df)

                st.download_button(
                    "Download POS STATUS by Sheet",
                    data=excel_data,
                    file_name=f"POS_STATUS_{uploaded_file.name}"
                )

            except Exception as e:
                st.error(f"Error: {e}")

    elif report_mode == "❌ NEGATIVE Status":
        render_page_header("❌ Negative Status", "Report Generator")

        uploaded_file = st.file_uploader("Upload CMS EXTRACTION file", type=["xlsx"])

        if uploaded_file:
            try:
                df = filter_negative_status(uploaded_file)

                if df.shape[0] == 0:
                    st.warning("No NEGATIVE rows found.")
                else:
                    st.dataframe(df.head(100), use_container_width=True)

                    st.download_button(
                        "Download Filtered File",
                        data=to_excel_bytes(df),
                        file_name=uploaded_file.name
                    )

            except Exception as e:
                st.error(f"Error: {e}")

    elif report_mode == "🏍️ FIELD RESULT":
        render_page_header("🏍️ FIELD RESULT", "Report Generator · BPI Cards XDays")

        uploaded_file = st.file_uploader("Upload FIELD RESULT file", type=["xlsx"])

        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="RESULT")
                df.columns = [col.strip().lower() for col in df.columns]

                if "date" in df.columns:
                    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%m/%d/%Y")

                if "bank" not in df.columns:
                    st.error("❌ Missing 'bank' column.")
                    return

                filtered_df = df[df["bank"].str.contains("bpi cards xdays", case=False, na=False)]

                columns_to_display = [
                    "chcode", "status", "sub status", "informant", "client number",
                    "dl received/unreceived", "message", "ptp-date", "ptp amount",
                    "field_name", "date", "bank"
                ]

                filtered_columns = [col for col in columns_to_display if col in filtered_df.columns]

                st.write("Filtered Data:")
                st.dataframe(filtered_df[filtered_columns], use_container_width=True)

                excel_data = convert_df_to_excel(filtered_df[filtered_columns])

                st.download_button(
                    label="Download Filtered Data as Excel",
                    data=excel_data.getvalue(),
                    file_name="filtered_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error: {e}")