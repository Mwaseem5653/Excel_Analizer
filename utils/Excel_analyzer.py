import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from utils.table_header_finder import read_excel_auto


def analyze_excel(file_path):
    """
    Professional Excel Analyzer
    Now includes:
    - Start & End Date for B-number usage
    - Start & End Date for Address usage
    """

    # -------------------- Read Excel --------------------
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext == ".csv":
        df = pd.read_csv(file_path)
        temp_excel_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
        df.to_excel(temp_excel_path, index=False)
        file_path = temp_excel_path
    else:
        df = read_excel_auto(file_path)

    # -------------------- Identify Columns --------------------
    possible_a_cols = ["A Number", "ANUMBER", "a number", "A party", "A_party", "Aparty"]
    possible_b_cols = ["B Number", "BNUMBER", "b number", "b party", "b_party", "CALL_DIALED_NUM", "BParty"]

    a_col = next((col for col in df.columns if col.strip().lower() in [x.lower() for x in possible_a_cols]), None)
    b_col = next((col for col in df.columns if col.strip().lower() in [x.lower() for x in possible_b_cols]), None)

    if not b_col:
        raise ValueError("❌ No valid B-number column found")

    # -------------------- Identify Date Column --------------------
    possible_date_cols = ["CALL_START_DT_TM", "Start Date", "Start Time", "Date", "STRT_TM", "Datetime"]
    date_col = next((col for col in df.columns if col.strip().lower() in [x.lower() for x in possible_date_cols]), None)

    if date_col:
        df["__DATE__"] = pd.to_datetime(df[date_col], errors="coerce")
    else:
        df["__DATE__"] = None

    # -------------------- Normalize Numbers --------------------
    def normalize_number(num):
        if pd.isna(num):
            return None
        num = re.sub(r"\D", "", str(num))
        if num.startswith("92"):
            num = num[2:]
        elif num.startswith("0"):
            num = num[1:]
        return num if re.fullmatch(r"3\d{9}", num) else None

    if a_col:
        df[a_col] = df[a_col].apply(normalize_number).apply(lambda x: f" {x}" if pd.notna(x) else None)

    df[b_col] = df[b_col].apply(normalize_number).apply(lambda x: f" {x}" if pd.notna(x) else None)

    # -------------------- MOBILE NUMBER SUMMARY WITH DATE RANGE --------------------
    mobile_df = df[[b_col, "__DATE__"]].dropna(subset=[b_col])
    mobile_group = mobile_df.groupby(b_col)

    mobile_count = pd.DataFrame({
        "Mobile Number": mobile_group.size().index,
        "Starting Date": mobile_group["__DATE__"].min().values,
        "Ending Date": mobile_group["__DATE__"].max().values,
        "Count": mobile_group.size().values
    }).sort_values(by="Count", ascending=False)

    # -------------------- ADDRESS SUMMARY WITH DATE RANGE --------------------
    possible_address_cols = ["Address", "Location", "Addr", "SITE_ADDRESS", "SiteLocation"]
    address_col = next((col for col in df.columns if col.strip().lower() in [x.lower() for x in possible_address_cols]), None)

    address_summary = None
    if address_col:
        address_df = df[[address_col, "__DATE__"]].dropna(subset=[address_col])
        group = address_df.groupby(address_col)

        address_summary = pd.DataFrame({
            address_col: group.size().index,
            "Starting Date": group["__DATE__"].min().values,
            "Ending Date": group["__DATE__"].max().values,
            "Count": group.size().values
        }).sort_values(by="Count", ascending=False)

    # -------------------- IMEI SUMMARY (Already Existing) --------------------
    possible_imei_cols = ["IMEI", "imei", "Imei number", "IMEI numbe"]
    imei_col = next((col for col in df.columns if col.strip().lower() in [x.lower() for x in possible_imei_cols]), None)

    imei_summary = None
    if imei_col:
        imei_df = df[[imei_col, "__DATE__"]].dropna(subset=[imei_col])
        imei_df[imei_col] = imei_df[imei_col].astype(str)
        group = imei_df.groupby(imei_col)

        imei_summary = pd.DataFrame({
            "IMEI Number": group.size().index,
            "Starting Date": group["__DATE__"].min().values,
            "Ending Date": group["__DATE__"].max().values,
            "Count": group.size().values
        }).sort_values(by="Count", ascending=False)

    # -------------------- Save Excel --------------------
    output_dir = "temp_uploads"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "analyzed_excel_formatted.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        mobile_count.to_excel(writer, sheet_name="Mobile Numbers", index=False)
        if address_summary is not None:
            address_summary.to_excel(writer, sheet_name="Addresses", index=False)
        if imei_summary is not None:
            imei_summary.to_excel(writer, sheet_name="IMEI Numbers", index=False)
        df.to_excel(writer, sheet_name="Formatted Data", index=False)

    # -------------------- Apply Formatting --------------------
    wb = load_workbook(output_path)
    fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    bold_font = Font(bold=True)

    for sheet in wb.sheetnames:
        ws = wb[sheet]

        # Header formatting
        for cell in ws[1]:
            cell.fill = fill
            cell.font = bold_font
            cell.alignment = Alignment(horizontal="center")

        for col in ws.columns:
            length = max((len(str(cell.value)) for cell in col if cell.value), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = length + 4

    wb.save(output_path)
    return output_path
