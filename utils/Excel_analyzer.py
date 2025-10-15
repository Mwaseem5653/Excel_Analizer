import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from utils.table_header_finder import read_excel_auto


def analyze_excel(file_path):
    """
    Enhanced Professional Excel Analyzer:
    - Clean & normalize mobile numbers (03XXXXXXXXX or 92XXXXXXXXXX or +92XXXXXXXXXX)
    - Add space before numeric-only columns (for Excel text format)
    - Keep only valid A/B numbers
    """

    # -------------------- Read Excel --------------------
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext == ".csv":
        df = pd.read_csv(file_path)
        # Save temporary Excel version for uniform processing
        temp_excel_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
        df.to_excel(temp_excel_path, index=False)
        file_path = temp_excel_path
    else:
        df = read_excel_auto(file_path)

    # -------------------- Identify A & B Number Columns --------------------
    possible_a_cols = ["A Number", "ANUMBER", "a number", "A party", "A_party" , "Aparty"]
    possible_b_cols = ["B Number", "BNUMBER", "b number", "b party", "b_party", "CALL_DIALED_NUM" ,"BParty"]

    a_col = None
    b_col = None

    for col in df.columns:
        clean_col = col.strip().lower()
        if clean_col in [c.lower() for c in possible_a_cols]:
            a_col = col
        if clean_col in [c.lower() for c in possible_b_cols]:
            b_col = col

    if not b_col:
        raise ValueError("No valid B Number column found (expected: B Number, b party, CALL_DIALED_NUM, etc.)")

    # -------------------- Smart Mobile Normalization --------------------
    def normalize_number(num):
        if pd.isna(num):
            return None
        num = str(num)
        num = re.sub(r"\D", "", num)  # remove non-digits

        # Handle formats: +92300..., 92300..., 0300...
        if num.startswith("92") and len(num) >= 12:
            num = num[2:]
        elif num.startswith("0") and len(num) >= 11:
            num = num[1:]

        # Now must start with 3 and be 10 digits
        return num if re.fullmatch(r"3\d{9}", num) else None

    # -------------------- Clean A & B Columns --------------------
    if a_col:
        df[a_col] = df[a_col].apply(normalize_number)
        df[a_col] = df[a_col].apply(lambda x: f" {x}" if pd.notna(x) else None)

    df[b_col] = df[b_col].apply(normalize_number)
    df[b_col] = df[b_col].apply(lambda x: f" {x}" if pd.notna(x) else None)

    # -------------------- Create Mobile Count Sheet --------------------
    mobile_series = df[b_col].dropna()
    mobile_count = mobile_series.value_counts().reset_index()
    mobile_count.columns = ["Mobile Number", "Count"]
    mobile_count = mobile_count.sort_values(by="Count", ascending=False)

    # -------------------- Address (Optional) --------------------
    possible_address_cols = ["Address", "Location", "Addr","SITE_ADDRESS","SiteLocation"] 
    address_col = next((col for col in df.columns if col.strip().lower() in [c.lower() for c in possible_address_cols]), None)

    # -------------------- IMEI SHEET --------------------
    possible_imei_cols = ["IMEI", "imei", "Imei number", "IMEI numbe"]
    possible_date_cols = ["CALL_START_DT_TM", "Start Date", "Start Time", "Date" ,"STRT_TM" ,"Datetime"]

    imei_col = next((col for col in df.columns if col.strip().lower() in [c.lower() for c in possible_imei_cols]), None)
    date_col = next((col for col in df.columns if col.strip().lower() in [c.lower() for c in possible_date_cols]), None)

    imei_summary = None
    if imei_col:
        imei_df = df[[imei_col]].copy()
        if date_col and date_col in df.columns:
            imei_df["Date"] = pd.to_datetime(df[date_col], errors="coerce")

        imei_df = imei_df.dropna(subset=[imei_col])
        imei_df[imei_col] = imei_df[imei_col].astype(str)

        imei_group = imei_df.groupby(imei_col)
        imei_summary = pd.DataFrame({
            "IMEI Number": imei_group.size().index,
            "Count": imei_group.size().values
        })

        if date_col and "Date" in imei_df.columns:
            imei_summary["Starting Date"] = imei_group["Date"].min().values
            imei_summary["Ending Date"] = imei_group["Date"].max().values
        else:
            imei_summary["Starting Date"] = None
            imei_summary["Ending Date"] = None

        imei_summary = imei_summary[["IMEI Number", "Starting Date", "Ending Date", "Count"]]
        imei_summary = imei_summary.sort_values(by="Count", ascending=False)

    # -------------------- Add space to numeric-only columns --------------------
    df_formatted = df.copy()
    for col in df_formatted.columns:
        # Skip A/B columns (already formatted)
        if col in [a_col, b_col]:
            continue
        # If column is numeric-only
        if pd.to_numeric(df_formatted[col], errors="coerce").notna().all():
            df_formatted[col] = df_formatted[col].apply(lambda x: f" {x}" if pd.notna(x) else x)

    # -------------------- Save analyzed Excel --------------------
    output_dir = "temp_uploads"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "analyzed_excel_formatted.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        mobile_count.to_excel(writer, sheet_name="Mobile Numbers", index=False)
        if address_col:
            address_df = df[[address_col]].dropna()
            address_count = address_df[address_col].value_counts().reset_index()
            address_count.columns = [address_col, "Count"]
            address_count = address_count.sort_values(by="Count", ascending=False)
            address_count.to_excel(writer, sheet_name="Addresses", index=False)
        if imei_summary is not None:
            imei_summary.to_excel(writer, sheet_name="IMEI Numbers", index=False)
        df_formatted.to_excel(writer, sheet_name="Formatted Data", index=False)

    # -------------------- Apply Professional Formatting --------------------
    wb = load_workbook(output_path)
    fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    bold_font = Font(bold=True)

    for sheet in wb.sheetnames:
        ws = wb[sheet]

        # Header style
        for cell in ws[1]:
            cell.fill = fill
            cell.font = bold_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Center align all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Auto width + Text format for IMEI / Mobile
        for col in ws.columns:
            max_length = 0
            column = get_column_letter(col[0].column)
            for cell in col:
                if isinstance(cell.value, (int, float)) and sheet in ["Mobile Numbers", "IMEI Numbers"]:
                    cell.number_format = "@"
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = (max_length + 4)
            ws.column_dimensions[column].width = adjusted_width

    wb.save(output_path)
    return output_path
