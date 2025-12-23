import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from utils.table_header_finder import read_excel_auto


def analyze_excel(file_path):

    # -------------------- Read Excel --------------------
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext == ".csv":
        df = pd.read_csv(file_path)
    else:
        df = read_excel_auto(file_path)

    df_original = df.copy()  # FORMATTED RAW BACKUP
    formatted_df = df_original.copy()

    for col in formatted_df.columns:
        if pd.api.types.is_numeric_dtype(formatted_df[col]):
            formatted_df[col] = formatted_df[col].apply(
                lambda x: f" {int(x)}" if pd.notna(x) else x
            )
    


    # -------------------- Column Detection --------------------
    def find_col(possible):
        return next((c for c in df.columns if c.strip().lower() in
                     [x.lower() for x in possible]), None)

    b_col = find_col(["B Number", "BNUMBER", "b number", "b party", "b_party", "CALL_DIALED_NUM", "BParty"])
    if not b_col:
        raise ValueError("❌ B Party column not found")

    calltype_col = find_col([
        "CallType", "CALL_TYPE", "Type"
    ])

    

    date_col = find_col([
        "CALL_START_DT_TM", "Start Date", "Datetime", "Date", "STRT_TM","Start Time"
    ])

    inbound = find_col([
        "INBOUND_OUTBOUND_IND","Direction"
    ])

    address_col = find_col(["Address", "Location", "Addr", "SITE_ADDRESS", "SiteLocation"])
    imei_col = find_col(["IMEI", "imei", "Imei number", "IMEI numbe"])

    # -------------------- Date --------------------
    if date_col:
        df["__DATE__"] = pd.to_datetime(df[date_col], errors="coerce")
    else:
        df["__DATE__"] = None

    # -------------------- CLEAN NUMBER (ANALYSIS ONLY) --------------------
    def normalize(num):
        if pd.isna(num):
            return None
        num = re.sub(r"\D", "", str(num))
        if num.startswith("92"):
            num = num[2:]
        elif num.startswith("0"):
            num = num[1:]
        return num if re.fullmatch(r"3\d{9}", num) else None

    df["__B_CLEAN__"] = df[b_col].apply(normalize)

    # -------------------- MOBILE SUMMARY --------------------
    mob = df.dropna(subset=["__B_CLEAN__"])
    g = mob.groupby("__B_CLEAN__")

    mobile_summary = pd.DataFrame({
        "Mobile Number": g.size().index,
        "Starting Date": g["__DATE__"].min().values,
        "Ending Date": g["__DATE__"].max().values,
        "Count": g.size().values
    }).sort_values("Count", ascending=False)

    # -------------------- ADDRESS SUMMARY --------------------
    address_summary = None
    if address_col:
        g = df.groupby(address_col)
        address_summary = pd.DataFrame({
            address_col: g.size().index,
            "Starting Date": g["__DATE__"].min().values,
            "Ending Date": g["__DATE__"].max().values,
            "Count": g.size().values
        }).sort_values("Count", ascending=False)

    # -------------------- IMEI SUMMARY --------------------
    imei_summary = None
    if imei_col:
        g = df.groupby(imei_col)
        imei_summary = pd.DataFrame({
            "IMEI Number": g.size().index,
            "Starting Date": g["__DATE__"].min().values,
            "Ending Date": g["__DATE__"].max().values,
            "Count": g.size().values
        }).sort_values("Count", ascending=False)
    if imei_summary is not None:
        imei_summary["IMEI Number"] = imei_summary["IMEI Number"].astype(str).apply(lambda x: " " + x)


    # -------------------- CALL LOGS SHEET --------------------
        call_df = df_original.copy()

       
        call_df["__B_CLEAN__"] = df["__B_CLEAN__"]
        call_df["__DATE__"] = df["__DATE__"]

        # normalize
        if inbound and calltype_col:
                # case: separate columns (sms/call + incoming/outgoing)
                call_df["CALLTYPE"] = (
                    call_df[inbound] + " " +
                    call_df[calltype_col]
                )
        else:
                # case: already combined column
                call_df["CALLTYPE"] = call_df[calltype_col]

       
        call_df["CALLTYPE"] = (call_df["CALLTYPE"].str.lower().str.replace("-", " ", regex=False).str.replace(r"\s+", " ", regex=True).str.strip())

        # helper conditions
        def is_in_call(x):
            return x in ["incoming", "incoming call", "incomingcall","call incoming", "callincomig","voice incoming","voiceincoming","incoming voice","incomingvoice"]

        def is_out_call(x):
            return x in ["outgoing", "outgoing call", "outgoingcall","call outgoing", "calloutgoing","voice outgoing","voiceoutgoing","outgoing voice","outgoingvoice"]

        def is_in_sms(x):
            return x in ["incoming sms","incomingsms","sms incoming", "smsincoming"]

        def is_out_sms(x):
            return x in ["outgoing sms", "outgoingsms","sms outgoing", "smsoutgoing"]

        summary = (
            call_df
            .dropna(subset=["__B_CLEAN__"])        # ✅ important
            .groupby("__B_CLEAN__")
            .apply(lambda x: pd.Series({
                "Same-Num-Count": len(x),
                "Starting Date": x["__DATE__"].min(),   # ✅ Start Date
                "Ending Date": x["__DATE__"].max(),     # ✅ End Date
                "In-SMS": x["CALLTYPE"].apply(is_in_sms).sum(),
                "Out-SMS": x["CALLTYPE"].apply(is_out_sms).sum(),
                "In-Call": x["CALLTYPE"].apply(is_in_call).sum(),
                "Out-Call": x["CALLTYPE"].apply(is_out_call).sum(),

               
            }))
            .reset_index()                       # 👈 drop=False (default)
            .rename(columns={"__B_CLEAN__": "B-party"})
            .sort_values(by="Same-Num-Count", ascending=False)  # 👈 Z → A
            .reset_index(drop=True)
        )



    # -------------------- SAVE --------------------
    out_dir = "temp_uploads"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "analyzed_excel_formatted.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        mobile_summary.to_excel(writer, "Mobile Numbers", index=False)
        if address_summary is not None:
            address_summary.to_excel(writer, "Addresses", index=False)
        if imei_summary is not None:
            imei_summary.to_excel(writer, "IMEI Numbers", index=False)
        summary.to_excel(writer, "Call Logs", index=False)
        formatted_df.to_excel(writer, "Formatted Data", index=False)

    # -------------------- FORMAT --------------------
    wb = load_workbook(out_path)
    fill = PatternFill("solid", start_color="ADD8E6")
    bold = Font(bold=True)

    for ws in wb.worksheets:
        for c in ws[1]:
            c.fill = fill
            c.font = bold
            c.alignment = Alignment(horizontal="center")

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)

            # ✅ Center align all cells
            for cell in col:
                cell.alignment = Alignment(horizontal="center")

            # ✅ Prevent scientific notation (force TEXT)
            ws.column_dimensions[col_letter].width = max(
                len(str(cell.value)) if cell.value else 10 for cell in col
            ) + 2
    ws = wb["Formatted Data"]

    for col in ws.columns:
        for cell in col:
            cell.number_format = "@"
            cell.alignment = Alignment(horizontal="center")

    wb.save(out_path)
    return out_path
