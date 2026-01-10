import pandas as pd
import os
import re
import requests
import base64
from io import BytesIO
from openpyxl.drawing.image import Image as XLImage
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from utils.table_header_finder import read_excel_auto
from utils.api_clients import get_phone_info, get_eyecon_info


def analyze_excel(file_path, top_n=15, enable_lookup=True, enable_eyecon_lookup=False, eyecon_top_n=15, include_eyecon_images=False):

    # -------------------- Read Excel --------------------
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext == ".csv":
        df = pd.read_csv(file_path)
    else:
        df = read_excel_auto(file_path)

    df_original = df.copy()  # FORMATTED RAW BACKUP
    
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

    # -------------------- PREPARE FORMATTED SHEET --------------------
    formatted_df = df_original.copy()
    
    # 1. Format Date Column to String (if exists)
    if date_col and date_col in formatted_df.columns:
        # Check if it's already datetime, if not convert
        if not pd.api.types.is_datetime64_any_dtype(formatted_df[date_col]):
             formatted_df[date_col] = pd.to_datetime(formatted_df[date_col], errors='coerce')
        
        # Format as string
        formatted_df[date_col] = formatted_df[date_col].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')

    # 2. Format Numeric Columns (Skip Date)
    for col in formatted_df.columns:
        if col == date_col:
            continue
            
        if pd.api.types.is_numeric_dtype(formatted_df[col]):
            formatted_df[col] = formatted_df[col].apply(
                lambda x: f" {int(x)}" if pd.notna(x) else x
            )

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

    # -------------------- FETCH API INFO --------------------
    lookup_cache = {} # Store results to use in Call Logs later

    if enable_lookup:
        mobile_summary["Name"] = ""
        mobile_summary["CNIC"] = ""
        mobile_summary["Address"] = ""

        # Iterate over top N
        for idx in mobile_summary.index[:top_n]:
            raw_num = str(mobile_summary.at[idx, "Mobile Number"]).strip()
            # Add 0 for 11-digit format
            query_num = "0" + raw_num
            
            try:
                # Direct function call (Serverless logic)
                data = get_phone_info(query_num)
                
                # Check if we got a valid list of records
                if isinstance(data, list) and len(data) > 0:
                    record = data[0]
                    # Try common keys
                    name_val = record.get("name") or record.get("Name") or ""
                    cnic_val = record.get("cnic") or record.get("CNIC") or ""
                    addr_val = record.get("address") or record.get("Address") or ""
                    
                    mobile_summary.at[idx, "Name"] = name_val
                    # Add space to prevent scientific notation in Excel
                    mobile_summary.at[idx, "CNIC"] = " " + str(cnic_val) if cnic_val else ""
                    mobile_summary.at[idx, "Address"] = addr_val
                    
                    # Cache for Call Logs
                    if raw_num not in lookup_cache:
                         lookup_cache[raw_num] = {}
                    
                    lookup_cache[raw_num]["Name"] = name_val
                    lookup_cache[raw_num]["CNIC"] = " " + str(cnic_val) if cnic_val else ""
                    lookup_cache[raw_num]["Address"] = addr_val

            except Exception as e:
                print(f"⚠️ Lookup Failed for {query_num}: {e}")

    # -------------------- FETCH EYECON INFO --------------------
    eyecon_cache = {}
    if enable_eyecon_lookup:
        mobile_summary["Eyecon Name"] = ""
        if include_eyecon_images:
            mobile_summary["Eyecon Image"] = ""
            mobile_summary["Facebook Link"] = ""
        
        def extract_eyecon_names(d):
            found = set()
            if not isinstance(d, dict): return found
            
            # Check direct keys
            fname = d.get("fullName") or d.get("name")
            if fname: found.add(fname)
            
            # Check otherNames list
            others = d.get("otherNames", [])
            if isinstance(others, list):
                for o in others:
                    if isinstance(o, dict):
                        n = o.get("name")
                        if n: found.add(n)
                    elif isinstance(o, str):
                        found.add(o)
            return found

        for idx in mobile_summary.index[:eyecon_top_n]:
            raw_num = str(mobile_summary.at[idx, "Mobile Number"]).strip()
            
            try:
                data = get_eyecon_info(raw_num)
                
                # Check specific API error flags
                if isinstance(data, dict) and "message" in data:
                    msg = data["message"].lower()
                    if "quota" in msg or "subscribe" in msg:
                        error_msg = "Quota Exceeded / Sub Req"
                        eyecon_cache[raw_num] = {"name": error_msg, "image": ""}
                        mobile_summary.at[idx, "Eyecon Name"] = error_msg
                        if include_eyecon_images:
                            mobile_summary.at[idx, "Eyecon Image"] = ""
                            mobile_summary.at[idx, "Facebook Link"] = ""
                        continue

                # Proceed if valid status
                if isinstance(data, dict) and (data.get("status") or data.get("fullName")):
                    all_names = extract_eyecon_names(data)
                    
                    # Check nested 'data' object if it exists
                    if isinstance(data.get("data"), dict):
                        all_names.update(extract_eyecon_names(data["data"]))
                    
                    final_name = ""
                    if all_names:
                        final_name = " | ".join(sorted(list(all_names)))
                    
                    # Extract Image
                    image_url = ""
                    
                    # Helper to get image from 'images' list
                    def get_images_url(d):
                        if isinstance(d.get("images"), list):
                            for img_entry in d["images"]:
                                if isinstance(img_entry, dict) and "pictures" in img_entry:
                                    pics = img_entry["pictures"]
                                    if isinstance(pics, dict):
                                        return pics.get("200") or pics.get("600") or next(iter(pics.values()), None)
                        return None

                    # 1. Check for Base64 (nested or direct)
                    raw_b64 = data.get("b64")
                    if not raw_b64 and isinstance(data.get("data"), dict):
                        raw_b64 = data["data"].get("b64")
                    
                    if raw_b64:
                        if not raw_b64.startswith("data:image"):
                            image_url = f"data:image/jpeg;base64,{raw_b64}"
                        else:
                            image_url = raw_b64
                    
                    # 2. Check for Photo URL (nested or direct)
                    if not image_url:
                        if isinstance(data.get("data"), dict) and data["data"].get("photo"):
                            image_url = data["data"].get("photo")
                        elif data.get("photo"):
                            image_url = data.get("photo")
                    
                    # 3. Check for Images List (Facebook etc)
                    if not image_url:
                         image_url = get_images_url(data)
                         if not image_url and isinstance(data.get("data"), dict):
                             image_url = get_images_url(data["data"])

                    # 4. Check otherNames
                    if not image_url and isinstance(data.get("otherNames"), list):
                         for item in data["otherNames"]:
                             if isinstance(item, dict) and item.get("photo"):
                                 image_url = item.get("photo")
                                 break
                    
                    # Extract Facebook URL
                    fb_url = ""
                    if isinstance(data.get("facebookID"), dict):
                        fb_url = data["facebookID"].get("url")
                    elif isinstance(data.get("data"), dict) and isinstance(data["data"].get("facebookID"), dict):
                        fb_url = data["data"]["facebookID"].get("url")
                        
                    mobile_summary.at[idx, "Eyecon Name"] = final_name
                    if include_eyecon_images:
                        mobile_summary.at[idx, "Eyecon Image"] = image_url
                        mobile_summary.at[idx, "Facebook Link"] = fb_url
                    
                    eyecon_cache[raw_num] = {
                        "name": final_name, 
                        "image": image_url if include_eyecon_images else "",
                        "fb_url": fb_url
                    }
                else:
                    # Optional: Log not found?
                    pass

            except Exception as e:
                print(f"⚠️ Eyecon Lookup Failed for {raw_num}: {e}")

    # Reorder columns: Mobile Number, [Eyecon Name], [Eyecon Image], [Name, CNIC, Address], ... others
    base_mob_cols = ["Mobile Number"]
    if enable_eyecon_lookup:
        base_mob_cols.append("Eyecon Name")
        if include_eyecon_images:
            base_mob_cols.append("Eyecon Image")
            base_mob_cols.append("Facebook Link")
    if enable_lookup:
        base_mob_cols.extend(["Name", "CNIC", "Address"])
        
    other_mob_cols = [c for c in mobile_summary.columns if c not in base_mob_cols]
    desired_order = base_mob_cols + ["Starting Date", "Ending Date", "Count"] 
    # Filter to ensure all exist
    final_order = [c for c in desired_order if c in mobile_summary.columns]
    
    mobile_summary = mobile_summary[final_order]

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
                "Starting Date": x["__DATE__"].min(),   # ✅ Start Date
                "Ending Date": x["__DATE__"].max(),     # ✅ End Date
                "In-SMS": x["CALLTYPE"].apply(is_in_sms).sum(),
                "Out-SMS": x["CALLTYPE"].apply(is_out_sms).sum(),
                "In-Call": x["CALLTYPE"].apply(is_in_call).sum(),
                "Out-Call": x["CALLTYPE"].apply(is_out_call).sum(),
                 "Same-Num-Count": len(x),

               
            }))
            .reset_index()                       # 👈 drop=False (default)
            .rename(columns={"__B_CLEAN__": "B-party"})
            .sort_values(by="Same-Num-Count", ascending=False)  # 👈 Z → A
            .reset_index(drop=True)
        )

        # -------------------- INJECT API INFO INTO CALL LOGS --------------------
        base_cols = ["B-party"]
        
        if enable_eyecon_lookup:
            summary["Eyecon Name"] = summary["B-party"].map(lambda x: eyecon_cache.get(str(x), {}).get("name", ""))
            base_cols.append("Eyecon Name")
            if include_eyecon_images:
                summary["Eyecon Image"] = summary["B-party"].map(lambda x: eyecon_cache.get(str(x), {}).get("image", ""))
                summary["Facebook Link"] = summary["B-party"].map(lambda x: eyecon_cache.get(str(x), {}).get("fb_url", ""))
                base_cols.append("Eyecon Image")
                base_cols.append("Facebook Link")
            
        if enable_lookup:
            summary["Name"] = summary["B-party"].map(lambda x: lookup_cache.get(str(x), {}).get("Name", ""))
            summary["CNIC"] = summary["B-party"].map(lambda x: lookup_cache.get(str(x), {}).get("CNIC", ""))
            summary["Address"] = summary["B-party"].map(lambda x: lookup_cache.get(str(x), {}).get("Address", ""))
            base_cols.extend(["Name", "CNIC", "Address"])

        # Reorder Summary Columns: B-party, [Eyecon Name], [Name, CNIC, Address], ... rest
        sum_cols = list(summary.columns)
        other_cols = [c for c in sum_cols if c not in base_cols]
        summary = summary[base_cols + other_cols]



    # -------------------- SAVE --------------------
    out_dir = "temp_uploads"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "analyzed_excel_formatted.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        mobile_summary.to_excel(writer, sheet_name="Mobile Numbers", index=False)
        if address_summary is not None:
            address_summary.to_excel(writer, sheet_name="Addresses", index=False)
        if imei_summary is not None:
            imei_summary.to_excel(writer, sheet_name="IMEI Numbers", index=False)
        summary.to_excel(writer, sheet_name="Call Logs", index=False)
        formatted_df.to_excel(writer, sheet_name="Formatted Data", index=False)

    # -------------------- FORMAT --------------------
    wb = load_workbook(out_path)
    fill = PatternFill("solid", start_color="ADD8E6")
    bold = Font(bold=True)

    for ws in wb.worksheets:
        eyecon_col_index = None
        eyecon_img_col_index = None
        fb_link_col_index = None

        for c in ws[1]:
            c.fill = fill
            c.font = bold
            c.alignment = Alignment(horizontal="center")
            if c.value == "Eyecon Name":
                eyecon_col_index = c.column
            if c.value == "Eyecon Image":
                eyecon_img_col_index = c.column
            if c.value == "Facebook Link":
                fb_link_col_index = c.column

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)

            # ✅ Center align all cells
            for cell in col:
                cell.alignment = Alignment(horizontal="center")

            # ✅ Prevent scientific notation (force TEXT)
            ws.column_dimensions[col_letter].width = max(
                len(str(cell.value)) if cell.value else 10 for cell in col
            ) + 2

        # Specific formatting for Eyecon Name
        if eyecon_col_index:
            col_letter = get_column_letter(eyecon_col_index)
            ws.column_dimensions[col_letter].width = 30 # Fixed width for wrapping
            for cell in ws[col_letter]:
                 # Keep existing horizontal/vertical but enable wrap_text
                 cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # -------------------- HYPERLINK FACEBOOK COLUMN --------------------
        if fb_link_col_index:
             col_letter = get_column_letter(fb_link_col_index)
             ws.column_dimensions[col_letter].width = 15
             
             for row in ws.iter_rows(min_row=2, min_col=fb_link_col_index, max_col=fb_link_col_index):
                 cell = row[0]
                 url = cell.value
                 if url and isinstance(url, str) and url.startswith("http"):
                     cell.value = "View Profile"
                     cell.hyperlink = url
                     cell.font = Font(color="0000FF", underline="single")
                     cell.alignment = Alignment(horizontal="center", vertical="center")

        # -------------------- EMBED EYECON IMAGES --------------------
        if include_eyecon_images and eyecon_img_col_index:
            col_letter = get_column_letter(eyecon_img_col_index)
            ws.column_dimensions[col_letter].width = 15  # Fixed width for image column
            
            # Iterate through rows in the image column
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, min_col=eyecon_img_col_index, max_col=eyecon_img_col_index), start=2):
                cell = row[0]
                url = cell.value
                
                if url and isinstance(url, str):
                    img_data = None
                    try:
                        if url.startswith("data:image"):
                            # Handle Base64
                            try:
                                header, encoded = url.split(",", 1)
                                img_bytes = base64.b64decode(encoded)
                                img_data = BytesIO(img_bytes)
                            except Exception as e:
                                print(f"⚠️ Failed to decode base64 image at row {row_idx}: {e}")
                        
                        elif url.startswith("http"):
                            # Handle URL
                            try:
                                res = requests.get(url, timeout=5)
                                if res.status_code == 200:
                                    img_data = BytesIO(res.content)
                                else:
                                    print(f"⚠️ Image download failed code {res.status_code} at row {row_idx}")
                            except Exception as e:
                                print(f"⚠️ Image download failed at row {row_idx}: {e}")

                        if img_data:
                            img = XLImage(img_data)
                            
                            # Resize image (approx 100x100 px)
                            img.width = 100
                            img.height = 100
                            
                            # Adjust row height to fit image
                            ws.row_dimensions[row_idx].height = 80
                            
                            # Clear URL and add image
                            cell.value = "" 
                            ws.add_image(img, cell.coordinate)
                            
                    except Exception as e:
                        print(f"⚠️ Failed to embed image for row {row_idx}: {e}")

    ws = wb["Formatted Data"]

    for col in ws.columns:
        for cell in col:
            cell.number_format = "@"
            cell.alignment = Alignment(horizontal="center")

    wb.save(out_path)
    return out_path
