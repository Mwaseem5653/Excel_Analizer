import os
import streamlit as st
from dotenv import load_dotenv
import time
import auth
import shutil
import uuid

# Load environment variables
load_dotenv()

# ---------- Page Config ----------
st.set_page_config(page_title="Urdu Police Application Extractor", layout="wide")

# ---------- Session State ----------
if "page" not in st.session_state:
    st.session_state.page = "app"  # default page
    if "session_id" not in st.session_state:
        st.session_state.session_id = str(uuid.uuid4())

SESSION_TEMP_DIR = os.path.join("temp_uploads", st.session_state.session_id)
os.makedirs(SESSION_TEMP_DIR, exist_ok=True)

# ---------- Cleanup ----------
def cleanup_session_files():
    if os.path.exists(SESSION_TEMP_DIR):
        shutil.rmtree(SESSION_TEMP_DIR, ignore_errors=True)

def normalize_phone_number(num: str):
    num = num.strip()
    # agar 10 digit ho aur 3 se start ho → 0 add karo
    if len(num) == 10 and num.startswith("3"):
        return "0" + num
    # agar already 11 digit aur 0 se start ho → ok
    if len(num) == 11 and num.startswith("0"):
        return num
    if len(num) == 12 and num.startswith("92"):
        return "0" + num[2:]
    return None

def handle_application_extractor():
    # Lazy imports to save memory
    import google.generativeai as genai
    import fitz  # PyMuPDF
    import asyncio
    from multi_file_handler import handle_files
    from utils.extract_fields import extract_fields_from_text
    from utils.excel_writer import save_to_excel
    
    st.title("📝 Urdu Police Application Extractor")
    
    genai.configure(api_key=os.getenv("GENAI_API_KEY"))

    model_options = {
         "Gemini 1.5 Flash": "models/gemini-flash-latest",
         "Gemini 2.0 Flash": "models/gemini-2.0-flash",
         "Gemini 2.5 Flash": "models/gemini-2.5-flash",
         "Gemini 2.5 Flash Image": "models/gemini-2.5-flash-image",
         "Gemini 2.5 Flash Image Preview": "models/gemini-2.5-flash-image-preview"
    }
    selected_model_key = st.selectbox("Select Gemini Model", list(model_options.keys()), index=1)
    selected_model_name = model_options[selected_model_key]
    
    uploaded_files = st.file_uploader(
        "Upload handwritten Urdu application image(s) or PDF(s):",
        type=["jpg", "jpeg", "png", "pdf"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.info("🔍 Processing files. Please wait...")
        os.makedirs("temp_uploads", exist_ok=True)

        class FakeMessage:
            def __init__(self, files):
                self.elements = []
                for f in files:
                    temp_path = os.path.join("temp_uploads", f.name)
                    with open(temp_path, "wb") as out_f:
                        out_f.write(f.getbuffer())
                    self.elements.append(type("Element", (), {"path": temp_path})())

        message = FakeMessage(uploaded_files)
        file_data = asyncio.run(handle_files(message))
        all_extracted_data = []

        for file in file_data:
            if "error" in file:
                st.error(f"❌ {file['file_name']}: {file['error']}")
                continue

            path = file["path"]
            ext = os.path.splitext(path)[1].lower()

            if ext == ".pdf":
                doc = fitz.open(path)
                page_num = int(file["file_name"].split("Page ")[1]) - 1
                pix = doc.load_page(page_num).get_pixmap()
                image_bytes = pix.tobytes("png")
            else:
                with open(path, "rb") as f:
                    image_bytes = f.read()

            st.info(f"Processing file: {file['file_name']}")

            prompt = (
                "From this handwritten Urdu police application image, extract ONLY the following fields. "
                "Translate the content into English if needed and follow fields Example stricly and return Plain Text only\n\n"
                "Fields Example:\n"
                """
                Name: Furqan Ur Rehman (only applicant name)
                Phone Number: 0313-0282098 (Mention in Last)
                IMEI Number: 354882089097706 354882089094534
                last Num Used: 0313-0282044 or None
                Mobile Model: Motrolla Edge Plus
                Other Property: None / Cash 3000 / wallet / bike  etc
                Date Of Offence: 29.06.2025 only use . instead /
                Time Of Offence: 08:00 PM
                Type: Snatched / Theft / Lost
                Police Station: ZamanTown"""
            )

            check_rate_limit()
            model = genai.GenerativeModel(selected_model_name)

            try:
                response = model.generate_content(
                    [prompt, {"mime_type": "image/jpeg", "data": image_bytes}]
                )
                raw_text = response.text
            except Exception as e:
                st.error(f"❌ Gemini Vision error: {str(e)}")
                continue

            st.text_area(f"📝 Extracted Text ({file['file_name']}):", raw_text, height=200)
            extracted_data = extract_fields_from_text(raw_text)
            all_extracted_data.append(extracted_data)

        excel_path = save_to_excel(all_extracted_data)
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📥 Download Extracted Excel",
                data=f,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                
            )
        
        cleanup_session_files()
def handle_excel_analyzer():
    # Lazy imports
    import zipfile
    import io
    from utils.Excel_analyzer import analyze_excel

    st.title("📈 Excel Analyzer")
    st.info("Upload multiple Excel/CSV files to analyze Mobile Numbers & Addresses.")

    uploaded_files = st.file_uploader(
        "Upload Excel/CSV files",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="analyzer_uploader"
    )

    if uploaded_files:
        os.makedirs("temp_uploads", exist_ok=True)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for uploaded_file in uploaded_files:
                temp_path = os.path.join("temp_uploads", uploaded_file.name)
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                try:
                    st.write(f"⏳ Processing file: **{uploaded_file.name}** ...")
                    analyzed_path = analyze_excel(temp_path)
                    st.success(f"✅ {uploaded_file.name} analyzed successfully!")
                    zipf.write(analyzed_path, arcname="(Analyzed)-" + uploaded_file.name)
                except Exception as e:
                    st.error(f"❌ Error in {uploaded_file.name}: {str(e)}")

        zip_buffer.seek(0)
        st.download_button(
            label="📦 Download All Analyzed Files (ZIP)",
            data=zip_buffer,
            file_name="Analyzed_Files.zip",
            mime="application/zip"
            
        )
    cleanup_session_files() 
def handle_pta_services():
    import io
    import pandas as pd
    from utils.operator_lookup import find_operators_and_download

    st.title("PTA Services - Operator Lookup and Sorted Export")
    st.subheader("Enter Phone Numbers for Lookup and Sorting")
    phone_numbers_input = st.text_area("Enter phone numbers, one per line:", key="phone_numbers_input_single_section")

    if st.button("Process Numbers for Operators and Sorted Export", key="process_numbers_button"):
        if phone_numbers_input:
            raw_numbers = [num.strip() for num in phone_numbers_input.split('\n') if num.strip()]
            phone_numbers = []
            for num in raw_numbers:
                normalized = normalize_phone_number(num)
                if normalized:
                    phone_numbers.append(normalized)

            phone_numbers = phone_numbers[:10]
            if phone_numbers:
                try:
                    excel_path_dummy, lookup_results_data = find_operators_and_download(phone_numbers) 
                    st.subheader("Processing Results")
                    
                    operator_groups = {
                        "Jazz Pakistan": [], "Zong Pakistan": [], "Telenor Pakistan": [],
                        "Ufone Pakistan": [], "Other": []
                    }
                    
                    for item in lookup_results_data:
                        operator = item["operator"]
                        if operator in operator_groups:
                            operator_groups[operator].append(item)
                        else:
                            operator_groups["Other"].append(item)
                    
                    sorted_results_list = []
                    for op_name in ["Jazz Pakistan", "Zong Pakistan", "Telenor Pakistan", "Ufone Pakistan", "Other"]:
                        sorted_results_list.extend(operator_groups[op_name])

                    standardized_results = []
                    for item in sorted_results_list:
                        number = item["number"]
                        if len(number) == 11 and number.startswith('0'):
                            standardized_number = '92' + number[1:]
                        else:
                            standardized_number = number
                        standardized_results.append({
                            "Phone Number": standardized_number,
                            "Detected Operator": item["operator"]
                        })

                    df_final_results = pd.DataFrame(standardized_results)
                    st.session_state.pta_results_for_cdr = standardized_results
                    st.dataframe(df_final_results)

                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        df_final_results.to_excel(writer, index=False, sheet_name='Sorted_Operator_Results')
                    excel_buffer.seek(0)

                    st.download_button(
                        label=f"📥 Download Sorted Operator Results (Excel)",
                        data=excel_buffer,
                        file_name="sorted_operator_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"An error occurred: {e}")
            else:
                st.warning("Please enter at least one valid phone number.")

def handle_cdr_format():
    import re

    st.title("📝 CDR Format")
    st.info("Select an HTML template to view and see numbers from PTA Services filtered by category.")

    if st.button("Clear Filtered Numbers", key="clear_filtered_numbers_top"):
        st.session_state.pta_results_for_cdr = []
        st.session_state.selected_cdr_html = None 
        st.rerun()

    html_files_to_process = {
        "Jazz CDR HTML": {"file": "jazz cdr 6 MONTH.html", "operator_key": "Jazz Pakistan"},
        "Telenor CDR HTML": {"file": "Telenor 6 month cdr.html", "operator_key": "Telenor Pakistan"},
        "Ufone Multi CDR HTML": {"file": "ufone 2 or more cdr 1 year.html", "operator_key": "Ufone Pakistan"},
        "Ufone Single CDR HTML": {"file": "ufone single cdr 1 year.html", "operator_key": "Ufone Pakistan"},
        "Zong CDR HTML": {"file": "zong cdr 6 MONTH.html", "operator_key": "Zong Pakistan"},
        "IMEI Format 3 Month HTML": {"file": "imei format 3 month.html", "operator_key": "All Operators"},
        "IMEI Format 6 Month HTML": {"file": "imei format 6 month.html", "operator_key": "All Operators"},
    }
    
    cols = st.columns(len(html_files_to_process))
    if "selected_cdr_html" not in st.session_state:
        st.session_state.selected_cdr_html = None

    for i, (button_label, file_info) in enumerate(html_files_to_process.items()):
        with cols[i]:
            if st.button(button_label, key=f"cdr_html_button_{button_label}"):
                st.session_state.selected_cdr_html = file_info

    if st.session_state.selected_cdr_html:
        selected_file_name = st.session_state.selected_cdr_html["file"]
        target_operator_key = st.session_state.selected_cdr_html["operator_key"]

        filtered_numbers_for_html = []
        if "pta_results_for_cdr" in st.session_state and st.session_state.pta_results_for_cdr:
            for item in st.session_state.pta_results_for_cdr:
                if target_operator_key.lower().replace(" pakistan", "") in item["Detected Operator"].lower().replace(" pakistan", ""):
                    filtered_numbers_for_html.append(item["Phone Number"])

        numbers_to_inject = "\n".join(filtered_numbers_for_html)

        try:
            with open(selected_file_name, "r", encoding="utf-8") as f:
                html_content = f.read()
            
            def replace_textarea_content(match):
                return match.group(1) + numbers_to_inject + match.group(3)

            modified_html_content = re.sub(
                r'(<textarea[^>]*id=["\']formatinput["\'][^>]*>)((?!</textarea>).*?)(</textarea>)',
                replace_textarea_content,
                html_content,
                flags=re.DOTALL | re.IGNORECASE
            )
            
            script_to_inject = """
            <script type="text/javascript">
            document.addEventListener('DOMContentLoaded', function() {
                setTimeout(function() {
                    if (typeof changeFormat === 'function') {
                        changeFormat();
                    }
                }, 100);
            });
            </script>
            """
            modified_html_content = modified_html_content.replace('</body>', f'{script_to_inject}</body>')
            st.components.v1.html(modified_html_content, height=750, scrolling=True)
        except Exception as e:
            st.error(f"Error: {e}")

def show_main_app():
    import pandas as pd
    from io import BytesIO
    from utils.api_clients import get_phone_info, get_vehicle_info

    st.title("🇵🇰 Information Extractor")
    search_type = st.selectbox("Select Search Type", ["SIM Info", "Vehicle Info"])

    if search_type == "SIM Info":
        st.header("SIM Owner Details")
        input_method = st.radio("Choose input method:", ("Manual Entry", "Upload Excel File"))

        if input_method == "Manual Entry":
            search_terms = st.text_area("Enter Phone Numbers / CNICs", placeholder="03001234567\n3520212345678")
            if st.button("🔍 Search SIM Info"):
                raw_items = search_terms.replace(",", "\n").split("\n")
                items = [i.strip() for i in raw_items if i.strip()]
                items = ["0" + i if len(i) == 10 and i.isdigit() else i for i in items]

                if items:
                    results = []
                    with st.spinner("Fetching SIM data..."):
                        for item in items:
                            data = get_phone_info(item)
                            if isinstance(data, list):
                                for record in data:
                                    results.append({"Input": item, "Name": record.get("name"), "Number": record.get("number"), "CNIC": record.get("cnic"), "Address": record.get("address"), "Status": "Found"})
                            else:
                                results.append({"Input": item, "Name": "", "Number": "", "CNIC": "", "Address": "", "Status": data.get("error", "Error")})
                    df = pd.DataFrame(results)
                    st.dataframe(df, use_container_width=True)
                    excel_buffer = BytesIO()
                    df.to_excel(excel_buffer, index=False)
                    st.download_button(label="⬇️ Download Excel", data=excel_buffer.getvalue(), file_name="sim_info_results.xlsx")

        elif input_method == "Upload Excel File":
            uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
            if uploaded_file:
                all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
                if "Mobile Numbers" in all_sheets:
                    df_upload = all_sheets["Mobile Numbers"]
                    if "Mobile Number" in df_upload.columns:
                        df_upload['Mobile Number'] = df_upload['Mobile Number'].astype(str).str.replace(r'\.0$', '', regex=True)
                        if st.button("🚀 Fetch Data"):
                            items = df_upload["Mobile Number"].dropna().head(10).tolist()
                            all_results = []
                            with st.spinner("Fetching..."):
                                for item in items:
                                    data = get_phone_info(item)
                                    if isinstance(data, list):
                                        for record in data:
                                            all_results.append({"Original Phone Number": item, "Name": record.get("name"), "Number": record.get("number"), "CNIC": record.get("cnic"), "Address": record.get("address")})
                                    else:
                                        all_results.append({"Original Phone Number": item, "Name": "Not Found"})
                            df_results = pd.DataFrame(all_results)
                            all_sheets["Search Results"] = df_results
                            excel_buffer = BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                for sheet_name, df in all_sheets.items():
                                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                            st.download_button(label="⬇️ Download Updated Excel", data=excel_buffer.getvalue(), file_name="updated_sim_info.xlsx")

    elif search_type == "Vehicle Info":
        st.header("Sindh Vehicle Details")
        vehicle_category = st.selectbox("Select Vehicle Category", ["", "2 wheeler", "4 wheeler"])
        reg_number = st.text_input("Enter Registration Number")
        if st.button("🔍 Search Vehicle Info"):
            if reg_number and vehicle_category:
                api_category = "2W" if vehicle_category == "2 wheeler" else "4W"
                data = get_vehicle_info(reg_number, api_category)
                if data.get("error"): st.error(data["error"])
                else: st.table(pd.DataFrame({"Attribute": list(data.keys()), "Value": list(data.values())}))

# ---------- Rate Limiting ----------
REQUEST_LIMIT = 5
TIME_WINDOW = 120
request_times = []

def check_rate_limit():
    global request_times
    now = time.time()
    request_times = [t for t in request_times if now - t < TIME_WINDOW]
    if len(request_times) >= REQUEST_LIMIT:
        wait_time = TIME_WINDOW - (now - request_times[0])
        st.warning(f"⏳ Rate limit! Waiting {int(wait_time)}s...")
        time.sleep(wait_time)
    request_times.append(time.time())

# ---------- Main App ----------
def main():
    if not auth.is_logged_in():
        auth.login()
        return

    with st.sidebar:
        st.image("Assets/app_icon.png", width=100)
        user_services = auth.get_user_services()
        service_map = {
            "Application Extractor": "app", "Excel Analyzer": "analyzer",
            "PTA Services": "pta_services", "CDR Format": "cdr_format",
            "Vehicle and Mobile": "vehicle_and_mobile", "Admin": "admin",
            "Settings / Future Tools": "settings"
        }
        for service in user_services:
            if service in service_map:
                if st.button(service):
                    st.session_state.page = service_map[service]
        if st.button("Logout"):
            auth.logout()

    page = st.session_state.get("page", "app")
    if page == "app": handle_application_extractor()
    elif page == "analyzer": handle_excel_analyzer()
    elif page == "pta_services": handle_pta_services()
    elif page == "cdr_format": handle_cdr_format()
    elif page == "vehicle_and_mobile": show_main_app()
    elif page == "admin": auth.admin_section()
    elif page == "settings":
        st.title("Settings / Tools")
        st.subheader("System Maintenance")
        if st.button("🗑️ Clear Temp Files"):
            cleanup_session_files()
            st.success("Temp files cleared successfully!")
        st.info("This section is under construction for more tools.")

if __name__ == "__main__":
    main()