import atexit
import sys
import os
import fitz  # PyMuPDF
import streamlit as st
from dotenv import load_dotenv
import google.generativeai as genai
from utils.extract_fields import extract_fields_from_text
from utils.excel_writer import save_to_excel
from multi_file_handler import handle_files
from utils.Excel_analyzer import analyze_excel
import asyncio
import time
import zipfile
import io
import pandas as pd
import auth
import requests
from io import BytesIO
import subprocess

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

    # baqi invalid
    return None


# Load Gemini API key
load_dotenv()
genai.configure(api_key=os.getenv("GENAI_API_KEY"))

# ---------- Page Config ----------
st.set_page_config(page_title="Urdu Police App & Excel Analyzer", layout="wide")

# ---------- Session State ----------
if "page" not in st.session_state:
    st.session_state.page = "app"  # default page

def handle_application_extractor():
    st.title("📝 Urdu Police Application Extractor")
    
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

        # Create temp folder
        os.makedirs("temp_uploads", exist_ok=True)

        # Convert Streamlit UploadedFile to compatible message format
        class FakeMessage:
            def __init__(self, files):
                self.elements = []
                for f in files:
                    temp_path = os.path.join("temp_uploads", f.name)
                    with open(temp_path, "wb") as out_f:
                        out_f.write(f.getbuffer())
                    self.elements.append(type("Element", (), {"path": temp_path})())

        message = FakeMessage(uploaded_files)

        # Handle files
        file_data = asyncio.run(handle_files(message))
        all_extracted_data = []

        for file in file_data:
            if "error" in file:
                st.error(f"❌ {file['file_name']}: {file['error']}")
                continue

            path = file["path"]
            ext = os.path.splitext(path)[1].lower()

            # Prepare image bytes
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
                """Name: Furqan Ur Rehman (only applicant name)
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

            # ✅ Apply Rate Limit before every request
            check_rate_limit()

            model = genai.GenerativeModel(selected_model_name)

            try:
                response = model.generate_content(
                    [
                        prompt,
                        {
                            "mime_type": "image/jpeg",
                            "data": image_bytes
                        }
                    ]
                )
                raw_text = response.text
            except Exception as e:
                st.error(f"❌ Gemini Vision error: {str(e)}")
                continue

            st.text_area(f"📝 Extracted Text ({file['file_name']}):", raw_text, height=200)

            extracted_data = extract_fields_from_text(raw_text)
            all_extracted_data.append(extracted_data)

        # Save to Excel
        excel_path = save_to_excel(all_extracted_data)

        # Download button
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📥 Download Extracted Excel",
                data=f,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def handle_excel_analyzer():
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

        # Final zip download
        zip_buffer.seek(0)
        st.download_button(
            label="📦 Download All Analyzed Files (ZIP)",
            data=zip_buffer,
            file_name="Analyzed_Files.zip",
            mime="application/zip"
        )

def handle_pta_services():
    st.title("PTA Services - Operator Lookup and Sorted Export")

    st.subheader("Enter Phone Numbers for Lookup and Sorting")
    phone_numbers_input = st.text_area("Enter phone numbers, one per line:", key="phone_numbers_input_single_section")

    if st.button("Process Numbers for Operators and Sorted Export", key="process_numbers_button"):
        if phone_numbers_input:
            from utils.operator_lookup import find_operators_and_download
            
            raw_numbers = [num.strip() for num in phone_numbers_input.split('\n') if num.strip()]

            phone_numbers = []
            invalid_numbers = []

            for num in raw_numbers:
                    normalized = normalize_phone_number(num)
                    if normalized:
                        phone_numbers.append(normalized)
                    else:
                        invalid_numbers.append(num)

            
            if phone_numbers:
                try:
                    # Perform operator lookup
                    # The original find_operators_and_download also saves to Excel, which is fine, but we'll sort here.
                    excel_path_dummy, lookup_results_data = find_operators_and_download(phone_numbers) 
                    
                    st.subheader("Processing Results")
                    
                    # Group results by operator
                    operator_groups = {
                        "Jazz Pakistan": [],
                        "Zong Pakistan": [],
                        "Telenor Pakistan": [],
                        "Ufone Pakistan": [],
                        "Other": []
                    }
                    
                    for item in lookup_results_data:
                        operator = item["operator"]
                        if operator in operator_groups:
                            operator_groups[operator].append(item)
                        else:
                            operator_groups["Other"].append(item)
                    
                    # Create a flat list of sorted results
                    sorted_results_list = []
                    for op_name in ["Jazz Pakistan", "Zong Pakistan", "Telenor Pakistan", "Ufone Pakistan", "Other"]:
                        sorted_results_list.extend(operator_groups[op_name])

                    # Standardize phone numbers to 92... format
                    standardized_results = []
                    for item in sorted_results_list:
                        number = item["number"]
                        # Apply '0' to '92' conversion for 11-digit numbers starting with '0'
                        if len(number) == 11 and number.startswith('0'):
                            standardized_number = '92' + number[1:]
                        else:
                            standardized_number = number # Keep as is if not matching criteria
                        standardized_results.append({
                            "Phone Number": standardized_number,
                            "Detected Operator": item["operator"]
                        })

                    # Convert to DataFrame for display and export
                    df_final_results = pd.DataFrame(standardized_results)
                    st.session_state.pta_results_for_cdr = standardized_results # Store for CDR Format section
                    st.dataframe(df_final_results)

                    # Prepare Excel for download
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        df_final_results.to_excel(writer, index=False, sheet_name='Sorted_Operator_Results')
                    excel_buffer.seek(0)

                    st.download_button(
                        label=f"📥 Download Sorted Operator Results (Excel)",
                        data=excel_buffer,
                        file_name="sorted_operator_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_sorted_operator_results"
                    )

                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
            else:
                st.warning("Please enter at least one valid phone number.")
        else:
            st.warning("Please enter phone numbers to process.")

def handle_cdr_format():
    st.title("📝 CDR Format")
    st.info("Select an HTML template to view and see numbers from PTA Services filtered by category.")

    # Move Clear Filtered Numbers button here
    st.markdown("---")
    if st.button("Clear Filtered Numbers", key="clear_filtered_numbers_top"): # New unique key
        st.session_state.pta_results_for_cdr = []
        st.session_state.selected_cdr_html = None 
        st.rerun()
    st.markdown("---")

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
            if st.button(button_label, key=f"cdr_html_button_{button_label}"): # Unique key for each button
                st.session_state.selected_cdr_html = file_info

    # The rest of the HTML display and processing logic remains the same,
    # but the clear button is no longer at the bottom.
    if st.session_state.selected_cdr_html:
        selected_file_name = st.session_state.selected_cdr_html["file"]
        target_operator_key = st.session_state.selected_cdr_html["operator_key"]

        st.markdown("---")
        st.subheader(f"Processing Numbers for {target_operator_key} using `{selected_file_name}`")

        # Get filtered numbers from PTA Services
        filtered_numbers_for_html = []
        if "pta_results_for_cdr" in st.session_state and st.session_state.pta_results_for_cdr:
            for item in st.session_state.pta_results_for_cdr:
                if target_operator_key.lower().replace(" pakistan", "") in item["Detected Operator"].lower().replace(" pakistan", ""):
                    filtered_numbers_for_html.append(item["Phone Number"])

        # Construct the numbers string for injection
        numbers_to_inject = "\n".join(filtered_numbers_for_html)

        # 1. Read the original HTML content
        try:
            with open(selected_file_name, "r", encoding="utf-8") as f:
                html_content = f.read()
            
            # 2. Dynamically inject numbers into the textarea and trigger changeFormat()
            # Use regex to find and inject content into the textarea
            import re
            
            # Define a replacement function for re.sub
            def replace_textarea_content(match):
                return match.group(1) + numbers_to_inject + match.group(3)

            modified_html_content = re.sub(
                r'(<textarea[^>]*id=["\']formatinput["\'][^>]*>)((?!</textarea>).*?)(</textarea>)',
                replace_textarea_content, # Pass the function as replacement
                html_content,
                flags=re.DOTALL | re.IGNORECASE
            )
            
            # Inject a script to call changeFormat() after the DOM is ready
            script_to_inject = """
            <script type="text/javascript">
            document.addEventListener('DOMContentLoaded', function() {
                setTimeout(function() {
                    if (typeof changeFormat === 'function') {
                        changeFormat();
                    }
                }, 100); // Small delay to ensure everything is rendered
            });
            </script>
            """
            modified_html_content = modified_html_content.replace('</body>', f'{script_to_inject}</body>')
            
            st.components.v1.html(modified_html_content, height=750, scrolling=True) # Use st.components.v1.html for better rendering
            
            if not filtered_numbers_for_html:
                st.info(f"No numbers found for '{target_operator_key}' in the PTA Services results to process. HTML displayed with empty input.")

        except FileNotFoundError:
            st.error(f"HTML file not found: `{selected_file_name}`. Please ensure it is in the root directory.")
        except Exception as e:
            st.error(f"Error reading or modifying HTML file: {e}")
        st.markdown("---") # Add a final markdown for spacing



# --- Server Management ---
import socket
from contextlib import closing

def find_free_port(start_port):
    for port in range(start_port, start_port + 100):
        with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return None

def start_server():
    """Starts the FastAPI server as a background process."""
    if "server_process" not in st.session_state:
        PORT_FILE = ".uvicorn_port"
        DEFAULT_PORT = 8000

        try:
            with open(PORT_FILE, "r") as f:
                port = int(f.read().strip())
        except (IOError, ValueError):
            port = DEFAULT_PORT
        
        free_port = find_free_port(port)
        
        if not free_port:
            st.error(f"Could not find a free port starting from {port}.")
            return

        try:
            with open(PORT_FILE, "w") as f:
                f.write(str(free_port))
        except IOError:
            # Non-critical, we can still proceed
            pass
        
        st.session_state["backend_port"] = free_port

        script_dir = os.path.dirname(os.path.abspath(__file__))
        venv_path = os.path.join(script_dir, "venv")
        
        command = []
        if os.path.exists(venv_path) and sys.platform == "win32":
            python_executable = os.path.join(venv_path, "Scripts", "python.exe")
            command = [
                python_executable, "-m", "uvicorn", "fastapi_server:app",
                "--host", "127.0.0.1", "--port", str(free_port)
            ]
        else:
            command = [
                "uvicorn", "fastapi_server:app", "--host", "0.0.0.0", "--port", str(free_port)
            ]

        try:
            st.session_state.server_process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            )
            
            time.sleep(3)

            if st.session_state.server_process.poll() is not None:
                st.error("Backend server failed to start. See logs below.")
                stdout, stderr = st.session_state.server_process.communicate()
                st.text("Server stdout:")
                st.code(stdout.decode('utf-8', errors='ignore'))
                st.text("Server stderr:")
                st.code(stderr.decode('utf-8', errors='ignore'))
                st.session_state.server_process = None

        except FileNotFoundError:
            st.error(f"Error: The command '{command[0]}' was not found.")
            st.error("Please ensure that your environment is set up correctly.")
            st.session_state.server_process = None
        except Exception as e:
            st.error(f"An unexpected error occurred while starting the server: {e}")
            st.session_state.server_process = None

def stop_server():
    """Stops the FastAPI server if it's running."""
    if "server_process" in st.session_state:
        st.session_state.server_process.terminate()
        st.session_state.server_process = None

# Start the server when the app starts
start_server()

# Register the stop_server function to be called on exit
atexit.register(stop_server)




def show_main_app():


    port = st.session_state.get("backend_port", 8000)
    SIM_BACKEND_URL = f"http://127.0.0.1:{port}/get-info/"
    VEHICLE_BACKEND_URL = f"http://127.0.0.1:{port}/get-vehicle-info/"

    st.title("🇵🇰 Information Extractor")
    st.markdown("Select the type of information you want to search for.")

    search_type = st.selectbox("Select Search Type", ["SIM Info", "Vehicle Info"])

    st.markdown("---")

    if search_type == "SIM Info":
        st.header("SIM Owner Details")

        # Option to either manually enter numbers or upload a file
        input_method = st.radio("Choose input method:", ("Manual Entry", "Upload Excel File"))

        if input_method == "Manual Entry":
            st.subheader("Multiple Search")
            search_terms = st.text_area(
                "Enter Phone Numbers / CNICs",
                placeholder="03001234567\n03111234567\n3520212345678",
                help="Enter multiple values separated by new line or comma"
            )

            if st.button("🔍 Search SIM Info"):
                raw_items = search_terms.replace(",", "\n").split("\n")
                items = [i.strip() for i in raw_items if i.strip()]

                # Add '0' to 10-digit numbers
                processed_items = []
                for item in items:
                    if len(item) == 10 and item.isdigit():
                        processed_items.append("0" + item)
                    else:
                        processed_items.append(item)
                items = processed_items

                if not items:
                    st.error("Please enter at least one phone number or CNIC.")
                else:
                    results = []
                    with st.spinner("Fetching SIM data..."):
                        for item in items:
                            if not item.isdigit() or len(item) not in (11, 13):
                                results.append({
                                    "Input": item, "Name": "Invalid", "Number": "Invalid",
                                    "CNIC": "Invalid", "Address": "Invalid", "Status": "Invalid Format"
                                })
                                continue
                            try:
                                response = requests.post(f"{SIM_BACKEND_URL}?phone_number={item}", timeout=10)
                                data = response.json()
                                if isinstance(data, dict) and data.get("error"):
                                    results.append({
                                        "Input": item, "Name": "", "Number": "", "CNIC": "",
                                        "Address": "", "Status": data["error"]
                                    })
                                elif isinstance(data, list):
                                    for record in data:
                                        results.append({
                                            "Input": item, "Name": record.get("name", ""), "Number": record.get("number", ""),
                                            "CNIC": record.get("cnic", ""), "Address": record.get("address", ""), "Status": "Found"
                                        })
                                else:
                                    results.append({
                                        "Input": item, "Name": "", "Number": "", "CNIC": "",
                                        "Address": "", "Status": "Unexpected response format"
                                    })
                            except Exception as e:
                                results.append({
                                    "Input": item, "Name": "", "Number": "", "CNIC": "",
                                    "Address": "", "Status": f"Request Failed: {e}"
                                })
                    df = pd.DataFrame(results)
                    st.success(f"Results found: {len(df)}")
                    st.dataframe(df, use_container_width=True)
                    excel_buffer = BytesIO()
                    df.to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)
                    st.download_button(
                        label="⬇️ Download Excel", data=excel_buffer,
                        file_name="sim_info_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        elif input_method == "Upload Excel File":
            st.subheader("Upload and Process Excel File")
            uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

            if uploaded_file is not None:
                try:
                    # Read all sheets from the uploaded Excel file
                    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)

                    if "Mobile Numbers" not in all_sheets:
                        st.error("The Excel file must contain a 'Mobile Numbers' sheet.")
                    else:
                        df_upload = all_sheets["Mobile Numbers"]
                        
                        if "Mobile Number" not in df_upload.columns:
                            st.error("The 'Mobile Numbers' sheet must contain a 'Mobile Number' column.")
                        else:
                            # Clean the 'Mobile Number' column
                            df_upload['Mobile Number'] = df_upload['Mobile Number'].astype(str).str.replace(r'\.0$', '', regex=True)
                            df_upload['Mobile Number'] = df_upload['Mobile Number'].apply(lambda x: '0' + x if len(x) == 10 and x.isdigit() else x)
                            
                            st.write("Original Data in 'Mobile Numbers' sheet:")
                            st.dataframe(df_upload)

                            if st.button("🚀 Fetch Data and Create New Sheet"):
                                items = (
                                df_upload["Mobile Number"]
                                .dropna()
                                .head(10)
                                .tolist()
                        )
                                
                                if not items:
                                    st.error("No phone numbers found in the 'Mobile Number' column.")
                                else:
                                    table_placeholder = st.empty()
                                    all_results = []
                                    
                                    with st.spinner("Fetching data for all numbers..."):
                                        for item in items:
                                            try:
                                                response = requests.post(f"{SIM_BACKEND_URL}?phone_number={item}", timeout=10)
                                                data = response.json()
                                                if isinstance(data, list) and data:
                                                    # If multiple records are found, append them as separate rows
                                                    for record in data:
                                                        all_results.append({
                                                            "Original Phone Number": item,
                                                            "Name": record.get("name"),
                                                            "Number": record.get("number"),
                                                            "CNIC": record.get("cnic"),
                                                            "Address": record.get("address")
                                                        })
                                                else:
                                                    all_results.append({
                                                        "Original Phone Number": item, "Name": "Not Found",
                                                        "Number": "", "CNIC": "", "Address": ""
                                                    })
                                            except Exception:
                                                all_results.append({
                                                    "Original Phone Number": item, "Name": "Request Failed",
                                                    "Number": "", "CNIC": "", "Address": ""
                                                })
                                            
                                            df_results_live = pd.DataFrame(all_results)
                                            table_placeholder.dataframe(df_results_live)

                                    df_results = pd.DataFrame(all_results)
                                    
                                    # Prepend space to 'Number' and 'CNIC' to prevent scientific notation
                                    if "Number" in df_results.columns:
                                        df_results["Number"] = df_results["Number"].astype(str).apply(lambda x: ' ' + x)
                                    if "CNIC" in df_results.columns:
                                        df_results["CNIC"] = df_results["CNIC"].astype(str).apply(lambda x: ' ' + x)
                                    
                                    # Add the results as a new sheet
                                    all_sheets["Search Results"] = df_results

                                    st.success("Data fetching complete!")
                                    st.write("New 'Search Results' sheet created with the fetched data.")
                                    
                                    # Write all sheets to a new Excel file in memory
                                    excel_buffer_updated = BytesIO()
                                    with pd.ExcelWriter(excel_buffer_updated, engine='xlsxwriter') as writer:
                                        for sheet_name, df in all_sheets.items():
                                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                                            # Auto-adjust columns
                                            worksheet = writer.sheets[sheet_name]
                                            for i, col in enumerate(df.columns):
                                                max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
                                                worksheet.set_column(i, i, max_len)

                                    
                                    excel_buffer_updated.seek(0)
                                    
                                    st.download_button(
                                        label="⬇️ Download Updated Excel File",
                                        data=excel_buffer_updated,
                                        file_name="updated_sim_info_with_results.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                except Exception as e:
                    st.error(f"An error occurred: {e}")

    elif search_type == "Vehicle Info":
        st.header("Sindh Vehicle Details")
        vehicle_category = st.selectbox("Select Vehicle Category", ["", "2 wheeler", "4 wheeler"])
        reg_number = st.text_input("Enter Registration Number", placeholder="ABC-123")
        if st.button("🔍 Search Vehicle Info"):
            if not reg_number or not vehicle_category:
                st.error("Please select category and enter registration number.")
            else:
                category_map = {"2 wheeler": "2W", "4 wheeler": "4W"}
                api_category = category_map[vehicle_category]
                with st.spinner("Fetching vehicle data..."):
                    try:
                        response = requests.post(
                            f"{VEHICLE_BACKEND_URL}?reg_no={reg_number}&category={api_category}",
                            timeout=10
                        )
                        data = response.json()
                        if data.get("error"):
                            st.error(data["error"])
                        else:
                            vehicle_data = {
                                "Attribute": [
                                    "Registration Number", "Owner Name", "Owner CNIC", "Model",
                                    "Model Year", "Color", "Engine Number", "Chassis Number",
                                    "Registration Date", "CPLC Status", "District", "Branch"
                                ],
                                "Value": [
                                    data.get("registrationNumber"), data.get("ownerName"), data.get("ownerCNIC"),
                                    f"{data.get('manufacturerName', '')} {data.get('modelName', '')}",
                                    data.get("modelYear"), data.get("color"), data.get("engineNumber"),
                                    data.get("chassisNumber"), data.get("registrationDate"),
                                    data.get("cplcStatus"), data.get("districtName"), data.get("branchName"),
                                ]
                            }
                            df = pd.DataFrame(vehicle_data)
                            st.table(df)
                    except Exception as e:
                        st.error(f"Server loaded try again later")
    st.markdown("""
    ---
    *Disclaimer: This tool uses third-party public APIs. Data accuracy is not guaranteed.*
    """)



# ---------- Rate Limiting ----------
REQUEST_LIMIT = 10   # max 10 requests
TIME_WINDOW = 60     # in seconds (1 min)
request_times = []   # store timestamps of last requests

def check_rate_limit():
    """Ensure only 10 requests per minute are sent."""
    global request_times
    now = time.time()
    # Keep only timestamps from the last 60s
    request_times = [t for t in request_times if now - t < TIME_WINDOW]

    if len(request_times) >= REQUEST_LIMIT:
        wait_time = TIME_WINDOW - (now - request_times[0])
        st.warning(f"⏳ Rate limit reached! Waiting {int(wait_time)}s before next request...")
        time.sleep(wait_time)

    # Add current request timestamp
    request_times.append(time.time())

def handle_settings():
    st.title("Settings / Future Tools")
    st.info("This section is under construction.")

# ---------- Main App ----------
def main():
    if not auth.is_logged_in():
        auth.login()
        return

    with st.sidebar:
        st.image("Assets/app_icon.png", width=100)
        st.markdown("### Menu")
        
        user_services = auth.get_user_services()
        
        service_map = {
            "Application Extractor": "app",
            "Excel Analyzer": "analyzer",
            "PTA Services": "pta_services",
            "CDR Format": "cdr_format",
            "Vehicle and Mobile": "vehicle_and_mobile",
            "Admin": "admin",
            "Settings / Future Tools": "settings"
        }

        service_icons = {
            "Application Extractor": "📝",
            "Excel Analyzer": "📈",
            "PTA Services": "📞",
            "CDR Format": "📄",
            "Vehicle and Mobile": "🔍",
            "Admin": "⚙️",
            "Settings / Future Tools": "🔮"
        }

        for service in user_services:
            if service in service_map:
                icon = service_icons.get(service, "➡️")
                if st.button(f"{icon} {service}"):
                    st.session_state.page = service_map[service]

        if st.button("Logout"):
            auth.logout()

    page = st.session_state.get("page", "app")

    if page == "app":
        handle_application_extractor()
    elif page == "analyzer":
        handle_excel_analyzer()
    elif page == "pta_services":
        handle_pta_services()
    elif page == "cdr_format":
        handle_cdr_format()
    elif page == "vehicle_and_mobile":
        show_main_app()
    elif page == "admin":
        auth.admin_section()
    elif page == "settings":
        handle_settings()

if __name__ == "__main__":
    main()
