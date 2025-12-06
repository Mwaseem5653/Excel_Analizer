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
import pandas as pd # Added for DataFrame handling

# Load Gemini API key
load_dotenv()
genai.configure(api_key=os.getenv("gemini_keys"))

# ---------- Page Config ----------
st.set_page_config(page_title="Urdu Police App & Excel Analyzer", layout="wide")

# ---------- Session State ----------
if "page" not in st.session_state:
    st.session_state.page = "app"  # default page

# ---------- Sidebar ----------
with st.sidebar:
    st.image("Assets/app_icon.png", width=100)
    st.markdown("### Menu")

    if st.button("📝 Application Extractor"):
        st.session_state.page = "app"
    if st.button("📈 Excel Analyzer"):
        st.session_state.page = "analyzer"
    if st.button("PTA Services"):
        st.session_state.page = "pta_services"
    if st.button("⚙️ Settings / Future Tools"):
        st.session_state.page = "settings"

# ---------- Page Logic ----------

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

# ---------- Page Logic ----------

# -------------------- Application Extractor --------------------
if st.session_state.page == "app":
    st.title("📝 Urdu Police Application Extractor")
    
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

            model = genai.GenerativeModel("gemini-2.0-flash")

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

# -------------------- Excel Analyzer --------------------
elif st.session_state.page == "analyzer":
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

elif st.session_state.page == "pta_services":
    st.title("PTA Services - Operator Lookup and Sorted Export")

    st.subheader("Enter Phone Numbers for Lookup and Sorting")
    phone_numbers_input = st.text_area("Enter phone numbers, one per line:", key="phone_numbers_input_single_section")

    if st.button("Process Numbers for Operators and Sorted Export", key="process_numbers_button"):
        if phone_numbers_input:
            from utils.operator_lookup import find_operators_and_download
            
            phone_numbers = [num.strip() for num in phone_numbers_input.split('\n') if num.strip()]
            
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



# -------------------- Settings / Future Tools --------------------
elif st.session_state.page == "settings":
    st.title("⚙️ Settings / Future Tools")
    st.info("Allah Pak Ka Huqam howa to yaha see or age qam karenge.")
