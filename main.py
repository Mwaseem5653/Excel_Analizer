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

# Load Gemini API key
load_dotenv()
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

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
    if st.button("⚙️ Settings / Future Tools"):
        st.session_state.page = "settings"

# ---------- Rate Limiting Setup ----------
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
                """Name: Furqan Ur Rehman
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
    st.info("Upload an Excel file to analyze Mobile Numbers & Addresses.")

    uploaded_excel = st.file_uploader("Upload Excel file", type=["xlsx","csv"], key="analyzer_uploader")

    if uploaded_excel:
        temp_path = os.path.join("temp_uploads", uploaded_excel.name)
        os.makedirs("temp_uploads", exist_ok=True)
        with open(temp_path, "wb") as f:
            f.write(uploaded_excel.getbuffer())

        try:
            st.write("⏳ Thora intezaar karein... Aapki file par kaam ho raha hai. Shukriya 😊")

            analyzed_path = analyze_excel(temp_path)
            st.success("✅ Excel analyzed successfully!")
            download_file_name = uploaded_excel.name
            with open(analyzed_path, "rb") as f:
                st.download_button(
                    label="📥 Download Analyzed Excel",
                    data=f,
                    file_name="(Analized)-"+download_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

# -------------------- Settings / Future Tools --------------------
elif st.session_state.page == "settings":
    st.title("⚙️ Settings / Future Tools")
    st.info("Allah Pak Ka Huqam howa to yaha see or age qam karenge.")
