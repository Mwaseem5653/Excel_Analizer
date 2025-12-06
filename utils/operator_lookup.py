import httpx
import openpyxl
import os

# --- Default API for direct access ---
# This will be used if no local server is detected.
# The external API URL without the /proxy part
EXTERNAL_API_BASE_URL = "https://easyload.com.pk/dingconnect.php"

# Get the FastAPI server URL from an environment variable.
# Use localhost as a default for local development.
LOCAL_API_URL = os.getenv("API_URL") # Try to get from environment variable first

if not LOCAL_API_URL: # If not set in environment, try to read from .uvicorn_port file
    uvicorn_port_file = ".uvicorn_port"
    if os.path.exists(uvicorn_port_file):
        try:
            with open(uvicorn_port_file, "r") as f:
                port = int(f.read().strip())
                LOCAL_API_URL = f"http://127.0.0.1:{port}"
                print(f"Using LOCAL_API_URL from .uvicorn_port file: {LOCAL_API_URL}")
        except (ValueError, IOError) as e:
            print(f"Error reading .uvicorn_port: {e}. Falling back to external API.")
    
# Determine the effective API_URL for this request
# Prefer local server if available, otherwise use external API directly
current_api_base = LOCAL_API_URL if LOCAL_API_URL else EXTERNAL_API_BASE_URL

def find_operators_and_download(phone_numbers):
    results = []
    with httpx.Client() as client:
        for number in phone_numbers:
            account_number = number
            if account_number.startswith('03'):
                account_number = '92' + account_number[1:]
            account_number = ''.join(filter(str.isdigit, account_number))
            
            print(f"Processing number: {number}")
            print(f"Formatted account number: {account_number}")
            
            # Construct the request URL based on whether we are using local proxy or direct external API
            if LOCAL_API_URL and current_api_base == LOCAL_API_URL:
                request_url = f"{current_api_base}/proxy?accountNumber={account_number}"
            else:
                # Direct call to easyload.com.pk, which already has dingconnect.php
                request_url = f"{EXTERNAL_API_BASE_URL}?action=GetProviders&accountNumber={account_number}"

            print(f"API Request URL: {request_url}")

            try:
                response = client.get(request_url)
                response.raise_for_status()
                data = response.json()
                print(f"API Response Status Code: {response.status_code}")
                print(f"API Response Body: {response.text}")
                
                operator = 'Not Found'
                if data and data.get("Items") and len(data["Items"]) > 0:
                    operator = data["Items"][0].get("Name", "Not Found")
                
                results.append({"number": number, "operator": operator})
            except (httpx.HTTPStatusError, httpx.RequestError) as e:
                print(f"Failed for {number}: {e}")
                results.append({"number": number, "operator": "Error"})

    # Create Excel file
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Operator Results"
    
    # Add headers
    sheet.append(["Phone Number", "Operator"])
    
    # Add data
    for result in results:
        sheet.append([result["number"], result["operator"]])
        
    # Save the workbook
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, "operator_results.xlsx")
    workbook.save(excel_path)
    
    return excel_path, results
