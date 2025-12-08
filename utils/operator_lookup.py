import httpx
import openpyxl
import os
import subprocess
import time
import requests

# --- Default API for direct access ---
# This will be used if no local server is detected.
# The external API URL without the /proxy part
EXTERNAL_API_BASE_URL = "https://easyload.com.pk/dingconnect.php"

# Get the FastAPI server URL from an environment variable.
# Use localhost as a default for local development.
LOCAL_API_URL = os.getenv("API_URL") # Try to get from environment variable first

FASTAPI_PROCESS = None # Global variable to hold the FastAPI subprocess

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

def start_fastapi_server_if_not_running():
    global FASTAPI_PROCESS, LOCAL_API_URL, current_api_base
    
    # If the process is already running (or we think it is), check its responsiveness
    if FASTAPI_PROCESS and FASTAPI_PROCESS.poll() is None:
        print("FastAPI server process seems to be running.")
        if LOCAL_API_URL: # Only check responsiveness if we have a known LOCAL_API_URL
            try:
                response = requests.get(f"{LOCAL_API_URL}/", timeout=1)
                if response.status_code == 200:
                    print(f"FastAPI server still responsive at {LOCAL_API_URL}. No action needed.")
                    current_api_base = LOCAL_API_URL
                    return
            except requests.exceptions.RequestException:
                print(f"FastAPI server process found but not responsive at {LOCAL_API_URL}. Will attempt to restart.")
                # Mark for restart
                FASTAPI_PROCESS = None 
                LOCAL_API_URL = None
        else: # Process is running, but LOCAL_API_URL is not set, meaning it's in an inconsistent state
            print("FastAPI process found but LOCAL_API_URL is not set. Assuming unresponsive and restarting.")
            FASTAPI_PROCESS = None
            LOCAL_API_URL = None

    # Check if a server is already responsive at a known LOCAL_API_URL
    if LOCAL_API_URL:
        try:
            response = requests.get(f"{LOCAL_API_URL}/", timeout=1)
            if response.status_code == 200:
                print(f"FastAPI server already responsive at {LOCAL_API_URL}")
                current_api_base = LOCAL_API_URL
                return
        except requests.exceptions.RequestException:
            print(f"FastAPI server not responsive at {LOCAL_API_URL}. Attempting to start...")
            LOCAL_API_URL = None # Reset if not responsive

    # If no responsive server, attempt to start fastapi_server.py
    if not LOCAL_API_URL:
        print("Attempting to start fastapi_server.py as a background process...")
        port_file_path = ".uvicorn_port"
        # Clean up any old .uvicorn_port file before starting
        if os.path.exists(port_file_path):
            print(f"Deleting stale {port_file_path}...")
            os.remove(port_file_path)
            
        try:
            # Start the server as a detached process
            # Use a platform-specific command to detach the process
            creation_flags = 0
            preexec_fn = None
            if os.name == 'nt':  # Windows
                creation_flags = subprocess.DETACHED_PROCESS
            else:  # Unix/Linux/macOS
                preexec_fn = os.setsid

            FASTAPI_PROCESS = subprocess.Popen(
                ["python", "fastapi_server.py"],
                stdout=subprocess.DEVNULL, # Redirect stdout to devnull
                stderr=subprocess.DEVNULL, # Redirect stderr to devnull
                stdin=subprocess.DEVNULL,  # Redirect stdin to devnull
                close_fds=True,            # Close file descriptors in child process
                shell=False,
                creationflags=creation_flags,
                preexec_fn=preexec_fn
            )
            print(f"FastAPI server started with PID: {FASTAPI_PROCESS.pid}")

            # Wait for .uvicorn_port file to be created and read the port
            port_file_path = ".uvicorn_port"
            retries = 10
            while retries > 0:
                if os.path.exists(port_file_path):
                    try:
                        with open(port_file_path, "r") as f:
                            port = int(f.read().strip())
                            LOCAL_API_URL = f"http://127.0.0.1:{port}"
                            current_api_base = LOCAL_API_URL
                            print(f"FastAPI server started successfully on {LOCAL_API_URL}")
                            return
                    except (ValueError, IOError) as e:
                        print(f"Error reading .uvicorn_port: {e}. Retrying...")
                time.sleep(1) # Wait for 1 second
                retries -= 1
            
            print("Failed to start FastAPI server or read its port after multiple retries. Falling back to external API.")
            LOCAL_API_URL = None # Fallback to external if starting fails

        except Exception as e:
            print(f"Error starting fastapi_server.py: {e}")
            FASTAPI_PROCESS = None
            LOCAL_API_URL = None # Fallback to external if starting fails
    
    # Update current_api_base at the very end to reflect the final decision
    current_api_base = LOCAL_API_URL if LOCAL_API_URL else EXTERNAL_API_BASE_URL
        
def find_operators_and_download(phone_numbers):
    start_fastapi_server_if_not_running()
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
