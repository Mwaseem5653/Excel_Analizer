import httpx
import openpyxl
import os

# --- Default API for direct access ---
EXTERNAL_API_BASE_URL = "https://easyload.com.pk/dingconnect.php"

def find_operators_and_download(phone_numbers):
    results = []
    with httpx.Client() as client:
        for number in phone_numbers:
            account_number = number
            if account_number.startswith('03'):
                account_number = '92' + account_number[1:]
            account_number = ''.join(filter(str.isdigit, account_number))
            
            # Direct call to easyload.com.pk
            request_url = f"{EXTERNAL_API_BASE_URL}?action=GetProviders&accountNumber={account_number}"

            try:
                response = client.get(request_url, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                operator = 'Not Found'
                if data and data.get("Items") and len(data["Items"]) > 0:
                    operator = data["Items"][0].get("Name", "Not Found")
                
                results.append({"number": number, "operator": operator})
            except Exception as e:
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
    output_dir = "temp_uploads"
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, "operator_results.xlsx")
    workbook.save(excel_path)
    
    return excel_path, results