import httpx
import openpyxl
import os

# --- Default API for direct access ---
EXTERNAL_API_BASE_URL = "https://easyload.com.pk/dingconnect.php"

def local_lookup_operators(phone_numbers):
    results = []
    for number in phone_numbers:
        # Standardize for prefix check (ensure starts with 03)
        prefix_num = number
        if prefix_num.startswith('92'):
            prefix_num = '0' + prefix_num[2:]
        
        operator = "Other"
        
        if len(prefix_num) >= 4:
            code = prefix_num[:4] # e.g., 0300
            
            # Jazz: 0300-0309, 0320-0329
            if code.startswith('030') or code.startswith('032'):
                operator = "Jazz Pakistan"
            # Zong: 0310-0319, 0370
            elif code.startswith('031') or code == '0370':
                operator = "Zong Pakistan"
            # Ufone: 0330-0339
            elif code.startswith('033'):
                operator = "Ufone Pakistan"
            # Telenor: 0340-0349
            elif code.startswith('034'):
                operator = "Telenor Pakistan"
                
        results.append({"number": number, "operator": operator})
    
    # Create Excel file (Same logic as online)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Operator Results"
    sheet.append(["Phone Number", "Operator"])
    for result in results:
        sheet.append([result["number"], result["operator"]])
        
    output_dir = "temp_uploads"
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, "operator_results_local.xlsx")
    workbook.save(excel_path)
    
    return excel_path, results

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