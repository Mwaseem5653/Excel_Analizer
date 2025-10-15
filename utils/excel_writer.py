import pandas as pd

def save_to_excel(all_data):
    """
    Save a list of dictionaries to Excel
    """
    # Ensure list of dicts
    df = pd.DataFrame(all_data)
    
    # Optional: clean IMEI column
    if "IMEI Number" in df.columns:
        df["IMEI Number"] = df["IMEI Number"].astype(str).str.strip()
    
    excel_path = "extracted_data.xlsx"
    df.to_excel(excel_path, index=False)
    return excel_path
