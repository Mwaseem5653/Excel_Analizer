import pandas as pd
import re

def read_excel_auto(file_path):
    """
    Automatically detects where the table starts,
    cleans headers, and formats numeric columns safely.
    """
    # Step 1: Read full sheet without header
    df_raw = pd.read_excel(file_path, header=None)

    # Step 2: Find the header row automatically
    header_row = None
    known_keywords = ['call', 'type', 'msisdn', 'bnumber', 'a number', 'imei', 'start', 'end']
    for i, row in df_raw.iterrows():
        values = [str(v).lower() for v in row if pd.notna(v)]
        if len(values) > 2 and any(any(k in v for k in known_keywords) for v in values):
            header_row = i
            break

    if header_row is None:
        raise ValueError("Table header not found automatically!")

    # Step 3: Re-read with correct header
    df = pd.read_excel(file_path, header=header_row)

    # Step 4: Clean column names
    df.columns = df.columns.astype(str).str.strip().str.lower()

    # Step 5: Detect A/B number columns
    a_col = next((c for c in df.columns if "a number" in c), None)
    b_col = next((c for c in df.columns if "b number" in c or "bnumber" in c), None)

    # Step 6: Normalize mobile numbers
    def normalize_mobile(num):
        if pd.isna(num): return None
        num = re.sub(r"\D", "", str(num))
        if num.startswith("+92"): num = num[3:]
        elif num.startswith("92"): num = num[2:]
        elif num.startswith("0"): num = num[1:]
        if re.fullmatch(r"3\d{9}", num):
            return " " + num  # add space for Excel safety
        return None

    if a_col: df[a_col] = df[a_col].apply(normalize_mobile)
    if b_col: df[b_col] = df[b_col].apply(normalize_mobile)

    # Step 7: Add space to long numeric values (IMEI etc.)
    for col in df.columns:
        try:
            values = df[col].dropna().astype(str)
            if values.str.match(r"^\d+$").mean() > 0.8:
                df[col] = df[col].apply(lambda x: f" {x}" if pd.notna(x) else x)
        except:
            pass

    return df

# Example:
# df = read_excel_auto("call_data.xlsx")
# df.to_excel("cleaned_output.xlsx", index=False)
