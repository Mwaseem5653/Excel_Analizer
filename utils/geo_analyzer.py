import pandas as pd
import re
from datetime import datetime

def normalize_geo_number(num):
    """Normalize phone number to 10 digits starting with 3 (e.g., 3001234567)."""
    if pd.isna(num) or num == "":
        return None
    num = str(num).strip().replace(".0", "")
    num = re.sub(r'\D', '', num)
    if num.startswith("92"):
        num = num[2:]
    if num.startswith("0"):
        num = num[1:]
    if len(num) == 10 and num.startswith("3"):
        return num
    return num 

def parse_datetime(dt_val):
    """Robustly parse date and time including Excel serial numbers."""
    if pd.isna(dt_val):
        return None
    
    # Handle Excel Serial Numbers (e.g., 45979.0014)
    if isinstance(dt_val, (int, float)):
        try:
            return pd.to_datetime(dt_val, unit='D', origin='1899-12-30')
        except:
            pass

    # If already a datetime-like object
    if isinstance(dt_val, (datetime, pd.Timestamp)):
        return dt_val

    dt_str = str(dt_val).strip()
    if not dt_str or dt_str.lower() == "nan":
        return None
    
    # Check if string is actually a numeric serial number
    if dt_str.replace('.','',1).isdigit():
        try:
            return pd.to_datetime(float(dt_str), unit='D', origin='1899-12-30')
        except:
            pass

    # Try specific formats
    formats = [
        "%m/%d/%Y %I:%M:%S %p", 
        "%d/%m/%Y %I:%M:%S %p",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%m-%d-%Y %I:%M:%S %p",
        "%d-%m-%Y %I:%M:%S %p",
        "%d-%b-%y %H.%M.%S",
        "%Y/%m/%d %H:%M:%S",
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(dt_str, fmt)
        except ValueError:
            continue
    
    try:
        res = pd.to_datetime(dt_str, errors='coerce')
        return res if not pd.isna(res) else None
    except:
        return None

def analyze_geo_fencing_data(file_path, start_time_str, end_time_str, include_b=True):
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    
    # Standardize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Keeping your updated aliases exactly as they were
    aliases = {
        'A_NUM': ['DLD_NO', 'MSISDN', 'A-Party', 'A_NUMBER', 'ORIGINATING_NUM', 'DLD NO', 'PHONE', 'NUMBER', 'MSISDN_A'],
        'B_NUM': ['DLG_NO', 'B-Party', 'RECEIVER', 'B_NUMBER', 'TERMINATING_NUM', 'DLG NO', 'MSISDN_B'],
        'STR_TM': ['Date And Time', 'START_TIME', 'CALL_TIME', 'DATETIME', 'STR TM', 'TIME', 'STRT_TM', 'CALL_START_DT_TM', 'DATE_TIME']
    }
    
    col_map = {}
    for standard_name, possible_names in aliases.items():
        for col in df.columns:
            if col.upper() in [p.upper() for p in possible_names]:
                col_map[standard_name] = col
                break
    
    if 'A_NUM' not in col_map or 'STR_TM' not in col_map:
        raise ValueError(f"Required columns not found. Please ensure your file has 'Number' and 'Time' columns. Columns found: {list(df.columns)}")
    
    # Process the single combined column
    df['parsed_dt'] = df[col_map['STR_TM']].apply(parse_datetime)
    df = df.dropna(subset=['parsed_dt'])
    
    df['A_NORM'] = df[col_map['A_NUM']].apply(normalize_geo_number)
    
    has_b_col = 'B_NUM' in col_map
    if has_b_col:
        df['B_NORM'] = df[col_map['B_NUM']].apply(normalize_geo_number)
    
    df = df.sort_values(by='parsed_dt', ascending=True)
    
    def get_time_obj(t_input):
        if hasattr(t_input, 'hour'): return t_input
        t_str = str(t_input).strip().upper()
        for fmt in ["%I:%M %p", "%I:%M%p", "%H:%M"]:
            try:
                return datetime.strptime(t_str, fmt).time()
            except ValueError:
                continue
        return datetime.strptime("00:00", "%H:%M").time()

    s_time = get_time_obj(start_time_str)
    e_time = get_time_obj(end_time_str)

    # Filter based on time of day
    window_mask = (df['parsed_dt'].dt.time >= s_time) & (df['parsed_dt'].dt.time <= e_time)
    
    if has_b_col and include_b:
        window_df = df[window_mask].drop_duplicates(subset=['A_NORM', 'B_NORM'])
    else:
        window_df = df[window_mask].drop_duplicates(subset=['A_NORM'])
    
    if window_df.empty:
        return None, f"No records found between {start_time_str} and {end_time_str}."

    results = []

    for _, row_win in window_df.iterrows():
        a_num = row_win['A_NORM']
        
        a_mask = (df['A_NORM'] == a_num) | (df['B_NORM'] == a_num) if has_b_col else (df['A_NORM'] == a_num)
        a_hist = df[a_mask]
        
        if not a_hist.empty:
            a_f = a_hist.iloc[0]
            a_l = a_hist.iloc[-1]
            
            row_data = {
                'A Number': a_num,
                'A Date': a_f['parsed_dt'].strftime('%d/%m/%Y'),
                'A First Time': a_f['parsed_dt'].strftime('%I:%M:%S %p'),
                'A Last Time': a_l['parsed_dt'].strftime('%I:%M:%S %p'),
                'A Count': len(a_hist)
            }
            
            if has_b_col and include_b:
                b_num = row_win['B_NORM']
                if pd.notna(b_num) and b_num != "":
                    b_hist = df[(df['A_NORM'] == b_num) | (df['B_NORM'] == b_num)]
                    if not b_hist.empty:
                        b_f = b_hist.iloc[0]
                        b_l = b_hist.iloc[-1]
                        row_data.update({
                            'B Number': b_num,
                            'B Date': b_f['parsed_dt'].strftime('%d/%m/%Y'),
                            'B First Time': b_f['parsed_dt'].strftime('%I:%M:%S %p'),
                            'B Last Time': b_l['parsed_dt'].strftime('%I:%M:%S %p'),
                            'B Count': len(b_hist)
                        })
                else:
                    row_data.update({ 'B Number': 'None', 'B Date': '', 'B First Time': '', 'B Last Time': '', 'B Count': 0 })

            results.append(row_data)
    
    res_df = pd.DataFrame(results)
    if not res_df.empty:
        res_df = res_df.sort_values(by='A First Time')
    
    return res_df, None