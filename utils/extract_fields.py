import re

def extract_fields_from_text(text):
    fields = {
        "Name": "",
        "Phone Number": "",
        "Police Station": "",
        "Other Property":"",
        "Mobile Model": "",
        "Type": "",
        "Date Of Offence": "",
        "Time Of Offence": "",
        "IMEI Number": "",
        "last Num Used": "",
        
        
        
        
        
    }
    
    
    lines = text.splitlines()

    for line in lines:
        line = line.strip()
        
        if re.match(r"(?i).*name[:：]", line):
            fields["Name"] = line.split(":", 1)[-1].strip()

        elif re.match(r"(?i).*Police Station[:：]", line):
            fields["Police Station"] = line.split(":", 1)[-1].strip()
        
        elif re.match(r"(?i).*Other Property[:：]", line):
            fields["Other Property"] = line.split(":", 1)[-1].strip()

        elif re.match(r"(?i).*last Num Used[:：]",line):
            fields["last Num Used"] = line.split(":" , 1)[-1].strip()

        elif re.match(r"(?i).*mobile model[:：]", line):
            fields["Mobile Model"] = line.split(":", 1)[-1].strip()

        # ✅ Updated IMEI logic: multiple IMEIs, single space, no Excel E+14 issue
        elif re.match(r"(?i).*imei number[:：]", line):
            imeis = re.findall(r"\b\d{14,17}\b", line)
            if imeis:
                safe_imeis = [f"{imei}" for imei in imeis]  # Prepend ' to stop Excel from converting
                fields["IMEI Number"] = " ".join(safe_imeis)

        elif re.match(r"(?i).*(phone|contact) number[:：]", line):
            fields["Phone Number"] = line.split(":", 1)[-1].strip()

        elif re.match(r"(?i).*Date Of Offence[:：]", line):
            fields["Date Of Offence"] = line.split(":", 1)[-1].strip()

        elif re.match(r"(?i).*Time Of Offence[:：]", line):
            match = re.search(r"(?i)Time Of Offence[:：]?\s*(\d{1,2}:\d{2}(?:\s?[APMapm]{2})?)", line)
            if match:
                fields["Time Of Offence"] = match.group(1).strip()

        elif re.match(r"(?i).*type[:：]", line):
            fields["Type"] = line.split(":", 1)[-1].strip()

    return fields
