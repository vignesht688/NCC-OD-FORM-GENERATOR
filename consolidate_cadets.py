import os
import re
import json

log_file_path = r"C:\Users\jayav\.gemini\antigravity-ide\brain\b0067744-fd02-4ccf-8e27-b9dca0c0bba3\.system_generated\tasks\task-356.log"
output_db_path = r"c:\Users\jayav\OneDrive\Desktop\od form for ncc\cadets.json"

def clean_shift(shift_str):
    s = (shift_str or "").strip().upper()
    if not s:
        return ""
    if "II" in s or "2" in s or "-II" in s:
        return "SHIFT-II"
    if "I" in s or "1" in s or "-I" in s:
        return "SHIFT-I"
    return ""

def is_valid_cadet_name(name):
    name = name.strip().upper()
    if not name or len(name) < 3:
        return False
        
    exclusions = [
        "S.NO", "REGISTER NO", "NAME", "DEPARTMENT", "SHIFT", "PARTICULARS", "ON DUTY",
        "B COM", "B.A ENGLISH", "BBA", "BBA CA", "BSC CS", "BCOM", "BCA", "BA", "BSC",
        "NCC SELECTION", "NCC", "CADETS", "TOTAL", "S.NO.", "REG NO", "REG.NO"
    ]
    if name in exclusions:
        return False
    
    # Exclude headings starting with I, II, III (e.g. "I BCOM", "II BSC CS", "III BCA")
    if re.match(r'^(I|II|III|IV)\s+(BCOM|BSC|BBA|BCA|BA|B\.COM|B\.SC|B|A|I)\b', name):
        return False

    # Exclude standalone department prefixes
    dep_codes = ["BCOM", "BSC", "BBA", "BCA", "BA", "B.COM", "B.SC", "B.B.A", "B.C.A", "B.A."]
    for code in dep_codes:
        if name == code or name.startswith(code + " ") or name.startswith(code + "."):
            return False
            
    if name.isdigit():
        return False
        
    return True

def main():
    if not os.path.exists(log_file_path):
        print(f"Log file not found: {log_file_path}")
        return

    cadets_map = {} # key -> cadet dict

    with open(log_file_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line_clean = line.strip()
            if not line_clean or "S.NO | REGISTER" in line_clean or "S.NO" in line_clean:
                continue
                
            # Skip logs
            if "=== FILE:" in line_clean or "Tables count:" in line_clean:
                continue

            parts = [p.strip() for p in line_clean.split("|")]
            if len(parts) < 3:
                continue
            
            # S.NO | REGISTER NO | NAME | DEPARTMENT | SHIFT | PARTICULARS | ON DUTY
            # parts[0] is S.NO (or empty if leading pipe)
            # parts[1] is REGISTER NO
            # parts[2] is NAME
            # parts[3] is DEPARTMENT
            # parts[4] is SHIFT
            reg_no = parts[1].strip() if len(parts) > 1 else ""
            name = parts[2].strip() if len(parts) > 2 else ""
            dept = parts[3].strip() if len(parts) > 3 else ""
            shift = parts[4].strip() if len(parts) > 4 else ""

            # Check if name is valid
            if not is_valid_cadet_name(name):
                continue

            # Normalization
            name = re.sub(r'\s+', ' ', name.upper())
            reg_no = re.sub(r'\D', '', reg_no) # digits only
            dept = re.sub(r'\s+', ' ', dept.upper())
            shift = clean_shift(shift)

            # Deduplication key
            key = reg_no if reg_no else name

            cadet = {
                "register_no": reg_no,
                "name": name,
                "department": dept,
                "shift": shift
            }

            if key in cadets_map:
                existing = cadets_map[key]
                if not existing["register_no"] and reg_no:
                    existing["register_no"] = reg_no
                if not existing["shift"] and shift:
                    existing["shift"] = shift
                if not existing["department"] and dept:
                    existing["department"] = dept
                if len(name) > len(existing["name"]):
                    existing["name"] = name
            else:
                cadets_map[key] = cadet

    # Convert to list and sort alphabetically by name
    cadets_list = list(cadets_map.values())
    cadets_list.sort(key=lambda c: c["name"])

    # Overwrite cadets.json
    with open(output_db_path, "w", encoding="utf-8") as out:
        json.dump(cadets_list, out, indent=2)

    print(f"Successfully processed {len(cadets_list)} unique cadets and saved to database.")

if __name__ == "__main__":
    main()
