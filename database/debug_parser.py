import re
import os

log_file_path = r"C:\Users\jayav\.gemini\antigravity-ide\brain\b0067744-fd02-4ccf-8e27-b9dca0c0bba3\.system_generated\tasks\task-356.log"

def is_valid_cadet_name(name):
    name = name.strip().upper()
    if not name or len(name) < 3:
        return False, "too short"
        
    exclusions = [
        "S.NO", "REGISTER NO", "NAME", "DEPARTMENT", "SHIFT", "PARTICULARS", "ON DUTY",
        "B COM", "B.A ENGLISH", "BBA", "BBA CA", "BSC CS", "BCOM", "BCA", "BA", "BSC",
        "NCC SELECTION", "NCC", "CADETS", "TOTAL", "S.NO.", "REG NO", "REG.NO"
    ]
    if name in exclusions:
        return False, "in exclusions"
    
    if re.match(r'^(I|II|III|IV)\s+(BCOM|BSC|BBA|BCA|BA|B\.COM|B\.SC|B|A|I)\b', name):
        return False, "matches Roman numeral header"

    dep_codes = ["BCOM", "BSC", "BBA", "BCA", "BA", "B.COM", "B.SC", "B.B.A", "B.C.A", "B.A."]
    for code in dep_codes:
        if name == code or name.startswith(code + " ") or name.startswith(code + "."):
            return False, f"matches dep_code {code}"
            
    if name.isdigit():
        return False, "is digit"
        
    return True, "valid"

def main():
    if not os.path.exists(log_file_path):
        print("Log not found")
        return
        
    print("Tracing log file parsing:")
    valid_count = 0
    invalid_reasons = {}

    with open(log_file_path, "r", encoding="utf-8", errors="ignore") as f:
        for idx, line in enumerate(f):
            line = line.strip()
            if not line:
                continue
            if "=== FILE:" in line or "Tables count:" in line:
                print(f"Log info: {line}")
                continue
                
            if " | " not in line:
                continue
                
            parts = [p.strip() for p in line.split(" | ")]
            if len(parts) < 3:
                continue
                
            name = parts[2].strip() if len(parts) > 2 else ""
            if "REGISTER NO" in line or "NAME" in name:
                continue
                
            valid, reason = is_valid_cadet_name(name)
            if valid:
                valid_count += 1
                if valid_count <= 10:
                    print(f"  VALID: {name} (Dept: {parts[3]} from line {idx})")
            else:
                invalid_reasons[reason] = invalid_reasons.get(reason, 0) + 1

    print("\nParsing Summary:")
    print(f"Total valid entries: {valid_count}")
    print("Invalid reasons count:")
    for reason, count in invalid_reasons.items():
        print(f"  - {reason}: {count}")

if __name__ == "__main__":
    main()
