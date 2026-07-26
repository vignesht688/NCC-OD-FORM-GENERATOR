import re
import os

log_file_path = r"C:\Users\jayav\.gemini\antigravity-ide\brain\b0067744-fd02-4ccf-8e27-b9dca0c0bba3\.system_generated\tasks\task-356.log"

def main():
    if not os.path.exists(log_file_path):
        print("Log not found")
        return
        
    print("Listing all names matching Roman numeral header pattern:")
    with open(log_file_path, "r", encoding="utf-8", errors="ignore") as f:
        for idx, line in enumerate(f):
            line = line.strip()
            if not line or " | " not in line:
                continue
            if "=== FILE:" in line or "Tables count:" in line:
                continue
                
            parts = [p.strip() for p in line.split(" | ")]
            if len(parts) < 3:
                continue
                
            name = parts[2].strip()
            if "REGISTER NO" in line or "NAME" in name:
                continue
                
            if re.match(r'^(I|II|III|IV)\s+(BCOM|BSC|BBA|BCA|BA|B\.COM|B\.SC|B|A|I)\b', name):
                print(f"  Line {idx}: name='{name}', full_line='{line}'")

if __name__ == "__main__":
    main()
