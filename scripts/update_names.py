import os

paths = [
    r"D:\Tygia-Tudong\scripts",
    r"C:\Users\THUAN\.gemini\antigravity\skills\vcb-exchange-rate\scripts"
]

replacements = [
    ("TyGia_Banking.xlsx", "TyGia_Banking.xlsx"),
    ("TheoDoi_USD", "TheoDoi_USD"),
    ("Data_TheoDoi_USD", "Data_TheoDoi_USD")
]

for base_dir in paths:
    if not os.path.exists(base_dir):
        continue
    print(f"Updating files in {base_dir}...")
    for f in os.listdir(base_dir):
        if f.endswith(".py") or f.endswith(".ps1"):
            fp = os.path.join(base_dir, f)
            try:
                with open(fp, "r", encoding="utf-8") as file:
                    content = file.read()
                
                updated_content = content
                changes_made = False
                for old, new in replacements:
                    if old in updated_content:
                        updated_content = updated_content.replace(old, new)
                        print(f"  Replaced '{old}' with '{new}' in {f}")
                        changes_made = True
                
                if changes_made:
                    with open(fp, "w", encoding="utf-8") as file:
                        file.write(updated_content)
            except Exception as e:
                print(f"  Error processing {f}: {e}")
