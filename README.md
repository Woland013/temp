# temp

import os
import zipfile
import pandas as pd
from io import BytesIO

try:
    import rarfile
    rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"
except ImportError:
    rarfile = None

def check_excel_for_resume(file_stream, file_name):
    """The 'DNA Test' - Refined for memory and engine safety."""
    try:
        # 1. Memory Safety: Only use BytesIO if it's a stream
        # (This is Point #1 from your friend)
        if hasattr(file_stream, 'read'):
            content = file_stream.read()
            # If the file is huge (>100MB), maybe skip or warn
            if len(content) > 100 * 1024 * 1024:
                print(f"  ⚠️ Skipping {file_name} - Too large ({len(content)//1024//1024} MB)")
                return False
            file_stream = BytesIO(content)

        # 2. Engine Safety: Explicitly handle .xls vs .xlsx
        # (Point #2 from your friend)
        engine = 'xlrd' if file_name.lower().endswith('.xls') else 'openpyxl'
        
        xl = pd.ExcelFile(file_stream, engine=engine)
        
        # 3. Efficiency: Use 'any' to stop as soon as we find a match
        # (Point #5 from your friend)
        if any(s.lower() == "resume" for s in xl.sheet_names):
            return True
            
    except Exception as e:
        print(f"  ❌ Error reading {file_name}: {e}")
    return False

def universal_scanner(path):
    print(f"--- 🛡️ Production Scan: {os.path.basename(path)} ---")
    matches = []

    if not os.path.exists(path):
        print("❌ Path not found.")
        return []

    # CASE 1: FOLDER
    if os.path.isdir(path):
        for root, _, files in os.walk(path):
            for file in files:
                if file.lower().endswith(('.xlsx', '.xls')):
                    full_path = os.path.join(root, file)
                    if check_excel_for_resume(full_path, file):
                        matches.append(full_path)

    # CASE 2: ZIP (With Corruption Handling - Point #3)
    elif path.lower().endswith('.zip'):
        try:
            with zipfile.ZipFile(path, 'r') as zf:
                for f_name in zf.namelist():
                    if f_name.lower().endswith(('.xlsx', '.xls')):
                        with zf.open(f_name) as f:
                            if check_excel_for_resume(f, f_name):
                                matches.append(f"{path} -> {f_name}")
        except zipfile.BadZipFile:
            print(f"❌ Corrupted ZIP: {path}")

    # CASE 3: RAR (With Corruption Handling)
    elif path.lower().endswith('.rar'):
        if not rarfile:
            print("❌ rarfile library missing.")
        else:
            try:
                with rarfile.RarFile(path) as rf:
                    for f_name in rf.namelist():
                        if f_name.lower().endswith(('.xlsx', '.xls')):
                            with rf.open(f_name) as f:
                                if check_excel_for_resume(f, f_name):
                                    matches.append(f"{path} -> {f_name}")
            except rarfile.Error as e:
                print(f"❌ RAR Error on {path}: {e}")

    # Results Report (Point #6 - Consistency)
    if matches:
        print(f"\n🎯 FOUND {len(matches)} MATCHES:")
        for m in matches: print(f" - {m}")
    else:
        print("\n❌ No 'Resume' sheet found.")
    
    return matches

# --- EXECUTION ---
# Path detection is now safe (Point #5)
user_home = os.environ['USERPROFILE']
desktop = os.path.join(user_home, 'Desktop')
if not os.path.exists(desktop): # Handle OneDrive
    desktop = os.path.join(user_home, 'OneDrive', 'Desktop')

target = os.path.join(desktop, "test.zip")
universal_scanner(target)
