import pandas as pd
import json
import os
from datetime import datetime
import shutil
import subprocess
import re

# --- CONFIGURATION ---
EXCEL_DIR = "EIA file"
BACKUP_DIR = os.path.join(EXCEL_DIR, "backup file")
OUTPUT_HTML = "index.html"
GITHUB_REPO_URL = "https://github.com/khmuang/EIA-dashboard.git"

FILES = {
    1: "1- IT Asset incomplete information.xlsx",
    2: "2.1 - Update OS - Replace.xlsx",
    3: "2.2 - Require Restart.xlsx",
    4: "3- Antivirus not Install.xlsx",
    5: "4- Built-in Firewall are not enable.xlsx",
    6: "5- Client devices are not joined to the domain.xlsx",
    7: "6- Privileged User management.xlsx",
    8: "7- Document request privileged user.xlsx"
}

# Total Audit Units per Team per Topic (Extracted from Stable Backup)
TOPIC_TOTALS = {
    1: {"Branch": 245, "DC": 38, "HO": 52},
    2: {"Branch": 7565, "DC": 863, "HO": 2691},
    3: {"Branch": 788, "DC": 229, "HO": 1367},
    4: {"Branch": 329, "DC": 34, "HO": 268},
    5: {"Branch": 4595, "DC": 919, "HO": 1446},
    6: {"Branch": 144, "DC": 18, "HO": 374},
    7: {"Branch": 1827, "DC": 363, "HO": 983},
    8: {"Branch": 3, "DC": 3, "HO": 3}
}

def backup_files():
    print(f"--- Backing up Excel files to '{BACKUP_DIR}' ---")
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for fid, name in FILES.items():
        src = os.path.join(EXCEL_DIR, name)
        if os.path.exists(src):
            dst = os.path.join(BACKUP_DIR, f"{timestamp}_{name}")
            shutil.copy2(src, dst)
    print("Backup completed successfully.")

def get_serviced_by_col(df):
    possible_names = ['Serviced By', 'Service By', 'serviced by', 'service by']
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def process_data():
    print(f"Reading files from '{EXCEL_DIR}'...")
    sections = []
    
    for fid, name in FILES.items():
        path = os.path.join(EXCEL_DIR, name)
        if not os.path.exists(path):
            print(f"Warning: {name} not found. Skipping.")
            continue
            
        # Topic 1 aggregation logic
        if fid == 1:
            sheets = ['No Company', 'No BU', 'No Group', 'No Location']
            all_df = []
            for s in sheets:
                try:
                    df_temp = pd.read_excel(path, sheet_name=s)
                    col = get_serviced_by_col(df_temp)
                    if col:
                        all_df.append(df_temp[[col]].rename(columns={col: 'Serviced By'}))
                except: continue
            df = pd.concat(all_df) if all_df else pd.DataFrame(columns=['Serviced By'])
        # Topic 5 special sheet
        elif fid == 5:
            try: df = pd.read_excel(path, sheet_name='No firewall')
            except: df = pd.read_excel(path)
        else:
            df = pd.read_excel(path)

        col = get_serviced_by_col(df)
        counts = df[col].value_counts() if col else pd.Series()
        
        details = []
        for team in ['Branch', 'HO', 'DC']:
            n_val = int(counts.get(team, 0))
            # Calculate Y based on Fixed Population constants
            total_for_team = TOPIC_TOTALS.get(fid, {}).get(team, 100)
            y_val = max(0, total_for_team - n_val)
            details.append({"Service Team": team, "Y": y_val, "N": n_val})
            
        sections.append({
            "id": fid,
            "title": name.replace(".xlsx", "").split("- ", 1)[-1] if "-" in name else name.replace(".xlsx", ""),
            "details": details
        })

    thai_year = datetime.now().year + 543
    timestamp_str = datetime.now().strftime(f"%d/%m/{thai_year} %H:%M:%S")
    
    data = {
        "timestamp": timestamp_str,
        "sections": sections
    }
    return data

def update_html(data):
    if not os.path.exists(OUTPUT_HTML):
        print(f"Error: {OUTPUT_HTML} template not found.")
        return

    with open(OUTPUT_HTML, 'r', encoding='utf-8') as f:
        content = f.read()

    json_data = json.dumps(data, ensure_ascii=False, indent=4)
    # Inject data into index.html
    updated_content = re.sub(r'const rawData = \{.*?\};', f'const rawData = {json_data};', content, flags=re.DOTALL)

    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(updated_content)
    print(f"Local {OUTPUT_HTML} updated with real data.")

def sync_to_github():
    print("\n--- Syncing to GitHub ---")
    try:
        subprocess.run(["git", "add", "index.html"], check=True)
        commit_msg = f"Auto-update Dashboard (Correct Logic): {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        subprocess.run(["git", "commit", "-m", commit_msg], check=True)
        subprocess.run(["git", "push", "origin", "main"], check=True)
        print("Success: Dashboard is now LIVE on GitHub Pages!")
    except Exception as e:
        print(f"Git Sync Failed: {e}")

if __name__ == "__main__":
    backup_files()
    data = process_data()
    update_html(data)
    sync_to_github()
