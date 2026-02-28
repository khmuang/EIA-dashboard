import pandas as pd
import json
import re
import os
import subprocess
from datetime import datetime

# === CONFIGURATION ===
EXCEL_FOLDER = 'EIA file'
HTML_FILE = 'index.html'

def get_path(filename):
    return os.path.join(EXCEL_FOLDER, filename)

def process_it_asset():
    file_path = get_path('1- IT Asset incomplete information.xlsx')
    if not os.path.exists(file_path): return []
    xls = pd.ExcelFile(file_path)
    sheets = ['No Company', 'No BU', 'No Group', 'No Location']
    summary_dict = {}
    for sn in sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sn, header=None)
            header_row, g_idx, s_idx = 0, 4, 28
            for i, row in df.head(5).iterrows():
                r_list = [str(x).strip() for x in row]
                if "Groups" in r_list or "Group" in r_list:
                    header_row = i
                    g_idx = r_list.index("Groups") if "Groups" in r_list else r_list.index("Group")
                    if "Update Status Y/N" in r_list: s_idx = r_list.index("Update Status Y/N")
                    break
            df = pd.read_excel(xls, sheet_name=sn, header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            g_col, s_col = df.columns[g_idx], df.columns[s_idx]
            df[s_col] = df[s_col].fillna('N').astype(str).str.strip().str.upper()
            df['Status'] = df[s_col].apply(lambda x: 'Y' if x == 'Y' else 'N')
            grp = df.groupby([g_col, 'Status']).size().unstack(fill_value=0).reset_index()
            if 'Y' not in grp.columns: grp['Y'] = 0
            if 'N' not in grp.columns: grp['N'] = 0
            for _, r in grp.iterrows():
                team = str(r[g_col])
                if team not in summary_dict: summary_dict[team] = {'Y': 0, 'N': 0}
                summary_dict[team]['Y'] += int(r['Y'])
                summary_dict[team]['N'] += int(r['N'])
        except: pass
    return [{"Service Team": k, "Y": v['Y'], "N": v['N']} for k, v in summary_dict.items()]

def process_generic(filename, sheet, skip_rows, team_idx, status_idx):
    file_path = get_path(filename)
    if not os.path.exists(file_path): return []
    try:
        df = pd.read_excel(file_path, sheet_name=sheet, header=None)
        data = df.iloc[skip_rows:]
        t_data = data[team_idx].fillna('Unknown')
        s_data = data[status_idx].fillna('N').astype(str).str.strip().str.upper()
        s_mapped = s_data.apply(lambda x: 'Y' if x == 'Y' else 'N')
        temp = pd.DataFrame({'Service Team': t_data, 'Status': s_mapped})
        grp = temp.groupby(['Service Team', 'Status']).size().unstack(fill_value=0).reset_index()
        if 'Y' not in grp.columns: grp['Y'] = 0
        if 'N' not in grp.columns: grp['N'] = 0
        return [{"Service Team": str(r['Service Team']), "Y": int(r['Y']), "N": int(r['N'])} for _, r in grp.iterrows() if str(r['Service Team']) != 'Service Team']
    except: return []

def process_os_replace():
    file_path = get_path('2.1 - Update OS - Replace.xlsx')
    if not os.path.exists(file_path): return []
    try:
        df = pd.read_excel(file_path, header=2)
        df.columns = [str(c).strip() for c in df.columns]
        df['Service Team'] = df['Service Team'].fillna('Unknown')
        df['Status'] = df['Updated or Replaced Y/N'].fillna('N').astype(str).str.strip().str.upper().apply(lambda x: 'Y' if x == 'Y' else 'N')
        grp = df.groupby(['Service Team', 'Status']).size().unstack(fill_value=0).reset_index()
        if 'Y' not in grp.columns: grp['Y'] = 0
        if 'N' not in grp.columns: grp['N'] = 0
        return [{"Service Team": str(r['Service Team']), "Y": int(r['Y']), "N": int(r['N'])} for _, r in grp.iterrows()]
    except: return []

def sync_to_github():
    print("\n--- Syncing to GitHub ---")
    try:
        # Move up to root if we are in 'EIA file' folder
        # But script should be at root. Let's assume it's at root.
        subprocess.run(["git", "add", "index.html", "update_dashboard.py", ".gitignore"], check=True)
        commit_msg = f"Dashboard Update: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        subprocess.run(["git", "commit", "-m", commit_msg], check=True)
        subprocess.run(["git", "push", "origin", "main"], check=True)
        print("\n[SUCCESS] Updated successfully.")
    except Exception as e: print(f"Sync error: {e}")

if __name__ == "__main__":
    print("Step 1: Processing data...")
    sections = [
        {"id": 1, "title": "1- IT Asset Incomplete Information", "details": process_it_asset()},
        {"id": 2, "title": "2.1 - Update OS - Replace", "details": process_os_replace()},
        {"id": 3, "title": "2.2 - OS Require Restart", "details": process_generic('2.2 - Require Restart.xlsx', 'Restart', 2, 2, 20)},
        {"id": 4, "title": "3- Antivirus Installation", "details": process_generic('3- Antivirus not Install.xlsx', 'No AV', 2, 2, 20)},
        {"id": 5, "title": "4- Built-in Firewall Enable", "details": process_generic('4- Built-in Firewall are not enable.xlsx', 'No firewall', 3, 2, 21)},
        {"id": 6, "title": "5- Client Joined Domain", "details": process_generic('5- Client devices are not joined to the domain.xlsx', 'Not join', 3, 2, 21)},
        {"id": 7, "title": "6- Privileged User management", "details": process_generic('6- Privileged User management.xlsx', 'Admin group', 3, 2, 21)},
        {"id": 8, "title": "7- Document Request Evidence", "details": process_generic('7- Document request privileged user.xlsx', 'Document request', 2, 2, 6)}
    ]
    new_raw = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "sections": sections}
    print("Step 2: Updating HTML...")
    with open(HTML_FILE, 'r', encoding='utf-8') as f: html = f.read()
    updated = re.sub(r'const rawData = \{.*?\};', f'const rawData = {json.dumps(new_raw, ensure_ascii=False, indent=4)};', html, flags=re.DOTALL)
    with open(HTML_FILE, 'w', encoding='utf-8') as f: f.write(updated)
    sync_to_github()
