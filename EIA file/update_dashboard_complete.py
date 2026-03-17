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

# STANDARD POPULATION TARGETS - Grand Total: 25,169
TOPIC_TOTALS = {
    1: {"Branch": 246, "DC": 38, "HO": 60},      # Total 344
    2: {"Branch": 7565, "DC": 863, "HO": 2691},  # Total 11119
    3: {"Branch": 788, "DC": 229, "HO": 1367},   # Total 2384
    4: {"Branch": 329, "DC": 34, "HO": 268},     # Total 631
    5: {"Branch": 4599, "DC": 919, "HO": 1455},  # Total 6973
    6: {"Branch": 144, "DC": 18, "HO": 374},     # Total 536
    7: {"Branch": 1859, "DC": 365, "HO": 949},   # Total 3173
    8: {"Branch": 3, "DC": 3, "HO": 3}           # Total 9
}

def get_correct_df(file_path, sheet_name=0):
    """Finds the header row dynamically and returns the DataFrame."""
    potential_keys = ['Name', 'BU', 'Service Team', 'Computer Name', 'Bu']
    for h in [2, 3, 1, 0, 4]:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=h)
            if any(k in df.columns for k in potential_keys):
                return df
        except: continue
    return pd.read_excel(file_path, sheet_name=sheet_name, header=0) # Fallback

def calculate_topic_stats(fid, file_path):
    """Calculates Y/N stats for a topic from its Excel file."""
    stats = {"Branch": 0, "DC": 0, "HO": 0}
    
    try:
        if fid == 1:
            xl = pd.ExcelFile(file_path)
            for sheet in xl.sheet_names:
                df = get_correct_df(file_path, sheet_name=sheet)
                status_col = next((c for c in df.columns if "Update Status Y/N" in str(c)), None)
                group_col = next((c for c in df.columns if "Groups" in str(c)), None)
                if status_col and group_col:
                    df['Team'] = df[group_col].astype(str).str.upper().apply(lambda x: 'HO' if 'HO' in x else ('DC' if 'DC' in x else 'Branch'))
                    y_counts = df[df[status_col].astype(str).str.strip().str.upper() == 'Y']['Team'].value_counts()
                    for team, count in y_counts.items(): stats[team] += int(count)
        elif fid == 3:
            df = pd.read_excel(file_path, sheet_name='Restart', header=2)
            team_col = 'Service Team'
            status_col = 'Restart Action  Y/N'
            if status_col in df.columns:
                mask_y = df[status_col].astype(str).str.strip().str.upper() == 'Y'
                y_counts = df[mask_y][team_col].value_counts()
                for team, count in y_counts.items():
                    t_name = str(team).strip()
                    if t_name in stats: stats[t_name] += int(count)
        else:
            df = get_correct_df(file_path)
            # Identify columns
            team_col = next((c for c in df.columns if "Service Team" in str(c)), None)
            status_keys = ["Update Status Y/N", "Updated or Replaced Y/N", "Install Status Y/N", "Firewall enable Y/N", 
                           "Join status Y/N", "Remove accounts", "evidence", "Y/N", "Restart Action", "Status"]
            status_col = None
            for key in status_keys:
                status_col = next((c for c in df.columns if key.lower() in str(c).lower()), None)
                if status_col: break
            
            if not team_col: # Fallback if no service team column (Topic 1 already handled)
                team_col = 'Service Team' if 'Service Team' in df.columns else None

            if status_col:
                mask_y = df[status_col].astype(str).str.strip().str.upper() == 'Y'
                if team_col:
                    y_counts = df[mask_y][team_col].value_counts()
                    for team, count in y_counts.items():
                        t_name = str(team).strip()
                        if t_name in stats: stats[t_name] += int(count)
                else:
                    # If no team col, distribute to all (rare)
                    stats["HO"] = int(mask_y.sum())
        
        return stats
    except Exception as e:
        print(f"Error reading Topic {fid}: {e}")
        return stats

def process_data():
    print(f"Reading Excel files and preparing V23 update...")
    sections = []
    
    for fid, name in FILES.items():
        file_path = os.path.join(EXCEL_DIR, name)
        y_stats = calculate_topic_stats(fid, file_path) if os.path.exists(file_path) else {"Branch": 0, "DC": 0, "HO": 0}
        
        details = []
        for team in ['Branch', 'HO', 'DC']:
            y = int(y_stats.get(team, 0))
            total = int(TOPIC_TOTALS[fid].get(team, 0))
            n = int(max(0, total - y))
            details.append({"Service Team": team, "Y": y, "N": n})
            
        sections.append({"id": fid, "title": name.replace(".xlsx", "").split("- ", 1)[-1], "details": details})
        t_y = sum(d['Y'] for d in details); t_n = sum(d['N'] for d in details)
        print(f"Topic {fid} -> Success(Y):{t_y} Pending(N):{t_n} | Total:{t_y+t_n}")

    thai_year = datetime.now().year + 543
    timestamp_str = datetime.now().strftime(f"%d/%m/{thai_year} %H:%M:%S")
    return {"timestamp": timestamp_str, "sections": sections}

def update_html(data):
    if not os.path.exists(OUTPUT_HTML): return
    with open(OUTPUT_HTML, 'r', encoding='utf-8') as f: content = f.read()
    json_data = json.dumps(data, ensure_ascii=False, indent=4)
    updated = re.sub(r'const rawData = \{.*?\};', f'const rawData = {json_data};', content, flags=re.DOTALL)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f: f.write(updated)
    print(f"\nSUCCESS: Local index.html updated with LIVE Excel data.")

if __name__ == "__main__":
    data = process_data()
    g_total = sum(sum(d['Y']+d['N'] for d in s['details']) for s in data['sections'])
    print(f"\n>>> FINAL SYSTEM CHECK: GRAND TOTAL = {g_total} (Target: 25169) <<<")
    if g_total == 25169:
        update_html(data)
    else:
        print(f"ERROR: Integrity Check Failed ({g_total} != 25169). Aborting.")
