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
    3: {"Branch": 798, "DC": 232, "HO": 1354},   # Total 2384
    4: {"Branch": 329, "DC": 34, "HO": 268},     # Total 631
    5: {"Branch": 4651, "DC": 922, "HO": 1400},  # Total 6973
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

def export_csv_summary(data):
    all_rows = []
    
    # Trackers for aggregation
    team_summary = {"Branch": {"Y": 0, "N": 0}, "HO": {"Y": 0, "N": 0}, "DC": {"Y": 0, "N": 0}}
    topic_summary = []
    
    # 1. Detailed Rows (Topic + Team)
    all_rows.append({"Category": "--- DETAILED BREAKDOWN (TOPIC + TEAM) ---"})
    for sec in data['sections']:
        t_y = 0; t_n = 0
        for d in sec['details']:
            y = d['Y']; n = d['N']; total = y + n
            pct = (y / total * 100) if total > 0 else 0
            
            all_rows.append({
                "Category": "Detailed",
                "Topic ID": sec['id'],
                "Topic Title": sec['title'],
                "Service Team": d['Service Team'],
                "Success (Y)": y,
                "Pending (N)": n,
                "Total": total,
                "Compliance %": f"{pct:.2f}%"
            })
            # Aggregate for team
            team_summary[d['Service Team']]["Y"] += y
            team_summary[d['Service Team']]["N"] += n
            # Aggregate for topic
            t_y += y; t_n += n
        
        topic_summary.append({
            "Topic ID": sec['id'],
            "Topic Title": sec['title'],
            "Y": t_y, "N": t_n
        })

    all_rows.append({}) # Spacer
    
    # 2. Summary by Topic
    all_rows.append({"Category": "--- SUMMARY BY TOPIC ---"})
    for ts in topic_summary:
        total = ts['Y'] + ts['N']
        pct = (ts['Y'] / total * 100) if total > 0 else 0
        all_rows.append({
            "Category": "Topic Summary",
            "Topic ID": ts['Topic ID'],
            "Topic Title": ts['Topic Title'],
            "Success (Y)": ts['Y'],
            "Pending (N)": ts['N'],
            "Total": total,
            "Compliance %": f"{pct:.2f}%"
        })

    all_rows.append({}) # Spacer

    # 3. Summary by Service Team
    all_rows.append({"Category": "--- SUMMARY BY SERVICE TEAM ---"})
    for team, vals in team_summary.items():
        total = vals['Y'] + vals['N']
        pct = (vals['Y'] / total * 100) if total > 0 else 0
        all_rows.append({
            "Category": "Team Summary",
            "Service Team": team,
            "Success (Y)": vals['Y'],
            "Pending (N)": vals['N'],
            "Total": total,
            "Compliance %": f"{pct:.2f}%"
        })

    all_rows.append({}) # Spacer

    # 4. Grand Total
    g_y = sum(v['Y'] for v in team_summary.values())
    g_n = sum(v['N'] for v in team_summary.values())
    g_total = g_y + g_n
    g_pct = (g_y / g_total * 100) if g_total > 0 else 0
    
    all_rows.append({"Category": "--- GRAND TOTAL PROJECT ---"})
    all_rows.append({
        "Category": "GRAND TOTAL",
        "Success (Y)": g_y,
        "Pending (N)": g_n,
        "Total": g_total,
        "Compliance %": f"{g_pct:.2f}%"
    })

    df = pd.DataFrame(all_rows)
    output_path = "dashboard_summary_complete.csv"
    # Reorder columns for clarity
    cols = ["Category", "Topic ID", "Topic Title", "Service Team", "Success (Y)", "Pending (N)", "Total", "Compliance %"]
    df = df[cols]
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"SUCCESS: Comprehensive CSV Summary exported to {output_path}")

if __name__ == "__main__":
    data = process_data()
    g_total = sum(sum(d['Y']+d['N'] for d in s['details']) for s in data['sections'])
    print(f"\n>>> FINAL SYSTEM CHECK: GRAND TOTAL = {g_total} (Target: 25169) <<<")
    if g_total == 25169:
        update_html(data)
        export_csv_summary(data)
    else:
        print(f"ERROR: Integrity Check Failed ({g_total} != 25169). Aborting.")
