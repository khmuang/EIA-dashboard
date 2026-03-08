import pandas as pd
import json
import os
from datetime import datetime
import shutil
import subprocess
import re
import mysql.connector

# --- CONFIGURATION ---
EXCEL_DIR = "EIA file"
BACKUP_DIR = os.path.join(EXCEL_DIR, "backup file")
OUTPUT_HTML = "index.html"
GITHUB_REPO_URL = "https://github.com/khmuang/EIA-dashboard.git"

# MySQL Config (XAMPP Default)
DB_CONFIG = {
    'user': 'root',
    'password': '',
    'host': '127.0.0.1',
    'database': 'eia_compliance'
}

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

def load_excel_with_auto_header(path, sheet_name=0):
    """Try to find the correct header row by looking for 'Serviced By' or 'Service Team'"""
    try:
        # Load first 20 rows to find header
        df_raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=20)
        header_row = 0
        
        # Support multiple possible column names
        possible_headers = ['serviced by', 'service by', 'service team', 'serviced team']
        
        for i, row in df_raw.iterrows():
            row_vals = [str(val).strip().lower() for val in row.values if pd.notnull(val)]
            if any(h in row_vals for h in possible_headers):
                header_row = i
                break
        
        # Reload with identified header row
        return pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    except Exception as e:
        print(f"Error reading {path}: {e}")
        return pd.DataFrame()

def get_serviced_by_col(df):
    possible_names = ['Serviced By', 'Service By', 'Service Team', 'Service team', 'Serviced team']
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
                df_temp = load_excel_with_auto_header(path, sheet_name=s)
                col = get_serviced_by_col(df_temp)
                if col:
                    all_df.append(df_temp[[col]].rename(columns={col: 'Service Team'}))
            df = pd.concat(all_df) if all_df else pd.DataFrame(columns=['Service Team'])
        # Topic 5 special sheet
        elif fid == 5:
            df = load_excel_with_auto_header(path, sheet_name='No firewall')
            if df.empty or get_serviced_by_col(df) is None:
                df = load_excel_with_auto_header(path)
        else:
            df = load_excel_with_auto_header(path)

        col = get_serviced_by_col(df)
        if col:
            # Normalize and count
            df[col] = df[col].astype(str).str.strip()
            counts = df[col].value_counts()
        else:
            print(f"Warning: Could not find team column in {name}")
            counts = pd.Series()
        
        details = []
        print(f"  > Topic {fid}: {name}")
        for team in ['Branch', 'HO', 'DC']:
            n_val = int(counts.get(team, 0))
            
            # Calculate Y based on Fixed Population constants
            total_for_team = TOPIC_TOTALS.get(fid, {}).get(team, 0)
            
            # Fallback if total is missing or too small
            if total_for_team < n_val:
                print(f"    [!] Warning: {team} pending ({n_val}) > total ({total_for_team}). Adjusting total to match pending.")
                total_for_team = n_val 
                
            y_val = max(0, total_for_team - n_val)
            print(f"    - {team}: Y={y_val}, N={n_val} (Total={total_for_team})")
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

def sync_to_mysql(data, quarter, year):
    print(f"\n--- Syncing Data to MySQL (Quarter: {quarter}, Year: {year}) ---")
    try:
        cnx = mysql.connector.connect(**DB_CONFIG)
        cursor = cnx.cursor()

        # Clean old records for this Q/Year
        delete_query = "DELETE FROM audit_data WHERE quarter = %s AND audit_year = %s"
        cursor.execute(delete_query, (quarter, year))

        insert_query = (
            "INSERT INTO audit_data (topic_id, topic_name, team_name, success_y, pending_n, quarter, audit_year) "
            "VALUES (%s, %s, %s, %s, %s, %s, %s)"
        )

        records = []
        for sec in data['sections']:
            for d in sec['details']:
                records.append((
                    sec['id'],
                    sec['title'],
                    d['Service Team'],
                    d['Y'],
                    d['N'],
                    quarter,
                    year
                ))
        
        cursor.executemany(insert_query, records)
        cnx.commit()
        print(f"Success: {len(records)} records synced to MySQL database.")
        
        cursor.close()
        cnx.close()
    except mysql.connector.Error as err:
        print(f"MySQL Error: {err}")

def update_html(data):
    if not os.path.exists(OUTPUT_HTML):
        print(f"Error: {OUTPUT_HTML} template not found.")
        return

    with open(OUTPUT_HTML, 'r', encoding='utf-8') as f:
        content = f.read()

    json_data = json.dumps(data, ensure_ascii=False, indent=4)
    updated_content = re.sub(r'const rawData = \{.*?\};', f'const rawData = {json_data};', content, flags=re.DOTALL)

    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(updated_content)
    print(f"Local {OUTPUT_HTML} updated with real data.")

def sync_to_github():
    print("\n--- Syncing to GitHub ---")
    try:
        subprocess.run(["git", "add", "index.html"], check=True)
        commit_msg = f"Auto-update Dashboard (Advanced DB Sync): {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        subprocess.run(["git", "commit", "-m", commit_msg], check=True)
        subprocess.run(["git", "push", "origin", "main"], check=True)
        print("Success: Dashboard is now LIVE on GitHub Pages!")
    except Exception as e:
        print(f"Git Sync Failed: {e}")

if __name__ == "__main__":
    # 1. Backup Excel
    backup_files()
    
    # 2. Process Data from Excel
    data = process_data()
    
    # 3. Update Local HTML (for immediate preview)
    update_html(data)
    
    # 4. Ask for Quarter/Year for MySQL Storage
    print("\n" + "="*30)
    print(" DATABASE SYNC CONFIGURATION")
    print("="*30)
    q_input = input("Enter Quarter (e.g., Q1, Q2, Q3, Q4) [Default: Q1]: ").strip().upper() or "Q1"
    y_input = input("Enter Year (e.g., 2026) [Default: 2026]: ").strip() or "2026"
    
    # 5. Sync to MySQL
    sync_to_mysql(data, q_input, int(y_input))
    
    # 6. Final Sync to GitHub (Ask user confirmation first - manually)
    print("\n" + "="*30)
    confirm = input("Push updates to GitHub now? (y/n): ").lower()
    if confirm == 'y':
        sync_to_github()
    else:
        print("Git Push cancelled by user.")
