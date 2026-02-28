import pandas as pd
import json
import re
from datetime import datetime

def process_it_asset():
    xls = pd.ExcelFile('1- IT Asset incomplete information.xlsx')
    sheets = ['No Company', 'No BU', 'No Group', 'No Location']
    all_details = []
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
                y, n = r['Y'], r['N']
                all_details.append({"Groups": team, "Y": int(y), "N": int(n), "Sheet": sn})
                if team not in summary_dict: summary_dict[team] = {'Y': 0, 'N': 0}
                summary_dict[team]['Y'] += int(y)
                summary_dict[team]['N'] += int(n)
        except: pass
    
    sec_details = [{"Service Team": k, "Y": v['Y'], "N": v['N']} for k, v in summary_dict.items()]
    return all_details, sec_details

def process_generic(file, sheet, skip_rows, team_idx, status_idx):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None)
        data = df.iloc[skip_rows:]
        t_data = data[team_idx].fillna('Unknown')
        s_data = data[status_idx].fillna('N').astype(str).str.strip().str.upper()
        s_mapped = s_data.apply(lambda x: 'Y' if x == 'Y' else 'N')
        temp = pd.DataFrame({'Team': t_data, 'Status': s_mapped})
        grp = temp.groupby(['Team', 'Status']).size().unstack(fill_value=0).reset_index()
        if 'Y' not in grp.columns: grp['Y'] = 0
        if 'N' not in grp.columns: grp['N'] = 0
        return [{"Service Team": str(r['Team']), "Y": int(r['Y']), "N": int(r['N'])} for _, r in grp.iterrows()]
    except: return []

def process_os_replace():
    try:
        df = pd.read_excel('2.1 - Update OS - Replace.xlsx', header=2)
        df.columns = [str(c).strip() for c in df.columns]
        df['Service Team'] = df['Service Team'].fillna('Unknown')
        df['Status'] = df['Updated or Replaced Y/N'].fillna('N').astype(str).str.strip().str.upper().apply(lambda x: 'Y' if x == 'Y' else 'N')
        grp = df.groupby(['Service Team', 'Status']).size().unstack(fill_value=0).reset_index()
        if 'Y' not in grp.columns: grp['Y'] = 0
        if 'N' not in grp.columns: grp['N'] = 0
        return [{"Service Team": str(r['Service Team']), "Y": int(r['Y']), "N": int(r['N'])} for _, r in grp.iterrows()]
    except: return []

print("Reading Excel files...")
it_details, it_sec = process_it_asset()
os_replace = process_os_replace()
os_restart = process_generic('2.2 - Require Restart.xlsx', 'Restart', 2, 2, 20)
antivirus = process_generic('3- Antivirus not Install.xlsx', 'No AV', 2, 2, 20)
firewall = process_generic('4- Built-in Firewall are not enable.xlsx', 'No firewall', 3, 2, 21)
domain = process_generic('5- Client devices are not joined to the domain.xlsx', 'Not join', 3, 2, 21)
privileged = process_generic('6- Privileged User management.xlsx', 'Admin group', 3, 2, 21)
doc_req = process_generic('7- Document request privileged user.xlsx', 'Document request', 2, 2, 6)

new_data = {
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "it_asset_details": it_details,
    "sections": [
        {"id": 1, "title": "1- IT Asset Incomplete Information", "details": it_sec},
        {"id": 2, "title": "2.1 - Update OS - Replace", "details": os_replace},
        {"id": 3, "title": "2.2 - OS Require Restart", "details": os_restart},
        {"id": 4, "title": "3- Antivirus Installation", "details": antivirus},
        {"id": 5, "title": "4- Built-in Firewall Enable", "details": firewall},
        {"id": 6, "title": "5- Client Joined Domain", "details": domain},
        {"id": 7, "title": "6- Privileged User management", "details": privileged},
        {"id": 8, "title": "7- Document Request Evidence", "details": doc_req}
    ]
}

print("Updating index.html...")
with open('index.html', 'r', encoding='utf-8') as f:
    html_content = f.read()

# Replace the rawData JS object in the HTML
json_str = json.dumps(new_data, ensure_ascii=False, indent=4)
pattern = re.compile(r'const rawData = \{.*?\};', re.DOTALL)
updated_html = pattern.sub(f'const rawData = {json_str};', html_content)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(updated_html)

print("Dashboard Updated Successfully! You can now open index.html")
