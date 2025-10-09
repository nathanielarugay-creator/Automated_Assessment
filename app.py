# app.py

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, render_template, request, make_response
from io import BytesIO
import os

app = Flask(__name__)

# --- Helper Functions ---

def authenticate_gsheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
    return gspread.authorize(creds)

def get_google_sheet_csv_url(url: str):
    if "docs.google.com/spreadsheets/d/" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return None

def to_excel_in_memory(df):
    output = BytesIO()
    df_cleaned = df.astype(str)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_cleaned.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return output

# --- Core Assessment Logic ---

def run_assessment_logic(df_nomination, df_inventory, choices={}):
    processed_rows = []
    for index, nom_row in df_nomination.iterrows():
        pla_id = str(nom_row['PLA ID'])
        matches = df_inventory[df_inventory['PLA ID'] == pla_id]
        
        selected_inventory_row = pd.Series(dtype=object)
        if not matches.empty:
            if len(matches) > 1 and pla_id in choices:
                selected_inventory_row = matches[matches['Transport NE'] == choices[pla_id]].iloc[0]
            else:
                selected_inventory_row = matches.iloc[0]
                
        combined_row = pd.concat([nom_row, selected_inventory_row.add_prefix('Inv_')])
        processed_rows.append(combined_row)
        
    df = pd.DataFrame(processed_rows)
    
    if 'Inv_MYCOM LOOP NORMAL UTILIZATION' in df:
        util_col = df['Inv_MYCOM LOOP NORMAL UTILIZATION'].astype(str).str.replace('%', '', regex=False)
        df['Inv_MYCOM LOOP NORMAL UTILIZATION'] = pd.to_numeric(util_col, errors='coerce').fillna(0)

    numeric_cols = ['GE Port Demand', '10GE Port Demand', 'Inv_GE_1G', 'Inv_GE_10G', 'Inv_25GE']
    for col in numeric_cols:
        if col in df:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # --- THIS FUNCTION HAS BEEN UPDATED WITH YOUR NEW LOGIC ---
    def get_node_assessment(row):
        # Rule 1: Override for 25GE ports
        if row.get('Inv_25GE', 0) > 2:
            return "With Headroom"

        ge_demand = row.get('GE Port Demand', 0)
        ten_ge_demand = row.get('10GE Port Demand', 0)

        # Rule 2: Check 1G port headroom. If demand exists and remaining ports are less than 2, it fails.
        if ge_demand > 0 and (row.get('Inv_GE_1G', 0) - ge_demand) < 2:
            return "Requires Port Augmentation"
        
        # Rule 3: Check 10G port headroom. If demand exists and remaining ports are less than 2, it fails.
        if ten_ge_demand > 0 and (row.get('Inv_GE_10G', 0) - ten_ge_demand) < 2:
            return "Requires Port Augmentation"
            
        # If no rules failed, determine if there was any demand
        if ge_demand > 0 or ten_ge_demand > 0:
            return "With Headroom"
        
        # If no demand at all
        return "No Port Demand"
    # --- END OF UPDATED FUNCTION ---

    def get_loop_assessment(row):
        return "Requires Loop Upgrade" if row.get('Inv_MYCOM LOOP NORMAL UTILIZATION', 0) >= 0.7 else "With Headroom"

    df['Node Assessment'] = df.apply(get_node_assessment, axis=1)
    df['Loop Assessment'] = df.apply(get_loop_assessment, axis=1)
    return df

# --- Data Loading (runs once on startup) ---

print("Loading master inventory data...")
try:
    gsheet_client = authenticate_gsheets()
    spreadsheet = gsheet_client.open_by_key('11B6VE-NJI_Xh6SEm7oerIXWoGD45IbEcDbrQmt1uzrQ')
    worksheet = spreadsheet.worksheet("Merged_Inventory_Data")
    df_inventory = pd.DataFrame(worksheet.get_all_records())
    if 'PLA ID' in df_inventory.columns:
        df_inventory['PLA ID'] = df_inventory['PLA ID'].astype(str)
    print("Master inventory data loaded successfully.")
except Exception as e:
    print(f"CRITICAL: Failed to load master inventory data. Error: {e}")
    df_inventory = pd.DataFrame()

# --- Web Routes ---

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

def handle_assessment_request(nomination_url, action='display'):
    if not nomination_url:
        return render_template('index.html', error="Nomination URL is required.")
    
    try:
        csv_url = get_google_sheet_csv_url(nomination_url)
        df_nomination = pd.read_csv(csv_url, dtype={'PLA ID': str})
        
        nominated_pla_ids = df_nomination['PLA ID'].unique()
        inventory_counts = df_inventory['PLA ID'].value_counts()
        duplicates_found = inventory_counts[inventory_counts > 1]
        
        duplicates_to_resolve = {}
        for pla_id in nominated_pla_ids:
            if pla_id in duplicates_found.index:
                duplicate_nes = df_inventory[df_inventory['PLA ID'] == pla_id]['Transport NE'].tolist()
                duplicates_to_resolve[pla_id] = duplicate_nes
                
        if duplicates_to_resolve:
            return render_template('index.html', duplicates_to_resolve=duplicates_to_resolve, nomination_url=nomination_url, action=action)
            
        df_result = run_assessment_logic(df_nomination, df_inventory)
        
        if action == 'display':
            return render_template('index.html', results_table=df_result.to_html(classes='table table-bordered table-hover results-table', index=False))
        else:
            excel_data = to_excel_in_memory(df_result)
            response = make_response(excel_data)
            response.headers['Content-Disposition'] = 'attachment; filename=Final_Assessment.xlsx'
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            return response

    except Exception as e:
        return render_template('index.html', error=f"An error occurred: {e}")

@app.route('/assess_and_display', methods=['POST'])
def assess_and_display():
    nomination_url = request.form.get('nomination_url')
    return handle_assessment_request(nomination_url, action='display')

@app.route('/assess_and_download', methods=['POST'])
def assess_and_download():
    nomination_url = request.form.get('nomination_url')
    return handle_assessment_request(nomination_url, action='download')

@app.route('/assess_with_choices', methods=['POST'])
def assess_with_choices():
    nomination_url = request.form.get('nomination_url')
    action = request.form.get('action')
    choices = {key: value for key, value in request.form.items() if key not in ['nomination_url', 'action']}
    
    try:
        csv_url = get_google_sheet_csv_url(nomination_url)
        df_nomination = pd.read_csv(csv_url, dtype={'PLA ID': str})
        df_result = run_assessment_logic(df_nomination, df_inventory, choices=choices)
        
        if action == 'display':
            return render_template('index.html', results_table=df_result.to_html(classes='table table-bordered table-hover results-table', index=False))
        else:
            excel_data = to_excel_in_memory(df_result)
            response = make_response(excel_data)
            response.headers['Content-Disposition'] = 'attachment; filename=Final_Assessment.xlsx'
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            return response
            
    except Exception as e:
        return render_template('index.html', error=f"An error occurred: {e}")

@app.route('/download_master')
def download_master():
    excel_data = to_excel_in_memory(df_inventory)
    response = make_response(excel_data)
    response.headers['Content-Disposition'] = 'attachment; filename=Service_Inventory_Data.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response

if __name__ == '__main__':
    app.run(debug=True)