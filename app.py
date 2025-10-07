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

# --- Core Assessment Logic (now accepts user choices) ---

def run_assessment_logic(df_nomination, df_inventory, choices={}):
    processed_rows = []
    for index, nom_row in df_nomination.iterrows():
        pla_id = nom_row['PLA ID']
        matches = df_inventory[df_inventory['PLA ID'] == pla_id]
        
        selected_inventory_row = pd.Series(dtype=object)
        if not matches.empty:
            if len(matches) > 1 and pla_id in choices:
                # Use the user's choice if provided
                selected_inventory_row = matches[matches['Transport NE'] == choices[pla_id]].iloc[0]
            else:
                # Default to the first match
                selected_inventory_row = matches.iloc[0]
                
        combined_row = pd.concat([nom_row, selected_inventory_row.add_prefix('Inv_')])
        processed_rows.append(combined_row)
        
    df = pd.DataFrame(processed_rows)

    numeric_cols = ['GE Port Demand', '10GE Port Demand', 'Inv_GE_1G', 'Inv_GE_10G', 'Inv_MYCOM LOOP NORMAL UTILIZATION']
    if 'Inv_MYCOM LOOP NORMAL UTILIZATION' in df and df['Inv_MYCOM LOOP NORMAL UTILIZATION'].dtype == 'object':
        df['Inv_MYCOM LOOP NORMAL UTILIZATION'] = df['Inv_MYCOM LOOP NORMAL UTILIZATION'].str.replace('%', '', regex=False).astype(float) / 100
    for col in numeric_cols:
        if col in df:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    def get_node_assessment(row):
        failures = []
        if row.get('GE Port Demand', 0) >= 1 and (row.get('Inv_GE_1G', 0) - row.get('GE Port Demand', 0)) <= 2: failures.append("Requires Port Augmentation")
        if row.get('10GE Port Demand', 0) >= 1 and (row.get('Inv_GE_10G', 0) - row.get('10GE Port Demand', 0)) <= 2: failures.append("Requires Port Augmentation")
        return " & ".join(failures) if failures else ("With Headroom" if row.get('GE Port Demand', 0) >= 1 or row.get('10GE Port Demand', 0) >= 1 else "No Port Demand")

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
    print("Master inventory data loaded successfully.")
except Exception as e:
    print(f"CRITICAL: Failed to load master inventory data. Error: {e}")
    df_inventory = pd.DataFrame()

# --- Web Routes ---

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

def handle_assessment_request(nomination_url, action='display'):
    """Helper to avoid code duplication for display and download actions."""
    if not nomination_url:
        return render_template('index.html', error="Nomination URL is required.")
    
    try:
        csv_url = get_google_sheet_csv_url(nomination_url)
        df_nomination = pd.read_csv(csv_url)
        
        # --- Pre-flight check for duplicates ---
        nominated_pla_ids = df_nomination['PLA ID'].unique()
        inventory_counts = df_inventory['PLA ID'].value_counts()
        duplicates_found = inventory_counts[inventory_counts > 1]
        
        duplicates_to_resolve = {}
        for pla_id in nominated_pla_ids:
            if pla_id in duplicates_found.index:
                duplicate_nes = df_inventory[df_inventory['PLA ID'] == pla_id]['Transport NE'].tolist()
                duplicates_to_resolve[pla_id] = duplicate_nes
                
        if duplicates_to_resolve:
            # If duplicates are found, stop and ask the user to resolve them
            return render_template('index.html', duplicates_to_resolve=duplicates_to_resolve, nomination_url=nomination_url, action=action)
            
        # If no duplicates, proceed with assessment
        df_result = run_assessment_logic(df_nomination, df_inventory)
        
        if action == 'display':
            return render_template('index.html', results_table=df_result.to_html(classes='table table-bordered table-hover results-table', index=False))
        else: # action == 'download'
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
    """Handles submission from the duplicate resolution form."""
    nomination_url = request.form.get('nomination_url')
    action = request.form.get('action')
    choices = {key: value for key, value in request.form.items() if key not in ['nomination_url', 'action']}
    
    try:
        csv_url = get_google_sheet_csv_url(nomination_url)
        df_nomination = pd.read_csv(csv_url)
        df_result = run_assessment_logic(df_nomination, df_inventory, choices=choices)
        
        if action == 'display':
            return render_template('index.html', results_table=df_result.to_html(classes='table table-bordered table-hover results-table', index=False))
        else: # action == 'download'
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