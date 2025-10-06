# app.py

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, render_template, request, make_response
from io import BytesIO
import os

app = Flask(__name__)
app.secret_key = os.urandom(24)

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
    # --- THIS IS THE CHANGED SECTION ---
    # Create a clean copy and convert all columns to strings to prevent export errors
    df_cleaned = df.astype(str)
    # --- END OF CHANGE ---
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Use the cleaned DataFrame for the export
        df_cleaned.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return output

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

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/assess', methods=['POST'])
def assess():
    nomination_url = request.form.get('nomination_url')
    if not nomination_url:
        return "Nomination URL is required.", 400
    try:
        csv_url = get_google_sheet_csv_url(nomination_url)
        if not csv_url:
            return "Invalid Google Sheet URL format.", 400
        df_nomination = pd.read_csv(csv_url)

        processed_rows = []
        for index, nom_row in df_nomination.iterrows():
            pla_id = nom_row['PLA ID']
            matches = df_inventory[df_inventory['PLA ID'] == pla_id]
            selected_inventory_row = matches.iloc[0] if not matches.empty else pd.Series(dtype=object)
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

        excel_data = to_excel_in_memory(df)
        response = make_response(excel_data)
        response.headers['Content-Disposition'] = 'attachment; filename=Final_Assessment.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    except Exception as e:
        return f"An error occurred: {e}", 500

@app.route('/download_master')
def download_master():
    excel_data = to_excel_in_memory(df_inventory)
    response = make_response(excel_data)
    response.headers['Content-Disposition'] = 'attachment; filename=Master_Inventory_Data.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response

# This part is for local development, Render will use the Start Command.
if __name__ == '__main__':
    app.run(debug=True)
