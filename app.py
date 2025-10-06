# app.py

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, HttpUrl

# --- Configuration & Helper Functions ---

# Define the data model for the incoming request
class NominationRequest(BaseModel):
    nomination_url: HttpUrl # FastAPI will automatically validate that this is a valid URL

# Authenticate with Google Sheets using a secret file
def authenticate_gsheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
    return gspread.authorize(creds)

# Transforms a public Google Sheet URL into a direct CSV export link
def get_google_sheet_csv_url(url: str):
    if "docs.google.com/spreadsheets/d/" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return None

# --- Data Loading (runs once on startup) ---

@cache
def load_master_inventory_data():
    """Connects to the private Google Sheet and loads the master inventory data."""
    try:
        gsheet_client = authenticate_gsheets()
        spreadsheet = gsheet_client.open_by_key('11B6VE-NJI_Xh6SEm7oerIXWoGD45IbEcDbrQmt1uzrQ')
        worksheet = spreadsheet.worksheet("Merged_Inventory_Data")
        return pd.DataFrame(worksheet.get_all_records())
    except Exception as e:
        # If the master data fails to load, the service cannot start.
        # In a real-world scenario, you'd add more robust logging here.
        raise RuntimeError(f"CRITICAL: Failed to load master inventory data on startup. Error: {e}")

df_inventory = load_master_inventory_data()

# --- FastAPI App Definition ---

app = FastAPI(
    title="Automated Assessment API",
    description="An API to process network nominations against a master inventory."
)

@app.post("/assess")
async def process_assessment(request: NominationRequest):
    """
    Accepts a nomination Google Sheet URL, processes it against the master inventory,
    and returns the final assessment as JSON.
    """
    try:
        # --- Process Nomination Data ---
        csv_url = get_google_sheet_csv_url(str(request.nomination_url))
        if not csv_url:
            raise HTTPException(status_code=400, detail="Invalid Google Sheet URL format.")
        
        df_nomination = pd.read_csv(csv_url)

        # --- Merge and Assess (Core Logic) ---
        processed_rows = []
        for index, nom_row in df_nomination.iterrows():
            pla_id = nom_row['PLA ID']
            matches = df_inventory[df_inventory['PLA ID'] == pla_id]
            selected_inventory_row = matches.iloc[0] if not matches.empty else pd.Series(dtype=object)
            combined_row = pd.concat([nom_row, selected_inventory_row.add_prefix('Inv_')])
            processed_rows.append(combined_row)
        
        df = pd.DataFrame(processed_rows)

        # --- Run Assessment Logic ---
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

        # Convert DataFrame to a list of dictionaries for JSON response
        result = df.to_dict(orient='records')
        return result

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An internal error occurred: {str(e)}")

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Automated Assessment API is running."}
