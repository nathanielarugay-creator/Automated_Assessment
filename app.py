# app.py

import pandas as pd
import streamlit as st
from io import BytesIO
import gspread
from google.oauth2.service_account import Credentials

# ==============================================================================
#  Helper Functions
# ==============================================================================

def to_excel(df):
    """Converts a DataFrame to an in-memory Excel file for downloading."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# Authenticate with Google Sheets using Streamlit's secrets
def authenticate_gsheets():
    """Uses st.secrets to authenticate and return a gspread client."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["google_credentials"], scopes=scopes
    )
    return gspread.authorize(creds)

# ==============================================================================
#  STREAMLIT WEB APP INTERFACE
# ==============================================================================

st.set_page_config(page_title="Automated Assessment Tool", layout="centered")
st.title("Transport Network Automated Assessment ⚙️")

# --- DATA LOADING SECTION ---
@st.cache_data(ttl=3600) # Cache the data for 1 hour
def load_premerged_data():
    """Connects to Google Sheets and loads the pre-merged inventory data."""
    try:
        gsheet_client = authenticate_gsheets()
        # Open the single, pre-merged spreadsheet
        spreadsheet = gsheet_client.open("Merged_Inventory_Data")
        worksheet = spreadsheet.sheet1
        df_inventory = pd.DataFrame(worksheet.get_all_records())
        return df_inventory
    except Exception as e:
        st.error(f"Failed to load source data from 'Merged_Inventory_Data'. Please ensure the sheet exists and is shared correctly. Error: {e}")
        return None

df_inventory = load_premerged_data()

if df_inventory is not None:
    st.success("Successfully loaded the latest merged inventory data.")

    # --- NEW: Download button for the merged source data ---
    st.download_button(
        label="📥 Download Merged Source Data",
        data=to_excel(df_inventory),
        file_name='Merged_Inventory_Data.xlsx',
        mime='application/vnd.ms-excel'
    )
    # --- END NEW ---

else:
    st.stop() # Stop the app if data loading fails

# ==============================================================================
#  Step 1: Process Nomination File
# ==============================================================================
with st.expander("▶️ Step 1: Process Nomination File", expanded=True):
    st.markdown("Upload your `Nomination.xlsx` file. The app will use the latest pre-merged source data.")

    uploaded_nomination_file = st.file_uploader(
        "Upload your 'Nomination.xlsx' file",
        type=['xlsx'],
        key="nomination_uploader"
    )

    if uploaded_nomination_file is not None:
        with st.spinner('Processing nomination...'):
            try:
                df_nomination = pd.read_excel(uploaded_nomination_file)
                processed_rows = []
                for index, nom_row in df_nomination.iterrows():
                    pla_id = nom_row['PLA ID']
                    matches = df_inventory[df_inventory['PLA ID'] == pla_id]
                    selected_inventory_row = matches.iloc[0] if not matches.empty else pd.Series(dtype=object)
                    combined_row = pd.concat([nom_row, selected_inventory_row.add_prefix('Inv_')])
                    processed_rows.append(combined_row)
                
                df_final = pd.DataFrame(processed_rows)
                st.session_state['processed_nomination_df'] = df_final
                st.success("✅ Nomination file processed successfully.")

            except Exception as e:
                st.error(f"An error occurred during processing: {e}")

# ==============================================================================
#  Step 2: Run Final Assessment & Download
# =================================
