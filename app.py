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

def get_google_sheet_csv_url(url):
    """Transforms a public Google Sheet URL into a direct CSV export link."""
    if "docs.google.com/spreadsheets/d/" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return None

# ==============================================================================
#  STREAMLIT WEB APP INTERFACE
# ==============================================================================

st.set_page_config(page_title="Automated Assessment Tool", layout="centered")
st.title("Transport Network Automated Assessment ⚙️")

# --- DATA LOADING SECTION ---
@st.cache_data(ttl=3600) # Cache the master data for 1 hour
def load_master_inventory_data():
    """Connects to the private Google Sheet and loads the master inventory data."""
    try:
        gsheet_client = authenticate_gsheets()
        # Open the specific spreadsheet by its ID/key
        spreadsheet = gsheet_client.open_by_key('11B6VE-NJI_Xh6SEm7oerIXWoGD45IbEcDbrQmt1uzrQ')
        # Open the specific tab by its name
        worksheet = spreadsheet.worksheet("Merged_Inventory_Data")
        df_inventory = pd.DataFrame(worksheet.get_all_records())
        return df_inventory
    except Exception as e:
        st.error(f"Failed to load master source data. Please ensure the 'Merged_Inventory_Data' sheet is shared with the service account. Error: {e}")
        return None

df_inventory = load_master_inventory_data()

if df_inventory is not None:
    st.success("Successfully loaded the latest master inventory data.")
    # --- Download button for the master data ---
    st.download_button(
        label="📥 Download Master Inventory Data",
        data=to_excel(df_inventory),
        file_name='Master_Inventory_Data.xlsx',
        mime='application/vnd.ms-excel'
    )
else:
    st.stop() # Stop the app if data loading fails

# ==============================================================================
#  Step 1: Process Nomination File from URL
# ==============================================================================
with st.expander("▶️ Step 1: Process Nomination from Google Sheet URL", expanded=True):
    st.markdown("Paste the URL of your nomination Google Sheet below. Ensure it is shared with **'Anyone with the link can view'**.")
    
    nomination_url = st.text_input("Enter your Nomination Google Sheet URL here:")

    if st.button("Process Nomination"):
        if not nomination_url:
            st.warning("Please enter a URL.")
        else:
            with st.spinner('Fetching your nomination data and processing...'):
                try:
                    csv_url = get_google_sheet_csv_url(nomination_url)
                    if csv_url:
                        df_nomination = pd.read_csv(csv_url)
                        
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
                        st.info("You can now proceed to Step 2.")
                    else:
                        st.error("Invalid Google Sheet URL format.")

                except Exception as e:
                    st.error(f"An error occurred. Please check your URL and sharing settings. Error: {e}")

# ==============================================================================
#  Step 2: Run Final Assessment & Download
# ==============================================================================
with st.expander("▶️ Step 2: Run Final Assessment", expanded=True):
    st.markdown("Run the final assessment on your processed nomination data.")

    if st.button("Run Final Assessment"):
        if 'processed_nomination_df' not in st.session_state:
            st.warning("⚠️ Please process a Nomination Sheet URL in Step 1 first.")
        else:
            with st.spinner('Running final assessment...'):
                try:
                    df = st.session_state['processed_nomination_df'].copy()
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

                    st.success("✅ Assessment complete! Your file is ready for download.")
                    
                    st.download_button(
                        label="📥 Download Final Assessment",
                        data=to_excel(df),
                        file_name='Final_Assessment.xlsx',
                        mime='application/vnd.ms-excel'
                    )
                    st.dataframe(df.head())

                except Exception as e:
                    st.error(f"An error occurred during assessment: {e}")
