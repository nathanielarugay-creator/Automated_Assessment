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

# Authenticate once and cache the client
try:
    gsheet_client = authenticate_gsheets()
    st.info("Successfully connected to Google Drive.")
except Exception as e:
    st.error("Failed to connect to Google Sheets. Please check your secrets configuration.")
    st.stop() # Stop the app if authentication fails

# ==============================================================================
#  Step 1: Process Nomination File
# ==============================================================================
with st.expander("▶️ Step 1: Process Nomination File", expanded=True):
    st.markdown("Upload your `Nomination.xlsx` file. The app will automatically fetch the latest source data from the private Google Drive.")

    uploaded_nomination_file = st.file_uploader(
        "Upload your 'Nomination.xlsx' file",
        type=['xlsx'],
        key="nomination_uploader"
    )

    if uploaded_nomination_file is not None:
        with st.spinner('Fetching source data and processing nomination...'):
            try:
                # --- AUTOMATED DATA FETCHING ---
                # Open the single spreadsheet by its name
                spreadsheet = gsheet_client.open("Automated_Assessment_Source")

                # Access each tab (worksheet) by its name
                wireless_sheet = spreadsheet.worksheet("Wireless")
                wireline_sheet = spreadsheet.worksheet("Wireline")
                port_an_sheet = spreadsheet.worksheet("Port_AN")
                port_ag_sheet = spreadsheet.worksheet("Port_AG")
                
                # Convert to DataFrames
                df_wireless = pd.DataFrame(wireless_sheet.get_all_records())
                df_wireline = pd.DataFrame(wireline_sheet.get_all_records())
                df_port_an = pd.DataFrame(port_an_sheet.get_all_records())
                df_port_ag = pd.DataFrame(port_ag_sheet.get_all_records())

                # Task 1: Merge Inventories
                combined_df = pd.concat([df_wireless, df_wireline], ignore_index=True)
                combined_df.drop_duplicates(subset=['Transport NE'], keep='first', inplace=True)
                
                final_columns = [
                    'Transport NE', 'PLA ID', 'Site Name', 'Territory', 'Network Type', 'Equipment Type/Model', 
                    'Equipment Status', 'AN CONFIG TYPE', 'LOOP NAME', 'MYCOM LOOP CATEGORY', 'LOOP CAPACITY CATEGORY', 
                    'MYCOM GW1 CAPACITY (GBPS)', 'MYCOM GW2 CAPACITY (GBPS)', 'MYCOM LOOP NORMAL UTILIZATION', 
                    'MYCOM LOOP NORMAL STATUS', 'MYCOM LOOP OUTAGE UTILIZATION', 'MYCOM LOOP OUTAGE STATUS', 
                    'AG1 HOMING NE NAME', 'AG2 HOMING NE NAME'
                ]
                existing_columns = [col for col in final_columns if col in combined_df.columns]
                df_main = combined_df[existing_columns]

                # Task 2: Add Port Data
                df_ports = pd.concat([df_port_an, df_port_ag], ignore_index=True)
                columns_to_copy = ['NE_Name', 'GE_1G', 'Total_of_GE_1G', 'GE_10G', 'Total_of_GE_10G', '25GE']
                port_data_to_merge = df_ports[columns_to_copy].drop_duplicates(subset=['NE_Name'], keep='first')
                df_inventory = pd.merge(df_main, port_data_to_merge, left_on='Transport NE', right_on='NE_Name', how='left').drop(columns=['NE_Name'])

                # --- NOMINATION PROCESSING ---
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
                st.success("✅ Nomination file processed successfully with the latest source data.")

            except Exception as e:
                st.error(f"An error occurred during processing: {e}")

# ==============================================================================
#  Step 2: Run Final Assessment & Download
# ==============================================================================
with st.expander("▶️ Step 2: Run Final Assessment", expanded=True):
    st.markdown("Run the final assessment on the processed nomination file.")

    if st.button("Run Final Assessment"):
        if 'processed_nomination_df' not in st.session_state:
            st.warning("⚠️ Please upload and process a Nomination file in Step 1 first.")
        else:
            with st.spinner('Running final assessment...'):
                try:
                    df = st.session_state['processed_nomination_df'].copy()
                    numeric_cols = ['GE Port Demand', '10GE Port Demand', 'Inv_GE_1G', 'Inv_GE_10G', 'Inv_MYCOM LOOP NORMAL UTILIZATION']
                    
                    if df['Inv_MYCOM LOOP NORMAL UTILIZATION'].dtype == 'object':
                        df['Inv_MYCOM LOOP NORMAL UTILIZATION'] = df['Inv_MYCOM LOOP NORMAL UTILIZATION'].str.replace('%', '', regex=False).astype(float) / 100
                    for col in numeric_cols:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    
                    def get_node_assessment(row):
                        failures = []
                        if row['GE Port Demand'] >= 1 and (row['Inv_GE_1G'] - row['GE Port Demand']) <= 2: failures.append("Requires Port Augmentation")
                        if row['10GE Port Demand'] >= 1 and (row['Inv_GE_10G'] - row['10GE Port Demand']) <= 2: failures.append("Requires Port Augmentation")
                        return " & ".join(failures) if failures else ("With Headroom" if row['GE Port Demand'] >= 1 or row['10GE Port Demand'] >= 1 else "No Port Demand")

                    def get_loop_assessment(row):
                        return "Requires Loop Upgrade" if row['Inv_MYCOM LOOP NORMAL UTILIZATION'] >= 0.7 else "With Headroom"

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
