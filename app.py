# app.py

import pandas as pd
import streamlit as st
from io import BytesIO

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

def get_google_sheet_csv_url(url):
    """Transforms a Google Sheet URL into a direct CSV export link."""
    if "docs.google.com/spreadsheets/d/" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return None

# ==============================================================================
#  STREAMLIT WEB APP INTERFACE
# ==============================================================================

st.set_page_config(page_title="Automated Assessment Tool", layout="centered")

st.title("Transport Network Automated Assessment ⚙️")

st.markdown("""
### **Pre-requisites:**
1.  Ensure your source Google Sheets are shared with **"Anyone with the link can view"**.
2.  Use the provided **Nomination.xlsx** template file for your nominations.
---
""")

# ==============================================================================
#  NEW: Step 1: Load Source Data from Google Sheets
# ==============================================================================
with st.expander("▶️ Step 1: Load Source Data from Google Sheets", expanded=True):
    st.markdown("Paste the share links for your Google Sheets below.")
    
    url_inv_wireless = st.text_input("Wireless Inventory Sheet URL")
    url_inv_wireline = st.text_input("Wireline Inventory Sheet URL")
    url_port_1 = st.text_input("Port Inventory Sheet 1 URL (e.g., AN)")
    url_port_2 = st.text_input("Port Inventory Sheet 2 URL (e.g., AG)")

    if st.button("1. Load and Prepare Source Data"):
        urls = [url_inv_wireless, url_inv_wireline, url_port_1, url_port_2]
        if not all(urls):
            st.error("❌ Please provide all four Google Sheet URLs.")
        else:
            with st.spinner('Fetching and processing data from Google Sheets...'):
                try:
                    # Task 1: Merge Inventories
                    df1 = pd.read_csv(get_google_sheet_csv_url(url_inv_wireless))
                    df2 = pd.read_csv(get_google_sheet_csv_url(url_inv_wireline))
                    combined_df = pd.concat([df1, df2], ignore_index=True)
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
                    st.success("✅ Raw inventories merged successfully.")

                    # Task 2: Add Port Data
                    df_port1 = pd.read_csv(get_google_sheet_csv_url(url_port_1))
                    df_port2 = pd.read_csv(get_google_sheet_csv_url(url_port_2))
                    df_ports = pd.concat([df_port1, df_port2], ignore_index=True)
                    
                    columns_to_copy = ['NE_Name', 'GE_1G', 'Total_of_GE_1G', 'GE_10G', 'Total_of_GE_10G', '25GE']
                    if not all(col in df_ports.columns for col in columns_to_copy):
                        missing = set(columns_to_copy) - set(df_ports.columns)
                        st.error(f"Port files are missing required columns: {missing}")
                    else:
                        port_data_to_merge = df_ports[columns_to_copy].drop_duplicates(subset=['NE_Name'], keep='first')
                        df_updated = pd.merge(df_main, port_data_to_merge, left_on='Transport NE', right_on='NE_Name', how='left').drop(columns=['NE_Name'])
                        
                        st.session_state['inventory_with_ports_df'] = df_updated
                        st.success("✅ Port data added. The prepared inventory is ready!")
                        st.info("You can now proceed to Step 2 below.")
                        
                except Exception as e:
                    st.error(f"An error occurred: {e}")
                    st.warning("Please check that your URLs are correct and the sheets are publicly shared.")


# ==============================================================================
#  CHANGED: Step 2 (was 3): Process Nomination File
# ==============================================================================
with st.expander("▶️ Step 2: Process Nomination File"):
    st.markdown("Upload your `Nomination.xlsx` file. This will be merged with the prepared data from Step 1.")
    
    uploaded_nomination_file = st.file_uploader(
        "Upload your 'Nomination.xlsx' file",
        type=['xlsx'],
        key="task3_uploader"
    )

    if st.button("2. Process Nomination"):
        if 'inventory_with_ports_df' not in st.session_state:
            st.warning("⚠️ Please complete Step 1 first by clicking the 'Load and Prepare' button.")
        elif uploaded_nomination_file is None:
            st.error("❌ Please upload the nomination file.")
        else:
            with st.spinner('Processing nominations...'):
                try:
                    df_inventory = st.session_state['inventory_with_ports_df']
                    df_nomination = pd.read_excel(uploaded_nomination_file)
                    processed_rows = []

                    for index, nom_row in df_nomination.iterrows():
                        pla_id = nom_row['PLA ID']
                        matches = df_inventory[df_inventory['PLA ID'] == pla_id]
                        
                        if not matches.empty:
                            selected_inventory_row = matches.iloc[0]
                        else:
                            selected_inventory_row = pd.Series(dtype=object)

                        combined_row = pd.concat([nom_row, selected_inventory_row.add_prefix('Inv_')])
                        processed_rows.append(combined_row)
                    
                    df_final = pd.DataFrame(processed_rows)
                    st.session_state['processed_nomination_df'] = df_final
                    st.success("✅ Step 2 complete! The nomination file has been processed.")

                except Exception as e:
                    st.error(f"An error occurred: {e}")

# ==============================================================================
#  CHANGED: Step 3 (was 4): Run Final Assessment & Download
# ==============================================================================
with st.expander("▶️ Step 3: Run Final Assessment"):
    st.markdown("Run the final assessment on the processed nomination file from Step 2 and download the result.")

    if st.button("3. Run Assessment"):
        if 'processed_nomination_df' not in st.session_state:
            st.warning("⚠️ Please complete Step 2 first.")
        else:
            with st.spinner('Running final assessment...'):
                try:
                    df = st.session_state['processed_nomination_df'].copy()

                    numeric_cols = ['GE Port Demand', '10GE Port Demand', 'Inv_GE_1G', 'Inv_GE_10G', 'Inv_MYCOM LOOP NORMAL UTILIZATION']
                    if not all(col in df.columns for col in numeric_cols):
                        missing = set(numeric_cols) - set(df.columns)
                        st.error(f"Processed file is missing required columns: {missing}")
                    else:
                        if df['Inv_MYCOM LOOP NORMAL UTILIZATION'].dtype == 'object':
                            df['Inv_MYCOM LOOP NORMAL UTILIZATION'] = df['Inv_MYCOM LOOP NORMAL UTILIZATION'].str.replace('%', '', regex=False).astype(float) / 100
                        for col in numeric_cols:
                            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        
                        def get_node_assessment(row):
                            failures = []
                            if row['GE Port Demand'] >= 1 and (row['Inv_GE_1G'] - row['GE Port Demand']) <= 2: failures.append("Requires Port Augmentation")
                            if row['10GE Port Demand'] >= 1 and (row['Inv_GE_10G'] - row['10GE Port Demand']) <= 2: failures.append("Requires Port Augmentation")
                            if not failures: return "With Headroom" if row['GE Port Demand'] >= 1 or row['10GE Port Demand'] >= 1 else "No Port Demand"
                            return " & ".join(failures)

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
                    st.error(f"An error occurred: {e}")
