# app.py

import pandas as pd
import streamlit as st
from io import BytesIO
import webbrowser

# ==============================================================================
#  Helper function to create a download link for DataFrames
# ==============================================================================
def to_excel(df):
    """Converts a DataFrame to an in-memory Excel file for downloading."""
    output = BytesIO()
    # Use xlsxwriter engine for compatibility
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# ==============================================================================
#  STREAMLIT WEB APP INTERFACE
# ==============================================================================

# Set the page title and layout
st.set_page_config(page_title="Automated Assessment Tool", layout="centered")

# --- App Title and Instructions ---
st.title("Transport Network Automated Assessment ⚙️")

st.markdown("""
### **Pre-requisites:**
1.  Download **Wireless** and **Wireline** Transport Details from the [AppSheet Link](https://www.appsheet.com/start/38f2ed05-fba9-408f-9a7a-abfa531a44ee).
2.  Download **XDB Port Inventory** for AN and AG from the TACTICS app.
3.  Use the provided **Nomination.xlsx** template file for your nominations.
---
""")


# ==============================================================================
#  TASK 1: Merge Raw Wireless and Wireline Inventory Files
# ==============================================================================
with st.expander("▶️ Step 1: Merge Inventory Files", expanded=True):
    st.markdown("Select the two raw inventory files (`Wireless.xlsx` and `Wireline.xlsx`) to combine them.")
    
    uploaded_inventory_files = st.file_uploader(
        "Upload the TWO source inventory files",
        type=['xlsx'],
        accept_multiple_files=True,
        key="task1_uploader"
    )

    if st.button("1. Merge Inventory", key="task1_button"):
        if len(uploaded_inventory_files) != 2:
            st.error("❌ Please upload exactly two inventory files to merge.")
        else:
            with st.spinner('Merging files...'):
                try:
                    df1 = pd.read_excel(uploaded_inventory_files[0])
                    df2 = pd.read_excel(uploaded_inventory_files[1])
                    combined_df = pd.concat([df1, df2], ignore_index=True)
                    
                    st.write(f"Original row count: `{len(combined_df)}`")
                    combined_df.drop_duplicates(subset=['Transport NE'], keep='first', inplace=True)
                    st.write(f"Row count after removing duplicates: `{len(combined_df)}`")

                    final_columns = [
                        'Transport NE', 'PLA ID', 'Site Name', 'Territory', 'Network Type', 'Equipment Type/Model', 
                        'Equipment Status', 'AN CONFIG TYPE', 'LOOP NAME', 'MYCOM LOOP CATEGORY', 'LOOP CAPACITY CATEGORY', 
                        'MYCOM GW1 CAPACITY (GBPS)', 'MYCOM GW2 CAPACITY (GBPS)', 'MYCOM LOOP NORMAL UTILIZATION', 
                        'MYCOM LOOP NORMAL STATUS', 'MYCOM LOOP OUTAGE UTILIZATION', 'MYCOM LOOP OUTAGE STATUS', 
                        'AG1 HOMING NE NAME', 'AG2 HOMING NE NAME'
                    ]
                    existing_columns = [col for col in final_columns if col in combined_df.columns]
                    final_df = combined_df[existing_columns]
                    
                    # Store the result in session state for the next step
                    st.session_state['merged_inventory_df'] = final_df
                    st.success("✅ Step 1 complete! The merged inventory is ready.")

                except Exception as e:
                    st.error(f"An error occurred during Task 1: {e}")

# ==============================================================================
#  TASK 2: Add Port Data to the Main Inventory
# ==============================================================================
with st.expander("▶️ Step 2: Add Port Data"):
    st.markdown("Upload the two port inventory files to add their data to the merged inventory from Step 1.")

    uploaded_port_files = st.file_uploader(
        "Upload the TWO port inventory files",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key="task2_uploader"
    )

    if st.button("2. Add Port Data", key="task2_button"):
        if 'merged_inventory_df' not in st.session_state:
            st.warning("⚠️ Please complete Step 1 first.")
        elif len(uploaded_port_files) != 2:
            st.error("❌ Please upload exactly two port files.")
        else:
            with st.spinner('Adding port data...'):
                try:
                    df_main = st.session_state['merged_inventory_df']
                    df_ports = pd.concat([pd.read_excel(f) for f in uploaded_port_files], ignore_index=True)
                    
                    columns_to_copy = ['NE_Name', 'GE_1G', 'Total_of_GE_1G', 'GE_10G', 'Total_of_GE_10G', '25GE']
                    if not all(col in df_ports.columns for col in columns_to_copy):
                        missing = set(columns_to_copy) - set(df_ports.columns)
                        st.error(f"Port files are missing required columns: {missing}")
                    else:
                        port_data_to_merge = df_ports[columns_to_copy].drop_duplicates(subset=['NE_Name'], keep='first')
                        df_updated = pd.merge(df_main, port_data_to_merge, left_on='Transport NE', right_on='NE_Name', how='left').drop(columns=['NE_Name'])
                        
                        # Store for the next step
                        st.session_state['inventory_with_ports_df'] = df_updated
                        st.success("✅ Step 2 complete! The inventory has been updated with port data.")
                        
                except Exception as e:
                    st.error(f"An error occurred during Task 2: {e}")

    if 'inventory_with_ports_df' in st.session_state:
        st.download_button(
            label="📥 Download Inventory (Optional)",
            data=to_excel(st.session_state['inventory_with_ports_df']),
            file_name='Inventory_with_Ports.xlsx',
            mime='application/vnd.ms-excel'
        )

# ==============================================================================
#  TASK 3: Process Nomination File
# ==============================================================================
with st.expander("▶️ Step 3: Process Nomination File"):
    st.markdown("Upload your `Nomination.xlsx` file. This will be merged with the result from Step 2.")
    
    uploaded_nomination_file = st.file_uploader(
        "Upload your 'Nomination.xlsx' file",
        type=['xlsx'],
        key="task3_uploader"
    )

    if st.button("3. Process Nomination", key="task3_button"):
        if 'inventory_with_ports_df' not in st.session_state:
            st.warning("⚠️ Please complete Step 2 first.")
        elif uploaded_nomination_file is None:
            st.error("❌ Please upload the nomination file.")
        else:
            with st.spinner('Processing nominations... This may take a moment.'):
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
                    
                    # Store for the final step
                    st.session_state['processed_nomination_df'] = df_final
                    st.success("✅ Step 3 complete! The nomination file has been processed.")

                except Exception as e:
                    st.error(f"An error occurred during Task 3: {e}")

# ==============================================================================
#  TASK 4: Run Final Assessment & Download
# ==============================================================================
with st.expander("▶️ Step 4: Run Final Assessment"):
    st.markdown("Run the final assessment on the processed nomination file from Step 3 and download the result.")

    if st.button("4. Run Assessment", key="task4_button"):
        if 'processed_nomination_df' not in st.session_state:
            st.warning("⚠️ Please complete Step 3 first.")
        else:
            with st.spinner('Running final assessment...'):
                try:
                    df = st.session_state['processed_nomination_df'].copy()

                    numeric_cols = ['GE Port Demand', '10GE Port Demand', 'Inv_GE_1G', 'Inv_GE_10G', 'Inv_MYCOM LOOP NORMAL UTILIZATION']
                    if not all(col in df.columns for col in numeric_cols):
                        missing = set(numeric_cols) - set(df.columns)
                        st.error(f"Processed file is missing required columns for assessment: {missing}")
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
                            label="📥 Download Final Assessment (Step 4)",
                            data=to_excel(df),
                            file_name='Final_Assessment.xlsx',
                            mime='application/vnd.ms-excel'
                        )
                        st.dataframe(df.head())

                except Exception as e:
                    st.error(f"An error occurred during Task 4: {e}")
