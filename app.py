import streamlit as st
import pandas as pd
from io import BytesIO
import time  # For progress bar simulation

# Function to enrich supermarket data and update ITEMMASTER
def process_and_enrich_data(supermarket_df, itemmaster_df, unique_identifier, desc_column, barcode_column=None):
    new_entries = []  # List to collect new entries for ITEMMASTER

    # Iterate over supermarket data
    for idx, row in supermarket_df.iterrows():
        unique_id = row.get(unique_identifier)
        barcode = row.get(barcode_column) if barcode_column else None

        # Check if the unique identifier or barcode exists in ITEMMASTER
        matched_row = itemmaster_df[(itemmaster_df['prodcode'] == unique_id) | 
                                    (itemmaster_df['Barcode(WITH GEN)'] == barcode)]

        if not matched_row.empty:
            # Match found in ITEMMASTER - enrich supermarket data with ITEMMASTER details
            for col in itemmaster_df.columns:
                if col in supermarket_df.columns:
                    supermarket_df.at[idx, col] = matched_row.iloc[0][col]
                else:
                    supermarket_df[col] = matched_row.iloc[0][col]  # Add column if missing in supermarket data
        else:
            # No match found - add this as a new entry for ITEMMASTER
            new_entry = {
                'prodcode': unique_id,
                'DESC': row.get(desc_column),
                'Barcode(WITH GEN)': barcode,
                'Brand': row.get('Brand', None)  # Include any available details
            }
            new_entries.append(new_entry)

    # Append new entries to ITEMMASTER
    if new_entries:
        new_entries_df = pd.DataFrame(new_entries)
        itemmaster_df = pd.concat([itemmaster_df, new_entries_df], ignore_index=True)

    return itemmaster_df, supermarket_df

def remove_duplicates(itemmaster_df):
    # Remove duplicates from ITEMMASTER based on 'prodcode' and 'Barcode(WITH GEN)'
    itemmaster_df.drop_duplicates(subset=['prodcode', 'Barcode(WITH GEN)'], keep='first', inplace=True)
    return itemmaster_df

# Streamlit App
st.title("Supermarket Data Processing App")

# Upload ITEMMASTER
if 'itemmaster_df' not in st.session_state:
    st.session_state.itemmaster_df = None

st.header("Upload ITEMMASTER File")
itemmaster_file = st.file_uploader("Upload ITEMMASTER Excel file", type="xlsx")

if itemmaster_file:
    itemmaster_df = pd.read_excel(itemmaster_file)
    st.session_state.itemmaster_df = itemmaster_df
    st.write("ITEMMASTER loaded successfully")

# Option to keep or upload new ITEMMASTER
if st.session_state.itemmaster_df is not None:
    if st.checkbox("Use existing ITEMMASTER in session"):
        itemmaster_df = st.session_state.itemmaster_df
    else:
        itemmaster_file_new = st.file_uploader("Upload a new ITEMMASTER Excel file", type="xlsx")
        if itemmaster_file_new:
            itemmaster_df = pd.read_excel(itemmaster_file_new)
            st.session_state.itemmaster_df = itemmaster_df

# Select supermarket and upload supermarket data
st.header("Upload Supermarket Data")
supermarket_name = st.selectbox("Select Supermarket", ["Naivas", "Quickmatt", "Carrefour", "Magunas", "Chandarana"])
supermarket_file = st.file_uploader("Upload supermarket Excel file", type="xlsx")

# Initialize a placeholder for the progress bar
progress_bar = st.empty()
start_processing = st.button("Start Processing")

# Process data only when all files are uploaded, and the user clicks the "Start Processing" button
if start_processing and supermarket_file and itemmaster_df is not None:
    with st.spinner("Processing data..."):
        # Display a progress bar
        progress_bar.progress(0)
        
        # Load the supermarket data
        supermarket_df = pd.read_excel(supermarket_file)
        st.write(f"{supermarket_name} data loaded successfully")

        # Map supermarket-specific columns and process data
        if supermarket_name == "Naivas":
            itemmaster_df, supermarket_df = process_and_enrich_data(supermarket_df, itemmaster_df, unique_identifier='Itemid', desc_column='Itemname')
        elif supermarket_name == "Quickmatt":
            itemmaster_df, supermarket_df = process_and_enrich_data(supermarket_df, itemmaster_df, unique_identifier='ITEM_CODE', desc_column='ITEM_NAME', barcode_column='BARCODE')
        elif supermarket_name == "Carrefour":
            itemmaster_df, supermarket_df = process_and_enrich_data(supermarket_df, itemmaster_df, unique_identifier='Item Code', desc_column='Item Name', barcode_column='Item Bar Code')
        elif supermarket_name == "Magunas":
            # Custom handling for Magunas' SKU description split
            for idx, row in supermarket_df.iterrows():
                sku_description = row.get('SKU-DESCRIPTION')
                if pd.notna(sku_description):
                    item_code, *description_parts = sku_description.split('-')
                    item_code = item_code.strip()
                    description = '-'.join(description_parts).strip()
                    supermarket_df.at[idx, 'Itemid'] = item_code
                    supermarket_df.at[idx, 'Itemname'] = description
            itemmaster_df, supermarket_df = process_and_enrich_data(supermarket_df, itemmaster_df, unique_identifier='Itemid', desc_column='Itemname')
        elif supermarket_name == "Chandarana":
            itemmaster_df, supermarket_df = process_and_enrich_data(supermarket_df, itemmaster_df, unique_identifier='Item Name', desc_column='Item Name', barcode_column='Barcode')

        # Remove duplicates in ITEMMASTER after adding new entries
        itemmaster_df = remove_duplicates(itemmaster_df)
        progress_bar.progress(100)  # Complete the progress bar

        # Update session with the latest ITEMMASTER
        st.session_state.itemmaster_df = itemmaster_df
        st.success("Processing complete!")

        # Display processed ITEMMASTER and supermarket data with enriched details
        st.write("Processed ITEMMASTER Data")
        st.dataframe(itemmaster_df)
        
        st.write(f"Enriched {supermarket_name} Data")
        st.dataframe(supermarket_df)

        # Download options for ITEMMASTER and enriched supermarket data
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            itemmaster_df.to_excel(writer, index=False, sheet_name='Updated ITEMMASTER')
        st.download_button(
            label="Download Processed ITEMMASTER",
            data=buffer.getvalue(),
            file_name="processed_itemmaster.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        supermarket_buffer = BytesIO()
        with pd.ExcelWriter(supermarket_buffer, engine='xlsxwriter') as writer:
            supermarket_df.to_excel(writer, index=False, sheet_name=f'Enriched {supermarket_name}')
        st.download_button(
            label=f"Download Enriched {supermarket_name} Data",
            data=supermarket_buffer.getvalue(),
            file_name=f"enriched_{supermarket_name.lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.warning("Please upload both the ITEMMASTER and supermarket files and then click 'Start Processing' to continue.")
