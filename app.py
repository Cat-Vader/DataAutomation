import streamlit as st
import pandas as pd
from io import BytesIO
import time  # For progress bar simulation

# Function to process each supermarket and update ITEMMASTER
def process_supermarket_data(supermarket_df, itemmaster_df, unique_identifier, desc_column, barcode_column=None, brand_column=None):
    new_entries = []
    for idx, row in supermarket_df.iterrows():
        unique_id = row.get(unique_identifier)
        description = row.get(desc_column)
        barcode = row.get(barcode_column) if barcode_column else None
        brand = row.get(brand_column) if brand_column else None
        if unique_id not in itemmaster_df['prodcode'].values:
            if barcode is None or barcode not in itemmaster_df['Barcode(WITH GEN)'].values:
                new_entry = {
                    'prodcode': unique_id,
                    'DESC': description,
                    'Barcode(WITH GEN)': barcode,
                    'Brand': brand
                }
                new_entries.append(new_entry)
    if new_entries:
        new_entries_df = pd.DataFrame(new_entries)
        itemmaster_df = pd.concat([itemmaster_df, new_entries_df], ignore_index=True)
    return itemmaster_df

def remove_duplicates(itemmaster_df):
    itemmaster_df.drop_duplicates(subset=['prodcode', 'Barcode(WITH GEN)'], keep='first', inplace=True)
    return itemmaster_df

def complete_missing_fields(itemmaster_df, supermarket_df):
    for idx, row in supermarket_df.iterrows():
        unique_id = row.get('Itemid') or row.get('ITEM_CODE') or row.get('Item Code') or row.get('SKU') or row.get('Barcode')
        if pd.notna(unique_id):
            mask = (itemmaster_df['prodcode'] == unique_id) | (itemmaster_df['Barcode(WITH GEN)'] == row.get('Barcode'))
            if mask.any():
                itemmaster_idx = itemmaster_df.index[mask][0]
                if pd.isna(itemmaster_df.at[itemmaster_idx, 'Brand']) and row.get('Brand'):
                    itemmaster_df.at[itemmaster_idx, 'Brand'] = row.get('Brand')
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

        # Processing data based on the selected supermarket
        if supermarket_name == "Naivas":
            itemmaster_df = process_supermarket_data(supermarket_df, itemmaster_df, unique_identifier='Itemid', desc_column='Itemname')
        elif supermarket_name == "Quickmatt":
            itemmaster_df = process_supermarket_data(supermarket_df, itemmaster_df, unique_identifier='ITEM_CODE', desc_column='ITEM_NAME', barcode_column='BARCODE')
        elif supermarket_name == "Carrefour":
            itemmaster_df = process_supermarket_data(supermarket_df, itemmaster_df, unique_identifier='Item Code', desc_column='Item Name', barcode_column='Item Bar Code')
        elif supermarket_name == "Magunas":
            for idx, row in supermarket_df.iterrows():
                sku_description = row.get('SKU-DESCRIPTION')
                if pd.notna(sku_description):
                    item_code, *description_parts = sku_description.split('-')
                    item_code = item_code.strip()
                    description = '-'.join(description_parts).strip()
                    if item_code not in itemmaster_df['prodcode'].values:
                        new_entry = {
                            'prodcode': item_code,
                            'DESC': description,
                            'Barcode(WITH GEN)': None,
                            'Brand': None
                        }
                        new_entries = [new_entry]
                        new_entries_df = pd.DataFrame(new_entries)
                        itemmaster_df = pd.concat([itemmaster_df, new_entries_df], ignore_index=True)
        elif supermarket_name == "Chandarana":
            for idx, row in supermarket_df.iterrows():
                item_name = row.get('Item Name')
                barcode = row.get('Barcode')
                if barcode not in itemmaster_df['Barcode(WITH GEN)'].values:
                    new_entry = {
                        'prodcode': None,
                        'DESC': item_name,
                        'Barcode(WITH GEN)': barcode,
                        'Brand': None
                    }
                    new_entries = [new_entry]
                    new_entries_df = pd.DataFrame(new_entries)
                    itemmaster_df = pd.concat([itemmaster_df, new_entries_df], ignore_index=True)

        # Remove duplicates
        itemmaster_df = remove_duplicates(itemmaster_df)
        progress_bar.progress(50)  # Update progress bar

        # Complete missing fields if necessary
        itemmaster_df = complete_missing_fields(itemmaster_df, supermarket_df)
        progress_bar.progress(100)  # Update progress bar

        # Update session with the latest ITEMMASTER
        st.session_state.itemmaster_df = itemmaster_df
        st.success("Processing complete!")

        # Display processed ITEMMASTER and download option
        st.write("Processed ITEMMASTER Data")
        st.dataframe(itemmaster_df)

        # Create an in-memory bytes buffer for the Excel file
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            st.session_state.itemmaster_df.to_excel(writer, index=False, sheet_name='Updated ITEMMASTER')
        
        # Download processed ITEMMASTER
        st.download_button(
            label="Download Processed ITEMMASTER",
            data=buffer.getvalue(),
            file_name="processed_itemmaster.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.warning("Please upload both the ITEMMASTER and supermarket files and then click 'Start Processing' to continue.")
