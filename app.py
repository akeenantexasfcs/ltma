#!/usr/bin/env python
# coding: utf-8

# In[1]:


import io
import json
import os
import pandas as pd
import streamlit as st
from Levenshtein import distance as levenshtein_distance
import re

# Define the initial lookup data
initial_lookup_data = {
    "Account": ["Cash and cash equivalents", "Line of credit", "Goodwill", 
                "Total Current Assets", "Total Assets", "Total Current Liabilities"],
    "Mnemonic": ["Cash & Cash Equivalents", "Short-Term Debt", "Goodwill", 
                 "Total Current Assets", "Total Assets", "Total Current Liabilities"],
    "CIQ": ["IQ_CASH_EQUIV", "IQ_ST_INVEST", "IQ_GW", 
            "IQ_TOTAL_CA", "IQ_TOTAL_ASSETS", "IQ_TOTAL_CL"]
}

# Define the file path for the data dictionary CSV file
data_dictionary_file = 'data_dictionary.csv'

# Load or initialize the lookup table
def load_or_initialize_lookup():
    if os.path.exists(data_dictionary_file):
        lookup_df = pd.read_csv(data_dictionary_file)
    else:
        lookup_df = pd.DataFrame(initial_lookup_data)
        lookup_df.to_csv(data_dictionary_file, index=False)
    return lookup_df

def save_lookup_table(df):
    df.to_csv(data_dictionary_file, index=False)

lookup_df = load_or_initialize_lookup()

def process_file(file):
    try:
        df = pd.read_excel(file, sheet_name=None)
        # Assuming the relevant sheet is the first one
        first_sheet_name = list(df.keys())[0]
        df = df[first_sheet_name]
        return df
    except Exception as e:
        st.error(f"Error processing file {file.name}: {e}")
        return None

def create_combined_df(dfs):
    combined_df = pd.DataFrame()
    for i, df in enumerate(dfs):
        final_mnemonic_col = 'Final Mnemonic Selection'
        if final_mnemonic_col not in df.columns:
            st.error(f"Column '{final_mnemonic_col}' not found in dataframe {i+1}")
            continue
        
        date_cols = [col for col in df.columns if col not in ['Label', 'Account', final_mnemonic_col, 'Mnemonic', 'Manual Selection']]
        if not date_cols:
            st.error(f"No date columns found in dataframe {i+1}")
            continue

        df_grouped = df.groupby([final_mnemonic_col, 'Label']).sum(numeric_only=True).reset_index()
        df_melted = df_grouped.melt(id_vars=[final_mnemonic_col, 'Label'], value_vars=date_cols, var_name='Date', value_name='Value')
        df_pivot = df_melted.pivot(index=['Label', final_mnemonic_col], columns='Date', values='Value')
        
        if combined_df.empty:
            combined_df = df_pivot
        else:
            combined_df = combined_df.join(df_pivot, how='outer')
    return combined_df.reset_index()

def aggregate_data(df):
    # Check for the presence of 'Label' and 'Account' columns dynamically
    if 'Label' not in df.columns or 'Account' not in df.columns:
        st.error("'Label' and/or 'Account' columns not found in the data.")
        return df
    
    # Example aggregation function: Pivoting the data
    pivot_table = df.pivot_table(index=['Label', 'Account'], 
                                 values=[col for col in df.columns if col not in ['Label', 'Account', 'Mnemonic', 'Manual Selection']], 
                                 aggfunc='sum').reset_index()
    return pivot_table

def clean_numeric_value(value):
    """
    Clean the given value to convert it to a numeric format.
    Removes special characters like $, commas, and parentheses.
    Converts to a float and handles negative numbers correctly.
    """
    value_str = str(value).strip()
    
    # Handle parentheses indicating negative numbers
    if value_str.startswith('(') and value_str.endswith(')'):
        value_str = '-' + value_str[1:-1]
    
    # Remove dollar signs and commas
    cleaned_value = re.sub(r'[$,]', '', value_str)
    
    try:
        return float(cleaned_value)
    except ValueError:
        return 0  # Return 0 if conversion fails

def sort_by_label(df):
    sort_order = {
        "Current Assets": 0,
        "Non Current Assets": 1,
        "Current Liabilities": 2,
        "Non Current Liabilities": 3,
        "Equity": 4,
        "Total Equity and Liabilities": 5
    }
    
    df['Label_Order'] = df['Label'].map(sort_order)
    df['Total_Order'] = df['Final Mnemonic Selection'].str.contains('Total', case=False).astype(int)
    
    df = df.sort_values(by=['Label_Order', 'Total_Order', 'Final Mnemonic Selection']).drop(columns=['Label_Order', 'Total_Order'])
    return df

def apply_unit_conversion(df, columns, factor):
    for selected_column in columns:
        if selected_column in df.columns:
            df[selected_column] = df[selected_column].apply(
                lambda x: x * factor if isinstance(x, (int, float)) else x)
    return df

def main():
    global lookup_df

    st.title("Table Extractor and Label Generators")

    # Define the tabs
    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Mnemonic Mapping", "Balance Sheet Data Dictionary", "Data Aggregation"])

    with tab1:
        uploaded_file = st.file_uploader("Choose a JSON file", type="json", key='json_uploader')
        if uploaded_file is not None:
            data = json.load(uploaded_file)
            tables = []
            for block in data['Blocks']:
                if block['BlockType'] == 'TABLE':
                    table = {}
                    if 'Relationships' in block:
                        for relationship in block['Relationships']:
                            if relationship['Type'] == 'CHILD':
                                for cell_id in relationship['Ids']:
                                    cell_block = next((b for b in data['Blocks'] if b['Id'] == cell_id), None)
                                    if cell_block:
                                        row_index = cell_block.get('RowIndex', 0)
                                        col_index = cell_block.get('ColumnIndex', 0)
                                        if row_index not in table:
                                            table[row_index] = {}
                                        cell_text = ''
                                        if 'Relationships' in cell_block:
                                            for rel in cell_block['Relationships']:
                                                if rel['Type'] == 'CHILD':
                                                    for word_id in rel['Ids']:
                                                        word_block = next((w for w in data['Blocks'] if w['Id'] == word_id), None)
                                                        if word_block and word_block['BlockType'] == 'WORD':
                                                            cell_text += ' ' + word_block.get('Text', '')
                                        table[row_index][col_index] = cell_text.strip()
                    table_df = pd.DataFrame.from_dict(table, orient='index').sort_index()
                    table_df = table_df.sort_index(axis=1)
                    tables.append(table_df)
            all_tables = pd.concat(tables, axis=0, ignore_index=True)
            if len(all_tables.columns) == 0:
                st.error("No columns found in the uploaded JSON file.")
                return

            column_a = all_tables.columns[0]
            all_tables.insert(0, 'Label', '')

            st.subheader("Data Preview")
            st.dataframe(all_tables)

            labels = ["Current Assets", "Non Current Assets", "Current Liabilities", 
                      "Non Current Liabilities", "Equity", "Total Equity and Liabilities"]
            selections = []

            for label in labels:
                st.subheader(f"Setting bounds for {label}")
                options = [''] + list(all_tables[column_a].dropna().unique())
                start_label = st.selectbox(f"Start Label for {label}", options, key=f"start_{label}")
                end_label = st.selectbox(f"End Label for {label}", options, key=f"end_{label}")
                selections.append((label, start_label, end_label))

            # Initialize new_column_names before the function is called
            new_column_names = {col: col for col in all_tables.columns}

            def update_labels(df):
                df['Label'] = ''
                account_column = new_column_names.get(column_a, column_a)
                for label, start_label, end_label in selections:
                    if start_label and end_label:
                        try:
                            start_index = df[df[account_column].eq(start_label)].index.min()
                            end_index = df[df[account_column].eq(end_label)].index.max()
                            if pd.notna(start_index) and pd.notna(end_index):
                                df.loc[start_index:end_index, 'Label'] = label
                            else:
                                st.error(f"Invalid label bounds for {label}. Skipping...")
                        except KeyError as e:
                            st.error(f"Error accessing column '{account_column}': {e}. Skipping...")
                    else:
                        st.info(f"No selections made for {label}. Skipping...")
                return df

            if st.button("Preview Setting Bounds ONLY", key="preview_setting_bounds"):
                preview_table = update_labels(all_tables.copy())
                st.subheader("Preview of Setting Bounds")
                st.dataframe(preview_table)

            # Adding column renaming functionality
            st.subheader("Rename Columns")
            quarter_options = [f"Q{i}-{year}" for year in range(2018, 2027) for i in range(1, 5)]
            ytd_options = [f"YTD {year}" for year in range(2018, 2027)]
            dropdown_options = [''] + ['Account'] + quarter_options + ytd_options

            for col in all_tables.columns:
                new_name_text = st.text_input(f"Rename '{col}' to:", value=col, key=f"rename_{col}_text")
                new_name_dropdown = st.selectbox(f"Or select predefined name for '{col}':", dropdown_options, key=f"rename_{col}_dropdown", index=0)
                new_column_names[col] = new_name_dropdown if new_name_dropdown else new_name_text
            
            all_tables.rename(columns=new_column_names, inplace=True)
            st.write("Updated Columns:", all_tables.columns.tolist())
            st.dataframe(all_tables)

            # Adding radio buttons for column removal
            st.subheader("Select columns to keep before export")
            columns_to_keep = []
            for col in all_tables.columns:
                if st.checkbox(f"Keep column '{col}'", value=True, key=f"keep_{col}"):
                    columns_to_keep.append(col)

            # Adding radio buttons for numerical column selection
            st.subheader("Select numerical columns")
            numerical_columns = []
            for col in all_tables.columns:
                if st.checkbox(f"Numerical column '{col}'", value=False, key=f"num_{col}"):
                    numerical_columns.append(col)

            # Ensure Label column is kept
            if 'Label' not in columns_to_keep:
                columns_to_keep.insert(0, 'Label')

            if 'Account' not in columns_to_keep:
                columns_to_keep.insert(1, 'Account')

            # Unit conversion functionality moved from Tab 4
            st.subheader("Convert Units")
            selected_columns = st.multiselect("Select columns for conversion", options=numerical_columns, key="columns_selection")
            selected_value = st.radio("Select conversion value", ["No Conversions Necessary", 1000, 1000000, 1000000000], index=0, key="conversion_value")

            if st.button("Apply Selected Labels and Generate Excel", key="apply_selected_labels_generate_excel_tab1"):
                updated_table = update_labels(all_tables.copy())
                updated_table = updated_table[[col for col in columns_to_keep if col in updated_table.columns]]  # Apply column removal

                # Convert selected numerical columns to numbers
                for col in numerical_columns:
                    if col in updated_table.columns:
                        updated_table[col] = updated_table[col].apply(clean_numeric_value)
                
                # Apply unit conversion if selected
                if selected_value != "No Conversions Necessary":
                    updated_table = apply_unit_conversion(updated_table, selected_columns, selected_value)

                # Convert all instances of '-' to '0'
                updated_table.replace('-', 0, inplace=True)

                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables_with_labels.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        uploaded_excel = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader_tab2')

        if uploaded_excel is not None:
            df = pd.read_excel(uploaded_excel)
            st.write("Columns in the uploaded file:", df.columns.tolist())

            if 'Account' not in df.columns:
                st.error("The uploaded file does not contain an 'Account' column.")
            else:
                def get_best_match(account):
                    best_score = float('inf')
                    best_match = None
                    for lookup_account in lookup_df['Account']:
                        account_str = str(account)
                        score = levenshtein_distance(account_str.lower(), lookup_account.lower()) / max(len(account_str), len(lookup_account))
                        if score < best_score:
                            best_score = score
                            best_match = lookup_account
                    return best_match, best_score

                df['Mnemonic'] = ''
                df['Manual Selection'] = ''
                for idx, row in df.iterrows():
                    account_value = row['Account']
                    label_value = row.get('Label', '')  # Get the label value if it exists
                    if pd.notna(account_value):
                        best_match, score = get_best_match(account_value)
                        if score < 0.2:
                            df.at[idx, 'Mnemonic'] = lookup_df.loc[lookup_df['Account'] == best_match, 'Mnemonic'].values[0]
                        else:
                            df.at[idx, 'Mnemonic'] = 'Human Intervention Required'
                    
                    if df.at[idx, 'Mnemonic'] == 'Human Intervention Required':
                        if label_value:
                            message = f"**Human Intervention Required for:** {account_value} [{label_value} - Index {idx}]"
                        else:
                            message = f"**Human Intervention Required for:** {account_value} - Index {idx}"
                        st.markdown(message)
                    
                    manual_selection = st.selectbox(
                        f"Select category for '{account_value}'",
                        options=[''] + lookup_df['Mnemonic'].tolist() + ['Other Category', 'REMOVE ROW'],
                        key=f"select_{idx}_tab2"
                    )
                    if manual_selection:
                        df.at[idx, 'Manual Selection'] = manual_selection.strip()

                st.dataframe(df[['Account', 'Mnemonic', 'Manual Selection']])

                if st.button("Generate Excel with Lookup Results", key="generate_excel_lookup_results_tab2"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'].strip() if row['Manual Selection'].strip() != '' else row['Mnemonic'], 
                        axis=1
                    )
                    final_output_df = df[df['Final Mnemonic Selection'].str.strip() != 'REMOVE ROW'].copy()
                    excel_file = io.BytesIO()
                    final_output_df.to_excel(excel_file, index=False)
                    excel_file.seek(0)
                    st.download_button("Download Excel", excel_file, "mnemonic_mapping_with_final_selection.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                if st.button("Update Data Dictionary with Manual Mappings", key="update_data_dictionary_tab2"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'] if row['Manual Selection'] not in ['Other Category', 'REMOVE ROW', ''] else row['Mnemonic'], 
                        axis=1
                    )
                    new_entries = []
                    for idx, row in df.iterrows():
                        manual_selection = row['Manual Selection']
                        final_mnemonic = row['Final Mnemonic Selection']
                        if manual_selection not in ['Other Category', 'REMOVE ROW', '']:
                            if row['Account'] not in lookup_df['Account'].values:
                                new_entries.append({'Account': row['Account'], 'Mnemonic': final_mnemonic, 'CIQ': ''})
                            else:
                                lookup_df.loc[lookup_df['Account'] == row['Account'], 'Mnemonic'] = final_mnemonic
                    if new_entries:
                        lookup_df = pd.concat([lookup_df, pd.DataFrame(new_entries)], ignore_index=True)
                    lookup_df.reset_index(drop=True, inplace=True)
                    save_lookup_table(lookup_df)
                    st.success("Data Dictionary Updated Successfully")

    with tab3:
        st.subheader("Balance Sheet Data Dictionary")

        # Upload feature
        uploaded_dict_file = st.file_uploader("Upload a new Data Dictionary CSV", type=['csv'], key='dict_uploader_tab3')
        if uploaded_dict_file is not None:
            new_lookup_df = pd.read_csv(uploaded_dict_file)
            lookup_df = new_lookup_df
            save_lookup_table(lookup_df)
            st.success("Data Dictionary uploaded and updated successfully!")

        # Display the data dictionary
        st.dataframe(lookup_df)

        # Record removal feature
        remove_indices = st.multiselect("Select rows to remove", lookup_df.index, key='remove_indices_tab3')
        if st.button("Remove Selected Rows", key="remove_selected_rows_tab3"):
            lookup_df = lookup_df.drop(remove_indices).reset_index(drop=True)
            save_lookup_table(lookup_df)
            st.success("Selected rows removed successfully!")
            st.dataframe(lookup_df)

        if st.button("Download Data Dictionary", key="download_data_dictionary_tab3"):
            excel_file = io.BytesIO()
            lookup_df.to_excel(excel_file, index=False)
            excel_file.seek(0)
            st.download_button("Download Excel", excel_file, "data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab4:
        st.subheader("Data Aggregation")

        currency_options = ["U.S. Dollar", "Euro", "British Pound Sterling", "Japanese Yen"]
        magnitude_options = ["MI standard", "Actuals", "Thousands", "Millions", "Billions", "Trillions"]

        selected_currency = st.selectbox("Select Currency", currency_options, key='currency_selection_tab4')
        selected_magnitude = st.selectbox("Select Magnitude", magnitude_options, key='magnitude_selection_tab4')

        uploaded_files = st.file_uploader("Upload your Excel files", type=['xlsx'], accept_multiple_files=True, key='xlsx_uploader_tab4')

        dfs = []
        if uploaded_files:
            dfs = [process_file(file) for file in uploaded_files if process_file(file) is not None]

        if dfs:
            # Combine the data into the "As Presented" sheet by stacking them vertically
            as_presented = pd.concat(dfs, ignore_index=True)

            # Filter to include 'Label', 'Account', and all other columns except 'Mnemonic', 'Manual Selection', and 'Final Mnemonic Selection'
            columns_to_include = [col for col in as_presented.columns if col not in ['Mnemonic', 'Manual Selection']]
            as_presented_filtered = as_presented[columns_to_include]

            # Aggregate data
            aggregated_table = aggregate_data(as_presented_filtered)

            # Ensure the columns are ordered as Label, Account, Final Mnemonic Selection, and then the remaining columns
            columns_order = ['Label', 'Account', 'Final Mnemonic Selection'] +                             [col for col in aggregated_table.columns if col not in ['Label', 'Account', 'Final Mnemonic Selection']]
            aggregated_table = aggregated_table[columns_order]

            # Combine the data into the "Standardized" sheet
            combined_df = create_combined_df(dfs)

            # Sort by Label and ensure 'Total' records are at the end within each label
            aggregated_table = sort_by_label(aggregated_table)
            combined_df = sort_by_label(combined_df)

            if st.button("Download Aggregated Excel", key="download_aggregated_excel_tab4"):
                excel_file = io.BytesIO()
                with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                    aggregated_table.to_excel(writer, sheet_name='As Presented', index=False)
                    combined_df.to_excel(writer, sheet_name='Standardized', index=False)
                    
                    # Create the "Cover" sheet with the selections
                    cover_df = pd.DataFrame({
                        'Selection': ['Currency', 'Magnitude'],
                        'Value': [selected_currency, selected_magnitude]
                    })
                    cover_df.to_excel(writer, sheet_name='Cover', index=False)
                
                excel_file.seek(0)

                st.download_button(
                    label="Download Aggregated Excel",
                    data=excel_file,
                    file_name="aggregated_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Please upload valid Excel files for aggregation.")

if __name__ == '__main__':
    main()

