#!/usr/bin/env python
# coding: utf-8

# In[1]:


import io
import json
import os
import pandas as pd
import streamlit as st
from Levenshtein import distance as levenshtein_distance
import xlsxwriter
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
    df = pd.read_excel(file, sheet_name=None)
    # Assuming the relevant sheet is the first one
    first_sheet_name = list(df.keys())[0]
    df = df[first_sheet_name]
    return df

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

        df_grouped = df.groupby(final_mnemonic_col).sum(numeric_only=True).reset_index()
        df_melted = df_grouped.melt(id_vars=[final_mnemonic_col], value_vars=date_cols, var_name='Date', value_name='Value')
        df_melted['Date'] = df_melted['Date'] + f'_{i+1}'
        df_pivot = df_melted.pivot(index=final_mnemonic_col, columns='Date', values='Value')
        
        if combined_df.empty:
            combined_df = df_pivot
        else:
            combined_df = combined_df.join(df_pivot, how='outer')
    return combined_df.reset_index()

def aggregate_data(df):
    # Check for the presence of 'Row Labels' and 'Account' columns dynamically
    if 'Label' not in df.columns or 'Account' not in df.columns:
        st.error("'Label' and/or 'Account' columns not found in the data.")
        return df
    
    # Example aggregation function: Pivoting the data
    pivot_table = df.pivot_table(index=['Label', 'Account'], 
                                 values=[col for col in df.columns if col not in ['Label', 'Account']], 
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

            def update_labels():
                all_tables['Label'] = ''
                for label, start_label, end_label in selections:
                    if start_label and end_label:
                        start_index = all_tables[all_tables[column_a].eq(start_label)].index.min()
                        end_index = all_tables[all_tables[column_a].eq(end_label)].index.max()
                        if pd.notna(start_index) and pd.notna(end_index):
                            all_tables.loc[start_index:end_index, 'Label'] = label
                        else:
                            st.error(f"Invalid label bounds for {label}. Skipping...")
                    else:
                        st.info(f"No selections made for {label}. Skipping...")
                return all_tables

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

            if st.button("Update Labels Preview"):
                updated_table = update_labels()
                st.subheader("Updated Data Preview")
                st.dataframe(updated_table[columns_to_keep])

            if st.button("Apply Selected Labels and Generate Excel"):
                updated_table = update_labels()
                updated_table = updated_table[columns_to_keep]  # Apply column removal
                
                # Convert selected numerical columns to numbers
                for col in numerical_columns:
                    updated_table[col] = updated_table[col].apply(clean_numeric_value)
                
                # Convert all instances of '-' to '0'
                updated_table.replace('-', 0, inplace=True)
                
                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables_with_labels.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        uploaded_excel = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader')

        if uploaded_excel is not None:
            df = pd.read_excel(uploaded_excel)
            st.write("Columns in the uploaded file:", df.columns.tolist())

            new_column_names = {}
            quarter_options = [f"Q{i}-{year}" for year in range(2018, 2027) for i in range(1, 5)]
            ytd_options = [f"YTD {year}" for year in range(2018, 2027)]
            dropdown_options = [''] + lookup_df['Mnemonic'].tolist() + quarter_options + ytd_options

            st.subheader("Rename Columns")
            for col in df.columns:
                new_name_text = st.text_input(f"Rename '{col}' to:", value=col, key=f"rename_{col}_text")
                new_name_dropdown = st.selectbox(f"Or select predefined name for '{col}':", dropdown_options, key=f"rename_{col}_dropdown")
                
                new_column_names[col] = new_name_dropdown if new_name_dropdown else new_name_text
            
            df.rename(columns=new_column_names, inplace=True)
            st.write("Updated Columns:", df.columns.tolist())
            st.dataframe(df)

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
                        options=dropdown_options,
                        key=f"select_{idx}"
                    )
                    if manual_selection:
                        df.at[idx, 'Manual Selection'] = manual_selection.strip()

                st.dataframe(df[['Account', 'Mnemonic', 'Manual Selection']])

                if st.button("Generate Excel with Lookup Results"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'].strip() if row['Manual Selection'].strip() != '' else row['Mnemonic'], 
                        axis=1
                    )
                    final_output_df = df[df['Final Mnemonic Selection'].str.strip() != 'REMOVE ROW'].copy()
                    excel_file = io.BytesIO()
                    final_output_df.to_excel(excel_file, index=False)
                    excel_file.seek(0)
                    st.download_button("Download Excel", excel_file, "mnemonic_mapping_with_final_selection.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                if st.button("Update Data Dictionary with Manual Mappings"):
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
        uploaded_dict_file = st.file_uploader("Upload a new Data Dictionary CSV", type=['csv'], key='dict_uploader')
        if uploaded_dict_file is not None:
            new_lookup_df = pd.read_csv(uploaded_dict_file)
            lookup_df = new_lookup_df
            save_lookup_table(lookup_df)
            st.success("Data Dictionary uploaded and updated successfully!")

        # Display the data dictionary
        st.dataframe(lookup_df)

        # Record removal feature
        remove_indices = st.multiselect("Select rows to remove", lookup_df.index)
        if st.button("Remove Selected Rows"):
            lookup_df = lookup_df.drop(remove_indices).reset_index(drop=True)
            save_lookup_table(lookup_df)
            st.success("Selected rows removed successfully!")
            st.dataframe(lookup_df)

        if st.button("Download Data Dictionary"):
            excel_file = io.BytesIO()
            lookup_df.to_excel(excel_file, index=False)
            excel_file.seek(0)
            st.download_button("Download Excel", excel_file, "data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab4:
        st.subheader("Data Aggregation")

        currency_options = ["U.S. Dollar", "Euro", "British Pound Sterling", "Japanese Yen"]
        magnitude_options = ["MI standard", "Actuals", "Thousands", "Millions", "Billions", "Trillions"]

        selected_currency = st.selectbox("Select Currency", currency_options, key='currency_selection')
        selected_magnitude = st.selectbox("Select Magnitude", magnitude_options, key='magnitude_selection')

        uploaded_files = st.file_uploader("Upload your Excel files", type=['xlsx'], accept_multiple_files=True, key='xlsx_uploader')

        if uploaded_files and st.button("Aggregate Data"):
            dfs = [process_file(file) for file in uploaded_files]

            # Combine the data into the "As Presented" sheet by stacking them vertically
            as_presented = pd.concat(dfs, ignore_index=True)
            
            # Filter to include 'Label', 'Account', and all other columns except 'Mnemonic', 'Manual Selection', and 'Final Mnemonic Selection'
            columns_to_include = [col for col in as_presented.columns if col not in ['Mnemonic', 'Manual Selection', 'Final Mnemonic Selection']]
            as_presented_filtered = as_presented[columns_to_include]

            # Aggregate data
            aggregated_table = aggregate_data(as_presented_filtered)

            # Combine the data into the "Combined" sheet
            combined_df = create_combined_df(dfs)

            excel_file = io.BytesIO()
            with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                aggregated_table.to_excel(writer, sheet_name='As Presented', index=False)
                combined_df.to_excel(writer, sheet_name='Combined', index=False)
                
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

if __name__ == '__main__':
    main()

