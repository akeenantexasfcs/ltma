#!/usr/bin/env python
# coding: utf-8

# In[1]:


import io
import json
import os
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from Levenshtein import distance as levenshtein_distance
import re

# Define the initial lookup data for Balance Sheet
initial_balance_sheet_lookup_data = {
    "Account": ["Cash and cash equivalents", "Line of credit", "Goodwill",
                "Total Current Assets", "Total Assets", "Total Current Liabilities"],
    "Mnemonic": ["Cash & Cash Equivalents", "Short-Term Debt", "Goodwill",
                 "Total Current Assets", "Total Assets", "Total Current Liabilities"],
    "CIQ": ["IQ_CASH_EQUIV", "IQ_ST_INVEST", "IQ_GW",
            "IQ_TOTAL_CA", "IQ_TOTAL_ASSETS", "IQ_TOTAL_CL"]
}

# Define the initial lookup data for Cash Flow
initial_cash_flow_lookup_data = {
    "Label": ["Operating Activities", "Investing Activities", "Financing Activities"],
    "Account": ["Net Cash Provided by Operating Activities", "Net Cash Used in Investing Activities", "Net Cash Provided by Financing Activities"],
    "Mnemonic": ["Operating Cash Flow", "Investing Cash Flow", "Financing Cash Flow"],
    "CIQ": ["IQ_OPER_CASH_FLOW", "IQ_INVEST_CASH_FLOW", "IQ_FIN_CASH_FLOW"]
}

# Define the file paths for the data dictionaries
balance_sheet_data_dictionary_file = 'balance_sheet_data_dictionary.csv'
cash_flow_data_dictionary_file = 'cash_flow_data_dictionary.csv'

# Load or initialize the lookup table
def load_or_initialize_lookup(file_path, initial_data):
    if os.path.exists(file_path):
        lookup_df = pd.read_csv(file_path)
    else:
        lookup_df = pd.DataFrame(initial_data)
        lookup_df.to_csv(file_path, index=False)
    return lookup_df

def save_lookup_table(df, file_path):
    df.to_csv(file_path, index=False)

# Initialize lookup tables for Balance Sheet and Cash Flow
balance_sheet_lookup_df = load_or_initialize_lookup(balance_sheet_data_dictionary_file, initial_balance_sheet_lookup_data)
cash_flow_lookup_df = load_or_initialize_lookup(cash_flow_data_dictionary_file, initial_cash_flow_lookup_data)

def process_file(file):
    try:
        df = pd.read_excel(file, sheet_name=None)
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
    if 'Label' not in df.columns or 'Account' not in df.columns:
        st.error("'Label' and/or 'Account' columns not found in the data.")
        return df
    
    pivot_table = df.pivot_table(index=['Label', 'Account'], 
                                 values=[col for col in df.columns if col not in ['Label', 'Account', 'Mnemonic', 'Manual Selection']], 
                                 aggfunc='sum').reset_index()
    return pivot_table

def clean_numeric_value(value):
    value_str = str(value).strip()
    if value_str.startswith('(') and value_str.endswith(')'):
        value_str = '-' + value_str[1:-1]
    cleaned_value = re.sub(r'[$,]', '', value_str)
    try:
        return float(cleaned_value)
    except ValueError:
        return 0

def sort_by_label_and_account(df):
    sort_order = {
        "Current Assets": 0,
        "Non Current Assets": 1,
        "Current Liabilities": 2,
        "Non Current Liabilities": 3,
        "Equity": 4,
        "Total Equity and Liabilities": 5
    }
    
    df['Label_Order'] = df['Label'].map(sort_order)
    df['Total_Order'] = df['Account'].str.contains('Total', case=False).astype(int)
    
    df = df.sort_values(by=['Label_Order', 'Label', 'Total_Order', 'Account']).drop(columns=['Label_Order', 'Total_Order'])
    return df

def sort_by_label_and_final_mnemonic(df):
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

def balance_sheet():
    global balance_sheet_lookup_df

    st.title("BALANCE SHEET LTMA")

    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Aggregate My Data", "Mappings and Data Aggregation", "Balance Sheet Data Dictionary"])

    with tab1:
        uploaded_file = st.file_uploader("Choose a JSON file", type="json", key='json_uploader')
        if uploaded_file is not None:
            data = json.load(uploaded_file)
            st.warning("PLEASE NOTE: In the Setting Bounds Preview Window, you will see only your respective labels. In the Updated Columns Preview Window, you will see only your renamed column headers. The labels from the Setting Bounds section will not appear in the Updated Columns Preview.")
            st.warning("PLEASE ALSO NOTE: An Account column must also be designated when you are in the Rename Columns section.")

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

            def get_unique_options(series):
                counts = series.value_counts()
                unique_options = []
                occurrence_counts = {}
                for item in series:
                    if counts[item] > 1:
                        if item not in occurrence_counts:
                            occurrence_counts[item] = 1
                        else:
                            occurrence_counts[item] += 1
                        unique_options.append(f"{item} {occurrence_counts[item]}")
                    else:
                        unique_options.append(item)
                return unique_options

            labels = ["Current Assets", "Non Current Assets", "Current Liabilities",
                      "Non Current Liabilities", "Equity", "Total Equity and Liabilities"]
            selections = []

            for label in labels:
                st.subheader(f"Setting bounds for {label}")
                options = [''] + get_unique_options(all_tables[column_a].dropna())
                start_label = st.selectbox(f"Start Label for {label}", options, key=f"start_{label}")
                end_label = st.selectbox(f"End Label for {label}", options, key=f"end_{label}")
                selections.append((label, start_label, end_label))

            new_column_names = {col: col for col in all_tables.columns}

            def update_labels(df):
                df['Label'] = ''
                account_column = new_column_names.get(column_a, column_a)
                for label, start_label, end_label in selections:
                    if start_label and end_label:
                        try:
                            start_label_base = " ".join(start_label.split()[:-1]) if start_label.split()[-1].isdigit() else start_label
                            start_index = df[df[account_column].str.contains(start_label_base)].index.min()
                            end_label_base = " ".join(end_label.split()[:-1]) if end_label.split()[-1].isdigit() else end_label
                            end_index = df[df[account_column].str.contains(end_label_base)].index.max()
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

            st.subheader("Rename Columns")
            quarter_options = [f"FQ{quarter}{year}" for year in range(2018, 2027) for quarter in range(1, 5)]
            ytd_options = [f"YTD{quarter}{year}" for year in range(2018, 2027) for quarter in range(1, 5)]
            dropdown_options = [''] + ['Account'] + quarter_options + ytd_options

            for col in all_tables.columns:
                new_name_text = st.text_input(f"Rename '{col}' to:", value=col, key=f"rename_{col}_text")
                new_name_dropdown = st.selectbox(f"Or select predefined name for '{col}':", dropdown_options, key=f"rename_{col}_dropdown", index=0)
                new_column_names[col] = new_name_dropdown if new_name_dropdown else new_name_text
            
            all_tables.rename(columns=new_column_names, inplace=True)
            st.write("Updated Columns:", all_tables.columns.tolist())
            st.dataframe(all_tables)

            st.subheader("Select columns to keep before export")
            columns_to_keep = []
            for col in all_tables.columns:
                if st.checkbox(f"Keep column '{col}'", value=True, key=f"keep_{col}"):
                    columns_to_keep.append(col)

            st.subheader("Select numerical columns")
            numerical_columns = []
            for col in all_tables.columns:
                if st.checkbox(f"Numerical column '{col}'", value=False, key=f"num_{col}"):
                    numerical_columns.append(col)

            if 'Label' not in columns_to_keep:
                columns_to_keep.insert(0, 'Label')

            if 'Account' not in columns_to_keep:
                columns_to_keep.insert(1, 'Account')

            st.subheader("Label Units")
            selected_columns = st.multiselect("Select columns for conversion", options=numerical_columns, key="columns_selection")
            selected_value = st.radio("Select conversion value", ["Actuals", "Thousands", "Millions", "Billions"], index=0, key="conversion_value")

            conversion_factors = {
                "Actuals": 1,
                "Thousands": 1000,
                "Millions": 1000000,
                "Billions": 1000000000
            }

            if st.button("Apply Selected Labels and Generate Excel", key="apply_selected_labels_generate_excel_tab1"):
                updated_table = update_labels(all_tables.copy())
                updated_table = updated_table[[col for col in columns_to_keep if col in updated_table.columns]]

                updated_table = updated_table[updated_table['Label'].str.strip() != '']
                updated_table = updated_table[updated_table['Account'].str.strip() != '']

                for col in numerical_columns:
                    if col in updated_table.columns:
                        updated_table[col] = updated_table[col].apply(clean_numeric_value)
                
                if selected_value != "No Conversions Necessary":
                    updated_table = apply_unit_conversion(updated_table, selected_columns, conversion_factors[selected_value])

                updated_table.replace('-', 0, inplace=True)

                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables_with_labels.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.subheader("Check for Duplicate Accounts")
            if 'Account' not in all_tables.columns:
                st.warning("The 'Account' column is missing. Please ensure your data includes an 'Account' column.")
            else:
                duplicated_accounts = all_tables[all_tables.duplicated(['Account'], keep=False)]
                if not duplicated_accounts.empty:
                    st.warning("Duplicates identified:")
                    st.dataframe(duplicated_accounts)
                else:
                    st.success("No duplicates identified")

    with tab2:
        st.subheader("Aggregate My Data")
        
        uploaded_files = st.file_uploader("Upload your Excel files from Tab 1", type=['xlsx'], accept_multiple_files=True, key='xlsx_uploader_tab2')

        dfs = []
        if uploaded_files:
            dfs = [process_file(file) for file in uploaded_files if process_file(file) is not None]

        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            st.dataframe(combined_df)

            aggregated_table = aggregate_data(combined_df)
            aggregated_table = sort_by_label_and_account(aggregated_table)

            st.subheader("Aggregated Data")
            st.dataframe(aggregated_table)

            if st.button("Download Aggregated Excel", key="download_aggregated_excel_tab2"):
                excel_file = io.BytesIO()
                with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                    aggregated_table.to_excel(writer, sheet_name='Aggregated Data', index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "aggregated_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("Please upload valid Excel files for aggregation.")

    with tab3:
        st.subheader("Mappings and Data Aggregation")

        uploaded_excel = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader_tab3_bs')

        currency_options = ["U.S. Dollar", "Euro", "British Pound Sterling", "Japanese Yen"]
        magnitude_options = ["Actuals", "MI standard", "Thousands", "Millions", "Billions", "Trillions"]

        selected_currency = st.selectbox("Select Currency", currency_options, key='currency_selection_tab3_bs')
        selected_magnitude = st.selectbox("Select Magnitude", magnitude_options, key='magnitude_selection_tab3_bs')

        if uploaded_excel is not None:
            df = pd.read_excel(uploaded_excel)
            st.write("Columns in the uploaded file:", df.columns.tolist())

            if 'Account' not in df.columns:
                st.error("The uploaded file does not contain an 'Account' column.")
            else:
                # Function to get the best match based on Label first, then Levenshtein distance on Account
                def get_best_match(label, account):
                    best_score = float('inf')
                    best_match = None
                    for _, lookup_row in balance_sheet_lookup_df.iterrows():
                        if lookup_row['Label'].strip().lower() == str(label).strip().lower():
                            lookup_account = lookup_row['Account']
                            account_str = str(account)
                            # Levenshtein distance for Account
                            score = levenshtein_distance(account_str.lower(), lookup_account.lower()) / max(len(account_str), len(lookup_account))
                            if score < best_score:
                                best_score = score
                                best_match = lookup_row
                    return best_match, best_score

                df['Mnemonic'] = ''
                df['Manual Selection'] = ''
                for idx, row in df.iterrows():
                    account_value = row['Account']
                    label_value = row.get('Label', '')
                    if pd.notna(account_value):
                        best_match, score = get_best_match(label_value, account_value)
                        if best_match is not None and score < 0.25:
                            df.at[idx, 'Mnemonic'] = best_match['Mnemonic']
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
                        options=[''] + balance_sheet_lookup_df['Mnemonic'].tolist() + ['REMOVE ROW'],
                        key=f"select_{idx}_tab3_bs"
                    )
                    if manual_selection:
                        df.at[idx, 'Manual Selection'] = manual_selection.strip()

                st.dataframe(df[['Label', 'Account', 'Mnemonic', 'Manual Selection']])

                if st.button("Generate Excel with Lookup Results", key="generate_excel_lookup_results_tab3_bs"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'].strip() if row['Manual Selection'].strip() != '' else row['Mnemonic'], 
                        axis=1
                    )
                    final_output_df = df[df['Final Mnemonic Selection'].str.strip() != 'REMOVE ROW'].copy()
                    
                    combined_df = create_combined_df([final_output_df])
                    combined_df = sort_by_label_and_final_mnemonic(combined_df)

                    # Add CIQ column based on lookup
                    def lookup_ciq(mnemonic):
                        if mnemonic == 'Human Intervention Required':
                            return 'CIQ IQ Required'
                        ciq_value = balance_sheet_lookup_df.loc[balance_sheet_lookup_df['Mnemonic'] == mnemonic, 'CIQ']
                        if ciq_value.empty:
                            return 'CIQ IQ Required'
                        return ciq_value.values[0]
                    
                    combined_df['CIQ'] = combined_df['Final Mnemonic Selection'].apply(lookup_ciq)

                    columns_order = ['Label', 'Final Mnemonic Selection', 'CIQ'] + [col for col in combined_df.columns if col not in ['Label', 'Final Mnemonic Selection', 'CIQ']]
                    combined_df = combined_df[columns_order]

                    # Include the "As Presented" sheet without the CIQ column, and with the specified column order
                    as_presented_df = final_output_df.drop(columns=['CIQ', 'Mnemonic', 'Manual Selection'], errors='ignore')
                    as_presented_columns_order = ['Label', 'Account', 'Final Mnemonic Selection'] + [col for col in as_presented_df.columns if col not in ['Label', 'Account', 'Final Mnemonic Selection']]
                    as_presented_df = as_presented_df[as_presented_columns_order]

                    excel_file = io.BytesIO()
                    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                        combined_df.to_excel(writer, sheet_name='Standardized', index=False)
                        as_presented_df.to_excel(writer, sheet_name='As Presented', index=False)
                        cover_df = pd.DataFrame({
                            'Selection': ['Currency', 'Magnitude'],
                            'Value': [selected_currency, selected_magnitude]
                        })
                        cover_df.to_excel(writer, sheet_name='Cover', index=False)
                    excel_file.seek(0)
                    st.download_button("Download Excel", excel_file, "mnemonic_mapping_with_aggregation.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                if st.button("Update Data Dictionary with Manual Mappings", key="update_data_dictionary_tab3_bs"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'] if row['Manual Selection'] not in ['REMOVE ROW', ''] else row['Mnemonic'], 
                        axis=1
                    )
                    new_entries = []
                    for idx, row in df.iterrows():
                        manual_selection = row['Manual Selection']
                        final_mnemonic = row['Final Mnemonic Selection']
                        if manual_selection == 'REMOVE ROW':
                            continue
                        ciq_value = balance_sheet_lookup_df.loc[balance_sheet_lookup_df['Mnemonic'] == final_mnemonic, 'CIQ'].values[0] if not balance_sheet_lookup_df.loc[balance_sheet_lookup_df['Mnemonic'] == final_mnemonic, 'CIQ'].empty else 'CIQ IQ Required'
                        
                        if manual_selection not in ['REMOVE ROW', '']:
                            if row['Account'] not in balance_sheet_lookup_df['Account'].values:
                                new_entries.append({'Account': row['Account'], 'Mnemonic': final_mnemonic, 'CIQ': ciq_value, 'Label': row['Label']})
                            else:
                                balance_sheet_lookup_df.loc[balance_sheet_lookup_df['Account'] == row['Account'], 'Mnemonic'] = final_mnemonic
                                balance_sheet_lookup_df.loc[balance_sheet_lookup_df['Account'] == row['Account'], 'Label'] = row['Label']
                                balance_sheet_lookup_df.loc[balance_sheet_lookup_df['Account'] == row['Account'], 'CIQ'] = ciq_value
                    if new_entries:
                        balance_sheet_lookup_df = pd.concat([balance_sheet_lookup_df, pd.DataFrame(new_entries)], ignore_index=True)
                    balance_sheet_lookup_df.reset_index(drop=True, inplace=True)
                    save_lookup_table(balance_sheet_lookup_df, balance_sheet_data_dictionary_file)
                    st.success("Data Dictionary Updated Successfully")

    with tab4:
        st.subheader("Balance Sheet Data Dictionary")

        uploaded_dict_file = st.file_uploader("Upload a new Data Dictionary CSV", type=['csv'], key='dict_uploader_tab4_bs')
        if uploaded_dict_file is not None:
            new_lookup_df = pd.read_csv(uploaded_dict_file)
            balance_sheet_lookup_df = new_lookup_df  # Overwrite the entire DataFrame
            save_lookup_table(balance_sheet_lookup_df, balance_sheet_data_dictionary_file)
            st.success("Data Dictionary uploaded and updated successfully!")

        st.dataframe(balance_sheet_lookup_df)

        remove_indices = st.multiselect("Select rows to remove", balance_sheet_lookup_df.index, key='remove_indices_tab4_bs')
        if st.button("Remove Selected Rows", key="remove_selected_rows_tab4_bs"):
            balance_sheet_lookup_df = balance_sheet_lookup_df.drop(remove_indices).reset_index(drop=True)
            save_lookup_table(balance_sheet_lookup_df, balance_sheet_data_dictionary_file)
            st.success("Selected rows removed successfully!")
            st.dataframe(balance_sheet_lookup_df)

        if st.button("Download Data Dictionary", key="download_data_dictionary_tab4_bs"):
            excel_file = io.BytesIO()
            balance_sheet_lookup_df.to_excel(excel_file, index=False)
            excel_file.seek(0)
            st.download_button("Download Excel", excel_file, "balance_sheet_data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def cash_flow_statement():
    global cash_flow_lookup_df

    st.title("CASH FLOW STATEMENT LTMA")
    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Aggregate My Data", "Mappings and Data Aggregation", "Cash Flow Data Dictionary"])

    with tab1:
        uploaded_file = st.file_uploader("Choose a JSON file", type="json", key='json_uploader_cfs')
        if uploaded_file is not None:
            data = json.load(uploaded_file)
            st.warning("PLEASE NOTE: In the Setting Bounds Preview Window, you will see only your respective labels. In the Updated Columns Preview Window, you will see only your renamed column headers. The labels from the Setting Bounds section will not appear in the Updated Columns Preview.")
            st.warning("PLEASE ALSO NOTE: An Account column must also be designated when you are in the Rename Columns section.")

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
                    table_df = table_df.sortindex(axis=1)
                    tables.append(table_df)
            all_tables = pd.concat(tables, axis=0, ignore_index=True)
            if len(all_tables.columns) == 0:
                st.error("No columns found in the uploaded JSON file.")
                return

            column_a = all_tables.columns[0]
            all_tables.insert(0, 'Label', '')

            st.subheader("Data Preview")
            st.dataframe(all_tables)

            def get_unique_options(series):
                counts = series.value_counts()
                unique_options = []
                occurrence_counts = {}
                for item in series:
                    if counts[item] > 1:
                        if item not in occurrence_counts:
                            occurrence_counts[item] = 1
                        else:
                            occurrence_counts[item] += 1
                        unique_options.append(f"{item} {occurrence_counts[item]}")
                    else:
                        unique_options.append(item)
                return unique_options

            labels = ["Operating Activities", "Investing Activities", "Financing Activities", "Cash Flow from Other", "Supplemental Cash Flow"]
            selections = []

            for label in labels:
                st.subheader(f"Setting bounds for {label}")
                options = [''] + get_unique_options(all_tables[column_a].dropna())
                start_label = st.selectbox(f"Start Label for {label}", options, key=f"start_{label}_cfs")
                end_label = st.selectbox(f"End Label for {label}", options, key=f"end_{label}_cfs")
                selections.append((label, start_label, end_label))

            new_column_names = {col: col for col in all_tables.columns}

            def update_labels(df):
                df['Label'] = ''
                account_column = new_column_names.get(column_a, column_a)
                for label, start_label, end_label in selections:
                    if start_label and end_label:
                        try:
                            start_label_base = " ".join(start_label.split()[:-1]) if start_label.split()[-1].isdigit() else start_label
                            start_index = df[df[account_column].str.contains(start_label_base)].index.min()
                            end_label_base = " ".join(end_label.split()[:-1]) if end_label.split()[-1].isdigit() else end_label
                            end_index = df[df[account_column].str.contains(end_label_base)].index.max()
                            if pd.notna(start_index) and pd.notna(end_index):
                                df.loc[start_index:end_index, 'Label'] = label
                            else:
                                st.error(f"Invalid label bounds for {label}. Skipping...")
                        except KeyError as e:
                            st.error(f"Error accessing column '{account_column}': {e}. Skipping...")
                    else:
                        st.info(f"No selections made for {label}. Skipping...")
                return df

            if st.button("Preview Setting Bounds ONLY", key="preview_setting_bounds_cfs"):
                preview_table = update_labels(all_tables.copy())
                st.subheader("Preview of Setting Bounds")
                st.dataframe(preview_table)

            st.subheader("Rename Columns")
            quarter_options = [f"FQ{quarter}{year}" for year in range(2018, 2027) for quarter in range(1, 5)]
            ytd_options = [f"YTD{quarter}{year}" for year in range(2018, 2027) for quarter in range(1, 5)]
            dropdown_options = [''] + ['Account'] + quarter_options + ytd_options

            for col in all_tables.columns:
                new_name_text = st.text_input(f"Rename '{col}' to:", value=col, key=f"rename_{col}_text_cfs")
                new_name_dropdown = st.selectbox(f"Or select predefined name for '{col}':", dropdown_options, key=f"rename_{col}_dropdown_cfs", index=0)
                new_column_names[col] = new_name_dropdown if new_name_dropdown else new_name_text
            
            all_tables.rename(columns=new_column_names, inplace=True)
            st.write("Updated Columns:", all_tables.columns.tolist())
            st.dataframe(all_tables)

            st.subheader("Select columns to keep before export")
            columns_to_keep = []
            for col in all_tables.columns:
                if st.checkbox(f"Keep column '{col}'", value=True, key=f"keep_{col}_cfs"):
                    columns_to_keep.append(col)

            st.subheader("Select numerical columns")
            numerical_columns = []
            for col in all_tables.columns:
                if st.checkbox(f"Numerical column '{col}'", value=False, key=f"num_{col}_cfs"):
                    numerical_columns.append(col)

            if 'Label' not in columns_to_keep:
                columns_to_keep.insert(0, 'Label')

            if 'Account' not in columns_to_keep:
                columns_to_keep.insert(1, 'Account')

            st.subheader("Convert Units")
            selected_columns = st.multiselect("Select columns for conversion", options=numerical_columns, key="columns_selection_cfs")
            selected_value = st.radio("Select conversion value", ["No Conversions Necessary", "Thousands", "Millions", "Billions"], index=0, key="conversion_value_cfs")

            conversion_factors = {
                "No Conversions Necessary": 1,
                "Thousands": 1000,
                "Millions": 1000000,
                "Billions": 1000000000
            }

            if st.button("Apply Selected Labels and Generate Excel", key="apply_selected_labels_generate_excel_tab1_cfs"):
                updated_table = update_labels(all_tables.copy())
                updated_table = updated_table[[col for col in columns_to_keep if col in updated_table.columns]]

                updated_table = updated_table[updated_table['Label'].str.strip() != '']
                updated_table = updated_table[updated_table['Account'].str.strip() != '']

                for col in numerical_columns:
                    if col in updated_table.columns:
                        updated_table[col] = updated_table[col].apply(clean_numeric_value)
                
                if selected_value != "No Conversions Necessary":
                    updated_table = apply_unit_conversion(updated_table, selected_columns, conversion_factors[selected_value])

                updated_table.replace('-', 0, inplace=True)

                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables_with_labels.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        st.subheader("Aggregate My Data")
        
        uploaded_files = st.file_uploader("Upload your Excel files from Tab 1", type=['xlsx'], accept_multiple_files=True, key='xlsx_uploader_tab2_cfs')

        dfs = []
        if uploaded_files:
            dfs = [process_file(file) for file in uploaded_files if process_file(file) is not None]

        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            st.dataframe(combined_df)

            aggregated_table = aggregate_data(combined_df)
            aggregated_table = sort_by_label_and_account(aggregated_table)

            st.subheader("Aggregated Data")
            st.dataframe(aggregated_table)

            if st.button("Download Aggregated Excel", key="download_aggregated_excel_tab2_cfs"):
                excel_file = io.BytesIO()
                with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                    aggregated_table.to_excel(writer, sheet_name='Aggregated Data', index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "aggregated_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("Please upload valid Excel files for aggregation.")

    with tab3:
        st.subheader("Mappings and Data Aggregation")

        uploaded_excel = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader_tab3_cfs')

        currency_options = ["U.S. Dollar", "Euro", "British Pound Sterling", "Japanese Yen"]
        magnitude_options = ["Actuals", "MI standard", "Thousands", "Millions", "Billions", "Trillions"]

        selected_currency = st.selectbox("Select Currency", currency_options, key='currency_selection_tab3_cfs')
        selected_magnitude = st.selectbox("Select Magnitude", magnitude_options, key='magnitude_selection_tab3_cfs')

        if uploaded_excel is not None:
            df = pd.read_excel(uploaded_excel)
            st.write("Columns in the uploaded file:", df.columns.tolist())

            if 'Account' not in df.columns:
                st.error("The uploaded file does not contain an 'Account' column.")
            else:
                # Function to get the best match based on Label first, then Levenshtein distance on Account
                def get_best_match(label, account):
                    best_score = float('inf')
                    best_match = None
                    for _, lookup_row in cash_flow_lookup_df.iterrows():
                        if lookup_row['Label'].strip().lower() == str(label).strip().lower():
                            lookup_account = lookup_row['Account']
                            account_str = str(account)
                            # Levenshtein distance for Account
                            score = levenshtein_distance(account_str.lower(), lookup_account.lower()) / max(len(account_str), len(lookup_account))
                            if score < best_score:
                                best_score = score
                                best_match = lookup_row
                    return best_match, best_score

                df['Mnemonic'] = ''
                df['Manual Selection'] = ''
                for idx, row in df.iterrows():
                    account_value = row['Account']
                    label_value = row.get('Label', '')
                    if pd.notna(account_value):
                        best_match, score = get_best_match(label_value, account_value)
                        if best_match is not None and score < 0.25:
                            df.at[idx, 'Mnemonic'] = best_match['Mnemonic']
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
                        options=[''] + cash_flow_lookup_df['Mnemonic'].tolist() + ['REMOVE ROW'],
                        key=f"select_{idx}_tab3_cfs"
                    )
                    if manual_selection:
                        df.at[idx, 'Manual Selection'] = manual_selection.strip()

                st.dataframe(df[['Label', 'Account', 'Mnemonic', 'Manual Selection']])  # Include 'Label' as the first column

                if st.button("Generate Excel with Lookup Results", key="generate_excel_lookup_results_tab3_cfs"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'].strip() if row['Manual Selection'].strip() != '' else row['Mnemonic'], 
                        axis=1
                    )
                    final_output_df = df[df['Final Mnemonic Selection'].str.strip() != 'REMOVE ROW'].copy()
                    
                    combined_df = create_combined_df([final_output_df])
                    combined_df = sort_by_label_and_final_mnemonic(combined_df)

                    # Add CIQ column based on lookup
                    def lookup_ciq(mnemonic):
                        if mnemonic == 'Human Intervention Required':
                            return 'CIQ IQ Required'
                        ciq_value = cash_flow_lookup_df.loc[cash_flow_lookup_df['Mnemonic'] == mnemonic, 'CIQ']
                        if ciq_value.empty:
                            return 'CIQ IQ Required'
                        return ciq_value.values[0]
                    
                    combined_df['CIQ'] = combined_df['Final Mnemonic Selection'].apply(lookup_ciq)

                    columns_order = ['Label', 'Final Mnemonic Selection', 'CIQ'] + [col for col in combined_df.columns if col not in ['Label', 'Final Mnemonic Selection', 'CIQ']]
                    combined_df = combined_df[columns_order]

                    # Include the "As Presented" sheet without the CIQ column, and with the specified column order
                    as_presented_df = final_output_df.drop(columns=['CIQ', 'Mnemonic', 'Manual Selection'], errors='ignore')
                    as_presented_columns_order = ['Label', 'Account', 'Final Mnemonic Selection'] + [col for col in as_presented_df.columns if col not in ['Label', 'Account', 'Final Mnemonic Selection']]
                    as_presented_df = as_presented_df[as_presented_columns_order]

                    excel_file = io.BytesIO()
                    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                        combined_df.to_excel(writer, sheet_name='Standardized', index=False)
                        as_presented_df.to_excel(writer, sheet_name='As Presented', index=False)
                        cover_df = pd.DataFrame({
                            'Selection': ['Currency', 'Magnitude'],
                            'Value': [selected_currency, selected_magnitude]
                        })
                        cover_df.to_excel(writer, sheet_name='Cover', index=False)
                    excel_file.seek(0)
                    st.download_button("Download Excel", excel_file, "mnemonic_mapping_with_aggregation.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                if st.button("Update Data Dictionary with Manual Mappings", key="update_data_dictionary_tab3_cfs"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'] if row['Manual Selection'] not in ['REMOVE ROW', ''] else row['Mnemonic'], 
                        axis=1
                    )
                    new_entries = []
                    for idx, row in df.iterrows():
                        manual_selection = row['Manual Selection']
                        final_mnemonic = row['Final Mnemonic Selection']
                        if manual_selection == 'REMOVE ROW':
                            continue
                        ciq_value = cash_flow_lookup_df.loc[cash_flow_lookup_df['Mnemonic'] == final_mnemonic, 'CIQ'].values[0] if not cash_flow_lookup_df.loc[cash_flow_lookup_df['Mnemonic'] == final_mnemonic, 'CIQ'].empty else 'CIQ IQ Required'
                        
                        if manual_selection not in ['REMOVE ROW', '']:
                            if row['Account'] not in cash_flow_lookup_df['Account'].values:
                                new_entries.append({'Account': row['Account'], 'Mnemonic': final_mnemonic, 'CIQ': ciq_value, 'Label': row['Label']})
                            else:
                                cash_flow_lookup_df.loc[cash_flow_lookup_df['Account'] == row['Account'], 'Mnemonic'] = final_mnemonic
                                cash_flow_lookup_df.loc[cash_flow_lookup_df['Account'] == row['Account'], 'Label'] = row['Label']
                                cash_flow_lookup_df.loc[cash_flow_lookup_df['Account'] == row['Account'], 'CIQ'] = ciq_value
                    if new_entries:
                        cash_flow_lookup_df = pd.concat([cash_flow_lookup_df, pd.DataFrame(new_entries)], ignore_index=True)
                    cash_flow_lookup_df.reset_index(drop=True, inplace=True)
                    save_lookup_table(cash_flow_lookup_df, cash_flow_data_dictionary_file)
                    st.success("Data Dictionary Updated Successfully")

    with tab4:
        st.subheader("Cash Flow Data Dictionary")

        uploaded_dict_file = st.file_uploader("Upload a new Data Dictionary CSV", type=['csv'], key='dict_uploader_tab4_cfs')
        if uploaded_dict_file is not None:
            new_lookup_df = pd.read_csv(uploaded_dict_file)
            cash_flow_lookup_df = new_lookup_df  # Overwrite the entire DataFrame
            save_lookup_table(cash_flow_lookup_df, cash_flow_data_dictionary_file)
            st.success("Data Dictionary uploaded and updated successfully!")

        st.dataframe(cash_flow_lookup_df)

        remove_indices = st.multiselect("Select rows to remove", cash_flow_lookup_df.index, key='remove_indices_tab4_cfs')
        if st.button("Remove Selected Rows", key="remove_selected_rows_tab4_cfs"):
            cash_flow_lookup_df = cash_flow_lookup_df.drop(remove_indices).reset_index(drop=True)
            save_lookup_table(cash_flow_lookup_df, cash_flow_data_dictionary_file)
            st.success("Selected rows removed successfully!")
            st.dataframe(cash_flow_lookup_df)

        if st.button("Download Data Dictionary", key="download_data_dictionary_tab4_cfs"):
            excel_file = io.BytesIO()
            cash_flow_lookup_df.to_excel(excel_file, index=False)
            excel_file.seek(0)
            st.download_button("Download Excel", excel_file, "cash_flow_data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


######################################INCOME STATEMENT##################################
# Global variables and functions
# Global variables and functions
# Global variables and functions
# Global variables and functions
# Global variables and functions
# Global variables and functions
income_statement_data_dictionary_file = 'income_statement_data_dictionary.xlsx'

conversion_factors = {
    "Actuals": 1,
    "Thousands": 1000,
    "Millions": 1000000,
    "Billions": 1000000000
}

def clean_numeric_value_IS(value):
    try:
        value_str = str(value).strip()
        if value_str.startswith('(') and value_str.endswith(')'):
            value_str = '-' + value_str[1:-1]
        cleaned_value = re.sub(r'[$,]', '', value_str)
        return float(cleaned_value)
    except (ValueError, TypeError):
        return value

def apply_unit_conversion_IS(df, columns, factor):
    for selected_column in columns:
        if selected_column in df.columns:
            df[selected_column] = df[selected_column].apply(
                lambda x: x * factor if isinstance(x, (int, float)) else x)
    return df

def save_lookup_table(df, file_path):
    df.to_excel(file_path, index=False)

def create_combined_df(dfs):
    combined_df = pd.DataFrame()
    for i, df in enumerate(dfs):
        final_mnemonic_col = 'Final Mnemonic Selection'
        if final_mnemonic_col not in df.columns:
            st.error(f"Column '{final_mnemonic_col}' not found in dataframe {i+1}")
            continue
        
        date_cols = [col for col in df.columns if col not in ['Account', final_mnemonic_col, 'Mnemonic', 'Manual Selection', 'Sort Index']]
        if not date_cols:
            st.error(f"No date columns found in dataframe {i+1}")
            continue

        df_grouped = df.groupby([final_mnemonic_col]).sum(numeric_only=True).reset_index()
        df_melted = df_grouped.melt(id_vars=[final_mnemonic_col], value_vars=date_cols, var_name='Date', value_name='Value')
        df_pivot = df_melted.pivot(index=[final_mnemonic_col], columns='Date', values='Value')
        
        if combined_df.empty:
            combined_df = df_pivot
        else:
            combined_df = combined_df.join(df_pivot, how='outer')
    return combined_df.reset_index()

def sort_by_sort_index(df):
    # Sort by Sort Index if it exists
    if 'Sort Index' in df.columns:
        df = df.sort_values(by=['Sort Index'])
    return df

def aggregate_data_tab2(files):
    dataframes = []
    unique_accounts = set()

    for i, file in enumerate(files):
        df = pd.read_excel(file)
        df.columns = [str(col).strip() for col in df.columns]  # Clean column names
        
        if 'Account' not in df.columns:
            st.error(f"Column 'Account' not found in file {file.name}")
            return None

        df['Sort Index'] = range(1, len(df) + 1)  # Add sort index starting from 1 for each file
        dataframes.append(df)
        unique_accounts.update(df['Account'].dropna().unique())

    # Concatenate dataframes while retaining Account names
    concatenated_df = pd.concat(dataframes, ignore_index=True)

    # Split the data into numeric and date rows
    statement_date_rows = concatenated_df[concatenated_df['Account'].str.contains('Statement Date:', na=False)]
    numeric_rows = concatenated_df[~concatenated_df['Account'].str.contains('Statement Date:', na=False)]

    # Clean numeric values
    for col in numeric_rows.columns:
        if col not in ['Account', 'Sort Index', 'Positive decrease NI']:
            numeric_rows[col] = numeric_rows[col].apply(clean_numeric_value_IS)

    # Fill missing numeric values with 0
    numeric_rows.fillna(0, inplace=True)

    # Ensure all numeric columns are actually numeric
    for col in numeric_rows.columns:
        if col not in ['Account', 'Sort Index', 'Positive decrease NI']:
            numeric_rows[col] = pd.to_numeric(numeric_rows[col], errors='coerce').fillna(0)

    # Aggregation logic
    aggregated_df = numeric_rows.groupby(['Account'], as_index=False).sum(min_count=1)

    # Handle Statement Date rows separately
    statement_date_rows['Sort Index'] = 100
    statement_date_rows = statement_date_rows.groupby('Account', as_index=False).first()

    # Combine numeric rows and statement date rows
    final_df = pd.concat([aggregated_df, statement_date_rows], ignore_index=True)

    # Add "Positive decrease NI" column
    final_df.insert(1, 'Positive decrease NI', False)

    # Move Sort Index to the last column
    sort_index_column = final_df.pop('Sort Index')
    final_df['Sort Index'] = sort_index_column

    # Ensure "Statement Date:" is always last
    final_df.sort_values('Sort Index', inplace=True)

    return final_df

def income_statement():
    global income_statement_lookup_df

    # Load the Income Statement Data Dictionary
    if 'income_statement_lookup_df' not in globals():
        if os.path.exists(income_statement_data_dictionary_file):
            income_statement_lookup_df = pd.read_excel(income_statement_data_dictionary_file)
        else:
            income_statement_lookup_df = pd.DataFrame()

    st.title("INCOME STATEMENT LTMA")
    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Aggregate My Data", "Mappings and Data Aggregation", "Income Statement Data Dictionary"])

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
                    table_df = pd.DataFrame.from_dict(table, orient='index').sortindex()
                    table_df = table_df.sortindex(axis=1)
                    tables.append(table_df)
            all_tables = pd.concat(tables, axis=0, ignore_index=True)
            column_a = all_tables.columns[0]

            st.subheader("Data Preview")
            st.dataframe(all_tables)

            # Adding column renaming functionality
            st.subheader("Rename Columns")
            new_column_names = {}
            quarter_options = [f"FQ{quarter}{year}" for year in range(2018, 2027) for quarter in range(1, 5)]
            ytd_options = [f"YTD{quarter}{year}" for year in range(2018, 2027) for quarter in range(1, 5)]
            dropdown_options = [''] + ['Account'] + quarter_options + ytd_options

            for col in all_tables.columns:
                new_name_text = st.text_input(f"Rename '{col}' to:", value=col, key=f"rename_{col}_text")
                new_name_dropdown = st.selectbox(f"Or select predefined name for '{col}':", dropdown_options, key=f"rename_{col}_dropdown")
                new_column_names[col] = new_name_dropdown if new_name_dropdown else new_name_text
            
            all_tables.rename(columns=new_column_names, inplace=True)
            st.write("Updated Columns:", all_tables.columns.tolist())
            st.dataframe(all_tables)

            # Adding interactive data editor for row removal
            st.subheader("Edit and Remove Rows")
            editable_df = st.experimental_data_editor(all_tables, num_rows="dynamic", use_container_width=True)

            # Adding checkboxes for numerical column selection
            st.subheader("Select numerical columns")
            numerical_columns = []
            for col in all_tables.columns:
                if st.checkbox(f"Numerical column '{col}'", value=False, key=f"num_{col}"):
                    numerical_columns.append(col)

            # Unit conversion functionality
            st.subheader("Convert Units")
            selected_columns = st.multiselect("Select columns for conversion", options=numerical_columns, key="columns_selection")
            selected_conversion_factor = st.radio("Select conversion factor", options=list(conversion_factors.keys()), key="conversion_factor")

            if st.button("Apply Selected Labels and Generate Excel", key="apply_selected_labels_generate_excel_tab1"):
                updated_table = editable_df  # Use the edited dataframe

                # Convert selected numerical columns to numbers
                for col in numerical_columns:
                    updated_table[col] = updated_table[col].apply(clean_numeric_value_IS)
                
                # Apply unit conversion if selected
                if selected_conversion_factor and selected_conversion_factor in conversion_factors:
                    conversion_factor = conversion_factors[selected_conversion_factor]
                    updated_table = apply_unit_conversion_IS(updated_table, selected_columns, conversion_factor)

                # Convert all instances of '-' to '0'
                updated_table.replace('-', 0, inplace=True)

                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    with tab2:
        st.subheader("Aggregate My Data")

        # File uploader for Excel files
        uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx'], accept_multiple_files=True, key='excel_uploader_amd')
        if uploaded_files:
            aggregated_df = aggregate_data_tab2(uploaded_files)
            if aggregated_df is not None:
                st.subheader("Aggregated Data Preview")

                # Make the aggregated data interactive
                editable_df = st.experimental_data_editor(aggregated_df, use_container_width=True)

                # Exclude the last row from numeric conversions and multiplication logic
                editable_df_excluded = editable_df.iloc[:-1]

                # Ensure all numeric columns are properly converted to numeric types
                for col in editable_df_excluded.columns:
                    if col not in ['Account', 'Sort Index', 'Positive decrease NI']:
                        editable_df_excluded[col] = pd.to_numeric(editable_df_excluded[col], errors='coerce').fillna(0)

                # Apply the multiplication logic just before export
                numeric_cols = editable_df_excluded.select_dtypes(include='number').columns.tolist()
                for index, row in editable_df_excluded.iterrows():
                    if row['Positive decrease NI'] and row['Sort Index'] != 100:
                        for col in numeric_cols:
                            if col not in ['Sort Index']:
                                editable_df_excluded.at[index, col] = row[col] * -1

                # Combine the processed rows with the excluded last row
                final_df = pd.concat([editable_df_excluded, editable_df.iloc[-1:]], ignore_index=True)

                # Drop 'Positive decrease NI' from export
                if 'Positive decrease NI' in final_df.columns:
                    final_df.drop(columns=['Positive decrease NI'], inplace=True)

                excel_file = io.BytesIO()
                final_df.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "aggregated_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab3:
        st.subheader("Mappings and Data Aggregation")

        uploaded_excel_is = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader_tab3_is')

        currency_options_is = ["U.S. Dollar", "Euro", "British Pound Sterling", "Japanese Yen"]
        magnitude_options_is = ["Actuals", "Thousands", "Millions", "Billions", "Trillions"]

        selected_currency_is = st.selectbox("Select Currency", currency_options_is, key='currency_selection_tab3_is')
        selected_magnitude_is = st.selectbox("Select Magnitude", magnitude_options_is, key='magnitude_selection_tab3_is')
        company_name_is = st.text_input("Enter Company Name", key='company_name_input_is')

        if uploaded_excel_is is not None:
            df_is = pd.read_excel(uploaded_excel_is)
            st.write("Columns in the uploaded file:", df_is.columns.tolist())

            if 'Account' not in df_is.columns:
                st.error("The uploaded file does not contain an 'Account' column.")
            else:
                # Ensure Sort Index is present
                if 'Sort Index' not in df_is.columns:
                    df_is['Sort Index'] = range(1, len(df_is) + 1)

                # Function to get the best match based on Account using Levenshtein distance
                def get_best_match_is(account):
                    best_score_is = float('inf')
                    best_match_is = None
                    for _, lookup_row in income_statement_lookup_df.iterrows():
                        lookup_account = lookup_row['Account']
                        account_str = str(account)
                        score_is = levenshtein_distance(account_str.lower(), lookup_account.lower()) / max(len(account_str), len(lookup_account))
                        if score_is < best_score_is:
                            best_score_is = score_is
                            best_match_is = lookup_row
                    return best_match_is, best_score_is

                df_is['Mnemonic'] = ''
                df_is['Manual Selection'] = ''
                for idx, row in df_is.iterrows():
                    account_value = row['Account']
                    if pd.notna(account_value):
                        best_match_is, score_is = get_best_match_is(account_value)
                        if best_match_is is not None and score_is < 0.25:
                            df_is.at[idx, 'Mnemonic'] = best_match_is['Mnemonic']
                        else:
                            df_is.at[idx, 'Mnemonic'] = 'Human Intervention Required'
                    
                    if df_is.at[idx, 'Mnemonic'] == 'Human Intervention Required':
                        message_is = f"**Human Intervention Required for:** {account_value} - Index {idx}"
                        st.markdown(message_is)
                    
                    manual_selection_is = st.selectbox(
                        f"Select category for '{account_value}'",
                        options=[''] + income_statement_lookup_df['Mnemonic'].tolist() + ['REMOVE ROW'],
                        key=f"select_{idx}_tab3_is"
                    )
                    if manual_selection_is:
                        df_is.at[idx, 'Manual Selection'] = manual_selection_is.strip()

                st.dataframe(df_is[['Account', 'Mnemonic', 'Manual Selection', 'Sort Index']])  # Include 'Sort Index' as a helper column

                if st.button("Generate Excel with Lookup Results", key="generate_excel_lookup_results_tab3_is"):
                    df_is['Final Mnemonic Selection'] = df_is.apply(
                        lambda row: row['Manual Selection'].strip() if row['Manual Selection'].strip() != '' else row['Mnemonic'], 
                        axis=1
                    )
                    final_output_df_is = df_is[df_is['Final Mnemonic Selection'].str.strip() != 'REMOVE ROW'].copy()
                    
                    combined_df_is = create_combined_df([final_output_df_is])
                    combined_df_is = sort_by_sort_index(combined_df_is)

                    # Add CIQ column based on lookup
                    def lookup_ciq_is(mnemonic):
                        if mnemonic == 'Human Intervention Required':
                            return 'CIQ IQ Required'
                        ciq_value_is = income_statement_lookup_df.loc[income_statement_lookup_df['Mnemonic'] == mnemonic, 'CIQ']
                        if ciq_value_is.empty:
                            return 'CIQ IQ Required'
                        return ciq_value_is.values[0]
                    
                    combined_df_is['CIQ'] = combined_df_is['Final Mnemonic Selection'].apply(lookup_ciq_is)

                    columns_order_is = ['Final Mnemonic Selection', 'CIQ'] + [col for col in combined_df_is.columns if col not in ['Final Mnemonic Selection', 'CIQ']]
                    combined_df_is = combined_df_is[columns_order_is]

                    # Include the "As Presented" sheet without the CIQ column, and with the specified column order
                    as_presented_df_is = final_output_df_is.drop(columns=['CIQ', 'Mnemonic', 'Manual Selection'], errors='ignore')
                    as_presented_df_is = sort_by_sort_index(as_presented_df_is)
                    as_presented_df_is = as_presented_df_is.drop(columns=['Sort Index'], errors='ignore')
                    as_presented_columns_order_is = ['Account', 'Final Mnemonic Selection'] + [col for col in as_presented_df_is.columns if col not in ['Account', 'Final Mnemonic Selection']]
                    as_presented_df_is = as_presented_df_is[as_presented_columns_order_is]

                    excel_file_is = io.BytesIO()
                    with pd.ExcelWriter(excel_file_is, engine='xlsxwriter') as writer:
                        combined_df_is.to_excel(writer, sheet_name='Standardized', index=False)
                        as_presented_df_is.to_excel(writer, sheet_name='As Presented', index=False)
                        cover_df_is = pd.DataFrame({
                            'Selection': ['Currency', 'Magnitude', 'Company Name'],
                            'Value': [selected_currency_is, selected_magnitude_is, company_name_is]
                        })
                        cover_df_is.to_excel(writer, sheet_name='Cover', index=False)
                    excel_file_is.seek(0)
                    st.download_button("Download Excel", excel_file_is, "mnemonic_mapping_with_aggregation_is.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                if st.button("Update Data Dictionary with Manual Mappings", key="update_data_dictionary_tab3_is"):
                    df_is['Final Mnemonic Selection'] = df_is.apply(
                        lambda row: row['Manual Selection'] if row['Manual Selection'] not in ['REMOVE ROW', ''] else row['Mnemonic'], 
                        axis=1
                    )
                    new_entries_is = []
                    for idx, row in df_is.iterrows():
                        manual_selection_is = row['Manual Selection']
                        final_mnemonic_is = row['Final Mnemonic Selection']
                        if manual_selection_is == 'REMOVE ROW':
                            continue
                        ciq_value_is = income_statement_lookup_df.loc[income_statement_lookup_df['Mnemonic'] == final_mnemonic_is, 'CIQ'].values[0] if not income_statement_lookup_df.loc[income_statement_lookup_df['Mnemonic'] == final_mnemonic_is, 'CIQ'].empty else 'CIQ IQ Required'
                        
                        if manual_selection_is not in ['REMOVE ROW', '']:
                            if row['Account'] not in income_statement_lookup_df['Account'].values:
                                new_entries_is.append({'Account': row['Account'], 'Mnemonic': final_mnemonic_is, 'CIQ': ciq_value_is})
                            else:
                                income_statement_lookup_df.loc[income_statement_lookup_df['Account'] == row['Account'], 'Mnemonic'] = final_mnemonic_is
                                income_statement_lookup_df.loc[income_statement_lookup_df['Account'] == row['Account'], 'CIQ'] = ciq_value_is
                    if new_entries_is:
                        income_statement_lookup_df = pd.concat([income_statement_lookup_df, pd.DataFrame(new_entries_is)], ignore_index=True)
                    income_statement_lookup_df.resetindex(drop=True, inplace=True)
                    save_lookup_table(income_statement_lookup_df, income_statement_data_dictionary_file)
                    st.success("Data Dictionary Updated Successfully")

    with tab4:
        st.subheader("Income Statement Data Dictionary")

        # Load most recent CSV in memory until a new CSV is uploaded
        if 'income_statement_data' not in st.session_state:
            if os.path.exists(income_statement_data_dictionary_file):
                st.session_state.income_statement_data = pd.read_excel(income_statement_data_dictionary_file)
            else:
                st.session_state.income_statement_data = pd.DataFrame()

        uploaded_dict_file_is = st.file_uploader("Upload a new Data Dictionary CSV", type=['csv'], key='dict_uploader_tab4_is')
        if uploaded_dict_file_is is not None:
            new_lookup_df_is = pd.read_csv(uploaded_dict_file_is)
            st.session_state.income_statement_data = new_lookup_df_is  # Update the session state with new DataFrame
            save_lookup_table(new_lookup_df_is, income_statement_data_dictionary_file)
            st.success("Data Dictionary uploaded and updated successfully!")

        # Use the data from the session state
        st.dataframe(st.session_state.income_statement_data)

        remove_indices_is = st.multiselect("Select rows to remove", st.session_state.income_statement_data.index, key='remove_indices_tab4_is')
        if st.button("Remove Selected Rows", key="remove_selected_rows_tab4_is"):
            st.session_state.income_statement_data = st.session_state.income_statement_data.drop(remove_indices_is).resetindex(drop=True)
            save_lookup_table(st.session_state.income_statement_data, income_statement_data_dictionary_file)
            st.success("Selected rows removed successfully!")
            st.dataframe(st.session_state.income_statement_data)

        if st.button("Download Data Dictionary", key="download_data_dictionary_tab4_is"):
            excel_file_is = io.BytesIO()
            st.session_state.income_statement_data.to_excel(excel_file_is, index=False)
            excel_file_is.seek(0)
            st.download_button("Download Excel", excel_file_is, "income_statement_data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

####################################### Populate CIQ Template ###################################
import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def populate_ciq_template():
    st.title("Populate CIQ Template")

    tab1 = st.tabs(["Upload Template"])[0]

    with tab1:
        uploaded_template = st.file_uploader("Upload Template", type=['xlsx', 'xlsm'], key='template_uploader')
        uploaded_income_statement = st.file_uploader("Upload Completed Income Statement", type=['xlsx', 'xlsm'], key='income_statement_uploader')

        if uploaded_template and uploaded_income_statement:
            # Read the uploaded template and income statement files
            template_book = load_workbook(uploaded_template)
            income_statement_df = pd.read_excel(uploaded_income_statement, sheet_name="Standardized")
            template_df = pd.read_excel(uploaded_template, sheet_name="Income Statement")

            if st.button("Populate Template"):
                # Extract CIQ Mnemonics and dates from the completed income statement
                ciq_mnemonics = income_statement_df.iloc[:, 1]
                income_statement_dates = income_statement_df.columns[2:]

                # Debug print statements
                st.write("Income Statement Dates:", income_statement_dates)

                # Extract mnemonics and dates from the template
                template_mnemonics = template_df.iloc[:, 8]
                template_dates = template_df.columns[3:7]

                # Debug print statements
                st.write("Template Dates:", template_dates)

                for i, mnemonic in enumerate(template_mnemonics):
                    if pd.notna(mnemonic):
                        # Find corresponding row in the income statement
                        income_statement_row = income_statement_df[ciq_mnemonics == mnemonic]
                        if not income_statement_row.empty:
                            for j, date in enumerate(template_dates):
                                if date in income_statement_dates.values:
                                    # Find corresponding column in the income statement
                                    income_statement_col = income_statement_dates.get_loc(date)
                                    # Populate the value in the template
                                    template_df.iat[i, 3 + j] = income_statement_row.iat[0, income_statement_col]

                # Update the 'Income Statement' sheet in the template workbook
                template_sheet = template_book["Income Statement"]
                for r_idx, row in enumerate(dataframe_to_rows(template_df, index=False, header=True), 1):
                    for c_idx, value in enumerate(row, 1):
                        template_sheet.cell(row=r_idx, column=c_idx, value=value)

                # Save the populated template to an in-memory file
                excel_file = io.BytesIO()
                template_book.save(excel_file)
                excel_file.seek(0)

                # Offer the populated template for download
                st.download_button(
                    label="Download Populated Template",
                    data=excel_file,
                    file_name="populated_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

def main():
    st.sidebar.title("Navigation")
    selection = st.sidebar.radio("Go to", ["Balance Sheet", "Cash Flow Statement", "Income Statement", "Populate CIQ Template"])

    if selection == "Balance Sheet":
        balance_sheet()
    elif selection == "Cash Flow Statement":
        cash_flow_statement()
    elif selection == "Income Statement":
        income_statement()
    elif selection == "Populate CIQ Template":
        populate_ciq_template()

if __name__ == '__main__':
    main()

