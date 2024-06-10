#!/usr/bin/env python
# coding: utf-8

# In[7]:


import io
import json
import os
import pandas as pd
import streamlit as st
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
            quarter_options = [f"Q{i}-{year}" for year in range(2018, 2027) for i in range(1, 5)]
            ytd_options = [f"YTD {year}" for year in range(2018, 2027)]
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



# Global variables and functions
income_statement_lookup_df = pd.DataFrame()
income_statement_data_dictionary_file = 'income_statement_data_dictionary.xlsx'
initial_upload_data = None  # Variable to store the initial uploaded data
initial_aggregated_data = None  # Variable to store the initial aggregated data

def save_lookup_table(df, file_path):
    df.to_excel(file_path, index=False)

def clean_numeric_value(value):
    value_str = str(value).strip()
    if value_str.startswith('(') and value_str.endswith(')'):
        value_str = '-' + value_str[1:-1]
    cleaned_value = re.sub(r'[$,]', '', value_str)
    try:
        return float(cleaned_value)
    except ValueError:
        return 0

def apply_unit_conversion(df, columns, factor):
    for selected_column in columns:
        if selected_column in df.columns:
            df[selected_column] = df[selected_column].apply(
                lambda x: x * factor if isinstance(x, (int, float)) else x)
    return df

def aggregate_data(files):
    dataframes = []
    for file in files:
        df = pd.read_excel(file)
        dataframes.append(df)
    concatenated_df = pd.concat(dataframes, ignore_index=True)
    
    # Aggregation logic
    if 'Account' in concatenated_df.columns:
        aggregated_df = concatenated_df.groupby('Account').sum().reset_index()
    else:
        st.error("Account column is not present in one or more files.")
        return None
    return aggregated_df

def income_statement():
    global income_statement_lookup_df, initial_upload_data, initial_aggregated_data

    st.title("INCOME STATEMENT LTMA")
    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Aggregate My Data", "Mappings and Data Aggregation", "Income Statement Data Dictionary"])

    with tab1:
        st.subheader("Table Extractor")

        uploaded_file = st.file_uploader("Choose a JSON file", type="json", key='json_uploader_is')
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

            all_tables["Remove"] = False
            all_tables["Remove"] = all_tables["Remove"].astype(bool)
            initial_upload_data = all_tables.copy()  # Store the initial upload data

            # Column Naming setup
            st.subheader("Rename Columns")
            quarter_options = [f"Q{i}-{year}" for year in range(2018, 2027) for i in range(1, 5)]
            ytd_options = [f"YTD {year}" for year in range(2018, 2027)]
            dropdown_options = [''] + ['Account'] + quarter_options + ytd_options + ['Remove']

            new_column_names = {col: col for col in all_tables.columns}
            for col in all_tables.columns:
                new_name_text = st.text_input(f"Rename '{col}' to:", value=col, key=f"rename_{col}_text_is")
                new_name_dropdown = st.selectbox(f"Or select predefined name for '{col}':", dropdown_options, key=f"rename_{col}_dropdown_is", index=0)
                new_column_names[col] = new_name_dropdown if new_name_dropdown else new_name_text

            all_tables.rename(columns=new_column_names, inplace=True)

            # Exclude columns that are renamed to 'Remove'
            columns_to_keep = [col for col in all_tables.columns if new_column_names.get(col) != 'Remove']

            # Columns to keep setup
            st.subheader("Select Columns to Keep Before Export")
            final_columns_to_keep = []
            for col in columns_to_keep:
                if st.checkbox(f"Keep column '{col}'", value=True, key=f"keep_{col}_is"):
                    final_columns_to_keep.append(col)

            # Select Numerical Columns conversion
            st.subheader("Select Numerical Columns")
            numerical_columns = []
            for col in all_tables.columns:
                if st.checkbox(f"Numerical column '{col}'", value=False, key=f"num_{col}_is"):
                    numerical_columns.append(col)

            # Add Statement Date
            st.subheader("Add Statement Date")
            statement_date_values = {}
            for col in final_columns_to_keep:
                if col == "Account":
                    statement_date_values[col] = "Statement Date:"
                else:
                    statement_date_values[col] = st.text_input(f"Statement date for '{col}'", key=f"statement_date_{col}")

            # Unit labeling
            st.subheader("Convert Units")
            selected_columns = st.multiselect("Select columns for conversion", options=numerical_columns, key="columns_selection_is")
            selected_value = st.radio("Select conversion value", ["No Conversions Necessary", "Thousands", "Millions", "Billions"], index=0, key="conversion_value_is")

            conversion_factors = {
                "No Conversions Necessary": 1,
                "Thousands": 1000,
                "Millions": 1000000,
                "Billions": 1000000000
            }

            if st.button("Apply and Generate Excel", key="apply_generate_excel_is"):
                updated_table = all_tables.copy()
                updated_table = updated_table[[col for col in final_columns_to_keep if col in updated_table.columns]]

                for col in numerical_columns:
                    if col in updated_table.columns:
                        updated_table[col] = updated_table[col].apply(clean_numeric_value)

                if selected_value != "No Conversions Necessary":
                    updated_table = apply_unit_conversion(updated_table, selected_columns, conversion_factors[selected_value])

                updated_table.replace('-', 0, inplace=True)

                # Move Statement Date row to the last row if it exists
                statement_date_row = updated_table[updated_table['Account'].str.contains('Statement Date:', na=False)]
                updated_table = updated_table[~updated_table['Account'].str.contains('Statement Date:', na=False)]
                updated_table = pd.concat([updated_table, statement_date_row], ignore_index=True)

                # Drop 'Increases NI' and 'Decreases NI' columns
                if 'Increases NI' in updated_table.columns:
                    updated_table.drop(columns=['Increases NI'], inplace=True)
                if 'Decreases NI' in updated_table.columns:
                    updated_table.drop(columns=['Decreases NI'], inplace=True)

                # Rename 'Increases NI' to 'Positive Number Increases Net Income'
                if 'Positive Number Increases Net Income' in updated_table.columns:
                    updated_table.rename(columns={'Positive Number Increases Net Income': 'Positive Number Increases Net Income'}, inplace=True)

                # Drop 'Positive Number Increases Net Income' from export
                if 'Positive Number Increases Net Income' in updated_table.columns:
                    updated_table.drop(columns=['Positive Number Increases Net Income'], inplace=True)

                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables_with_labels.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.subheader("Data Preview")
            st.dataframe(all_tables.astype(str))

            # Create data editor for row removal
            edited_table = st.experimental_data_editor(all_tables, num_rows="dynamic", key="data_editor_is")

            rows_to_remove = edited_table[edited_table["Remove"] == True].index.tolist()
            if rows_to_remove:
                all_tables = all_tables.drop(rows_to_remove).reset_index(drop=True)
                st.subheader("Updated Data Preview")
                st.dataframe(all_tables.astype(str))

    with tab2:
        st.subheader("Aggregate My Data")

        # File uploader for Excel files
        uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx'], accept_multiple_files=True, key='excel_uploader_amd')
        if uploaded_files:
            aggregated_df = aggregate_data(uploaded_files)
            if aggregated_df is not None:
                initial_aggregated_data = aggregated_df.copy()  # Store the initial aggregated data

                st.subheader("Aggregated Data Preview")
                st.dataframe(aggregated_df)

                # Adding statement intent columns
                st.subheader("Interactive Data Frame")
                aggregated_df["Positive Number Increases Net Income"] = False
                aggregated_df["Statement Intent"] = ""

                edited_df = st.experimental_data_editor(
                    aggregated_df,
                    use_container_width=True,
                    num_rows="dynamic"  # This enables dynamic row handling
                )

                # Ensure only one of "Positive Number Increases Net Income" or "Decreases NI" can be selected at a time
                def update_selection(df, index):
                    if df.at[index, "Positive Number Increases Net Income"]:
                        df.at[index, "Statement Intent"] = "Positive Number Increases Net Income"
                    else:
                        df.at[index, "Statement Intent"] = ""

                for i in edited_df.index:
                    update_selection(edited_df, i)

                # Dynamic column ordering
                columns = edited_df.columns.tolist()
                ordered_columns = st.multiselect("Reorder Columns", options=columns, default=columns, key='reorder_columns')

                edited_df = edited_df[ordered_columns]

                st.dataframe(edited_df)

                if st.button("Download Aggregated Data", key='download_aggregated_data_amd'):
                    filtered_df = edited_df

                    # Move Statement Date row to the last row if it exists
                    statement_date_row = filtered_df[filtered_df['Account'].str.contains('Statement Date:', na=False)]
                    filtered_df = filtered_df[~filtered_df['Account'].str.contains('Statement Date:', na=False)]
                    filtered_df = pd.concat([filtered_df, statement_date_row], ignore_index=True)

                    # Drop 'Positive Number Increases Net Income' from export
                    if 'Positive Number Increases Net Income' in filtered_df.columns:
                        filtered_df.drop(columns=['Positive Number Increases Net Income'], inplace=True)

                    excel_file = io.BytesIO()
                    filtered_df.to_excel(excel_file, index=False)
                    excel_file.seek(0)
                    st.download_button("Download Excel", excel_file, "aggregated_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # Button to revert to initial upload data
                if st.button("Revert to Initial Data"):
                    if initial_aggregated_data is not None:
                        aggregated_df = initial_aggregated_data.copy()
                        st.experimental_rerun()

    with tab3:
        st.subheader("Mappings and Data Aggregation")

        # Example of potential content for this tab
        st.write("This section can be used to define and manage mappings for data aggregation.")
        st.write("You can upload mappings, define rules, and manage aggregated data here.")

        uploaded_mapping_file = st.file_uploader("Upload Mapping File", type=['xlsx', 'csv'], key='mapping_uploader_mda')
        if uploaded_mapping_file:
            if uploaded_mapping_file.name.endswith('.xlsx'):
                mapping_df = pd.read_excel(uploaded_mapping_file)
            else:
                mapping_df = pd.read_csv(uploaded_mapping_file)

            st.write("Uploaded Mapping Data")
            st.dataframe(mapping_df)

            # Here you can add more functionality such as editing the mapping, applying it to data, etc.

    with tab4:
        st.subheader("Income Statement Data Dictionary")

        uploaded_dict_file = st.file_uploader("Upload a new Data Dictionary CSV", type=['csv'], key='dict_uploader_tab4_is')
        if uploaded_dict_file is not None:
            new_lookup_df = pd.read_csv(uploaded_dict_file)
            income_statement_lookup_df = new_lookup_df  # Overwrite the entire DataFrame
            save_lookup_table(income_statement_lookup_df, income_statement_data_dictionary_file)
            st.success("Data Dictionary uploaded and updated successfully!")

        st.dataframe(income_statement_lookup_df)

        remove_indices = st.multiselect("Select rows to remove", income_statement_lookup_df.index, key='remove_indices_tab4_is')
        if st.button("Remove Selected Rows", key="remove_selected_rows_tab4_is"):
            income_statement_lookup_df = income_statement_lookup_df.drop(remove_indices).reset_index(drop=True)
            save_lookup_table(income_statement_lookup_df, income_statement_data_dictionary_file)
            st.success("Selected rows removed successfully!")
            st.dataframe(income_statement_lookup_df)

        if st.button("Download Data Dictionary", key="download_data_dictionary_tab4_is"):
            excel_file = io.BytesIO()
            income_statement_lookup_df.to_excel(excel_file, index=False)
            excel_file.seek(0)
            st.download_button("Download Excel", excel_file, "income_statement_data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def main():
    st.sidebar.title("Navigation")
    selection = st.sidebar.radio("Go to", ["Balance Sheet", "Cash Flow Statement", "Income Statement"])

    if selection == "Balance Sheet":
        balance_sheet()
    elif selection == "Cash Flow Statement":
        cash_flow_statement()
    elif selection == "Income Statement":
        income_statement()

if __name__ == '__main__':
    main()

