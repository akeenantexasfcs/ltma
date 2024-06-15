#!/usr/bin/env python
# coding: utf-8

# In[19]:


import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import json
from collections import defaultdict
from Levenshtein import distance as levenshtein_distance

# Global variables and functions
income_statement_lookup_df = pd.DataFrame()
income_statement_data_dictionary_file = 'income_statement_data_dictionary.xlsx'

conversion_factors = {
    "Actuals": 1,
    "Thousands": 1000,
    "Millions": 1000000,
    "Billions": 1000000000
}

def save_lookup_table(df, file_path):
    df.to_excel(file_path, index=False)

def clean_numeric_value(value):
    try:
        value_str = str(value).strip()
        if value_str.startswith('(') and value_str.endswith(')'):
            value_str = '-' + value_str[1:-1]
        cleaned_value = re.sub(r'[$,]', '', value_str)
        return float(cleaned_value)
    except (ValueError, TypeError):
        return value

def apply_unit_conversion(df, columns, factor):
    for selected_column in columns:
        if selected_column in df.columns:
            df[selected_column] = df[selected_column].apply(
                lambda x: x * factor if isinstance(x, (int, float)) else x)
    return df

def aggregate_data(files):
    dataframes = []
    unique_accounts = set()

    for i, file in enumerate(files):
        df = pd.read_excel(file)
        df.columns = [str(col).strip() for col in df.columns]  # Clean column names
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
            numeric_rows[col] = numeric_rows[col].apply(clean_numeric_value)

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

def income_statement():
    global income_statement_lookup_df

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

            st.subheader("Data Preview")
            all_tables["Remove"] = False
            all_tables["Remove"] = all_tables["Remove"].astype(bool)

            st.dataframe(all_tables.astype(str))

            # Create data editor for row removal
            edited_table = st.experimental_data_editor(all_tables, num_rows="dynamic", key="data_editor_is")

            rows_to_remove = edited_table[edited_table["Remove"] == True].index.tolist()
            if rows_to_remove:
                all_tables = all_tables.drop(rows_to_remove).reset_index(drop=True)
                st.subheader("Updated Data Preview")
                st.dataframe(all_tables.astype(str))

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
            st.write("Updated Columns:", all_tables.columns.tolist())
            st.dataframe(all_tables.astype(str))

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

    with tab2:
        st.subheader("Aggregate My Data")

        # File uploader for Excel files
        uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx'], accept_multiple_files=True, key='excel_uploader_amd')
        if uploaded_files:
            aggregated_df = aggregate_data(uploaded_files)
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
                    
                    # Safeguard: Check if 'Mnemonic' column exists in the dataframe
                    if 'Mnemonic' in income_statement_lookup_df.columns:
                        mnemonic_options = income_statement_lookup_df['Mnemonic'].tolist() + ['REMOVE ROW']
                    else:
                        st.error("'Mnemonic' column not found in lookup dataframe.")
                        mnemonic_options = ['REMOVE ROW']

                    manual_selection_is = st.selectbox(
                        f"Select category for '{account_value}'",
                        options=[''] + mnemonic_options,
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
                    income_statement_lookup_df.reset_index(drop=True, inplace=True)
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
            st.session_state.income_statement_data = st.session_state.income_statement_data.drop(remove_indices_is).reset_index(drop=True)
            save_lookup_table(st.session_state.income_statement_data, income_statement_data_dictionary_file)
            st.success("Selected rows removed successfully!")
            st.dataframe(st.session_state.income_statement_data)

        if st.button("Download Data Dictionary", key="download_data_dictionary_tab4_is"):
            excel_file_is = io.BytesIO()
            st.session_state.income_statement_data.to_excel(excel_file_is, index=False)
            excel_file_is.seek(0)
            st.download_button("Download Excel", excel_file_is, "income_statement_data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def main():
    st.sidebar.title("Navigation")
    selection = st.sidebar.radio("Go to", ["Balance Sheet", "Cash Flow Statement", "Income Statement"])

    if selection == "Balance Sheet":
        st.write("Balance Sheet functionality not implemented yet.")
    elif selection == "Cash Flow Statement":
        st.write("Cash Flow Statement functionality not implemented yet.")
    elif selection == "Income Statement":
        income_statement()

if __name__ == '__main__':
    main()

