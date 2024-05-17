#!/usr/bin/env python
# coding: utf-8

# In[3]:


import io
import json
import pandas as pd
import streamlit as st
from Levenshtein import distance as levenshtein_distance

# Lookup table
lookup_data = {
    "Account": ["Cash and cash equivalents", "Line of credit", "Goodwill", 
                "Total Current Assets", "Total Assets", "Total Current Liabilities"],
    "Mnemonic": ["Cash & Cash Equivalents", "Short-Term Debt", "Goodwill", 
                 "Total Current Assets", "Total Assets", "Total Current Liabilities"],
    "CIQ": ["IQ_CASH_EQUIV", "IQ_ST_INVEST", "IQ_GW", 
            "IQ_TOTAL_CA", "IQ_TOTAL_ASSETS", "IQ_TOTAL_CL"]
}
lookup_df = pd.DataFrame(lookup_data)

def main():
    st.title("Table Extractor and Label Generators")

    # Define the tabs
    tab1, tab2, tab3 = st.tabs(["Table Extractor", "Mnemonic Mapping", "Balance Sheet Data Dictionary"])

    with tab1:
        # File uploader for the Table Extractor
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

            # Display the data preview after JSON conversion
            st.subheader("Data Preview")
            st.dataframe(all_tables)

            labels = ["Current Assets", "Non Current Assets", "Total Assets", "Current Liabilities", 
                      "Non Current Liabilities", "Total Liabilities", "Shareholder's Equity", 
                      "Total Equity", "Total Equity and Liabilities"]
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

            # Add an update button to apply the changes and update the preview
            if st.button("Update Labels Preview"):
                updated_table = update_labels()
                st.subheader("Updated Data Preview")
                st.dataframe(updated_table)

            if st.button("Apply Selected Labels and Generate Excel"):
                updated_table = update_labels()
                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables_with_labels.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        # File uploader for Mnemonic Mapping
        uploaded_excel = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader')

        if uploaded_excel is not None:
            df = pd.read_excel(uploaded_excel)
            st.write("Columns in the uploaded file:", df.columns.tolist())  # Display the columns of the uploaded DataFrame

            # Allow the user to rename columns
            new_column_names = {}
            st.subheader("Rename Columns")
            for col in df.columns:
                new_name = st.text_input(f"Rename '{col}' to:", value=col, key=f"rename_{col}")
                new_column_names[col] = new_name
            
            # Apply the new column names
            df.rename(columns=new_column_names, inplace=True)
            st.write("Updated Columns:", df.columns.tolist())

            # Display the updated DataFrame
            st.dataframe(df)

            if 'Account' not in df.columns:
                st.error("The uploaded file does not contain an 'Account' column.")
            else:
                def get_best_match(account):
                    best_score = float('inf')
                    best_match = None
                    for lookup_account in lookup_df['Account']:
                        # Ensure account is a string
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
                    if pd.notna(account_value):  # Ensure the account value is not NaN
                        best_match, score = get_best_match(account_value)
                        if score < 0.2:
                            df.at[idx, 'Mnemonic'] = lookup_df.loc[lookup_df['Account'] == best_match, 'Mnemonic'].values[0]
                        else:
                            df.at[idx, 'Mnemonic'] = 'Human Intervention Required'
                    if df.at[idx, 'Mnemonic'] == 'Human Intervention Required':
                        st.markdown(f"**Human Intervention Required for:** {account_value}")
                    df.at[idx, 'Manual Selection'] = st.selectbox(
                        f"Select category for '{account_value}'",
                        options=[''] + lookup_df['Account'].tolist() + ['Other Category', 'REMOVE ROW'],
                        key=f"select_{idx}"
                    ).strip()  # Strip any whitespace

                # Display the dataframe with the Mnemonic and Manual Selection columns for user interaction
                st.dataframe(df[['Account', 'Mnemonic', 'Manual Selection']])

                if st.button("Generate Excel with Lookup Results"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'].strip() if row['Mnemonic'] == 'Human Intervention Required' else row['Mnemonic'], 
                        axis=1
                    )
                    # Remove rows where 'Final Mnemonic Selection' is 'REMOVE ROW'
                    final_output_df = df[df['Final Mnemonic Selection'].str.strip() != 'REMOVE ROW'].copy()
                    excel_file = io.BytesIO()
                    final_output_df.to_excel(excel_file, index=False)
                    excel_file.seek(0)
                    st.download_button("Download Excel", excel_file, "mnemonic_mapping_with_final_selection.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                if st.button("Update Data Dictionary with Manual Mappings"):
                    for idx, row in df.iterrows():
                        manual_selection = row['Manual Selection']
                        if manual_selection not in ['Other Category', 'REMOVE ROW', '']:
                            if manual_selection not in lookup_df['Account'].values:
                                new_row = {'Account': manual_selection, 'Mnemonic': row['Final Mnemonic Selection'], 'CIQ': ''}
                                lookup_df = lookup_df.append(new_row, ignore_index=True)
                            else:
                                lookup_df.loc[lookup_df['Account'] == manual_selection, 'Mnemonic'] = row['Final Mnemonic Selection']
                    lookup_df.reset_index(drop=True, inplace=True)
                    st.success("Data Dictionary Updated Successfully")

    with tab3:
        st.subheader("Balance Sheet Data Dictionary")
        st.dataframe(lookup_df)

        if st.button("Download Data Dictionary"):
            excel_file = io.BytesIO()
            lookup_df.to_excel(excel_file, index=False)
            excel_file.seek(0)
            st.download_button("Download Excel", excel_file, "data_dictionary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    main()

