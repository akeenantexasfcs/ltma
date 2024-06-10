#!/usr/bin/env python
# coding: utf-8

# In[2]:


import streamlit as st
import pandas as pd
import io
import json

def clean_numeric_value(val):
    try:
        return float(val)
    except ValueError:
        return val

def apply_unit_conversion(df, columns, factor):
    for col in columns:
        df[col] = df[col].apply(lambda x: x / factor if isinstance(x, (int, float)) else x)
    return df

def process_file(file):
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error processing file {file.name}: {e}")
        return None

def aggregate_data(df):
    aggregated_df = df.groupby("Account").sum(numeric_only=True).reset_index()
    return aggregated_df

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
            edited_table = st.data_editor(all_tables, num_rows="dynamic", key="data_editor_is")

            rows_to_remove = edited_table[edited_table["Remove"] == True].index.tolist()
            if rows_to_remove:
                all_tables = all_tables.drop(rows_to_remove).reset_index(drop=True)
                st.subheader("Updated Data Preview")
                st.dataframe(all_tables.astype(str))

            # Column Naming setup
            st.subheader("Rename Columns")
            quarter_options = [f"Q{i}-{year}" for year in range(2018, 2027) for i in range(1, 4)]
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

                # Append Statement Date row
                statement_date_row = pd.DataFrame({col: [statement_date_values.get(col, "")] for col in updated_table.columns})
                updated_table = pd.concat([updated_table, statement_date_row], ignore_index=True)

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
            combined_df["Statement Intent"] = ""

            st.subheader("Data Preview with Statement Intent")
            edited_combined_df = st.data_editor(combined_df, num_rows="dynamic", key="data_editor_combined_df")

            statement_intent_options = ["", "Increase NI", "Decrease NI", "Remove"]
            for index, row in edited_combined_df.iterrows():
                statement_intent = st.selectbox(f"What is the statement intent of {row['Account']}?", options=statement_intent_options, key=f"statement_intent_{index}")
                edited_combined_df.at[index, "Statement Intent"] = statement_intent

            # Remove rows with Statement Intent as 'Remove'
            final_df = edited_combined_df[edited_combined_df["Statement Intent"] != "Remove"]

            st.dataframe(final_df)

            aggregated_table = aggregate_data(final_df)
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
        st.subheader("Placeholder for Mappings and Data Aggregation")

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

