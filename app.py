#!/usr/bin/env python
# coding: utf-8

# In[29]:


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
    "Label": ["Current Assets", "Current Assets", "Current Assets", "Current Assets", "Current Assets", "Current Assets", 
              "Current Assets", "Current Assets", "Current Assets", "Current Assets", "Current Liabilities", "Current Liabilities", 
              "Current Liabilities", "Current Liabilities", "Current Liabilities", "Current Liabilities", "Current Liabilities", 
              "Equity", "Equity", "Equity", "Equity", "Equity", "Equity", "Equity", "Non Current Assets", "Non Current Assets", 
              "Non Current Assets", "Non Current Assets", "Non Current Assets", "Non Current Assets", "Non Current Liabilities", 
              "Non Current Liabilities", "Non Current Liabilities", "Non Current Liabilities", "Non Current Liabilities", "Non Current Liabilities", 
              "Total Equity and Liabilities"],
    "Account": ["Cash and Equivalents", "Short Term Investments", "Trading Asset Securities", "Accounts Receivable", 
                "Other Receivables", "Inventory", "Prepaid Exp.", "Restricted Cash", "Other Current Assets", "Total Current Assets", 
                "Accounts Payable", "Accrued Exp.", "Short-term Borrowings", "Current Portion of Long Term Debt", 
                "Curr. Portion of Leases", "Other Current Liabilities", "Total Current Liabilities", "Total Pref. Equity", 
                "Common Equity", "Additional Paid In Capital", "Retained Earnings", "Treasury Stock", 
                "Comprehensive Inc. and Other", "Minority Interest", "Net Property, Plant & Equipment", 
                "Long-term Investments", "Goodwill", "Other Intangibles", "Right-of-Use Asset-Net", 
                "Other Long-Term Assets", "Long-Term Debt", "Long-Term Leases", 
                "Pension & Other Post-Retire. Benefits", "Def. Tax Liability, Non-Curr.", 
                "Other Non-Current Liabilities", "Total Liabilities", "Total Liabilities And Equity"],
    "Mnemonic": ["Cash and Equivalents", "Short Term Investments", "Trading Asset Securities", "Accounts Receivable", 
                 "Other Receivables", "Inventory", "Prepaid Exp.", "Restricted Cash", "Other Current Assets", "Total Current Assets", 
                 "Accounts Payable", "Accrued Exp.", "Short-term Borrowings", "Current Portion of Long Term Debt", 
                 "Curr. Portion of Leases", "Other Current Liabilities", "Total Current Liabilities", "Total Pref. Equity", 
                 "Common Equity", "Additional Paid In Capital", "Retained Earnings", "Treasury Stock", 
                 "Comprehensive Inc. and Other", "Minority Interest", "Net Property, Plant & Equipment", 
                 "Long-term Investments", "Goodwill", "Other Intangibles", "Right-of-Use Asset-Net", 
                 "Other Long-Term Assets", "Long-Term Debt", "Long-Term Leases", 
                 "Pension & Other Post-Retire. Benefits", "Def. Tax Liability, Non-Curr.", 
                 "Other Non-Current Liabilities", "Total Liabilities", "Total Liabilities And Equity"],
    "CIQ": ["IQ_CASH_EQUIV", "IQ_ST_INVEST", "IQ_TRADING_ASSETS", "IQ_AR", 
            "IQ_OTHER_RECEIV", "IQ_INVENTORY", "IQ_PREPAID_EXP", "IQ_RESTRICTED_CASH", "IQ_OTHER_CA_SUPPL", "IQ_TOTAL_CA", 
            "IQ_AP", "IQ_AE", "IQ_ST_DEBT", "IQ_CURRENT_PORT_DEBT", 
            "IQ_CURRENT_PORT_LEASES", "IQ_OTHER_CL_SUPPL", "IQ_TOTAL_CL", "IQ_PREF_EQUITY", 
            "IQ_COMMON", "IQ_APIC", "IQ_RE", "IQ_TREASURY", 
            "IQ_OTHER_EQUITY", "IQ_MINORITY_INTEREST", "IQ_NPPE", 
            "IQ_LT_INVEST", "IQ_GW", "IQ_OTHER_INTAN", "IQ_RUA_NET", 
            "IQ_OTHER_LT_ASSETS", "IQ_LT_DEBT", "IQ_LONG_TERM_LEASES", 
            "IQ_PENSION", "IQ_DEF_TAX_LIAB_LT", 
            "IQ_OTHER_LIAB_LT", "IQ_TOTAL_LIAB", "IQ_TOTAL_LIAB_EQUITY"]
}

# Define the initial lookup data for Cash Flow
initial_cash_flow_lookup_data = {
    "Label": ["Operating Activities", "Operating Activities", "Operating Activities", "Operating Activities", "Operating Activities", 
              "Operating Activities", "Operating Activities", "Operating Activities", "Operating Activities", "Operating Activities", 
              "Operating Activities", "Operating Activities", "Operating Activities", "Operating Activities", "Operating Activities", 
              "Operating Activities", "Operating Activities", "Investing Activities", "Investing Activities", "Investing Activities", 
              "Investing Activities", "Investing Activities", "Financing Activities", "Financing Activities", "Financing Activities", 
              "Financing Activities", "Financing Activities", "Financing Activities", "Financing Activities", "Financing Activities", 
              "Financing Activities", "Financing Activities", "Financing Activities", "Cash from other", "Cash from other"],
    "Account": ["Net Income", "Depreciation & Amort.", "Amort. of Goodwill and Intangibles", "Other Amortization", 
                "(Gain) Loss From Sale Of Assets", "(Gain) Loss On Sale Of Invest.", "Asset Writedown & Restructuring Costs", 
                "Stock-Based Compensation", "Net Cash From Discontinued Ops.", "Other Operating Activities", 
                "Change in Trad. Asset Securities", "Change in Acc. Receivable", "Change In Inventories", 
                "Change in Acc. Payable", "Change in Unearned Rev.", "Change in Inc. Taxes", "Change in Def. Taxes", 
                "Change in Other Net Operating Assets", "Capital Expenditure", "Sale of Property, Plant, and Equipment", 
                "Cash Acquisitions", "Divestitures", "Other Investing Activities", "Short Term Debt Issued", 
                "Long-Term Debt Issued", "Short Term Debt Repaid", "Long-Term Debt Repaid", "Issuance of Common Stock", 
                "Repurchase of Common Stock", "Issuance of Pref. Stock", "Repurchase of Preferred Stock", 
                "Common and/or Pref. Dividends Paid", "Special Dividend Paid", "Other Financing Activities", 
                "Foreign Exchange Rate Adj.", "Misc. Cash Flow Adj."],
    "Mnemonic": ["Net Income", "Depreciation & Amort.", "Amort. of Goodwill and Intangibles", "Other Amortization", 
                 "(Gain) Loss From Sale Of Assets", "(Gain) Loss On Sale Of Invest.", "Asset Writedown & Restructuring Costs", 
                 "Stock-Based Compensation", "Net Cash From Discontinued Ops.", "Other Operating Activities", 
                 "Change in Trad. Asset Securities", "Change in Acc. Receivable", "Change In Inventories", 
                 "Change in Acc. Payable", "Change in Unearned Rev.", "Change in Inc. Taxes", "Change in Def. Taxes", 
                 "Change in Other Net Operating Assets", "Capital Expenditure", "Sale of Property, Plant, and Equipment", 
                 "Cash Acquisitions", "Divestitures", "Other Investing Activities", "Short Term Debt Issued", 
                 "Long-Term Debt Issued", "Short Term Debt Repaid", "Long-Term Debt Repaid", "Issuance of Common Stock", 
                 "Repurchase of Common Stock", "Issuance of Pref. Stock", "Repurchase of Preferred Stock", 
                 "Common and/or Pref. Dividends Paid", "Special Dividend Paid", "Other Financing Activities", 
                 "Foreign Exchange Rate Adj.", "Misc. Cash Flow Adj."],
    "CIQ": ["IQ_NI_CF", "IQ_DA_SUPPL_CF", "IQ_GW_INTAN_AMORT_CF", "IQ_OTHER_AMORT", "IQ_GAIN_ASSETS_CF", 
            "IQ_GAIN_INVEST_CF", "IQ_ASSET_WRITEDOWN_CF", "IQ_STOCK_BASED_CF", "IQ_DO_CF", "IQ_OTHER_OPER_ACT", 
            "IQ_CHANGE_TRADING_ASSETS", "IQ_CHANGE_AR", "IQ_CHANGE_INVENTORY", "IQ_CHANGE_AP", "IQ_CHANGE_UNEARN_REV", 
            "IQ_CHANGE_INC_TAX", "IQ_CHANGE_DEF_TAX", "IQ_CHANGE_OTHER_NET_OPER_ASSETS", "IQ_CAPEX", 
            "IQ_SALE_PPE_CF", "IQ_CASH_ACQUIRE_CF", "IQ_DIVEST_CF", "IQ_OTHER_INVEST_ACT_SUPPL", "IQ_ST_DEBT_ISSUED", 
            "IQ_LT_DEBT_ISSUED", "IQ_ST_DEBT_REPAID", "IQ_LT_DEBT_REPAID", "IQ_COMMON_ISSUED", "IQ_COMMON_REP", 
            "IQ_PREF_ISSUED", "IQ_PREF_REP", "IQ_COMMON_PREF_DIV_CF", "IQ_SPECIAL_DIV_CF", "IQ_OTHER_FINANCE_ACT_SUPPL", 
            "IQ_FX", "IQ_MISC_ADJUST_CF"]
}

# Define the file paths for the data dictionaries
balance_sheet_data_dictionary_file = 'balance_sheet_data_dictionary.csv'
cash_flow_data_dictionary_file = 'cash_flow_data_dictionary.csv'
income_statement_data_dictionary_file = 'income_statement_data_dictionary.xlsx'

# Initialize lookup tables for Balance Sheet and Cash Flow
def load_or_initialize_lookup(file_path, initial_data):
    if os.path.exists(file_path):
        lookup_df = pd.read_csv(file_path)
    else:
        lookup_df = pd.DataFrame(initial_data)
        lookup_df.to_csv(file_path, index=False)
    return lookup_df

balance_sheet_lookup_df = load_or_initialize_lookup(balance_sheet_data_dictionary_file, initial_balance_sheet_lookup_data)
cash_flow_lookup_df = load_or_initialize_lookup(cash_flow_data_dictionary_file, initial_cash_flow_lookup_data)

# Function to save lookup tables as CSV
def save_lookup_table_csv(df, file_path):
    df.to_csv(file_path, index=False)

# Function to save lookup tables as Excel
def save_lookup_table_excel(df, file_path):
    df.to_excel(file_path, index=False)

# Utility functions
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

def create_combined_df(dfs):
    combined_df = pd.DataFrame()
    for i, df in enumerate(dfs):
        final_mnemonic_col = 'Final Mnemonic Selection'
        if final_mnemonic_col not in df.columns:
            st.error(f"Column '{final_mnemonic_col}' not found in dataframe {i+1}")
            st.write(df.columns.tolist())  # Output the columns for debugging
            continue
        
        # Identify date columns
        date_cols = [col for col in df.columns if col not in ['Label', 'Account', final_mnemonic_col, 'Mnemonic', 'Manual Selection']]
        if not date_cols:
            st.error(f"No date columns found in dataframe {i+1}")
            st.write(df.columns.tolist())  # Output the columns for debugging
            continue

        df_grouped = df.groupby([final_mnemonic_col, 'Label']).sum(numeric_only=True).reset_index()
        st.write(f"Grouped DataFrame for dataframe {i+1}:")
        st.write(df_grouped.head())  # Output grouped DataFrame for debugging

        # Print the columns of df_grouped for debugging
        st.write(f"Columns before melting dataframe {i+1}: {df_grouped.columns.tolist()}")

        # Verify that date columns exist in the grouped DataFrame
        missing_date_cols = [col for col in date_cols if col not in df_grouped.columns]
        if missing_date_cols:
            st.error(f"Missing date columns in dataframe {i+1}: {missing_date_cols}")
            st.write(df_grouped.columns.tolist())  # Output the columns for debugging
            continue

        try:
            df_melted = df_grouped.melt(id_vars=[final_mnemonic_col, 'Label'], value_vars=date_cols, var_name='Date', value_name='Value')
            st.write(f"Melted DataFrame for dataframe {i+1}:")
            st.write(df_melted.head())  # Output melted DataFrame for debugging
        except KeyError as e:
            st.error(f"Error melting dataframe {i+1}: {e}")
            st.write(df_grouped.columns.tolist())  # Output the columns for debugging
            continue
        
        df_pivot = df_melted.pivot(index=['Label', final_mnemonic_col], columns='Date', values='Value')
        st.write(f"Pivoted DataFrame for dataframe {i+1}:")
        st.write(df_pivot.head())  # Output pivoted DataFrame for debugging
        
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

def balance_sheet():
    global balance_sheet_lookup_df

    if 'balance_sheet_lookup_df' not in globals():
        if os.path.exists(balance_sheet_data_dictionary_file):
            balance_sheet_lookup_df = pd.read_csv(balance_sheet_data_dictionary_file)
        else:
            balance_sheet_lookup_df = pd.DataFrame()

    st.title("BALANCE SHEET LTMA")
    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Aggregate My Data", "Mappings and Data Consolidation", "Balance Sheet Data Dictionary"])

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

            st.subheader("Data Preview")
            st.dataframe(all_tables)

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

            st.subheader("Select numerical columns")
            numerical_columns = []
            for col in all_tables.columns:
                if st.checkbox(f"Numerical column '{col}'", value=False, key=f"num_{col}"):
                    numerical_columns.append(col)

            st.subheader("Convert Units")
            selected_columns = st.multiselect("Select columns for conversion", options=numerical_columns, key="columns_selection")
            selected_conversion_factor = st.radio("Select conversion factor", options=list(conversion_factors.keys()), key="conversion_factor")

            if st.button("Apply Selected Labels and Generate Excel", key="apply_selected_labels_generate_excel_tab1"):
                updated_table = all_tables.copy()

                for col in numerical_columns:
                    updated_table[col] = updated_table[col].apply(clean_numeric_value)
                
                if selected_conversion_factor and selected_conversion_factor in conversion_factors:
                    conversion_factor = conversion_factors[selected_conversion_factor]
                    updated_table = apply_unit_conversion(updated_table, selected_columns, conversion_factor)

                updated_table.replace('-', 0, inplace=True)

                excel_file = io.BytesIO()
                updated_table.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "extracted_combined_tables.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        st.subheader("Aggregate My Data")
        uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx'], accept_multiple_files=True, key='excel_uploader_amd')
        if uploaded_files:
            aggregated_df = aggregate_data_tab2(uploaded_files)
            if aggregated_df is not None:
                st.subheader("Aggregated Data Preview")
                editable_df = st.experimental_data_editor(aggregated_df, use_container_width=True)
                st.dataframe(editable_df)

                excel_file = io.BytesIO()
                editable_df.to_excel(excel_file, index=False)
                excel_file.seek(0)
                st.download_button("Download Excel", excel_file, "aggregated_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab3:
        st.subheader("Mappings and Data Consolidation")

        uploaded_excel = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader_tab3_bs')

        currency_options = ["U.S. Dollar", "Euro", "British Pound Sterling", "Japanese Yen"]
        magnitude_options = ["Actuals", "Thousands", "Millions", "Billions", "Trillions"]

        selected_currency = st.selectbox("Select Currency", currency_options, key='currency_selection_tab3_bs')
        selected_magnitude = st.selectbox("Select Magnitude", magnitude_options, key='magnitude_selection_tab3_bs')
        company_name = st.text_input("Enter Company Name", key='company_name_input_bs')

        if uploaded_excel is not None:
            df = pd.read_excel(uploaded_excel)
            st.write("Columns in the uploaded file:", df.columns.tolist())

            if 'Account' not in df.columns or 'Label' not in df.columns:
                st.error("The uploaded file does not contain the required 'Account' or 'Label' columns.")
            else:
                if 'Sort Index' not in df.columns:
                    df['Sort Index'] = range(1, len(df) + 1)

                # Function to get the best match based on Label and Account using Levenshtein distance
                def get_best_match(label, account):
                    best_score = float('inf')
                    best_match = None
                    for _, lookup_row in balance_sheet_lookup_df.iterrows():
                        if lookup_row['Label'].strip().lower() == str(label).strip().lower():
                            lookup_account = lookup_row['Account']
                            account_str = str(account)
                            score = levenshtein_distance(account_str.lower(), lookup_account.lower()) / max(len(account_str), len(lookup_account))
                            if score < best_score:
                                best_score = score
                                best_match = lookup_row
                    return best_match, best_score

                df['Mnemonic'] = ''
                df['Manual Selection'] = ''
                for idx, row in df.iterrows():
                    account_value = row['Account']
                    label_value = row['Label']
                    if pd.notna(account_value) and pd.notna(label_value):
                        best_match, score = get_best_match(label_value, account_value)
                        if best_match is not None and score < 0.25:
                            df.at[idx, 'Mnemonic'] = best_match['Mnemonic']
                        else:
                            df.at[idx, 'Mnemonic'] = 'Human Intervention Required'
                    
                    if df.at[idx, 'Mnemonic'] == 'Human Intervention Required':
                        message = f"**Human Intervention Required for:** {account_value} [{label_value} - Index {idx}]"
                        st.markdown(message)
                    
                    manual_selection = st.selectbox(
                        f"Select category for '{account_value}'",
                        options=[''] + balance_sheet_lookup_df['Mnemonic'].tolist() + ['REMOVE ROW'],
                        key=f"select_{idx}_tab3_bs"
                    )
                    if manual_selection:
                        df.at[idx, 'Manual Selection'] = manual_selection.strip()

                st.dataframe(df[['Label', 'Account', 'Mnemonic', 'Manual Selection', 'Sort Index']])  # Include 'Sort Index' as a helper column

                if st.button("Generate Excel with Lookup Results", key="generate_excel_lookup_results_tab3_bs"):
                    df['Final Mnemonic Selection'] = df.apply(
                        lambda row: row['Manual Selection'].strip() if row['Manual Selection'].strip() != '' else row['Mnemonic'], 
                        axis=1
                    )
                    final_output_df = df[df['Final Mnemonic Selection'].str.strip() != 'REMOVE ROW'].copy()
                    
                    combined_df = create_combined_df([final_output_df])
                    combined_df = combined_df.sort_values(by=['Final Mnemonic Selection'])

                    # Add CIQ column based on lookup
                    def lookup_ciq(mnemonic):
                        if mnemonic == 'Human Intervention Required':
                            return 'CIQ IQ Required'
                        ciq_value = balance_sheet_lookup_df.loc[balance_sheet_lookup_df['Mnemonic'] == mnemonic, 'CIQ']
                        if ciq_value.empty:
                            return 'CIQ IQ Required'
                        return ciq_value.values[0]
                    
                    combined_df['CIQ'] = combined_df['Final Mnemonic Selection'].apply(lookup_ciq)

                    columns_order = ['Final Mnemonic Selection', 'CIQ'] + [col for col in combined_df.columns if col not in ['Final Mnemonic Selection', 'CIQ']]
                    combined_df = combined_df[columns_order]

                    # Include the "As Presented" sheet without the CIQ column, and with the specified column order
                    as_presented_df = final_output_df.drop(columns=['CIQ', 'Mnemonic', 'Manual Selection'], errors='ignore')
                    as_presented_df = as_presented_df.sort_values(by=['Sort Index'])
                    as_presented_df = as_presented_df.drop(columns=['Sort Index'], errors='ignore')
                    as_presented_columns_order = ['Label', 'Account', 'Final Mnemonic Selection'] + [col for col in as_presented_df.columns if col not in ['Label', 'Account', 'Final Mnemonic Selection']]
                    as_presented_df = as_presented_df[as_presented_columns_order]

                    excel_file = io.BytesIO()
                    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                        combined_df.to_excel(writer, sheet_name='Standardized', index=False)
                        as_presented_df.to_excel(writer, sheet_name='As Presented', index=False)
                        cover_df = pd.DataFrame({
                            'Selection': ['Currency', 'Magnitude', 'Company Name'],
                            'Value': [selected_currency, selected_magnitude, company_name]
                        })
                        cover_df.to_excel(writer, sheet_name='Cover', index=False)
                    excel_file.seek(0)
                    st.download_button("Download Excel", excel_file, "mnemonic_mapping_with_aggregation_bs.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
                                new_entries.append({'Label': row['Label'], 'Account': row['Account'], 'Mnemonic': final_mnemonic, 'CIQ': ciq_value})
                            else:
                                balance_sheet_lookup_df.loc[(balance_sheet_lookup_df['Account'] == row['Account']) & (balance_sheet_lookup_df['Label'] == row['Label']), 'Mnemonic'] = final_mnemonic
                                balance_sheet_lookup_df.loc[(balance_sheet_lookup_df['Account'] == row['Account']) & (balance_sheet_lookup_df['Label'] == row['Label']), 'CIQ'] = ciq_value
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
            csv_file = io.BytesIO()
            balance_sheet_lookup_df.to_csv(csv_file, index=False)
            csv_file.seek(0)
            st.download_button("Download CSV", csv_file, "balance_sheet_data_dictionary.csv", "text/csv")




            
#######################################################Cash Flow Statement########################################################
def cash_flow_statement():
    global cash_flow_lookup_df

    st.title("CASH FLOW STATEMENT LTMA")
    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Aggregate My Data", "Mappings and Data Consolidation", "Cash Flow Data Dictionary"])

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
            selected_value = st.radio("Select conversion value", ["Actuals", "Thousands", "Millions", "Billions"], index=0, key="conversion_value_cfs")

            conversion_factors = {
                "Actuals": 1,
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
                
                if selected_value != "Actuals":
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
            dfs = [pd.read_excel(file) for file in uploaded_files]

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
        st.subheader("Mappings and Data Consolidation")

        uploaded_excel = st.file_uploader("Upload your Excel file for Mnemonic Mapping", type=['xlsx'], key='excel_uploader_tab3_cfs')

        currency_options = ["U.S. Dollar", "Euro", "British Pound Sterling", "Japanese Yen"]
        magnitude_options = ["Actuals", "Thousands", "Millions", "Billions", "Trillions"]

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
                    save_lookup_table_csv(cash_flow_lookup_df, cash_flow_data_dictionary_file)
                    st.success("Data Dictionary Updated Successfully")

    with tab4:
        st.subheader("Cash Flow Data Dictionary")

        uploaded_dict_file = st.file_uploader("Upload a new Data Dictionary CSV", type=['csv'], key='dict_uploader_tab4_cfs')
        if uploaded_dict_file is not None:
            new_lookup_df = pd.read_csv(uploaded_dict_file)
            cash_flow_lookup_df = new_lookup_df  # Overwrite the entire DataFrame
            save_lookup_table_csv(cash_flow_lookup_df, cash_flow_data_dictionary_file)
            st.success("Data Dictionary uploaded and updated successfully!")

        st.dataframe(cash_flow_lookup_df)

        remove_indices = st.multiselect("Select rows to remove", cash_flow_lookup_df.index, key='remove_indices_tab4_cfs')
        if st.button("Remove Selected Rows", key="remove_selected_rows_tab4_cfs"):
            cash_flow_lookup_df = cash_flow_lookup_df.drop(remove_indices).reset_index(drop=True)
            save_lookup_table_csv(cash_flow_lookup_df, cash_flow_data_dictionary_file)
            st.success("Selected rows removed successfully!")
            st.dataframe(cash_flow_lookup_df)

        if st.button("Download Data Dictionary", key="download_data_dictionary_tab4_cfs"):
            csv_file = io.BytesIO()
            cash_flow_lookup_df.to_csv(csv_file, index=False)
            csv_file.seek(0)
            st.download_button("Download CSV", csv_file, "cash_flow_data_dictionary.csv", "text/csv")



######################################INCOME STATEMENT##################################

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
    tab1, tab2, tab3, tab4 = st.tabs(["Table Extractor", "Aggregate My Data", "Mappings and Data Consolidation", "Income Statement Data Dictionary"])

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
        st.subheader("Mappings and Data Consolidation")

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

####################################### Populate CIQ Template ###################################
import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import json

def copy_sheet(source_book, target_book, sheet_name, tab_color="00FF00"):
    source_sheet = source_book[sheet_name]
    target_sheet = target_book.create_sheet(title=sheet_name)

    for row in source_sheet.iter_rows():
        for cell in row:
            target_cell = target_sheet[cell.coordinate]
            target_cell.value = cell.value

            # Copy styles manually
            if cell.has_style:
                target_cell.font = cell.font.copy()
                target_cell.border = cell.border.copy()
                target_cell.fill = cell.fill.copy()
                target_cell.number_format = cell.number_format
                target_cell.protection = cell.protection.copy()
                target_cell.alignment = cell.alignment.copy()

            # Copy hyperlinks and comments
            if cell.hyperlink:
                target_cell.hyperlink = cell.hyperlink
            if cell.comment:
                target_cell.comment = cell.comment

    target_sheet.sheet_properties.tabColor = tab_color

    # Copy merged cells
    for merged_cell in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_cell))

def get_cell_value_as_string(sheet, cell_address):
    cell_value = sheet[cell_address].value
    return str(cell_value) if cell_value is not None else ""

def populate_ciq_template():
    st.title("Populate CIQ Template")

    tab1 = st.tabs(["Final Output"])[0]

    with tab1:
        uploaded_template = st.file_uploader("Upload Template", type=['xlsx', 'xlsm'], key='template_uploader')
        uploaded_income_statement = st.file_uploader("Upload Completed Income Statement", type=['xlsx', 'xlsm'], key='income_statement_uploader')
        uploaded_balance_sheet = st.file_uploader("Upload Completed Balance Sheet", type=['xlsx', 'xlsm'], key='balance_sheet_uploader')
        uploaded_cash_flow_statement = st.file_uploader("Upload Completed Cash Flow Statement", type=['xlsx', 'xlsm'], key='cash_flow_statement_uploader')

        if uploaded_template and (uploaded_income_statement or uploaded_balance_sheet or uploaded_cash_flow_statement):
            try:
                file_extension = uploaded_template.name.split('.')[-1]
                template_book = load_workbook(uploaded_template, data_only=False, keep_vba=True if file_extension == 'xlsm' else False)
                if uploaded_income_statement:
                    income_statement_book = load_workbook(uploaded_income_statement, data_only=False)
                    income_statement_df = pd.read_excel(uploaded_income_statement, sheet_name="Standardized")
                    template_income_statement_df = pd.read_excel(uploaded_template, sheet_name="Income Statement")
                if uploaded_balance_sheet:
                    balance_sheet_book = load_workbook(uploaded_balance_sheet, data_only=False)
                    balance_sheet_df = pd.read_excel(uploaded_balance_sheet, sheet_name="Standardized")
                    template_balance_sheet_df = pd.read_excel(uploaded_template, sheet_name="Balance Sheet")
                if uploaded_cash_flow_statement:
                    cash_flow_statement_book = load_workbook(uploaded_cash_flow_statement, data_only=False)
                    cash_flow_statement_df = pd.read_excel(uploaded_cash_flow_statement, sheet_name="Standardized")
                    template_cash_flow_statement_df = pd.read_excel(uploaded_template, sheet_name="Cash Flow")
            except Exception as e:
                st.error(f"Error reading files: {e}")
                return

            errors = []

            if st.button("Populate Template"):
                if uploaded_income_statement:
                    try:
                        ciq_mnemonics_income = income_statement_df.iloc[:, 1]
                        income_statement_dates = income_statement_df.columns[2:]
                    except Exception as e:
                        st.error(f"Error processing income statement data: {e}")
                        return

                    st.write("Income Statement Dates:", list(income_statement_dates))

                    try:
                        template_mnemonics_income = template_income_statement_df.iloc[11:89, 8]
                    except Exception as e:
                        st.error(f"Error processing template data: {e}")
                        return

                    try:
                        template_income_sheet = template_book["Income Statement"]
                        template_income_dates = [
                            get_cell_value_as_string(template_income_sheet, "D10"),
                            get_cell_value_as_string(template_income_sheet, "E10"),
                            get_cell_value_as_string(template_income_sheet, "F10"),
                            get_cell_value_as_string(template_income_sheet, "G10")
                        ]
                    except Exception as e:
                        st.error(f"Error reading dates from template income statement: {e}")
                        return

                    st.write("Template Income Statement Dates:", template_income_dates)

                    for i, mnemonic in enumerate(template_mnemonics_income):
                        if pd.notna(mnemonic):
                            try:
                                income_statement_row = income_statement_df[ciq_mnemonics_income == mnemonic]
                                if not income_statement_row.empty:
                                    for j, date in enumerate(template_income_dates):
                                        if date in income_statement_dates.values:
                                            try:
                                                income_statement_col = income_statement_dates.get_loc(date)
                                                st.write(f"Populating template for mnemonic {mnemonic} at row {i + 12}, column {3 + j} with value from income statement column {income_statement_col + 2}")
                                                if 3 + j not in [10, 11, 12, 13]:  # Columns J, K, L, M are 10, 11, 12, 13
                                                    template_income_statement_df.iat[i + 11, 3 + j] = income_statement_row.iat[0, income_statement_col + 2]
                                            except Exception as e:
                                                errors.append(f"Error at mnemonic {mnemonic}, row {i + 12}, column {3 + j}: {e}")
                            except Exception as e:
                                errors.append(f"Error processing row for mnemonic {mnemonic}: {e}")

                    try:
                        for r_idx, row in enumerate(dataframe_to_rows(template_income_statement_df, index=False, header=True), 1):
                            if r_idx >= 12 and r_idx <= 90:
                                for c_idx, value in enumerate(row, 1):
                                    if c_idx >= 4 and c_idx <= 7:
                                        cell = template_income_sheet.cell(row=r_idx, column=c_idx)
                                        if cell.column not in [10, 11, 12, 13]:  # Skip columns J, K, L, M
                                            if not (cell.value and isinstance(cell.value, str) and cell.value.startswith('=')):
                                                for merge_cell in template_income_sheet.merged_cells.ranges:
                                                    if cell.coordinate in merge_cell:
                                                        template_income_sheet.unmerge_cells(str(merge_cell))
                                                        break
                                                cell.value = value
                    except Exception as e:
                        st.error(f"Error updating template income sheet at cell {cell.coordinate}: {e}")
                        return

                    try:
                        copy_sheet(income_statement_book, template_book, "As Presented - Income Stmt")
                    except Exception as e:
                        st.error(f"Error copying 'As Presented - Income Stmt' sheet: {e}")
                        return

                if uploaded_balance_sheet:
                    try:
                        ciq_mnemonics_balance = balance_sheet_df.iloc[:, 1]
                        balance_sheet_dates = balance_sheet_df.columns[2:]
                    except Exception as e:
                        st.error(f"Error processing balance sheet data: {e}")
                        return

                    st.write("Balance Sheet Dates:", list(balance_sheet_dates))

                    try:
                        template_mnemonics_balance = template_balance_sheet_df.iloc[11:89, 8]
                    except Exception as e:
                        st.error(f"Error processing template data: {e}")
                        return

                    try:
                        template_balance_sheet = template_book["Balance Sheet"]
                        template_balance_dates = [
                            get_cell_value_as_string(template_balance_sheet, "D10"),
                            get_cell_value_as_string(template_balance_sheet, "E10"),
                            get_cell_value_as_string(template_balance_sheet, "F10"),
                            get_cell_value_as_string(template_balance_sheet, "G10")
                        ]
                    except Exception as e:
                        st.error(f"Error reading dates from template balance sheet: {e}")
                        return

                    st.write("Template Balance Sheet Dates:", template_balance_dates)

                    for i, mnemonic in enumerate(template_mnemonics_balance):
                        if pd.notna(mnemonic):
                            try:
                                balance_sheet_row = balance_sheet_df[ciq_mnemonics_balance == mnemonic]
                                if not balance_sheet_row.empty:
                                    for j, date in enumerate(template_balance_dates):
                                        if date in balance_sheet_dates.values:
                                            try:
                                                balance_sheet_col = balance_sheet_dates.get_loc(date)
                                                st.write(f"Populating template for mnemonic {mnemonic} at row {i + 12}, column {3 + j} with value from balance sheet column {balance_sheet_col + 2}")
                                                if 3 + j not in [10, 11, 12, 13]:  # Columns J, K, L, M are 10, 11, 12, 13
                                                    template_balance_sheet_df.iat[i + 11, 3 + j] = balance_sheet_row.iat[0, balance_sheet_col + 2]
                                            except Exception as e:
                                                errors.append(f"Error at mnemonic {mnemonic}, row {i + 12}, column {3 + j}: {e}")
                            except Exception as e:
                                errors.append(f"Error processing row for mnemonic {mnemonic}: {e}")

                    try:
                        for r_idx, row in enumerate(dataframe_to_rows(template_balance_sheet_df, index=False, header=True), 1):
                            if r_idx >= 12 and r_idx <= 90:
                                for c_idx, value in enumerate(row, 1):
                                    if c_idx >= 4 and c_idx <= 7:
                                        cell = template_balance_sheet.cell(row=r_idx, column=c_idx)
                                        if cell.column not in [10, 11, 12, 13]:  # Skip columns J, K, L, M
                                            if not (cell.value and isinstance(cell.value, str) and cell.value.startswith('=')):
                                                for merge_cell in template_balance_sheet.merged_cells.ranges:
                                                    if cell.coordinate in merge_cell:
                                                        template_balance_sheet.unmerge_cells(str(merge_cell))
                                                        break
                                                cell.value = value
                    except Exception as e:
                        st.error(f"Error updating template balance sheet at cell {cell.coordinate}: {e}")
                        return

                    try:
                        copy_sheet(balance_sheet_book, template_book, "As Presented - Balance Sheet")
                    except Exception as e:
                        st.error(f"Error copying 'As Presented - Balance Sheet' sheet: {e}")
                        return

                if uploaded_cash_flow_statement:
                    try:
                        ciq_mnemonics_cash_flow = cash_flow_statement_df.iloc[:, 1]
                        cash_flow_statement_dates = cash_flow_statement_df.columns[2:]
                    except Exception as e:
                        st.error(f"Error processing cash flow statement data: {e}")
                        return

                    st.write("Cash Flow Statement Dates:", list(cash_flow_statement_dates))

                    try:
                        template_mnemonics_cash_flow = template_cash_flow_statement_df.iloc[11:89, 8]
                    except Exception as e:
                        st.error(f"Error processing template data: {e}")
                        return

                    try:
                        template_cash_flow_sheet = template_book["Cash Flow"]
                        template_cash_flow_dates = [
                            get_cell_value_as_string(template_cash_flow_sheet, "D10"),
                            get_cell_value_as_string(template_cash_flow_sheet, "E10"),
                            get_cell_value_as_string(template_cash_flow_sheet, "F10"),
                            get_cell_value_as_string(template_cash_flow_sheet, "G10")
                        ]
                    except Exception as e:
                        st.error(f"Error reading dates from template cash flow statement: {e}")
                        return

                    st.write("Template Cash Flow Statement Dates:", template_cash_flow_dates)

                    for i, mnemonic in enumerate(template_mnemonics_cash_flow):
                        if pd.notna(mnemonic):
                            try:
                                cash_flow_statement_row = cash_flow_statement_df[ciq_mnemonics_cash_flow == mnemonic]
                                if not cash_flow_statement_row.empty:
                                    for j, date in enumerate(template_cash_flow_dates):
                                        if date in cash_flow_statement_dates.values:
                                            try:
                                                cash_flow_statement_col = cash_flow_statement_dates.get_loc(date)
                                                st.write(f"Populating template for mnemonic {mnemonic} at row {i + 12}, column {3 + j} with value from cash flow statement column {cash_flow_statement_col + 2}")
                                                if 3 + j not in [10, 11, 12, 13]:  # Columns J, K, L, M are 10, 11, 12, 13
                                                    template_cash_flow_statement_df.iat[i + 11, 3 + j] = cash_flow_statement_row.iat[0, cash_flow_statement_col + 2]
                                            except Exception as e:
                                                errors.append(f"Error at mnemonic {mnemonic}, row {i + 12}, column {3 + j}: {e}")
                            except Exception as e:
                                errors.append(f"Error processing row for mnemonic {mnemonic}: {e}")

                    try:
                        for r_idx, row in enumerate(dataframe_to_rows(template_cash_flow_statement_df, index=False, header=True), 1):
                            if r_idx >= 12 and r_idx <= 90:
                                for c_idx, value in enumerate(row, 1):
                                    if c_idx >= 4 and c_idx <= 7:
                                        cell = template_cash_flow_sheet.cell(row=r_idx, column=c_idx)
                                        if cell.column not in [10, 11, 12, 13]:  # Skip columns J, K, L, M
                                            if not (cell.value and isinstance(cell.value, str) and cell.value.startswith('=')):
                                                for merge_cell in template_cash_flow_sheet.merged_cells.ranges:
                                                    if cell.coordinate in merge_cell:
                                                        template_cash_flow_sheet.unmerge_cells(str(merge_cell))
                                                        break
                                                cell.value = value
                    except Exception as e:
                        st.error(f"Error updating template cash flow sheet at cell {cell.coordinate}: {e}")
                        return

                    try:
                        copy_sheet(cash_flow_statement_book, template_book, "As Presented - Cash Flow")
                    except Exception as e:
                        st.error(f"Error copying 'As Presented - Cash Flow' sheet: {e}")
                        return

                try:
                    output_file_name = f"populated_template.{file_extension}"
                    excel_file = io.BytesIO()
                    template_book.save(excel_file)
                    excel_file.seek(0)
                except Exception as e:
                    st.error(f"Error saving the populated template: {e}")
                    return

                if errors:
                    st.error("Errors encountered during processing:")
                    for error in errors:
                        st.error(error)

                mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if file_extension == 'xlsm' else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                st.download_button(
                    label="Download Populated Template",
                    data=excel_file,
                    file_name=output_file_name,
                    mime=mime_type
                )

import json
import pandas as pd
import streamlit as st
from io import BytesIO

def json_conversion():
    st.title("JSON Conversion")

    uploaded_file = st.file_uploader("Choose a JSON file", type="json", key='json_uploader')
    if uploaded_file is not None:
        try:
            # Read the uploaded file as a string
            file_contents = uploaded_file.read().decode('utf-8')

            # Load the JSON data
            data = json.loads(file_contents)
            
            # Display file size for debugging
            st.text(f"File size: {len(file_contents)} bytes")

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

            all_tables.insert(0, 'Label', '')

            st.subheader("Data Preview")
            st.dataframe(all_tables)

            # Button to export data to Excel
            def to_excel(df):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                writer.close()
                processed_data = output.getvalue()
                return processed_data

            excel_data = to_excel(all_tables)

            st.download_button(label='📥 Download Excel file',
                               data=excel_data,
                               file_name='converted_data.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except json.JSONDecodeError:
            st.error("The uploaded file is not a valid JSON.")
        except Exception as e:
            st.error(f"An error occurred: {e}")

def main():
    st.sidebar.title("Navigation")
    selection = st.sidebar.radio("Go to", ["Balance Sheet", "Cash Flow Statement", "Income Statement", "Populate CIQ Template", "Extras"])

    if selection == "Balance Sheet":
        balance_sheet()
    elif selection == "Cash Flow Statement":
        cash_flow_statement()
    elif selection == "Income Statement":
        income_statement()
    elif selection == "Populate CIQ Template":
        populate_ciq_template()
    elif selection == "Extras":
        extras_tab()

def extras_tab():
    st.sidebar.title("Extras")
    extra_selection = st.sidebar.radio("Select Extra Function", ["JSON Conversion"])

    if extra_selection == "JSON Conversion":
        json_conversion()

if __name__ == '__main__':
    main()

