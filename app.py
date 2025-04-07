#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import io
import json
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from Levenshtein import distance as levenshtein_distance
import re
import anthropic
import numpy as np
from concurrent.futures import ThreadPoolExecutor
import logging
from sklearn.feature_extraction.text import TfidfVectorizer
import time
from word2number import w2n
import os
import boto3
import tempfile
import botocore
from io import BytesIO

# Set up logging
logging.basicConfig(level=logging.INFO)

# Global vectorizer
tfidf_vectorizer = None

# Load model function with TF-IDF only (no SentenceTransformer)
@st.cache_resource
def load_model(max_retries=3, base_wait=5):
    global tfidf_vectorizer
    # Initialize fallback vectorizer
    tfidf_vectorizer = TfidfVectorizer(stop_words='english')
    
    # Log that we're using TF-IDF for text similarity
    logging.info("Using TF-IDF for text similarity")
    return None

# Get embedding function using TF-IDF
@st.cache_data
def get_embedding(text):
    global tfidf_vectorizer
    
    # Ensure we have some text to process
    if not text or not isinstance(text, str):
        text = str(text) if text is not None else ""
    
    # Fit on financial terms + the input text to create a relevant vocabulary
    corpus = [
        text, "revenue", "assets", "liabilities", "cash flow", 
        "income", "net profit", "expenses", "equity", "tax", 
        "depreciation", "amortization", "current assets",
        "non-current assets", "operating activities"
    ]
    
    # Fit and transform
    tfidf_vectorizer.fit(corpus)
    vector = tfidf_vectorizer.transform([text]).toarray()[0]
    return vector

# Cosine similarity function that handles TF-IDF vectors
def cosine_similarity(a, b):
    # Convert inputs to numpy arrays
    a_array = np.array(a)
    b_array = np.array(b)
    
    # Guard against zero vectors
    norm_a = np.linalg.norm(a_array)
    norm_b = np.linalg.norm(b_array)
    
    if norm_a == 0 or norm_b == 0:
        return 0
    
    return np.dot(a_array, b_array) / (norm_a * norm_b)

# Load the model
model = load_model()

# Set up the Anthropic client with error handling
@st.cache_resource
def setup_anthropic_client():
    try:
        return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    except KeyError:
        st.error("Anthropic API key not found in secrets. Please check your configuration.")
        st.stop()

client = setup_anthropic_client()

# Function to generate a response from Claude
@st.cache_data
def generate_response(prompt, max_tokens=10000):
    try:
        logging.info("Generating response from Claude...")
        response = client.messages.create(
            model="claude-3-sonnet-20240229",
            max_tokens=max_tokens,
            temperature=0.2,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        return "I'm sorry, but I encountered an error while processing your request."
    
def get_ai_suggested_mapping_BS(label, account, balance_sheet_lookup_df, nearby_rows):
    logging.info("Computing AI suggested mapping...")
    prompt = f"""Given the following account information:
    Label: {label}
    Account: {account}

    Nearby rows:
    {nearby_rows}

    And the following balance sheet lookup data:
    {balance_sheet_lookup_df.to_string()}

    What is the most appropriate Mnemonic mapping for this account based on Label and Account combination?"""

    suggested_mnemonic = generate_response(prompt).strip()

    account_embedding = get_embedding(f"{label} {account}")
    similarities = balance_sheet_lookup_df.apply(lambda row: cosine_similarity(account_embedding, get_embedding(f"{row['Label']} {row['Account']}")), axis=1)

    top_3_similar = similarities.nlargest(3)
    scores = {}
    for idx in top_3_similar.index:
        row = balance_sheet_lookup_df.loc[idx]
        score = 0
        if row['Label'].lower() == label.lower():
            score += 2
        if row['Mnemonic'] == suggested_mnemonic:
            score += 3
        score += top_3_similar[idx] * 5
        scores[row['Mnemonic']] = score

    best_mnemonic = max(scores, key=scores.get)
    logging.info(f"Suggested mnemonic: {best_mnemonic}")
    return best_mnemonic

# Define the initial lookup data for Balance Sheet
initial_balance_sheet_lookup_data = {
    "Label": ["Current Assets", "Non Current Assets", "Non Current Assets", "Non Current Assets", "Non Current Assets", "Non Current Assets", "Non Current Assets"],
    "Account": ["Gross Property, Plant & Equipment", "Accumulated Depreciation", "Net Property, Plant & Equipment", "Long-term Investments", "Goodwill", "Other Intangibles", "Right-of-Use Asset-Net"],
    "Mnemonic": ["Gross Property, Plant & Equipment", "Accumulated Depreciation", "Net Property, Plant & Equipment", "Long-term Investments", "Goodwill", "Other Intangibles", "Right-of-Use Asset-Net"],
    "CIQ": ["IQ_GPPE", "IQ_AD", "IQ_NPPE", "IQ_LT_INVEST", "IQ_GW", "IQ_OTHER_INTAN", "IQ_RUA_NET"]
}

# Define the file path for the Balance Sheet data dictionary
balance_sheet_data_dictionary_file = 'balance_sheet_data_dictionary.xlsx'

# Load or initialize the lookup table
@st.cache_data(ttl=600)
def load_balance_sheet_data():
    logging.info("Loading or initializing balance sheet data...")
    try:
        return pd.read_excel(balance_sheet_data_dictionary_file)
    except FileNotFoundError:
        return pd.DataFrame(initial_balance_sheet_lookup_data)

def save_and_update_balance_sheet_data(df):
    logging.info("Saving and updating balance sheet data...")
    st.session_state.balance_sheet_data = df
    save_lookup_table(df, balance_sheet_data_dictionary_file)

def save_lookup_table(df, file_path):
    logging.info("Writing data to Excel...")
    df.to_excel(file_path, index=False)
    
# General Utility Functions
def process_file(file):
    logging.info(f"Processing file: {file.name}")
    try:
        df = pd.read_excel(file, sheet_name=None)
        first_sheet_name = list(df.keys())[0]
        df = df[first_sheet_name]
        return df
    except Exception as e:
        st.error(f"Error processing file {file.name}: {e}")
        return None

def create_combined_df(dfs):
    logging.info("Combining data frames...")
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
    logging.info("Aggregating data...")
    if 'Label' not in df.columns or 'Account' not in df.columns:
        st.error("'Label' and/or 'Account' columns not found in the data.")
        return df

    pivot_table = df.pivot_table(index=['Label', 'Account'], 
                                 values=[col for col in df.columns if col not in ['Label', 'Account', 'Mnemonic', 'Manual Selection']], 
                                 aggfunc='sum').reset_index()
    return pivot_table

def clean_numeric_value(value):
    logging.info(f"Cleaning numeric value: {value}")
    try:
        value_str = str(value).strip()
        
        # Remove any number of spaces between $ and (
        value_str = re.sub(r'\$\s*\(', '$(', value_str)
        
        # Handle negative values in parentheses with or without dollar sign
        if (value_str.startswith('$(') and value_str.endswith(')')) or (value_str.startswith('(') and value_str.endswith(')')):
            value_str = '-' + value_str.lstrip('$(').rstrip(')')
        
        # Remove dollar signs and commas
        cleaned_value = re.sub(r'[$,]', '', value_str)
        
        # Convert text to number
        try:
            cleaned_value = w2n.word_to_num(cleaned_value)
        except ValueError:
            pass
        
        return float(cleaned_value)
    except (ValueError, TypeError) as e:
        logging.error(f"Error converting value: {value} with error: {e}")
        return 0

def sort_by_label_and_account(df):
    logging.info("Sorting by label and account...")
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
    logging.info("Sorting by label and final mnemonic selection...")
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
    logging.info("Applying unit conversion...")
    for selected_column in columns:
        if selected_column in df.columns:
            df[selected_column] = df[selected_column].apply(
                lambda x: x * factor if isinstance(x, (int, float)) else x)
    return df

def check_all_zeroes(df):
    logging.info("Checking for all zeroes in data frame...")
    zeroes = (df.iloc[:, 2:] == 0).all(axis=1)
    return zeroes

def get_ai_suggested_mapping_CF(label, account, cash_flow_lookup_df, nearby_rows):
    prompt = f"""Given the following account information:
    Label: {label}
    Account: {account}

    Nearby rows:
    {nearby_rows}

    And the following cash flow lookup data:
    {cash_flow_lookup_df.to_string()}

    What is the most appropriate Mnemonic mapping for this account based on Label and Account combination? Please consider the following:
    1. The account's position in the cash flow structure (e.g., Operating Activities, Investing Activities, Financing Activities)
    2. The semantic meaning of the account name and its relationship to standard financial statement line items
    3. The nearby rows to understand the context of this account
    4. Common financial reporting standards and practices

    Please provide only the value from the 'Mnemonic' column in the Cash Flow Data Dictionary data frame based on Label and Account combination, without any explanation. Ensure that the suggested Mnemonic is appropriate for the given Label."""

    suggested_mnemonic = generate_response(prompt).strip()

    account_embedding = get_embedding(f"{label} {account}")
    similarities = cash_flow_lookup_df.apply(lambda row: cosine_similarity(account_embedding, get_embedding(f"{row['Label']} {row['Account']}")), axis=1)

    top_3_similar = similarities.nlargest(3)
    scores = {}
    for idx in top_3_similar.index:
        row = cash_flow_lookup_df.loc[idx]
        score = 0
        if row['Label'].lower() == label.lower():
            score += 2
        if row['Mnemonic'] == suggested_mnemonic:
            score += 3
        score += top_3_similar[idx] * 5
        scores[row['Mnemonic']] = score

    if "total" in account.lower() and not any("total" in mnemonic.lower() for mnemonic in scores):
        total_mnemonics = cash_flow_lookup_df[cash_flow_lookup_df['Mnemonic'].str.contains('Total', case=False)]['Mnemonic']
        if not total_mnemonics.empty:
            scores[total_mnemonics.iloc[0]] = max(scores.values()) + 1

    best_mnemonic = max(scores, key=scores.get)
    return best_mnemonic

def get_ai_suggested_mapping_IS(account, income_statement_lookup_df, nearby_rows):
    prompt = f"""Given the following account information:
    Account: {account}

    Nearby rows:
    {nearby_rows}

    And the following income statement lookup data:
    {income_statement_lookup_df.to_string()}

    What is the most appropriate Mnemonic mapping for this account based on the Account name and context provided by the nearby rows? Please consider the following:
    1. The semantic meaning of the account name and its relationship to standard financial statement line items
    2. The nearby rows to understand the context of this account
    3. Common financial reporting standards and practices for income statements

    Please provide only the value from the 'Mnemonic' column in the Income Statement Data Dictionary data frame, without any explanation."""

    suggested_mnemonic = generate_response(prompt).strip()

    # Calculate embedding similarities
    account_embedding = get_embedding(account)
    similarities = income_statement_lookup_df.apply(lambda row: cosine_similarity(account_embedding, get_embedding(row['Account'])), axis=1)

    # Get top 3 most similar entries
    top_three_similar = similarities.nlargest(3)

    # Scoring system
    scores = {}
    for idx in top_three_similar.index:
        row = income_statement_lookup_df.loc[idx]
        score = 0
        if row['Mnemonic'] == suggested_mnemonic:
            score += 3
        score += top_three_similar[idx] * 5  # Weight similarity score
        scores[row['Mnemonic']] = score

    best_mnemonic = max(scores, key=scores.get)
    return best_mnemonic

def clean_numeric_value_IS(value):
    try:
        value_str = str(value).strip()
        
        # Remove any number of spaces between $ and (
        value_str = re.sub(r'\$\s*\(', '$(', value_str)
        
        # Handle negative values in parentheses with or without dollar sign
        if (value_str.startswith('$(') and value_str.endswith(')')) or (value_str.startswith('(') and value_str.endswith(')')):
            value_str = '-' + value_str.lstrip('$(').rstrip(')')
        
        # Remove dollar signs and commas
        cleaned_value = re.sub(r'[$,]', '', value_str)
        
        # Convert text to number
        try:
            cleaned_value = w2n.word_to_num(cleaned_value)
        except ValueError:
            pass
        
        return float(cleaned_value)
    except (ValueError, TypeError):
        return value

def apply_unit_conversion_IS(df, columns, factor):
    for selected_column in columns:
        if (selected_column in df.columns) and (not df[selected_column].isnull().all()):
            df[selected_column] = df[selected_column].apply(
                lambda x: x * factor if isinstance(x, (int, float)) else x)
    return df

def sort_by_sort_index(df):
    if 'Sort Index' in df.columns:
        df = df.sort_values(by=['Sort Index'])
    return df

def update_negative_values(df):
    criteria = [
        "IQ_COGS",
        "IQ_SGA_SUPPL",
        "IQ_RD_EXP",
        "IQ_DA_SUPPL",
        "IQ_STOCK_BASED",
        "IQ_OTHER_OPER",
        "IQ_INC_TAX"
    ]
    
    for index, row in df.iterrows():
        if row['CIQ'] in criteria:
            for col in df.columns[2:]:  # Start from column index 2 (skip 'Final Mnemonic Selection' and 'CIQ')
                if isinstance(row[col], (int, float)) and row[col] < 0:
                    df.at[index, col] = row[col] * 1
    return df

def main():
    st.sidebar.title("Navigation")

    # HTML code to style the image position in the sidebar
    logo_html = """
        <div style="position: fixed; 
                    bottom: 30px; 
                    left: 30px;">
            <img src="https://raw.githubusercontent.com/akeenantexasfcs/ltma/main/TFCLogo.jpg" width="250"/>
        </div>
        """
    # Display the logo in the sidebar using HTML
    st.sidebar.markdown(logo_html, unsafe_allow_html=True)

    selection = st.sidebar.radio("Go to", ["Balance Sheet", "Cash Flow Statement", "Income Statement", "Populate CIQ Template", "Extras"])

    if selection == "Balance Sheet":
        balance_sheet_BS()
    elif selection == "Cash Flow Statement":
        cash_flow_statement_CF()
    elif selection == "Income Statement":
        income_statement()
    elif selection == "Populate CIQ Template":
        populate_ciq_template_pt()
    elif selection == "Extras":
        extras_tab()

if __name__ == '__main__':
    main()

