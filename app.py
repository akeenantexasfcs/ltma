#!/usr/bin/env python
# coding: utf-8

# In[3]:


import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np

# Function to get futures data
def get_futures_data(ticker_symbol, start_date, end_date):
    ticker_data = yf.Ticker(ticker_symbol)
    ticker_df = ticker_data.history(period='1d', start=start_date, end=end_date)
    return ticker_df

# Define the functions for the Altman Z Score section
def ratio_x_1(ticker):
    df = ticker.balance_sheet
    working_capital = df.loc['Current Assets'].iloc[0] - df.loc['Current Liabilities'].iloc[0]
    total_assets = df.loc['Total Assets'].iloc[0]
    return working_capital / total_assets

def ratio_x_2(ticker):
    df = ticker.balance_sheet
    retained_earnings = df.loc['Retained Earnings'].iloc[0]
    total_assets = df.loc['Total Assets'].iloc[0]
    return retained_earnings / total_assets

def ratio_x_3(ticker):
    df = ticker.financials
    ebit = df.loc['EBIT'].iloc[0]
    total_assets = ticker.balance_sheet.loc['Total Assets'].iloc[0]
    return ebit / total_assets

def ratio_x_4(ticker):
    equity_market_value = ticker.info['sharesOutstanding'] * ticker.info['currentPrice']
    total_liabilities = ticker.balance_sheet.loc['Total Liabilities Net Minority Interest'].iloc[0]
    return equity_market_value / total_liabilities

def ratio_x_5(ticker):
    df = ticker.financials
    sales = df.loc['Total Revenue'].iloc[0]
    total_assets = ticker.balance_sheet.loc['Total Assets'].iloc[0]
    return sales / total_assets

def z_score(ticker):
    try:
        ratio_1 = ratio_x_1(ticker)
        ratio_2 = ratio_x_2(ticker)
        ratio_3 = ratio_x_3(ticker)
        ratio_4 = ratio_x_4(ticker)
        ratio_5 = ratio_x_5(ticker)
        zscore = 1.2 * ratio_1 + 1.4 * ratio_2 + 3.3 * ratio_3 + 0.6 * ratio_4 + 1.0 * ratio_5
        return zscore
    except Exception as e:
        return np.nan

# Streamlit App
st.sidebar.title('Navigation')
option = st.sidebar.selectbox('Select a section:', ['Altman Z Score', 'Futures Pricing'])

if option == 'Altman Z Score':
    st.title('Altman Z-Score Calculator')
    
    # Define the number of input slots
    NUM_INPUTS = 10
    
    # Input fields for ticker symbols
    tickers = []
    for i in range(NUM_INPUTS):
        ticker = st.text_input(f'Ticker {i+1}', '')
        if ticker:
            tickers.append(ticker.upper())
    
    # Dictionary to hold scores for each symbol
    symbol_to_score = {}
    
    # Calculate Z-Scores
    if st.button('Calculate Z-Scores'):
        for symbol in tickers:
            ticker = yf.Ticker(symbol)
            symbol_to_score[symbol] = z_score(ticker)
    
        # Categorize
        distress = [''] * len(tickers)
        grey = [''] * len(tickers)
        safe = [''] * len(tickers)
    
        for idx, symbol in enumerate(tickers):
            zscore = symbol_to_score[symbol]
            if zscore <= 1.8:
                distress[idx] = zscore
            elif zscore > 1.8 and zscore <= 2.99:
                grey[idx] = zscore
            else:
                safe[idx] = zscore
    
        # Create DataFrame
        data_dict = {'Symbol': tickers, 'Distress Zone': distress, 'Grey Zone': grey, 'Safe Zone': safe}
        df = pd.DataFrame.from_dict(data_dict)
    
        # Apply styles
        def highlight_distress(val):
            return 'background-color: indianred' if val != '' else ""
    
        def highlight_grey(val):
            return 'background-color: grey' if val != '' else ""
    
        def highlight_safe(val):
            return 'background-color: green' if val != '' else ""
    
        def format_score(val):
            try:
                return '{:.2f}'.format(float(val))
            except:
                return ''
    
        def make_pretty(styler):
            # No index
            styler.hide(axis='index')
            
            # Column formatting
            styler.format(format_score, subset=['Distress Zone', 'Grey Zone', 'Safe Zone'])
    
            # Left text alignment for some columns
            styler.set_properties(subset=['Symbol', 'Distress Zone', 'Grey Zone', 'Safe Zone'], **{'text-align': 'center', 'width': '100px'})
    
            # Apply highlight methods to columns
            styler.applymap(highlight_grey, subset=['Grey Zone'])
            styler.applymap(highlight_safe, subset=['Safe Zone'])
            styler.applymap(highlight_distress, subset=['Distress Zone'])
            return styler
    
        # Apply styles to DataFrame
        styles = [
            dict(selector='td', props=[('font-size', '10pt'), ('border-style', 'solid'), ('border-width', '1px')]),
            dict(selector='th.col_heading', props=[('font-size', '11pt'), ('text-align', 'center')]),
            dict(selector='caption', props=[('text-align', 'center'), ('font-size', '14pt'), ('font-weight', 'bold')])
        ]
    
        df_styled = df.style.pipe(make_pretty).set_caption('Altman Z Score').set_table_styles(styles)
        st.dataframe(df_styled)

elif option == 'Futures Pricing':
    st.title('Futures Pricing')
    ticker = st.text_input('Enter the Futures Ticker Symbol (e.g., CL=F):', 'CL=F')
    start_date = st.date_input('Start Date', value=pd.to_datetime('2014-01-01'))
    end_date = st.date_input('End Date', value=pd.to_datetime('2024-04-08'))
    
    if st.button('Get Data'):
        data = get_futures_data(ticker, start_date, end_date)
        st.write(data)
        
        # Export to CSV
        csv_file_name = f'historical_prices_{ticker}.csv'
        data.to_csv(csv_file_name)
        st.write(f'Data exported to {csv_file_name}')
        st.download_button(
            label="Download CSV",
            data=data.to_csv().encode('utf-8'),
            file_name=csv_file_name,
            mime='text/csv',
        )

