#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import streamlit as st
import boto3
import tempfile
import json
import botocore
import pandas as pd
import time
import io

def check_aws_credentials(access_key, secret_key, region):
    try:
        session = boto3.Session(
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key,
            region_name=region
        )
        sts = session.client('sts')
        sts.get_caller_identity()
        return True
    except botocore.exceptions.ClientError:
        return False

def safe_get(dict_obj, key, default=None):
    """Safely get a value from a dictionary."""
    return dict_obj.get(key, default)

def extract_table_data(table_blocks, blocks_map):
    rows = []
    for relationship in safe_get(table_blocks, 'Relationships', []):
        if safe_get(relationship, 'Type') == 'CHILD':
            for child_id in safe_get(relationship, 'Ids', []):
                cell = blocks_map.get(child_id, {})
                if safe_get(cell, 'BlockType') == 'CELL':
                    row_index = safe_get(cell, 'RowIndex', 1) - 1
                    col_index = safe_get(cell, 'ColumnIndex', 1) - 1
                    while len(rows) <= row_index:
                        rows.append([])
                    while len(rows[row_index]) <= col_index:
                        rows[row_index].append('')
                    cell_text = safe_get(cell, 'Text', '')
                    if not cell_text:
                        # If 'Text' is not present, try to get it from child relationships
                        for cell_relationship in safe_get(cell, 'Relationships', []):
                            if safe_get(cell_relationship, 'Type') == 'CHILD':
                                for word_id in safe_get(cell_relationship, 'Ids', []):
                                    word = blocks_map.get(word_id, {})
                                    cell_text += safe_get(word, 'Text', '') + ' '
                        cell_text = cell_text.strip()
                    rows[row_index][col_index] = cell_text
    return rows

def upload_to_s3(s3_client, file_bytes, bucket_name, object_name):
    try:
        s3_client.put_object(Body=file_bytes, Bucket=bucket_name, Key=object_name)
        st.success(f"Successfully uploaded file to S3: {bucket_name}/{object_name}")
    except botocore.exceptions.ClientError as e:
        st.error(f"Failed to upload file to S3: {str(e)}")
        raise

def start_document_analysis(textract_client, bucket_name, object_name):
    try:
        response = textract_client.start_document_analysis(
            DocumentLocation={'S3Object': {'Bucket': bucket_name, 'Name': object_name}},
            FeatureTypes=['TABLES']
        )
        return response['JobId']
    except botocore.exceptions.ClientError as e:
        st.error(f"Failed to start document analysis: {str(e)}")
        raise

def get_document_analysis(textract_client, job_id):
    pages = []
    next_token = None
    max_attempts = 60  # Increase this for larger documents
    attempt = 0
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    while attempt < max_attempts:
        try:
            if next_token:
                response = textract_client.get_document_analysis(JobId=job_id, NextToken=next_token)
            else:
                response = textract_client.get_document_analysis(JobId=job_id)
            
            status = response['JobStatus']
            status_text.text(f"Processing document... Status: {status}")
            
            if status == 'SUCCEEDED':
                pages.append(response)
                next_token = response.get('NextToken')
                if not next_token:
                    status_text.text("Document processing complete!")
                    progress_bar.progress(1.0)
                    break
            elif status == 'FAILED':
                raise Exception(f"Textract job failed: {response.get('StatusMessage')}")
            else:  # IN_PROGRESS
                time.sleep(5)
                attempt += 1
                progress_bar.progress(attempt / max_attempts)
        except Exception as e:
            st.error(f"An error occurred while getting document analysis: {str(e)}")
            raise
    
    if attempt >= max_attempts:
        raise Exception("Textract job timed out")
    
    return pages

def process_document(file_path, textract_client, s3_client, bucket_name):
    try:
        with open(file_path, 'rb') as document:
            file_bytes = document.read()
        
        _, file_extension = os.path.splitext(file_path)
        object_name = os.path.basename(file_path)
        
        if file_extension.lower() == '.pdf':
            st.info("Uploading document to S3...")
            upload_to_s3(s3_client, file_bytes, bucket_name, object_name)
            st.info("Starting Textract job...")
            job_id = start_document_analysis(textract_client, bucket_name, object_name)
            st.info(f"Textract job started with ID: {job_id}")
            response_pages = get_document_analysis(textract_client, job_id)
        else:
            response = textract_client.analyze_document(
                Document={'Bytes': file_bytes},
                FeatureTypes=['TABLES']
            )
            response_pages = [response]
        
        if not response_pages:
            raise Exception("Document analysis failed")
        
        # Extract and combine 'Blocks' from all pages
        all_blocks = []
        for page in response_pages:
            all_blocks.extend(page.get('Blocks', []))
        
        # Create a simplified JSON structure
        simplified_response = {'Blocks': all_blocks}
        
        # Save the simplified response as a JSON file
        response_json_path = file_path + '.json'
        with open(response_json_path, 'w') as json_file:
            json.dump(simplified_response, json_file, indent=4)
        
        # Extract tables
        tables = []

        # Process all blocks
        blocks_map = {safe_get(block, 'Id'): block for block in all_blocks}
        for block in all_blocks:
            block_type = safe_get(block, 'BlockType')
            if block_type == 'TABLE':
                tables.append(extract_table_data(block, blocks_map))
        
        return tables, response_json_path, simplified_response

    except Exception as e:
        st.error(f"An error occurred during document processing: {str(e)}")
        raise

# Initialize session state
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'credentials_valid' not in st.session_state:
    st.session_state.credentials_valid = False

def aws_textract_tab():
    st.title("AWS Textract with Streamlit - Table Extraction")
    st.write("Enter your AWS credentials and upload an image or PDF file to extract tables using AWS Textract.")

    # AWS Credentials Input
    aws_access_key = st.text_input("AWS Access Key ID", type="password", key="aws_access_key")
    aws_secret_key = st.text_input("AWS Secret Access Key", type="password", key="aws_secret_key")
    aws_region = st.selectbox("AWS Region", ["us-east-2", "us-east-1", "us-west-1", "us-west-2"], index=0)
    s3_bucket_name = st.text_input("S3 Bucket Name")

    if st.button("Confirm Credentials"):
        if check_aws_credentials(aws_access_key, aws_secret_key, aws_region):
            st.success("AWS credentials are valid!")
            st.session_state.credentials_valid = True
        else:
            st.error("Invalid AWS credentials. Please check and try again.")
            st.session_state.credentials_valid = False

    # File Upload (only show if credentials are valid)
    if st.session_state.credentials_valid:
        uploaded_file = st.file_uploader("Choose an image or PDF file", type=["jpg", "jpeg", "png", "pdf"])

        if uploaded_file is not None and not st.session_state.processed:
            temp_file_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as temp_file:
                    temp_file.write(uploaded_file.getvalue())
                    temp_file_path = temp_file.name
                
                textract_client = boto3.client('textract',
                                               aws_access_key_id=aws_access_key,
                                               aws_secret_access_key=aws_secret_key,
                                               region_name=aws_region)
                
                s3_client = boto3.client('s3',
                                         aws_access_key_id=aws_access_key,
                                         aws_secret_access_key=aws_secret_key,
                                         region_name=aws_region)
                
                with st.spinner("Processing document..."):
                    tables, response_json_path, simplified_response = process_document(temp_file_path, textract_client, s3_client, s3_bucket_name)
                
                st.session_state.processed = True
                st.session_state.response_json_path = response_json_path
                st.session_state.simplified_response = simplified_response
                st.session_state.tables = tables

            except botocore.exceptions.ClientError as e:
                st.error(f"AWS Error: {str(e)}")
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
                st.error("Please check the AWS credentials, S3 bucket permissions, and try again.")
            finally:
                # Clean up the temporary files
                if temp_file_path and os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)

    # Show download button and table extraction results only if processing is complete
    if st.session_state.processed:
        response_json_path = st.session_state.response_json_path
        simplified_response = st.session_state.simplified_response
        tables = st.session_state.tables
        
        # JSON file rename input
        st.subheader("Download Full JSON Response:")
        json_file_name = st.text_input("Enter JSON file name (without .json extension)", value='textract_response')
        
        with open(response_json_path, 'rb') as f:
            st.download_button(
                label="Download JSON",
                data=f,
                file_name=f"{json_file_name}.json" if json_file_name else 'textract_response.json',
                mime='application/json'
            )

        # Now display the extracted information
        st.subheader("Detected Tables:")
        st.write(f"Number of tables detected: {len(tables)}")
        if tables:
            for i, table in enumerate(tables):
                st.write(f"Table {i+1}:")
                if table:
                    df = pd.DataFrame(table)
                    st.dataframe(df)
                else:
                    st.write("Empty table detected")
        else:
            st.write("No tables detected")

        # Debug information
        st.subheader("Debug Information:")
        st.write(f"Number of blocks processed: {len(simplified_response['Blocks'])}")
        st.json(simplified_response)

        # Display structure of the first few blocks
        st.subheader("Structure of First Few Blocks:")
        first_few_blocks = simplified_response['Blocks'][:10]  # Display first 10 blocks
        for i, block in enumerate(first_few_blocks):
            st.write(f"Block {i}:")
            st.json(block)

        # Add a button to reset the process
        if st.button("Process another document"):
            st.session_state.processed = False
            st.experimental_rerun()

def extras_tab():
    st.title("Extras")
    selection = st.radio("Select an option", ["Backup Data Dictionaries", "AWS Textract"])

    if selection == "Backup Data Dictionaries":
        if st.button("Backup Data Dictionaries"):
            # Load data dictionaries
            balance_sheet_data = pd.read_csv('balance_sheet_data_dictionary.csv')
            cash_flow_data = pd.read_csv('cash_flow_data_dictionary.csv')
            income_statement_data = pd.read_excel('income_statement_data_dictionary.xlsx')

            # Create a new Excel writer object
            with pd.ExcelWriter("data_dictionaries_backup.xlsx", engine='xlsxwriter') as writer:
                # Write each DataFrame to a different sheet
                balance_sheet_data.to_excel(writer, sheet_name='Balance Sheet', index=False)
                cash_flow_data.to_excel(writer, sheet_name='Cash Flow', index=False)
                income_statement_data.to_excel(writer, sheet_name='Income Statement', index=False)

            # Read the file into a BytesIO object for download
            with open("data_dictionaries_backup.xlsx", "rb") as file:
                backup_file = io.BytesIO(file.read())

            st.download_button(
                label="Download Backup",
                data=backup_file,
                file_name="data_dictionaries_backup.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    elif selection == "AWS Textract":
        aws_textract_tab()

def main():
    st.sidebar.title("Navigation")
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

