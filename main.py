import streamlit as st
import pandas as pd
import openpyxl
import shutil
import datetime
import os
import json
import logging
from typing import List, Dict, Any, Optional
from openpyxl.utils import get_column_letter
from pathlib import Path
from google_drive_service import GoogleDriveService

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('ExcelTransferBot')

# Health check endpoint for cloud deployment
@st.cache_data
def health_check():
    """Health check endpoint for cloud deployment"""
    try:
        return {
            "status": "healthy",
            "timestamp": datetime.datetime.now().isoformat(),
            "service": "Excel Data Transfer Bot",
            "environment": "cloud"
        }
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        return {
            "status": "unhealthy",
            "error": str(e),
            "timestamp": datetime.datetime.now().isoformat()
        }

def main():
    try:
        # Set page configuration
        st.set_page_config(
            page_title="Excel Data Transfer Bot",
            layout="wide",
            initial_sidebar_state="expanded"
        )

        # Display health check status
        if st.sidebar.checkbox("Show Health Status", value=False):
            st.sidebar.json(health_check())

        # Title
        st.title('Excel Data Transfer Bot')

        # Initialize Google Drive service
        drive_service = GoogleDriveService()
        auth_success, auth_message = drive_service.authenticate()
        
        if not auth_success:
            st.error(f"Failed to authenticate with Google Drive: {auth_message}")
            return

        # Step 1: User inputs keyword
        keyword = st.text_input('Enter the keyword for the output filename:', key='keyword_input')

        # Step 2: User uploads CSV file
        uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"], key='file_upload')

        # Step 3: Generate button for user to initiate generation
        if st.button("Generate", key='generate_button'):
            if not keyword:
                st.warning("Please enter a keyword before generating.")
            elif not uploaded_file:
                st.warning("Please upload a CSV file before generating.")
            else:
                try:
                    # Create output directory if it doesn't exist
                    output_dir = Path("generated_files")
                    output_dir.mkdir(exist_ok=True)

                    # Extract today's date for the output file name
                    today = datetime.datetime.now().strftime("%Y-%m-%d")
                    output_filename = f"{keyword}_{today}.xlsx"
                    output_filepath = output_dir / output_filename

                    # Load the input CSV file into a pandas DataFrame
                    input_df = pd.read_csv(uploaded_file)
                    
                    # Copy the template file to create a new file
                    template_path = Path("BC CALC (4).xlsx")
                    shutil.copy(template_path, output_filepath)
                    
                    # Open the new file using openpyxl
                    workbook = openpyxl.load_workbook(output_filepath)
                    worksheet = workbook.active

                    # Helper function to extract data based on column name
                    def extract_data(df, possible_column_names, start_row, column_letter):
                        for column_name in possible_column_names:
                            matching_columns = [col for col in df.columns if column_name in col.lower()]
                            if matching_columns:
                                data = df[matching_columns[0]].iloc[:10].tolist()
                                for i, value in enumerate(data):
                                    worksheet[f"{column_letter}{start_row + i}"] = value
                                return

                    # Extract and write data
                    extract_data(input_df, ['product details', 'product'], 4, 'F')
                    extract_data(input_df, ['brand'], 4, 'G')
                    extract_data(input_df, ['price'], 4, 'H')
                    extract_data(input_df, ['revenue'], 4, 'I')

                    # Format revenue cells
                    for i in range(4, 14):
                        cell = worksheet[f"I{i}"]
                        if cell.value is not None:
                            try:
                                cleaned_value = str(cell.value).replace(",", "")
                                numeric_value = float(cleaned_value)
                                cell.value = numeric_value
                                cell.number_format = '$#,##0.00'
                            except ValueError:
                                st.warning(f"Cell I{i} contains non-numeric data: {cell.value}")

                    # Save the workbook
                    workbook.save(output_filepath)

                    # Upload to Google Drive
                    file_id = drive_service.upload_file(str(output_filepath), output_filename)

                    if file_id:
                        st.success(f"✅ File generated and uploaded successfully!")
                        
                        # Provide download button
                        with open(output_filepath, "rb") as file:
                            st.download_button(
                                label="Download Excel File",
                                data=file,
                                file_name=output_filename,
                                key='download_button'
                            )
                    else:
                        st.warning("File generated but couldn't be uploaded to Google Drive")
                        
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")

        # Display log in sidebar
        st.sidebar.title("Generated Files Log")
        log_path = Path("generated_files_log.json")
        if log_path.exists():
            with open(log_path, "r") as log_file:
                log_data = json.load(log_file)
            for entry in reversed(log_data):
                st.sidebar.write(f"Keyword: {entry['keyword']} | File: {entry['filename']} | Date: {entry['timestamp']}")
        else:
            st.sidebar.write("No files have been generated yet.")

    except Exception as e:
        logger.error(f"Application error: {str(e)}")
        st.error(f"An unexpected error occurred: {str(e)}")

if __name__ == "__main__":
    main()
