import streamlit as st
import pandas as pd
import openpyxl
import shutil
import datetime
import os
import json
import logging
from typing import List, Dict, Any, Optional, Tuple
from google_drive_service import GoogleDriveService

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    force=True
)
logger = logging.getLogger('ExcelTransferBot')

# Get the current directory of the script
current_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(current_dir, 'BC CALC (4).xlsx')
GENERATED_FILES_DIR = os.path.join(current_dir, 'generated_files')
LOG_PATH = os.path.join(current_dir, 'generated_files_log.json')

# Create generated_files directory if it doesn't exist
os.makedirs(GENERATED_FILES_DIR, exist_ok=True)

def process_excel_file(input_df: pd.DataFrame, keyword: str) -> Tuple[bool, str, Optional[str]]:
    """Process the input DataFrame and generate Excel file"""
    try:
        # Create output filename with timestamp
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        output_filename = f"{keyword}_{today}.xlsx"
        output_filepath = os.path.join(GENERATED_FILES_DIR, output_filename)

        # Copy template file
        shutil.copy(TEMPLATE_PATH, output_filepath)
        
        # Open the new file using openpyxl
        workbook = openpyxl.load_workbook(output_filepath)
        worksheet = workbook.active
        
        # Helper function to extract data based on column name
        def extract_data(df: pd.DataFrame, possible_column_names: List[str], start_row: int, column_letter: str):
            for column_name in possible_column_names:
                matching_columns = [col for col in df.columns if column_name.lower() in col.lower()]
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

        # Format revenue cells as currency
        for i in range(4, 14):
            cell = worksheet[f"I{i}"]
            if cell.value is not None:
                try:
                    cleaned_value = str(cell.value).replace(",", "")
                    numeric_value = float(cleaned_value)
                    cell.value = numeric_value
                    cell.number_format = '$#,##0.00'
                except ValueError:
                    logger.warning(f"Cell I{i} contains non-numeric data: {cell.value}")

        # Save the workbook
        workbook.save(output_filepath)
        
        return True, output_filepath, output_filename

    except Exception as e:
        error_msg = f"Error processing Excel file: {str(e)}"
        logger.error(error_msg)
        return False, error_msg, None

def update_log_file(log_entry: Dict[str, Any]) -> bool:
    """Update the log file with new entry"""
    try:
        if os.path.exists(LOG_PATH):
            with open(LOG_PATH, "r") as log_file:
                log_data = json.load(log_file)
        else:
            log_data = []
            
        log_data.append(log_entry)
        
        with open(LOG_PATH, "w") as log_file:
            json.dump(log_data, log_file, indent=4)
            
        return True
    except Exception as e:
        logger.error(f"Error updating log file: {str(e)}")
        return False

def main():
    """Main application entry point"""
    try:
        logger.info("Starting Excel Data Transfer Bot application")
        
        # Configure Streamlit page
        st.set_page_config(
            page_title="Excel Data Transfer Bot",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        # Display main application content
        st.title("Excel Data Transfer Bot")
        st.write("Upload your CSV file and transfer data to Excel template.")

        # Initialize Google Drive service
        drive_service = GoogleDriveService()
        auth_success, auth_message = drive_service.authenticate()
        
        if not auth_success:
            st.error(f"Failed to connect to Google Drive: {auth_message}")
            return
            
        st.success("✓ Connected to Google Drive")

        # File upload and keyword input
        keyword = st.text_input('Enter the keyword for the output filename:', key='keyword_input')
        uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"], key='file_upload')

        if st.button("Generate", key='generate_button'):
            if not keyword:
                st.warning("Please enter a keyword before generating.")
            elif not uploaded_file:
                st.warning("Please upload a CSV file before generating.")
            else:
                try:
                    with st.spinner("Processing file and uploading to Google Drive..."):
                        # Process the uploaded file
                        input_df = pd.read_csv(uploaded_file)
                        success, filepath, filename = process_excel_file(input_df, keyword)
                        
                        if success:
                            # Upload to Google Drive
                            upload_success, upload_message, file_id = drive_service.upload_file(filepath, filename)
                            
                            if upload_success:
                                # Create log entry
                                log_entry = {
                                    "keyword": keyword,
                                    "filename": filename,
                                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d"),
                                    "storage_path": filepath,
                                    "drive_file_id": file_id
                                }
                                
                                # Update log file
                                if update_log_file(log_entry):
                                    st.success(f"File generated and uploaded successfully: {filename}")
                                    
                                    # Provide download button
                                    with open(filepath, "rb") as file:
                                        st.download_button(
                                            label="Download Excel File",
                                            data=file,
                                            file_name=filename,
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key='download_button'
                                        )
                                else:
                                    st.warning("File processed but log update failed")
                            else:
                                st.error(f"Failed to upload file to Google Drive: {upload_message}")
                        else:
                            st.error(f"Failed to generate file: {filepath}")
                            
                except Exception as e:
                    st.error(f"Error processing file: {str(e)}")
                    logger.error(f"File processing error: {str(e)}")
        
        # Display log in sidebar
        st.sidebar.title("Generated Files Log")
        if os.path.exists(LOG_PATH):
            with open(LOG_PATH, "r") as log_file:
                log_data = json.load(log_file)
            for entry in reversed(log_data):
                st.sidebar.write(
                    f"Keyword: {entry['keyword']} | "
                    f"File: {entry['filename']} | "
                    f"Date: {entry['timestamp']}"
                )
        else:
            st.sidebar.write("No files have been generated yet.")
        
    except Exception as e:
        logger.error(f"Application error: {str(e)}")
        st.error(f"An unexpected error occurred: {str(e)}")

if __name__ == "__main__":
    logger.info("Application startup")
    main()
