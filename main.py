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
                    # Process the uploaded file
                    input_df = pd.read_csv(uploaded_file)
                    success, result, output_filename = process_excel_file(input_df, keyword)
                    
                    if success:
                        st.success(f"File generated successfully: {output_filename}")
                        
                        # Provide download button
                        with open(result, "rb") as file:
                            st.download_button(
                                label="Download Excel File",
                                data=file,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key='download_button'
                            )
                    else:
                        st.error(f"Failed to generate file: {result}")
                        
                except Exception as e:
                    st.error(f"Error processing file: {str(e)}")
                    logger.error(f"File processing error: {str(e)}")
        
    except Exception as e:
        logger.error(f"Application error: {str(e)}")
        st.error(f"An unexpected error occurred: {str(e)}")

if __name__ == "__main__":
    logger.info("Application startup")
    main()
