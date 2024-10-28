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

# Health check endpoint
def health_check():
    return {"status": "healthy", "timestamp": datetime.datetime.now().isoformat()}

# Error handler for server startup
def handle_server_startup():
    try:
        # Set page configuration first, before any other Streamlit commands
        st.set_page_config(
            page_title="Excel Data Transfer Bot",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        return True
    except Exception as e:
        logger.error(f"Failed to start Streamlit server: {str(e)}")
        return False

# Initialize Google Drive Service
drive_service = GoogleDriveService()

# Load custom CSS
def load_css() -> None:
    try:
        css_file = Path(__file__).parent / "styles.css"
        with open(css_file) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except Exception as e:
        logger.error(f"Failed to load CSS: {str(e)}")
        st.warning("Custom styling could not be loaded")

# Get the current directory of the script
current_dir = Path(__file__).parent
TEMPLATE_PATH = current_dir / 'BC CALC (4).xlsx'
LOG_PATH = current_dir / 'generated_files_log.json'
STORAGE_DIR = current_dir / 'generated_files'

# Create storage directory if it doesn't exist
STORAGE_DIR.mkdir(exist_ok=True)
logger.info(f"Storage directory initialized at: {STORAGE_DIR}")

def get_stored_file_path(filename: str) -> Path:
    """Get the persistent path for a stored file."""
    return STORAGE_DIR / filename

def extract_data(df: pd.DataFrame, possible_column_names: List[str], 
                start_row: int, column_letter: str, worksheet: Any) -> None:
    """Extract data from DataFrame and write to worksheet."""
    try:
        for column_name in possible_column_names:
            matching_columns = [col for col in df.columns if column_name in col.lower()]
            if matching_columns:
                data = df[matching_columns[0]].iloc[:10].tolist()
                for i, value in enumerate(data):
                    worksheet[f"{column_letter}{start_row + i}"] = value
                return
    except Exception as e:
        logger.error(f"Error in extract_data: {str(e)}")
        raise

def format_currency_cells(worksheet: Any, column: str, start_row: int, end_row: int) -> None:
    """Format cells as currency."""
    try:
        for i in range(start_row, end_row):
            cell = worksheet[f"{column}{i}"]
            if cell.value is not None:
                try:
                    cleaned_value = str(cell.value).replace(",", "")
                    numeric_value = float(cleaned_value)
                    cell.value = numeric_value
                    cell.number_format = '$#,##0.00'
                except ValueError:
                    logger.warning(f"Cell {column}{i} contains non-numeric data: {cell.value}")
                    st.warning(f"⚠️ Cell {column}{i} contains non-numeric data: {cell.value}")
    except Exception as e:
        logger.error(f"Error in format_currency_cells: {str(e)}")
        raise

def update_log(keyword: str, output_filename: str, today: str, file_id: Optional[str] = None) -> None:
    """Update the log file with new entry."""
    try:
        storage_path = str(get_stored_file_path(output_filename))
        logger.info(f"Updating log with file: {output_filename}, storage path: {storage_path}")
        
        log_entry = {
            "keyword": keyword,
            "filename": output_filename,
            "timestamp": today,
            "storage_path": storage_path
        }
        if file_id:
            log_entry["drive_file_id"] = file_id
            
        log_data = []
        if LOG_PATH.exists():
            with open(LOG_PATH, "r") as log_file:
                log_data = json.load(log_file)
        log_data.append(log_entry)
        with open(LOG_PATH, "w") as log_file:
            json.dump(log_data, log_file, indent=4)
        logger.info(f"Log updated successfully with entry: {log_entry}")
    except Exception as e:
        logger.error(f"Error updating log: {str(e)}")
        raise

def create_download_button(filepath: Path, filename: str, key_suffix: str = "") -> None:
    """Create a download button for a file."""
    try:
        logger.info(f"Creating download button for file: {filepath}")
        if not isinstance(filepath, Path):
            filepath = Path(filepath)
            
        if not filepath.is_file():
            logger.error(f"File not found or is not a file: {filepath}")
            return
            
        with open(filepath, "rb") as file:
            unique_key = f"download_{filename}_{key_suffix}" if key_suffix else f"download_{filename}"
            st.download_button(
                label="📥 Download Excel File",
                data=file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=unique_key
            )
        logger.info(f"Download button created successfully for file: {filename}")
    except Exception as e:
        logger.error(f"Error creating download button: {str(e)}")
        st.error(f"❌ Error creating download button: {str(e)}")

def main():
    # Initialize server and handle startup errors
    if not handle_server_startup():
        st.error("Failed to start the application. Please try again later.")
        return

    try:
        # Load custom CSS
        load_css()

        # Add health check endpoint
        if st.experimental_get_query_params().get("health") == ["check"]:
            st.json(health_check())
            return

        # Title only
        st.title('Excel Data Transfer Bot')

        # Authenticate with Google Drive
        auth_success, auth_message = drive_service.authenticate()
        if not auth_success:
            st.error(f"❌ Google Drive Authentication Failed: {auth_message}")
            return

        # User inputs
        col1, col2 = st.columns(2)
        with col1:
            keyword = st.text_input('Enter the keyword for the output filename:',
                                help="This will be used in the output filename")
        
        with col2:
            uploaded_file = st.file_uploader(
                "Upload your CSV file",
                type=["csv"],
                help="Upload a CSV file containing your data"
            )

        # Generate button
        if st.button("Generate", key='generate_button'):
            if not keyword:
                st.error("⚠️ Please enter a keyword before generating.")
                return
            if not uploaded_file:
                st.error("⚠️ Please upload a CSV file before generating.")
                return
            
            try:
                # Process the file
                today = datetime.datetime.now().strftime("%Y-%m-%d")
                output_filename = f"{keyword}_{today}.xlsx"
                output_filepath = get_stored_file_path(output_filename)
                
                logger.info(f"Processing file: {output_filename}")
                logger.info(f"Output filepath: {output_filepath}")

                # Load and process data
                input_df = pd.read_csv(uploaded_file)
                shutil.copy(TEMPLATE_PATH, output_filepath)
                logger.info(f"Template copied to: {output_filepath}")
                
                workbook = openpyxl.load_workbook(output_filepath)
                worksheet = workbook.active

                # Extract and write data
                extract_data(input_df, ['product details', 'product'], 4, 'F', worksheet)
                extract_data(input_df, ['brand'], 4, 'G', worksheet)
                extract_data(input_df, ['price'], 4, 'H', worksheet)
                extract_data(input_df, ['revenue'], 4, 'I', worksheet)

                # Format currency cells
                format_currency_cells(worksheet, 'I', 4, 14)
                
                # Save workbook to persistent storage
                workbook.save(output_filepath)
                logger.info(f"Excel file saved to persistent storage: {output_filepath}")
                
                # Upload to Google Drive
                logger.info("Starting Google Drive upload")
                file_id = drive_service.upload_file(str(output_filepath), output_filename)
                
                if not file_id:
                    st.error("❌ Failed to upload file to Google Drive")
                    logger.error("Failed to upload file to Google Drive")
                else:
                    # Update log with persistent storage path and file ID
                    update_log(keyword, output_filename, today, file_id)
                    
                    # Show success message and create single download button
                    st.success('✅ File generated and uploaded successfully!')
                    create_download_button(output_filepath, output_filename, key_suffix=today)

            except Exception as e:
                error_msg = f"❌ An error occurred while processing the file: {str(e)}"
                logger.error(error_msg)
                st.error(error_msg)
                st.error("Please check your CSV file format and try again.")

        # Simplified sidebar with log (no download buttons)
        st.sidebar.title("📋 Generated Files Log")
        if LOG_PATH.exists():
            with open(LOG_PATH, "r") as log_file:
                log_data = json.load(log_file)
            for entry in reversed(log_data):
                st.sidebar.markdown(
                    f"""
                    **File:** {entry['filename']}  
                    **Date:** {entry['timestamp']}
                    ---
                    """
                )
        else:
            st.sidebar.info("No files have been generated yet.")

    except Exception as e:
        error_msg = f"❌ An unexpected error occurred: {str(e)}"
        logger.error(error_msg)
        st.error(error_msg)

if __name__ == "__main__":
    main()
