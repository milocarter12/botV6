import streamlit as st
import pandas as pd
import openpyxl
import shutil
import datetime
import os
import json
import logging
from typing import List, Dict, Any, Optional, Tuple
from openpyxl.utils import get_column_letter
from pathlib import Path
from google_drive_service import GoogleDriveService

# Configure logging with more detailed format
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    force=True
)
logger = logging.getLogger('ExcelTransferBot')

def format_private_key(private_key: str) -> str:
    """Format private key string correctly with proper line breaks."""
    try:
        if private_key and not private_key.startswith('-----BEGIN PRIVATE KEY-----'):
            # Remove any existing escape characters
            private_key = private_key.replace('\\n', '\n')
            # Add header and footer if missing
            private_key = f"-----BEGIN PRIVATE KEY-----\n{private_key}\n-----END PRIVATE KEY-----"
        return private_key
    except Exception as e:
        logger.error(f"Error formatting private key: {str(e)}")
        return private_key

def log_environment_status():
    """Log the status of all required environment variables"""
    required_vars = [
        "GOOGLE_SERVICE_ACCOUNT_TYPE",
        "GOOGLE_PROJECT_ID",
        "GOOGLE_PRIVATE_KEY_ID",
        "GOOGLE_PRIVATE_KEY",
        "GOOGLE_CLIENT_EMAIL",
        "GOOGLE_CLIENT_ID",
        "GOOGLE_AUTH_URI",
        "GOOGLE_TOKEN_URI",
        "GOOGLE_AUTH_PROVIDER_X509_CERT_URL",
        "GOOGLE_CLIENT_X509_CERT_URL"
    ]
    
    logger.info("Checking environment variables...")
    missing_vars = []
    
    try:
        for var in required_vars:
            value = os.environ.get(var)
            if value:
                # Log presence without exposing sensitive data
                logger.info(f"✓ {var} is set")
            else:
                logger.error(f"✗ {var} is missing")
                missing_vars.append(var)
        
        # Special handling for private key
        if "GOOGLE_PRIVATE_KEY" in os.environ:
            os.environ["GOOGLE_PRIVATE_KEY"] = format_private_key(os.environ["GOOGLE_PRIVATE_KEY"])
            
        return missing_vars
    except Exception as e:
        logger.error(f"Error checking environment variables: {str(e)}")
        return required_vars

@st.cache_data(show_spinner=False)
def health_check():
    """Health check endpoint for cloud deployment"""
    try:
        # Check environment variables
        missing_vars = log_environment_status()
        env_vars_set = len(missing_vars) == 0
        
        # Test Google Drive authentication
        drive_service = GoogleDriveService()
        auth_success, auth_message = drive_service.authenticate()
        
        health_status = {
            "status": "healthy" if (env_vars_set and auth_success) else "error",
            "timestamp": datetime.datetime.now().isoformat(),
            "service": "Excel Data Transfer Bot",
            "environment": "cloud",
            "env_vars_set": env_vars_set,
            "google_drive_auth": auth_success,
            "missing_vars": missing_vars if missing_vars else None,
            "auth_message": auth_message if not auth_success else None
        }
        
        logger.info(f"Health check status: {health_status['status']}")
        return health_status
    except Exception as e:
        error_msg = f"Health check failed: {str(e)}"
        logger.error(error_msg)
        return {
            "status": "unhealthy",
            "error": str(e),
            "timestamp": datetime.datetime.now().isoformat()
        }

def initialize_app():
    """Initialize the Streamlit application with proper error handling"""
    try:
        # Configure Streamlit page
        st.set_page_config(
            page_title="Excel Data Transfer Bot",
            layout="wide",
            initial_sidebar_state="expanded"
        )

        # Perform health check
        with st.spinner("Checking application health..."):
            health_status = health_check()
            
        # Display health status in sidebar
        with st.sidebar:
            st.json(health_status)
            
            if health_status["status"] != "healthy":
                st.error("⚠️ Application Configuration Error")
                if health_status.get("missing_vars"):
                    st.error(f"Missing environment variables: {', '.join(health_status['missing_vars'])}")
                if health_status.get("auth_message"):
                    st.error(f"Authentication error: {health_status['auth_message']}")
                return False
            
            st.success("✓ Application is healthy")
        
        return True
        
    except Exception as e:
        logger.error(f"Application initialization error: {str(e)}")
        st.error(f"Failed to initialize application: {str(e)}")
        return False

def main():
    """Main application entry point with improved error handling"""
    try:
        logger.info("Starting Excel Data Transfer Bot application")
        
        # Initialize application
        if not initialize_app():
            return
        
        # Display main application content
        st.title("Excel Data Transfer Bot")
        st.write("Upload your CSV file and transfer data to Excel template.")

        # Initialize Google Drive service (already verified in health check)
        drive_service = GoogleDriveService()
        drive_service.authenticate()
        st.success("✓ Connected to Google Drive")

        # Rest of your existing application code...
        
    except Exception as e:
        logger.error(f"Application error: {str(e)}")
        st.error(f"An unexpected error occurred: {str(e)}")

if __name__ == "__main__":
    logger.info("Application startup")
    main()
