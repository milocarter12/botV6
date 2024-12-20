from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import os
import io
import streamlit as st
from typing import Optional, Tuple
from datetime import datetime
from pathlib import Path
import logging
import json
import base64
import re

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('GoogleDriveService')

class GoogleDriveService:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/drive.file']
        self.credentials = None
        self.drive_service = None
        self.root_folder_name = 'Excel Data Transfer Bot'
        self.root_folder_id = None
        self.user_email = 'milocarter12@gmail.com'

    def _validate_private_key(self, private_key: str) -> Tuple[bool, str]:
        """Validate the format of the private key."""
        try:
            logger.info("Validating private key format")
            if not private_key:
                return False, "Private key is missing"

            # Remove any escaped newlines and replace with actual newlines
            private_key = private_key.replace('\\n', '\n')

            # Check if the key has the correct header and footer
            if not (private_key.startswith('-----BEGIN PRIVATE KEY-----') and 
                   private_key.endswith('-----END PRIVATE KEY-----')):
                
                # Try to fix the key format if it's just the base64 part
                if re.match(r'^[A-Za-z0-9+/=\n]+$', private_key.strip()):
                    private_key = (
                        '-----BEGIN PRIVATE KEY-----\n' +
                        private_key.strip() +
                        '\n-----END PRIVATE KEY-----'
                    )
                else:
                    return False, "Invalid private key format: Missing header/footer"

            # Verify the key can be decoded as base64
            try:
                key_parts = private_key.split('-----')
                if len(key_parts) < 3:
                    return False, "Invalid private key structure"
                
                base64_part = key_parts[2].strip()
                base64.b64decode(base64_part)
            except Exception as e:
                return False, f"Invalid base64 encoding in private key: {str(e)}"

            logger.info("Private key validation successful")
            return True, private_key
        except Exception as e:
            error_msg = f"Error validating private key: {str(e)}"
            logger.error(error_msg)
            return False, error_msg

    def _check_required_env_vars(self) -> Tuple[bool, str]:
        """Check if all required environment variables are present."""
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

        missing_vars = [var for var in required_vars if not os.environ.get(var)]
        
        if missing_vars:
            error_msg = f"Missing required environment variables: {', '.join(missing_vars)}"
            logger.error(error_msg)
            return False, error_msg
            
        logger.info("All required environment variables are present")
        return True, "All required environment variables are present"

    def authenticate(self) -> Tuple[bool, str]:
        """Authenticate using service account credentials from environment variables."""
        try:
            logger.info("Starting Google Drive authentication")
            
            # Check required environment variables
            env_vars_ok, env_vars_msg = self._check_required_env_vars()
            if not env_vars_ok:
                return False, env_vars_msg

            # Validate private key
            private_key = os.environ.get("GOOGLE_PRIVATE_KEY", "")
            is_valid_key, key_result = self._validate_private_key(private_key)
            
            if not is_valid_key:
                return False, f"Private key validation failed: {key_result}"

            # Create service account info dictionary
            service_account_info = {
                "type": os.environ["GOOGLE_SERVICE_ACCOUNT_TYPE"],
                "project_id": os.environ["GOOGLE_PROJECT_ID"],
                "private_key_id": os.environ["GOOGLE_PRIVATE_KEY_ID"],
                "private_key": key_result,
                "client_email": os.environ["GOOGLE_CLIENT_EMAIL"],
                "client_id": os.environ["GOOGLE_CLIENT_ID"],
                "auth_uri": os.environ["GOOGLE_AUTH_URI"],
                "token_uri": os.environ["GOOGLE_TOKEN_URI"],
                "auth_provider_x509_cert_url": os.environ["GOOGLE_AUTH_PROVIDER_X509_CERT_URL"],
                "client_x509_cert_url": os.environ["GOOGLE_CLIENT_X509_CERT_URL"]
            }

            logger.info("Service account info prepared successfully")

            # Create credentials
            try:
                self.credentials = service_account.Credentials.from_service_account_info(
                    service_account_info,
                    scopes=self.SCOPES
                )
                logger.info("Successfully created credentials")
            except Exception as e:
                error_msg = f"Failed to create credentials: {str(e)}"
                logger.error(error_msg)
                return False, error_msg

            # Initialize Drive service
            try:
                self.drive_service = build('drive', 'v3', credentials=self.credentials)
                logger.info("Successfully initialized Drive API service")
            except Exception as e:
                error_msg = f"Failed to initialize Drive service: {str(e)}"
                logger.error(error_msg)
                return False, error_msg

            return True, "Authentication successful"

        except Exception as e:
            error_msg = f"Authentication failed: {str(e)}"
            logger.error(error_msg)
            return False, error_msg

    def create_or_get_folder(self) -> Tuple[bool, str, Optional[str]]:
        """Create or get the root folder ID for the application."""
        try:
            logger.info(f"Looking for root folder: {self.root_folder_name}")
            
            # Search for existing folder
            query = f"name='{self.root_folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
            results = self.drive_service.files().list(q=query, spaces='drive').execute()
            items = results.get('files', [])

            if items:
                folder_id = items[0]['id']
                logger.info(f"Found existing folder with ID: {folder_id}")
                self.root_folder_id = folder_id
                return True, "Folder found", folder_id

            # Create new folder if it doesn't exist
            folder_metadata = {
                'name': self.root_folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            
            file = self.drive_service.files().create(
                body=folder_metadata,
                fields='id'
            ).execute()
            
            folder_id = file.get('id')
            logger.info(f"Created new folder with ID: {folder_id}")
            
            # Share the folder with the user
            self._share_file(folder_id)
            
            self.root_folder_id = folder_id
            return True, "Folder created", folder_id

        except Exception as e:
            error_msg = f"Error creating/getting folder: {str(e)}"
            logger.error(error_msg)
            return False, error_msg, None

    def _share_file(self, file_id: str) -> bool:
        """Share a file with the specified user email."""
        try:
            permission = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': self.user_email
            }
            
            self.drive_service.permissions().create(
                fileId=file_id,
                body=permission,
                sendNotificationEmail=False
            ).execute()
            
            logger.info(f"Successfully shared file {file_id} with {self.user_email}")
            return True
            
        except Exception as e:
            logger.error(f"Error sharing file {file_id}: {str(e)}")
            return False

    def upload_file(self, file_path: str, file_name: str) -> Tuple[bool, str, Optional[str]]:
        """Upload a file to Google Drive and share it."""
        try:
            logger.info(f"Starting upload for file: {file_name}")
            
            # Ensure we have a root folder
            folder_success, folder_msg, folder_id = self.create_or_get_folder()
            if not folder_success:
                return False, f"Failed to get/create root folder: {folder_msg}", None
            
            # Prepare file metadata
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            
            # Upload file
            try:
                with open(file_path, 'rb') as file:
                    media = MediaIoBaseUpload(
                        file,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        resumable=True
                    )
                    
                    file = self.drive_service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields='id'
                    ).execute()
                    
                    file_id = file.get('id')
                    logger.info(f"File uploaded successfully with ID: {file_id}")
                    
                    # Share the file
                    if self._share_file(file_id):
                        return True, "File uploaded and shared successfully", file_id
                    else:
                        return False, "File uploaded but sharing failed", file_id
                    
            except Exception as e:
                error_msg = f"Error uploading file: {str(e)}"
                logger.error(error_msg)
                return False, error_msg, None
                
        except Exception as e:
            error_msg = f"Error in upload process: {str(e)}"
            logger.error(error_msg)
            return False, error_msg, None
