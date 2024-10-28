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
logging.basicConfig(level=logging.INFO)
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
            # Check if key is None or empty
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
                # Extract the base64 part (between header and footer)
                key_parts = private_key.split('-----')
                if len(key_parts) < 3:
                    return False, "Invalid private key structure"
                
                base64_part = key_parts[2].strip()
                base64.b64decode(base64_part)
            except Exception as e:
                return False, f"Invalid base64 encoding in private key: {str(e)}"

            return True, private_key
        except Exception as e:
            return False, f"Error validating private key: {str(e)}"

    def _get_service_account_info(self) -> Tuple[bool, dict, str]:
        """Get and validate service account information from environment variables."""
        try:
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

            # Check for missing environment variables
            missing_vars = [var for var in required_vars if not os.environ.get(var)]
            if missing_vars:
                return False, {}, f"Missing required environment variables: {', '.join(missing_vars)}"

            # Validate private key
            private_key = os.environ.get("GOOGLE_PRIVATE_KEY", "")
            is_valid_key, key_result = self._validate_private_key(private_key)
            
            if not is_valid_key:
                return False, {}, f"Private key validation failed: {key_result}"

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

            return True, service_account_info, "Service account info validated successfully"
        except Exception as e:
            return False, {}, f"Error getting service account info: {str(e)}"

    def authenticate(self) -> Tuple[bool, str]:
        """Authenticate using service account credentials from environment variables."""
        try:
            logger.info("Starting Google Drive authentication")
            
            # Get and validate service account info
            success, service_account_info, message = self._get_service_account_info()
            if not success:
                logger.error(f"Service account validation failed: {message}")
                return False, message

            logger.info("Service account info validated successfully")

            # Create credentials from service account info
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

            # Initialize the Drive API service
            try:
                self.drive_service = build('drive', 'v3', credentials=self.credentials)
                logger.info("Successfully initialized Drive API service")
            except Exception as e:
                error_msg = f"Failed to initialize Drive service: {str(e)}"
                logger.error(error_msg)
                return False, error_msg
            
            # Create or get root folder
            self.root_folder_id = self._get_or_create_folder(self.root_folder_name)
            if not self.root_folder_id:
                error_msg = "Failed to create or get root folder"
                logger.error(error_msg)
                return False, error_msg
            
            logger.info(f"Root folder ID: {self.root_folder_id}")
            return True, "Authentication successful"

        except Exception as e:
            error_msg = f"Authentication failed: {str(e)}"
            logger.error(error_msg)
            return False, error_msg

    def _get_or_create_folder(self, folder_name: str, parent_id: Optional[str] = None) -> Optional[str]:
        """Get or create a folder in Google Drive."""
        try:
            logger.info(f"Getting or creating folder: {folder_name}")
            # Search for existing folder
            query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}'"
            if parent_id:
                query += f" and '{parent_id}' in parents"
            logger.info(f"Search query: {query}")
            
            results = self.drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)'
            ).execute()

            files = results.get('files', [])
            logger.info(f"Found {len(files)} matching folders")
            
            # Return existing folder ID if found
            if files:
                folder_id = files[0]['id']
                logger.info(f"Using existing folder with ID: {folder_id}")
                if not self._share_with_user(folder_id):
                    logger.warning(f"Failed to share existing folder {folder_id}, but continuing anyway")
                return folder_id
            
            # Create new folder if not found
            logger.info("Creating new folder")
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            if parent_id:
                folder_metadata['parents'] = [parent_id]
            
            folder = self.drive_service.files().create(
                body=folder_metadata,
                fields='id'
            ).execute()
            
            folder_id = folder.get('id')
            if folder_id:
                logger.info(f"Created new folder with ID: {folder_id}")
                if not self._share_with_user(folder_id):
                    logger.warning(f"Failed to share new folder {folder_id}, but continuing anyway")
                return folder_id

            logger.error("Failed to get folder ID after creation")
            return None

        except Exception as e:
            logger.error(f"Error in _get_or_create_folder: {str(e)}")
            return None

    def _share_with_user(self, file_id: str) -> bool:
        """Share a file or folder with the user with writer access."""
        try:
            logger.info(f"Sharing file/folder {file_id} with {self.user_email}")
            permission = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': self.user_email
            }
            
            self.drive_service.permissions().create(
                fileId=file_id,
                body=permission,
                sendNotificationEmail=True
            ).execute()
            logger.info("Permission created successfully")
            return True

        except Exception as e:
            logger.error(f"Error in _share_with_user: {str(e)}")
            return False

    def upload_file(self, file_path: str, file_name: str) -> Optional[str]:
        """Upload a file to Google Drive and share it with the user."""
        try:
            logger.info(f"Starting file upload: {file_name}")
            if not self.drive_service or not self.root_folder_id:
                logger.error("Drive service or root folder ID not initialized")
                st.error("Google Drive service not properly initialized")
                return None

            # Create date folder
            today = datetime.now().strftime("%Y-%m-%d")
            logger.info(f"Creating/getting date folder: {today}")
            date_folder_id = self._get_or_create_folder(today, self.root_folder_id)
            
            if not date_folder_id:
                error_msg = "Failed to create or get date folder"
                logger.error(error_msg)
                st.error(error_msg)
                return None

            # Prepare file metadata
            file_metadata = {
                'name': file_name,
                'parents': [date_folder_id]
            }
            logger.info(f"File metadata prepared: {file_metadata}")
            
            # Upload file
            logger.info("Starting file upload to Drive")
            with open(file_path, 'rb') as f:
                media = MediaIoBaseUpload(
                    io.BytesIO(f.read()),
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    resumable=True
                )
                
                file = self.drive_service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id'
                ).execute()
                
                file_id = file.get('id')
                if file_id:
                    logger.info(f"File uploaded successfully with ID: {file_id}")
                    if not self._share_with_user(file_id):
                        logger.warning(f"Failed to share uploaded file {file_id}, but continuing anyway")
                    st.success('✅ File uploaded successfully to Google Drive')
                    return file_id

                logger.error("Failed to get file ID after upload")
                st.error("Failed to upload file to Google Drive")
                return None

        except Exception as e:
            error_msg = f"Error uploading file: {str(e)}"
            logger.error(error_msg)
            st.error(error_msg)
            return None
