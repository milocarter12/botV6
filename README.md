# Excel Data Transfer Bot

A Streamlit application that automates the process of transferring CSV data to formatted Excel templates with Google Drive integration.

## Environment Setup

### Required Environment Variables

The application requires the following environment variables to be set:

#### Google Drive API Credentials
- `GOOGLE_DRIVE_CLIENT_ID`: OAuth 2.0 Client ID
- `GOOGLE_DRIVE_CLIENT_SECRET`: OAuth 2.0 Client Secret

#### Google Service Account Credentials
- `GOOGLE_SERVICE_ACCOUNT_TYPE`: Type of service account (usually "service_account")
- `GOOGLE_PROJECT_ID`: Your Google Cloud project ID
- `GOOGLE_PRIVATE_KEY_ID`: Service account private key ID
- `GOOGLE_PRIVATE_KEY`: Service account private key (PEM format)
- `GOOGLE_CLIENT_EMAIL`: Service account email address
- `GOOGLE_CLIENT_ID`: Service account client ID
- `GOOGLE_AUTH_URI`: Authentication URI (usually "https://accounts.google.com/o/oauth2/auth")
- `GOOGLE_TOKEN_URI`: Token URI (usually "https://oauth2.googleapis.com/token")
- `GOOGLE_AUTH_PROVIDER_X509_CERT_URL`: Auth provider x509 certificate URL
- `GOOGLE_CLIENT_X509_CERT_URL`: Client x509 certificate URL

### Setting Up Google Service Account

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Google Drive API for your project
4. Go to "APIs & Services" > "Credentials"
5. Click "Create Credentials" > "Service Account"
6. Fill in the service account details and click "Create"
7. Click on the newly created service account
8. Go to the "Keys" tab
9. Click "Add Key" > "Create New Key"
10. Choose "JSON" and click "Create"
11. Save the downloaded JSON file securely
12. Extract the values from the JSON file and add them to your environment variables

### Environment Variables Setup

1. Copy `template.env` to `.env`
2. Fill in all the required values from your Google service account JSON file
3. Do not commit the `.env` file to version control

## Development Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Set up environment variables as described above
4. Run the application:
   ```bash
   streamlit run main.py
   ```

## Security Notes

- Never commit sensitive credentials to version control
- Keep your `.env` file secure and private
- Regularly rotate your service account keys
- Follow the principle of least privilege when setting up service account permissions
