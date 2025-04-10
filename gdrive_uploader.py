from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Put the correct Google Drive folder ID here (from the shared folder URL)
FOLDER_ID = "1cTlWq4kOPq8PbybqKZi17f6QYz0xAoZ1"

# Path to your service account credentials
SERVICE_ACCOUNT_FILE = "/etc/secrets/gdrive-creds.json"
SCOPES = ['https://www.googleapis.com/auth/drive']

# Set up the API client
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=SCOPES
)
drive_service = build('drive', 'v3', credentials=credentials)

def upload_file_to_gdrive(local_file_path: str, filename: str):
    file_metadata = {
        'name': filename,
        'parents': [FOLDER_ID]
    }
    media = MediaFileUpload(local_file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    uploaded = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink, webContentLink'
    ).execute()

    return uploaded.get("webViewLink"), uploaded.get("webContentLink")
