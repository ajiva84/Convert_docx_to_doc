from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import os.path

def convert_files_in_folder(folder_id):
    SCOPES = ['https://www.googleapis.com/auth/drive']

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)

    # List all files in the folder
    results = service.files().list(q=f"'{folder_id}' in parents", pageSize=100, fields="files(id, name, mimeType)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
        return

    for item in items:
        print(f"Processing file: {item['name']} ({item['id']}) with mimeType {item['mimeType']}")
        
        if item['name'].startswith('~$') or not item['name'].endswith('.docx'):
            print(f"Skipping file: {item['name']} as it's not a valid .docx file.")
            continue

        # If the file is not already a Google Docs file, convert it
        if item['mimeType'] != 'application/vnd.google-apps.document':
            file_id = item['id']
            service.files().copy(fileId=file_id, body={'mimeType': 'application/vnd.google-apps.document'}).execute()
            print(f"Converted {item['name']} to Google Docs format.")
        else:
            print(f"{item['name']} is already in Google Docs format.")

if __name__ == '__main__':
    folder_id = 'Your Google Folder ID here'
    convert_files_in_folder(folder_id)
