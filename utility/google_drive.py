import os
from datetime import datetime
#, UTC)
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

class GoogleDrive:
    def __init__(self, service_account_file):
        SCOPES = ['https://www.googleapis.com/auth/drive.file']
        self.credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
        self._service = None

    def __getattr__(self, key):
        if key == 'service':
            if self._service is None:
                self._service = build('drive', 'v3', credentials=self.credentials)
            return self._service
        else:
            return super().__getattr__(key)

    def upload(self, local_file, upstream_file=None):
        if upstream_file is None:
            upstream_file = os.path.basename(local_file)
        file_metadata = {'name': upstream_file}
        media = MediaFileUpload(local_file, mimetype='application/octet-stream')
        file = self.service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        file_id = file.get('id')
        return file_id

    def change_permission(self, file_id, type, role, **kwargs):
        # type == anyone, role == reader: normal configurations
        body = {'type': type, 'role': role}
        if kwargs.get('expirationTime') is not None:
            expiration_time = (datetime.utcnow() + kwargs['expirationTime']).isoformat() + 'Z'
            body['expirationTime'] = expiration_time
        elif kwargs.get('emailAddress') is not None:
            body['emailAddress'] = kwargs['emailAddress']
        self.service.permissions().create(
            fileId=file_id,
            body=body
        ).execute()

    def upload_for_someone_read(self, local_file, upstream_file=None, expirationTime=None, emailAddress=None):
        file_id = self.upload(local_file, upstream_file)
        if emailAddress is None:
            self.change_permission(file_id, 'anyone', 'reader', expirationTime=expirationTime)
        else:
            self.change_permission(file_id, 'user', 'reader', expirationTime=expirationTime, emailAddress=emailAddress)
        download_link = self.enclose_file(file_id)
        return download_link

    def download_file(self, file_id, local_path):
        request = self.service.files().get_media(fileId=file_id)
        with open(local_path, 'wb') as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while done is False:
                _, done = downloader.next_chunk()

    def delete_file(self, file_id):
        self.service.files().delete(fileId=file_id).execute()

    def list_all_files(self):
        results = self.service.files().list(pageSize=100, fields="nextPageToken, files(id, name)").execute()
        return results.get('files', [])

    def enclose_file(self, file_id):
        return f"https://drive.google.com/uc?id={file_id}&export=download"
