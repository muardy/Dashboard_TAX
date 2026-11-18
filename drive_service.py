import os
import io
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

def get_drive_service():
    # 1. Prioritas Production (Render): baca langsung dari Environment Variable GOOGLE_TOKEN_JSON
    if os.getenv("GOOGLE_TOKEN_JSON"):
        info = json.loads(os.getenv("GOOGLE_TOKEN_JSON"))
        creds = Credentials.from_authorized_user_info(info)
    else:
        # 2. Fallback Lokal: baca dari file token.json (atau token_new.json)
        token_path = os.getenv("GOOGLE_TOKEN_FILE", "token.json")
        if not os.path.exists(token_path):
            raise FileNotFoundError(f"File {token_path} tidak ditemukan dan GOOGLE_TOKEN_JSON belum diset!")
        creds = Credentials.from_authorized_user_file(token_path)

    # Auto-refresh token jika masa aktif 1 jamnya habis
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        # Hanya simpan ulang ke file fisik jika sedang berjalan di lokal
        if not os.getenv("GOOGLE_TOKEN_JSON"):
            token_path = os.getenv("GOOGLE_TOKEN_FILE", "token.json")
            with open(token_path, "w") as token_file:
                token_file.write(creds.to_json())

    return build("drive", "v3", credentials=creds, cache_discovery=False)

def get_or_create_client_folder(service, client_name, parent_folder_id):
    # Rapikan nama klien & hindari error tanda kutip tunggal pada query Drive API
    clean_name = client_name.strip()
    escaped_name = clean_name.replace("'", "\\'")

    query = (
        f"name = '{escaped_name}' and "
        f"'{parent_folder_id}' in parents and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )

    results = service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name)'
    ).execute()

    folders = results.get('files', [])

    # Jika folder klien sudah ada, langsung pakai ID-nya
    if folders:
        return folders[0]['id']

    # Jika belum ada, buat sub-folder baru di dalam folder induk
    folder_metadata = {
        'name': clean_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    folder = service.files().create(
        body=folder_metadata,
        fields='id'
    ).execute()

    return folder.get('id')

def upload_file_to_drive(file_storage, filename, client_name, parent_folder_id):
    service = get_drive_service()

    # 1. Dapatkan atau buat sub-folder sesuai nama klien
    client_folder_id = get_or_create_client_folder(service, client_name, parent_folder_id)

    # 2. Baca stream file ke memori untuk di-upload
    file_bytes = io.BytesIO(file_storage.read())
    media = MediaIoBaseUpload(
        file_bytes,
        mimetype=file_storage.mimetype or 'application/octet-stream',
        resumable=True
    )

    file_metadata = {
        'name': filename,
        'parents': [client_folder_id]
    }

    uploaded_file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink'
    ).execute()

    return {
        'file_id': uploaded_file.get('id'),
        'web_link': uploaded_file.get('webViewLink')
    }