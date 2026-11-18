import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Kita pakai scope ini biar aplikasi punya akses nulis/upload ke Drive
SCOPES = ['https://www.googleapis.com/auth/drive']
def otentikasi_akun():
    creds = None
    
    # 1. Deteksi otomatis lokasi pasti folder script ini (Absolute Path)
    folder_script = os.path.dirname(os.path.abspath(__file__))
    token_path = os.path.join(folder_script, 'token.json')
    
    # PASTIKAN nama file JSON rahasia lo sesuai di sini (misal 'xxxxxx.json' atau 'client_secret.json')
    client_secret_path = os.path.join(folder_script, 'client_secret_285668422515-m6s07jj76342q1m8n78du4kqj09hnpnp.apps.googleusercontent.com.json') 

    # Mengecek apakah token sudah pernah dibuat sebelumnya
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    
    # Kalau belum ada atau tokennya expired, minta login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                client_secret_path, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Simpan token dengan alamat pasti, nggak bakal error lagi
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
            
    return creds


def upload_file_ke_drive(lokasi_file_lokal, folder_id_tujuan):
    if not os.path.exists(lokasi_file_lokal):
        print(f"Error: File '{lokasi_file_lokal}' nggak ketemu bro!")
        return

    creds = otentikasi_akun()
    drive_service = build('drive', 'v3', credentials=creds)

    file_metadata = {
        'name': 'Upload_Pertama_Portal_OAuth.pdf',
        'parents': [folder_id_tujuan]
    }
    
    media = MediaFileUpload(lokasi_file_lokal, resumable=True)
    
    print(f"Lagi proses upload '{lokasi_file_lokal}' ke Drive...")
    file_terupload = drive_service.files().create(
        body=file_metadata, 
        media_body=media, 
        fields='id'
    ).execute()
    
    print(f"Sukses Bro! File ID di Google Drive: {file_terupload.get('id')}")

if __name__ == '__main__':
    # GANTI STRING DI BAWAH PAKE FOLDER ID ASLI LO
    FOLDER_ID_KLIEN = '1HQdWx6dQwjxj20novAoBE9jQgsOEEvyX' 
    
    # Lokasi file lokal lo
    FILE_LOKAL_YANG_MAU_DIUPLOAD = r'C:\Users\ASUS\Downloads\Muardy_Muhammad_Junius_new_2024 (3).pdf'
    
    upload_file_ke_drive(FILE_LOKAL_YANG_MAU_DIUPLOAD, FOLDER_ID_KLIEN)