# -*- coding: utf-8 -*-
"""
Hàm dùng chung để xác thực & thao tác với Google Drive API bằng Service
Account, dùng cho cả gdrive_download.py (tải file .xlsx từ Drive về trước
khi cập nhật) và gdrive_upload.py (đẩy file .xlsx đã cập nhật lên lại Drive).

Cách thiết lập (làm 1 lần):
  1. Vào https://console.cloud.google.com/ -> tạo (hoặc chọn) 1 Project.
  2. Vào "APIs & Services" -> "Library" -> bật "Google Drive API".
  3. Vào "APIs & Services" -> "Credentials" -> "Create credentials" ->
     "Service account" -> đặt tên tuỳ ý -> Create -> bỏ qua phần gán role
     (không cần) -> Done.
  4. Bấm vào Service Account vừa tạo -> tab "Keys" -> "Add Key" -> "Create
     new key" -> chọn JSON -> tải file JSON key về máy (giữ kín, không
     commit vào repo).
  5. Mở file JSON đó, copy field "client_email" (dạng
     xxx@xxx.iam.gserviceaccount.com).
  6. Vào Google Drive, mở thư mục GỐC chứa các thư mục con theo mã
     (vd FPT/, HPG/,... mỗi thư mục con chứa 1 file <MÃ>.xlsx)
     -> Share -> dán email ở bước 5 vào -> chọn quyền "Editor" (bắt buộc
     phải là Editor nếu muốn upload/ghi đè, không được để Viewer) -> Share.
     Lưu ý: chỉ cần share thư mục GỐC, các thư mục con bên trong sẽ tự
     động được kế thừa quyền, không cần share riêng từng thư mục con.
  7. Copy ID của thư mục Drive GỐC đó (phần cuối URL khi mở thư mục:
     https://drive.google.com/drive/folders/<FOLDER_ID>).
  8. Trong GitHub repo -> Settings -> Secrets and variables -> Actions ->
     New repository secret:
       - GDRIVE_SA_KEY   = dán TOÀN BỘ nội dung file JSON key ở bước 4
       - GDRIVE_FOLDER_ID = ID thư mục ở bước 7
"""

import io
import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive"]
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_FOLDER = "application/vnd.google-apps.folder"


def get_drive_service():
    """Tạo Google Drive API client từ Service Account key lưu trong biến môi
    trường GDRIVE_SA_KEY (nội dung JSON, không phải đường dẫn file)."""
    raw_key = os.getenv("GDRIVE_SA_KEY", "").strip()
    if not raw_key:
        raise SystemExit(
            "Thiếu biến môi trường GDRIVE_SA_KEY (nội dung JSON key của "
            "Service Account). Xem hướng dẫn thiết lập ở đầu file gdrive_utils.py."
        )
    try:
        info = json.loads(raw_key)
    except json.JSONDecodeError as e:
        raise SystemExit(f"GDRIVE_SA_KEY không phải JSON hợp lệ: {e}")

    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def get_folder_id():
    folder_id = os.getenv("GDRIVE_FOLDER_ID", "").strip()
    if not folder_id:
        raise SystemExit("Thiếu biến môi trường GDRIVE_FOLDER_ID (ID thư mục Google Drive).")
    return folder_id


def list_xlsx_files(service, folder_id):
    """Trả về dict {tên_file: file_id} cho mọi file .xlsx nằm trực tiếp
    trong thư mục Drive `folder_id` (không đệ quy vào thư mục con).
    Dùng cho 1 thư mục con theo mã (vd thư mục "HPG/"), KHÔNG dùng trực
    tiếp cho thư mục gốc vì các file giờ nằm trong thư mục con theo mã."""
    files = {}
    page_token = None
    query = (
        f"'{folder_id}' in parents and trashed = false and "
        f"(mimeType = '{MIME_XLSX}' or name contains '.xlsx')"
    )
    while True:
        resp = service.files().list(
            q=query,
            spaces="drive",
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
        ).execute()
        for f in resp.get("files", []):
            files[f["name"]] = f["id"]
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def list_subfolders(service, parent_folder_id):
    """Trả về dict {tên_thư_mục_con: folder_id} cho mọi thư mục con nằm
    trực tiếp trong `parent_folder_id`. Ứng với cấu trúc Drive:
    <thư mục gốc>/<MÃ>/<MÃ>.xlsx -> mỗi thư mục con ở đây
    tương ứng 1 mã chứng khoán (vd "FPT", "HPG")."""
    folders = {}
    page_token = None
    query = (
        f"'{parent_folder_id}' in parents and trashed = false and "
        f"mimeType = '{MIME_FOLDER}'"
    )
    while True:
        resp = service.files().list(
            q=query,
            spaces="drive",
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
        ).execute()
        for f in resp.get("files", []):
            folders[f["name"]] = f["id"]
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return folders


def get_or_create_subfolder(service, parent_folder_id, name, existing_subfolders=None):
    """Trả về folder_id của thư mục con tên `name` nằm trong
    `parent_folder_id`. Nếu chưa tồn tại -> tạo mới. `existing_subfolders`
    (dict trả về từ list_subfolders) có thể truyền vào để tránh gọi lại
    API nhiều lần khi xử lý nhiều mã liên tiếp."""
    if existing_subfolders is not None and name in existing_subfolders:
        return existing_subfolders[name]

    query = (
        f"'{parent_folder_id}' in parents and trashed = false and "
        f"mimeType = '{MIME_FOLDER}' and name = '{name}'"
    )
    resp = service.files().list(
        q=query, spaces="drive", fields="files(id, name)"
    ).execute()
    found = resp.get("files", [])
    if found:
        folder_id = found[0]["id"]
    else:
        meta = {
            "name": name,
            "mimeType": MIME_FOLDER,
            "parents": [parent_folder_id],
        }
        created = service.files().create(body=meta, fields="id").execute()
        folder_id = created["id"]

    if existing_subfolders is not None:
        existing_subfolders[name] = folder_id
    return folder_id


def download_file(service, file_id, dest_path):
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    with open(dest_path, "wb") as f:
        f.write(buf.getvalue())


def upload_or_replace_file(service, folder_id, existing_files, local_path, drive_name):
    """Nếu drive_name đã tồn tại trong existing_files -> ghi đè nội dung
    (giữ nguyên file_id, không tạo file trùng). Nếu chưa có -> tạo mới
    trong đúng thư mục folder_id."""
    media = MediaFileUpload(local_path, mimetype=MIME_XLSX, resumable=True)
    if drive_name in existing_files:
        file_id = existing_files[drive_name]
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id, "updated"
    else:
        meta = {"name": drive_name, "parents": [folder_id]}
        created = service.files().create(body=meta, media_body=media, fields="id").execute()
        return created["id"], "created"
