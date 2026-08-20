# -*- coding: utf-8 -*-
"""
Đẩy các file financials/<MÃ>/<MÃ>.xlsx (đã được
merge_financials.py cập nhật 3 sheet báo cáo, giữ nguyên sheet chỉ số tự
tạo) lên lại thư mục Google Drive (GDRIVE_FOLDER_ID). Nếu file cùng tên đã
tồn tại trên Drive -> ghi đè nội dung (giữ nguyên file, giữ nguyên link
chia sẻ cũ). Nếu chưa có -> tạo mới.

Cấu trúc thư mục trên Google Drive (GDRIVE_FOLDER_ID là thư mục GỐC):
    <thư mục gốc>/
        FPT/
            FPT.xlsx
        HPG/
            HPG.xlsx
        ...
Với mỗi mã, script tìm (hoặc tạo mới nếu chưa có) thư mục con cùng tên mã
ngay trong thư mục gốc, rồi upload/ghi đè file <MÃ>.xlsx vào
đúng thư mục con đó.

Chạy SAU merge_financials.py.

Chạy:
    python gdrive_upload.py                    # upload tất cả mã có file .xlsx trong financials/
    python gdrive_upload.py HPG,TCB,FPT          # chỉ upload các mã chỉ định
"""

import sys
import os

from gdrive_utils import (
    get_drive_service,
    get_folder_id,
    list_subfolders,
    list_xlsx_files,
    get_or_create_subfolder,
    upload_or_replace_file,
)

FINANCIALS_DIR = "financials"


def main():
    if len(sys.argv) > 1:
        symbols = [s.strip().upper() for s in sys.argv[1].split(",") if s.strip()]
    else:
        symbols = sorted(
            d for d in os.listdir(FINANCIALS_DIR)
            if os.path.isdir(os.path.join(FINANCIALS_DIR, d))
        ) if os.path.isdir(FINANCIALS_DIR) else []

    if not symbols:
        print("Không có mã nào để upload.")
        return

    service = get_drive_service()
    folder_id = get_folder_id()
    symbol_folders = list_subfolders(service, folder_id)

    uploaded = 0
    for symbol in symbols:
        name = f"{symbol}.xlsx"
        local_path = os.path.join(FINANCIALS_DIR, symbol, name)
        if not os.path.exists(local_path):
            print(f"  Bỏ qua {symbol}: không thấy {local_path}")
            continue

        sub_folder_id = get_or_create_subfolder(
            service, folder_id, symbol, existing_subfolders=symbol_folders
        )
        existing_files = list_xlsx_files(service, sub_folder_id)

        file_id, action = upload_or_replace_file(
            service, sub_folder_id, existing_files, local_path, name
        )
        print(f"  {action.upper()} {symbol}/{name} (Drive id={file_id})")
        uploaded += 1

    print(f"Hoàn tất: đã upload {uploaded} file lên Google Drive.")


if __name__ == "__main__":
    main()
