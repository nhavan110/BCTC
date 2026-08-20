# -*- coding: utf-8 -*-
"""
Tải các file <MÃ>.xlsx từ thư mục Google Drive (GDRIVE_FOLDER_ID)
về đúng vị trí financials/<MÃ>/<MÃ>.xlsx, để dùng làm bản NỀN cho
merge_financials.py (giữ nguyên sheet do bạn tự tạo, vd "chi_so_tai_chinh",
và giữ lịch sử dữ liệu nhiều năm đã có).

Cấu trúc thư mục trên Google Drive (GDRIVE_FOLDER_ID là thư mục GỐC):
    <thư mục gốc>/
        FPT/
            FPT.xlsx
        HPG/
            HPG.xlsx
        ...
Mỗi mã có 1 thư mục con cùng tên, bên trong chứa đúng 1 file
<MÃ>.xlsx. Script này đệ quy vào từng thư mục con để tìm file
tương ứng với mã đó.

Chạy TRƯỚC fetch_full_financials.py và merge_financials.py.

Chạy:
    python gdrive_download.py                    # tải tất cả mã có thư mục con trên Drive
    python gdrive_download.py HPG,TCB,FPT          # chỉ tải các mã chỉ định (nếu có trên Drive)
"""

import sys
import os

from gdrive_utils import (
    get_drive_service,
    get_folder_id,
    list_subfolders,
    list_xlsx_files,
    download_file,
)

FINANCIALS_DIR = "financials"


def main():
    wanted_symbols = None
    if len(sys.argv) > 1:
        wanted_symbols = {s.strip().upper() for s in sys.argv[1].split(",") if s.strip()}

    service = get_drive_service()
    folder_id = get_folder_id()
    symbol_folders = list_subfolders(service, folder_id)

    if not symbol_folders:
        print("Không tìm thấy thư mục con nào (theo mã) trong thư mục Drive gốc. Bỏ qua bước tải xuống.")
        return

    downloaded = 0
    for symbol, sub_folder_id in sorted(symbol_folders.items()):
        symbol_upper = symbol.upper()
        if wanted_symbols and symbol_upper not in wanted_symbols:
            continue

        name = f"{symbol_upper}.xlsx"
        files_in_subfolder = list_xlsx_files(service, sub_folder_id)

        file_id = files_in_subfolder.get(name)
        if file_id is None:
            # Dự phòng: thư mục con có thể chỉ chứa đúng 1 file .xlsx
            # nhưng đặt tên không khớp tuyệt đối "<MÃ>.xlsx".
            if len(files_in_subfolder) == 1:
                only_name, file_id = next(iter(files_in_subfolder.items()))
                print(f"  CẢNH BÁO: thư mục '{symbol}' không có file '{name}', dùng tạm '{only_name}'.")
            else:
                print(f"  Bỏ qua '{symbol}': không tìm thấy '{name}' trong thư mục con.")
                continue

        dest_path = os.path.join(FINANCIALS_DIR, symbol_upper, name)
        print(f"  Tải {symbol}/{name} (Drive id={file_id}) -> {dest_path}")
        download_file(service, file_id, dest_path)
        downloaded += 1

    print(f"Hoàn tất: đã tải {downloaded} file từ Google Drive.")


if __name__ == "__main__":
    main()
