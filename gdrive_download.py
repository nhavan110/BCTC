# -*- coding: utf-8 -*-
"""
Tải các file <MÃ>.xlsx (dữ liệu năm) VÀ <MÃ>_Q.xlsx (dữ liệu quý, nếu có)
từ thư mục Google Drive (GDRIVE_FOLDER_ID) về đúng vị trí
financials/<MÃ>/<MÃ>.xlsx và financials/<MÃ>/<MÃ>_Q.xlsx, để dùng làm bản
NỀN cho merge_financials.py (giữ nguyên sheet do bạn tự tạo, vd
"chi_so_tai_chinh", và giữ lịch sử dữ liệu đã có).

Cấu trúc thư mục trên Google Drive (GDRIVE_FOLDER_ID là thư mục GỐC):
    <thư mục gốc>/
        FPT/
            FPT.xlsx
        HPG/
            HPG.xlsx
        MWG/
            MWG.xlsx
            MWG_Q.xlsx        (tự tạo thêm nếu muốn theo dõi theo quý)
        ...
Mỗi mã có 1 thư mục con cùng tên; bên trong có thể chứa file <MÃ>.xlsx
(năm) và/hoặc file <MÃ>_Q.xlsx (quý). File nào không có trên Drive thì
được bỏ qua (không lỗi) — vd nếu chưa tạo <MÃ>_Q.xlsx cho mã nào thì
script chỉ tải phần dữ liệu năm cho mã đó.

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

# "" -> file năm (<MÃ>.xlsx), "_Q" -> file quý (<MÃ>_Q.xlsx).
FILE_SUFFIXES = ["", "_Q"]


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

        files_in_subfolder = list_xlsx_files(service, sub_folder_id)

        found_any = False
        for suffix in FILE_SUFFIXES:
            name = f"{symbol_upper}{suffix}.xlsx"
            file_id = files_in_subfolder.get(name)
            if file_id is None:
                continue
            found_any = True
            dest_path = os.path.join(FINANCIALS_DIR, symbol_upper, name)
            print(f"  Tải {symbol}/{name} (Drive id={file_id}) -> {dest_path}")
            download_file(service, file_id, dest_path)
            downloaded += 1

        if not found_any:
            print(f"  Bỏ qua '{symbol}': không tìm thấy file .xlsx nào trong thư mục con.")

    print(f"Hoàn tất: đã tải {downloaded} file từ Google Drive.")


if __name__ == "__main__":
    main()
