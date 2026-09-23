"""
Gộp 3 file CSV (balance_sheet, income_statement, cash_flow) của mỗi mã trong
thư mục financials/<MÃ>/ thành 1 file Excel financials/<MÃ>/<MÃ>.xlsx
với 3 sheet cùng tên.

Cách gộp dữ liệu qua thời gian (để không mất các năm cũ khi vnstock chỉ trả
về vài năm gần nhất mỗi lần fetch):
    - Dữ liệu mới được đối chiếu với dữ liệu cũ đã có trong file Excel theo
      cột "item" (tên khoản mục).
    - Nếu "item" TRÙNG KHỚP với 1 dòng đã có -> merge các cột năm mới vào
      dòng đó (đè giá trị nếu năm đã tồn tại, thêm cột nếu là năm mới) ->
      qua nhiều lần chạy sẽ tích luỹ dữ liệu nhiều năm.
    - Nếu "item" KHÔNG TRÙNG KHỚP (khoản mục hoàn toàn mới, chưa từng có
      trong dữ liệu cũ) -> KHÔNG merge dòng đó, dữ liệu cũ giữ nguyên.
    - Nếu file Excel chưa tồn tại (lần chạy đầu tiên) -> dùng toàn bộ CSV
      hiện tại làm dữ liệu nền ban đầu.

Nếu file Excel đã có thêm sheet khác (vd "chi_so_tai_chinh" tự thêm sau
này, có thể chứa công thức link tới workbook khác), sheet đó được giữ
nguyên 100%, không bị đụng tới: cách làm là MỞ THẲNG workbook cũ bằng
openpyxl rồi chỉ xoá/ghi lại đúng 3 sheet báo cáo TRONG CHÍNH workbook đó,
không tạo workbook mới rồi chép nội dung các sheet khác sang (cách chép
thủ công từng phần dễ bỏ sót các thứ nằm ở cấp WORKBOOK chứ không phải cấp
sheet - ví dụ external links khi 1 công thức tham chiếu sang file Excel
khác - làm Excel báo lỗi "unreadable content" dù openpyxl vẫn đọc được).

Hỗ trợ cả 2 kỳ báo cáo, chạy độc lập, ra 2 file Excel riêng biệt cho mỗi mã:
    - period=year (mặc định)  -> đọc CSV "<MÃ>_<report>.csv"      -> ghi "<MÃ>.xlsx"
    - period=quarter          -> đọc CSV "<MÃ>_Q_<report>.csv"    -> ghi "<MÃ>_Q.xlsx"
Cách merge dữ liệu qua thời gian (theo cột "item") áp dụng như nhau cho cả
2 kỳ; với period=quarter, tên cột kỳ có dạng "2024-Q1", "2024-Q2"... (sắp
xếp tăng dần đúng thứ tự thời gian nhờ định dạng chuỗi cố định độ dài).

Chạy:
    python merge_financials.py                          # gộp tất cả mã, period=year
    python merge_financials.py HPG                       # 1 mã, period=year
    python merge_financials.py HPG,TCB,FPT,PNJ            # nhiều mã, period=year
    python merge_financials.py HPG,TCB,FPT,PNJ quarter     # nhiều mã, period=quarter
"""

import sys
import os
import tempfile
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

FINANCIALS_DIR = "financials"

FONT_NAME = "Tahoma"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Tên sheet <-> hậu tố tên file CSV, theo đúng thứ tự hiển thị trong Excel.
STATEMENT_SHEETS = {
    "balance_sheet": "balance_sheet",
    "income_statement": "income_statement",
    "cash_flow": "cash_flow",
}

DROP_COLUMNS = ["item_en", "item_id"]

# period="year" -> hậu tố rỗng (giữ nguyên hành vi cũ: "<MÃ>_<report>.csv" -> "<MÃ>.xlsx").
# period="quarter" -> hậu tố "_Q" ("<MÃ>_Q_<report>.csv" -> "<MÃ>_Q.xlsx").
PERIOD_FILE_SUFFIX = {"year": "", "quarter": "_Q"}


def _load_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="utf-8-sig")
    drop_cols = [c for c in DROP_COLUMNS if c in df.columns]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


def _sort_year_columns(cols):
    """Sắp xếp các cột năm tăng dần (cũ nhất bên trái, mới nhất bên phải),
    cột không phải số năm (nếu có) được đẩy xuống cuối, giữ nguyên thứ tự gốc."""

    def key(c):
        try:
            return (0, int(c))
        except (TypeError, ValueError):
            return (1, str(c))

    return sorted(cols, key=key)


def merge_sheet(old_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    """Merge new_df vào old_df theo cột 'item'. Xem quy tắc merge ở đầu file."""
    if old_df is None or old_df.empty:
        # Chưa có dữ liệu cũ -> dùng dữ liệu mới làm nền.
        merged = new_df.copy()
    else:
        merged = old_df.copy()
        year_cols_new = [c for c in new_df.columns if c != "item"]

        # Đảm bảo các cột năm mới tồn tại trong merged (dù chưa có dòng nào khớp).
        for yc in year_cols_new:
            if yc not in merged.columns:
                merged[yc] = pd.NA

        item_to_index = {}
        for idx, item_val in merged["item"].items():
            item_to_index.setdefault(item_val, idx)

        for _, row in new_df.iterrows():
            item_val = row["item"]
            if item_val not in item_to_index:
                # Không trùng khớp -> bỏ qua, không merge, giữ nguyên dữ liệu cũ.
                continue
            old_idx = item_to_index[item_val]
            for yc in year_cols_new:
                merged.loc[old_idx, yc] = row[yc]

    year_cols = _sort_year_columns([c for c in merged.columns if c != "item"])
    merged = merged[["item"] + year_cols]
    return merged


def _write_dataframe(ws, df: pd.DataFrame) -> None:
    """Ghi DataFrame vào worksheet openpyxl trống (tương đương df.to_excel
    nhưng thao tác trực tiếp trên 1 sheet của workbook đang mở, không qua
    pandas ExcelWriter/workbook riêng). NaN/NA -> ô trống (giống to_excel)."""
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in row])


def _autofit_number_column_width(ws, col_idx, min_width=10, max_width=24, padding=2):
    """Tính độ rộng vừa đủ cho 1 cột SỐ (năm) dựa trên nội dung thực tế
    (số đã format dấu phân cách nghìn), để không cần mở file chỉnh tay."""
    max_len = 4  # tối thiểu bằng độ dài tiêu đề năm, vd "2022"
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
        for cell in row:
            if cell.value is None:
                continue
            if isinstance(cell.value, (int, float)):
                text = f"{cell.value:,.0f}"
            else:
                text = str(cell.value)
            max_len = max(max_len, len(text))
    return max(min_width, min(max_width, max_len + padding))


def _format_statement_sheet(ws) -> None:
    """Định dạng chuẩn cho 1 sheet báo cáo (balance_sheet/income_statement/
    cash_flow) vừa được ghi: font Tahoma, dòng tiêu đề (năm) in đậm/nền
    màu, số có dấu phân cách nghìn, cố định dòng 1 + cột A, viền mảnh cho
    toàn vùng dữ liệu.
    Độ rộng cột: cột A (khoản mục) giữ cố định vì tên khoản mục dài, không
    autofit theo nội dung (sẽ ra cột quá rộng); các cột năm (số) thì autofit
    theo nội dung thực tế để hiển thị đủ luôn, không cần mở file chỉnh tay."""
    if ws.max_row < 1 or ws.max_column < 1:
        return

    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    for col_idx in range(2, ws.max_column + 1):
        width = _autofit_number_column_width(ws, col_idx)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = BORDER
            if cell.row == 1:
                cell.font = Font(name=FONT_NAME, bold=True, color="FFFFFF")
                cell.fill = HEADER_FILL
                cell.alignment = Alignment(horizontal="center" if cell.column > 1 else "left")
            else:
                cell.font = Font(name=FONT_NAME, size=10)
                if cell.column > 1:
                    cell.number_format = "#,##0;(#,##0);\"-\""

    ws.freeze_panes = "B2"


def process_symbol(symbol: str, file_suffix: str = "") -> bool:
    sym_dir = os.path.join(FINANCIALS_DIR, symbol)
    out_path = os.path.join(sym_dir, f"{symbol}{file_suffix}.xlsx")

    if not os.path.isdir(sym_dir):
        print(f"  Bỏ qua {symbol}: không tìm thấy thư mục {sym_dir}")
        return False

    # Chỉ đọc bằng pandas 3 sheet báo cáo gốc (dữ liệu thuần, không công
    # thức) để lấy dữ liệu NỀN cho merge_sheet(). Việc GHI thì làm trực
    # tiếp trên workbook openpyxl mở từ out_path (xem bên dưới) chứ không
    # dùng pandas ExcelWriter, để các sheet/thứ khác (external links, named
    # ranges, comment, v.v.) không hề bị workbook mới "quên" mất.
    existing_sheets = {}
    if os.path.exists(out_path):
        try:
            existing_sheets = pd.read_excel(
                out_path, sheet_name=list(STATEMENT_SHEETS.keys()), engine="openpyxl")
        except Exception as e:
            print(f"  CẢNH BÁO: không đọc được file Excel cũ {out_path} ({e}) -> tạo mới.")
            existing_sheets = {}

    tmp_path = None
    try:
        output_sheets = {}
        any_written = False

        for sheet_name, suffix in STATEMENT_SHEETS.items():
            csv_path = os.path.join(sym_dir, f"{symbol}{file_suffix}_{suffix}.csv")
            if not os.path.exists(csv_path):
                print(f"  Bỏ qua sheet '{sheet_name}': không tìm thấy {csv_path}")
                continue

            new_df = _load_csv(csv_path)
            old_df = existing_sheets.get(sheet_name)
            output_sheets[sheet_name] = merge_sheet(old_df, new_df)
            any_written = True

        if not any_written:
            print(f"  Không có CSV nào cho {symbol}, bỏ qua.")
            return False

        ordered_names = [n for n in STATEMENT_SHEETS if n in output_sheets]

        # Mở NGUYÊN workbook cũ (nếu có) bằng openpyxl - giữ nguyên 100%
        # mọi thứ (sheet khác, công thức, external links, named ranges,
        # conditional formatting, comment...) vì chưa hề đụng tới. Nếu chưa
        # có file cũ hoặc file cũ đọc lỗi -> tạo workbook mới trống.
        wb = None
        if os.path.exists(out_path):
            try:
                wb = openpyxl.load_workbook(out_path)
            except Exception as e:
                print(f"  CẢNH BÁO: không mở được workbook cũ {out_path} bằng openpyxl "
                      f"({e}) -> tạo workbook mới (CÁC SHEET TỰ TẠO KHÁC SẼ MẤT).")
                wb = None
        is_new_wb = wb is None
        if wb is None:
            wb = openpyxl.Workbook()

        # Ghi lại đúng 3 sheet báo cáo: nếu đã tồn tại -> xoá rồi tạo lại
        # ĐÚNG VỊ TRÍ CŨ (giữ nguyên thứ tự sheet); nếu chưa có -> thêm mới
        # theo đúng thứ tự STATEMENT_SHEETS.
        for name in ordered_names:
            if name in wb.sheetnames:
                idx = wb.sheetnames.index(name)
                del wb[name]
                ws = wb.create_sheet(name, idx)
            else:
                ws = wb.create_sheet(name)
            _write_dataframe(ws, output_sheets[name])
            _format_statement_sheet(ws)

        # Xoá sheet mặc định "Sheet" nếu là workbook mới tạo trống.
        if is_new_wb and "Sheet" in wb.sheetnames and "Sheet" not in ordered_names:
            del wb["Sheet"]

        # Đảm bảo 3 sheet báo cáo luôn đứng đầu, đúng thứ tự cố định; các
        # sheet khác (tự tạo, vd "chi_so_tai_chinh") theo sau, giữ nguyên
        # thứ tự tương đối cũ của chúng.
        desired_order = list(ordered_names) + [n for n in wb.sheetnames if n not in ordered_names]
        wb._sheets = [wb[n] for n in desired_order]

        # Ghi an toàn (atomic write): lưu ra file tạm CÙNG THƯ MỤC với
        # out_path rồi mới os.replace ở bước cuối (atomic, cùng
        # filesystem). Nếu có lỗi giữa chừng, out_path thật không hề bị
        # đụng tới -> không bao giờ bị hỏng/dở dang.
        tmp_fd, tmp_path = tempfile.mkstemp(
            suffix=".xlsx", prefix=f".{symbol}{file_suffix}_", dir=sym_dir)
        os.close(tmp_fd)
        wb.save(tmp_path)
        os.replace(tmp_path, out_path)
        tmp_path = None  # đã đổi tên thành out_path, không còn để dọn dẹp
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)

    print(f"  Đã ghi {out_path}  (sheets: {', '.join(desired_order)})")
    return True


def main():
    if len(sys.argv) > 1:
        symbols = [s.strip().upper() for s in sys.argv[1].split(",") if s.strip()]
    else:
        # Mặc định: gộp tất cả các mã đang có sẵn trong financials/
        if os.path.isdir(FINANCIALS_DIR):
            symbols = sorted(
                d for d in os.listdir(FINANCIALS_DIR)
                if os.path.isdir(os.path.join(FINANCIALS_DIR, d))
            )
        else:
            symbols = []

    if not symbols:
        print(f"Không tìm thấy mã nào trong '{FINANCIALS_DIR}/'.")
        sys.exit(1)

    period = sys.argv[2].strip().lower() if len(sys.argv) > 2 else "year"
    if period not in PERIOD_FILE_SUFFIX:
        print(f"Tham số period không hợp lệ: '{period}'. Chỉ chấp nhận 'year' hoặc 'quarter'.")
        sys.exit(1)
    file_suffix = PERIOD_FILE_SUFFIX[period]

    any_success = False
    failed_symbols = []
    for symbol in symbols:
        print(f"\n=== Mã {symbol} ({period}) ===")
        try:
            if process_symbol(symbol, file_suffix):
                any_success = True
        except Exception as e:
            # Cô lập lỗi theo từng mã: 1 mã lỗi (CSV bất thường, dữ liệu
            # thiếu cột "item"...) không được làm dừng cả script hay ảnh
            # hưởng tới file .xlsx của các mã khác/đã ghi trước đó.
            print(f"  LỖI khi xử lý {symbol}: {e} -> bỏ qua mã này, giữ nguyên file cũ.")
            failed_symbols.append(symbol)

    if failed_symbols:
        print(f"\nCác mã lỗi, đã bỏ qua: {', '.join(failed_symbols)}")

    if not any_success:
        print("\nKhông gộp được dữ liệu cho bất kỳ mã nào.")
        sys.exit(1)

    print("\nHoàn tất.")


if __name__ == "__main__":
    main()
