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

Nếu file Excel đã có thêm sheet khác (vd "financial_ratios" tự thêm sau
này), sheet đó được giữ nguyên, không bị đụng tới.

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
import copy
import shutil
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
    cash_flow) vừa được pandas ghi ra: font Tahoma, dòng tiêu đề (năm) in
    đậm/nền màu, số có dấu phân cách nghìn, cố định dòng 1 + cột A, viền
    mảnh cho toàn vùng dữ liệu.
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

    # Chỉ đọc bằng pandas 3 sheet báo cáo gốc (dữ liệu thuần, không công thức).
    # KHÔNG dùng pd.read_excel cho toàn bộ workbook: pandas chỉ đọc được giá
    # trị đã tính sẵn (cached value) của ô công thức chứ không đọc được công
    # thức, nên nếu nạp rồi ghi lại các sheet khác (vd "chi_so_tai_chinh")
    # qua pandas thì mọi công thức trong đó sẽ bị "phẳng hoá" thành số tĩnh.
    # Các sheet ngoài 3 sheet báo cáo được giữ nguyên 100% (kể cả công thức,
    # định dạng) bằng cách copy trực tiếp qua openpyxl ở bước bên dưới.
    existing_sheets = {}
    other_sheet_names = []
    if os.path.exists(out_path):
        try:
            existing_sheets = pd.read_excel(
                out_path, sheet_name=list(STATEMENT_SHEETS.keys()), engine="openpyxl")
        except Exception as e:
            print(f"  CẢNH BÁO: không đọc được file Excel cũ {out_path} ({e}) -> tạo mới.")
            existing_sheets = {}
        try:
            other_sheet_names = [
                s for s in openpyxl.load_workbook(out_path, read_only=True).sheetnames
                if s not in STATEMENT_SHEETS
            ]
        except Exception:
            other_sheet_names = []

    backup_path = out_path + ".bak_other_sheets.xlsx"
    tmp_path = None

    # Toàn bộ phần còn lại (đọc CSV, merge, ghi file) nằm trong try/finally
    # để bảo đảm 2 file phụ (backup_path, tmp_path) LUÔN được dọn dẹp dù lỗi
    # xảy ra ở bất kỳ bước nào (kể cả trước khi tmp_path được tạo) -> không
    # để sót file rác, và out_path thật không bao giờ bị đụng tới cho tới
    # khi mọi bước phía trên đã thành công.
    try:
        if other_sheet_names and os.path.exists(out_path):
            shutil.copyfile(out_path, backup_path)

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

        # Ghi an toàn (atomic write): ghi ra file tạm CÙNG THƯ MỤC với
        # out_path (đảm bảo os.replace ở cuối là atomic, cùng filesystem),
        # chỉ thay thế file thật sau khi mọi bước (ghi 3 sheet + copy sheet
        # khác) đã thành công. Nếu có lỗi giữa chừng (dữ liệu bất thường,
        # tiến trình bị ngắt...) file thật out_path không hề bị đụng tới ->
        # không bao giờ bị hỏng/dở dang.
        tmp_fd, tmp_path = tempfile.mkstemp(
            suffix=".xlsx", prefix=f".{symbol}{file_suffix}_", dir=sym_dir)
        os.close(tmp_fd)

        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            for name in ordered_names:
                output_sheets[name].to_excel(writer, sheet_name=name, index=False)
                _format_statement_sheet(writer.sheets[name])

        # Copy nguyên trạng (công thức + định dạng) các sheet khác từ file cũ
        # (vd "chi_so_tai_chinh") sang file TẠM vừa ghi ở trên (không đụng
        # tới out_path thật cho tới khi mọi thứ xong xuôi).
        final_order = list(ordered_names)
        if other_sheet_names:
            _copy_other_sheets(tmp_path, backup_path, other_sheet_names)
            final_order += other_sheet_names

        os.replace(tmp_path, out_path)  # atomic, chỉ 1 bước "chuyển giao" cuối cùng
        tmp_path = None  # đã đổi tên thành out_path, không còn để dọn dẹp
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
        if os.path.exists(backup_path):
            os.remove(backup_path)

    print(f"  Đã ghi {out_path}  (sheets: {', '.join(final_order)})")
    return True


def _copy_other_sheets(out_path, backup_path, sheet_names):
    """Copy NGUYÊN TRẠNG (giá trị, công thức, style, độ rộng cột/dòng, ô
    merge, freeze panes, Conditional Formatting, Comment, Data Validation,
    Hyperlink) các sheet có tên trong `sheet_names` từ BẢN CŨ (backup tạm
    lấy trước khi pandas ghi đè) sang file `out_path` vừa được pandas ghi
    lại (chỉ chứa 3 sheet báo cáo). Đây là các sheet do người dùng tự tạo
    (vd sheet chỉ số tài chính tự nhập công thức) -> không được đụng tới,
    kể cả các phần định dạng nâng cao mà trước đây (bản cũ) từng bị bỏ sót
    khi copy (conditional formatting/comment/data validation/hyperlink)."""
    if not os.path.exists(backup_path):
        return
    src_wb = openpyxl.load_workbook(backup_path, data_only=False)
    dst_wb = openpyxl.load_workbook(out_path)
    for name in sheet_names:
        if name not in src_wb.sheetnames:
            continue
        src_ws = src_wb[name]
        dst_ws = dst_wb.create_sheet(name)
        dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
        for col, dim in src_ws.column_dimensions.items():
            dst_ws.column_dimensions[col].width = dim.width
        for row_dim_idx, row_dim in src_ws.row_dimensions.items():
            if row_dim.height:
                dst_ws.row_dimensions[row_dim_idx].height = row_dim.height
        for merged_range in src_ws.merged_cells.ranges:
            dst_ws.merge_cells(str(merged_range))
        for row in src_ws.iter_rows():
            for cell in row:
                new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = cell.font.copy()
                    new_cell.fill = cell.fill.copy()
                    new_cell.border = cell.border.copy()
                    new_cell.alignment = cell.alignment.copy()
                    new_cell.number_format = cell.number_format
                if cell.comment:
                    new_cell.comment = copy.copy(cell.comment)
                if cell.hyperlink:
                    new_cell.hyperlink = copy.copy(cell.hyperlink)
        if src_ws.freeze_panes:
            dst_ws.freeze_panes = src_ws.freeze_panes
        # Conditional Formatting (vd tô màu theo ngưỡng giá trị đặt trong Excel).
        for cf in src_ws.conditional_formatting:
            for rule in cf.rules:
                dst_ws.conditional_formatting.add(str(cf.sqref), copy.copy(rule))
        # Data Validation (vd dropdown list, ràng buộc nhập liệu).
        for dv in src_ws.data_validations.dataValidation:
            dst_ws.add_data_validation(copy.copy(dv))
    dst_wb.save(out_path)


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
