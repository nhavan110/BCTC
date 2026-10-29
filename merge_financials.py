"""
Gộp 3 file CSV (balance_sheet, income_statement, cash_flow) của mỗi mã trong
thư mục financials/<MÃ>/ thành 1 file Excel financials/<MÃ>/<MÃ>.xlsx
với 3 sheet cùng tên.

Cách gộp dữ liệu qua thời gian (để không mất các năm cũ khi vnstock chỉ trả
về vài năm gần nhất mỗi lần fetch):
    - Dữ liệu mới được đối chiếu với dữ liệu cũ đã có trong file Excel theo
      cột "item" (tên khoản mục). Khi so khớp, KHÔNG phân biệt hoa/thường và
      bỏ qua khoảng trắng thừa; mọi khác biệt khác đều coi là KHÔNG khớp.
    - Nếu "item" KHỚP với 1 dòng đã có -> merge các cột năm mới vào dòng đó
      (đè giá trị nếu năm đã tồn tại, thêm cột nếu là năm mới). Giá trị mới
      rỗng (NaN) không đè lên số cũ.
    - Nếu "item" KHÔNG KHỚP -> thêm nguyên dòng đó từ dữ liệu mới thành 1
      dòng MỚI ở CUỐI sheet (không chèn giữa để không làm lệch địa chỉ ô mà
      các sheet công thức tự tạo đang tham chiếu). Dòng cũ giữ nguyên. Script
      in cảnh báo kèm gợi ý "có thể là đổi tên của ..." để đối chiếu thủ công.
    - Tên khoản mục trùng nhau trong cùng 1 báo cáo được ghép theo thứ tự
      xuất hiện (dòng trùng thứ n bên mới ghép với dòng trùng thứ n bên cũ).
    - Nếu file Excel chưa tồn tại (lần chạy đầu tiên) -> dùng toàn bộ CSV
      hiện tại làm dữ liệu nền ban đầu.

Chỉ ghi lại file Excel của 1 mã khi dữ liệu thực sự thay đổi. Mã không có gì
mới thì file giữ nguyên và không cần upload lên Google Drive. Danh sách mã có
thay đổi được ghi vào $GITHUB_OUTPUT (changed=HPG,TCB) để workflow quyết định
có upload hay không; cảnh báo/tóm tắt ghi vào $GITHUB_STEP_SUMMARY.

Nếu file Excel đã có thêm sheet khác (vd "financial_ratios" tự thêm sau
này), sheet đó được giữ nguyên, không bị đụng tới.

Chạy:
    python merge_financials.py                    # gộp tất cả mã có trong financials/
    python merge_financials.py HPG                 # 1 mã
    python merge_financials.py HPG,TCB,FPT,PNJ      # nhiều mã, cách nhau dấu phẩy
"""

import sys
import os
import copy
import difflib
import shutil
import numpy as np
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

# Số dòng tối đa in ra / ghi vào summary cho mỗi danh sách cảnh báo.
MAX_LISTED = 30


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


def _norm_item(value) -> str:
    """Khoá so khớp tên khoản mục: không phân biệt hoa/thường, gộp khoảng
    trắng thừa (kể cả NBSP). Mọi khác biệt khác đều là KHÔNG khớp."""
    return " ".join(str(value).replace("\xa0", " ").split()).casefold()


def _keys_with_occurrence(items):
    """Gán cho mỗi dòng khoá (tên chuẩn hoá, lần xuất hiện thứ k) để các
    khoản mục trùng tên trong cùng 1 báo cáo ghép được theo thứ tự thay vì
    cùng đè lên 1 dòng."""
    counter = {}
    keys = []
    for v in items:
        n = _norm_item(v)
        k = counter.get(n, 0)
        counter[n] = k + 1
        keys.append((n, k))
    return keys


def merge_sheet(old_df: pd.DataFrame, new_df: pd.DataFrame):
    """Merge new_df vào old_df theo cột 'item'. Xem quy tắc ở đầu file.

    Trả về (merged_df, info) với info = {
        "appended": [(tên khoản mục mới không khớp, gợi ý tên cũ hoặc None)],
        "orphans":  [tên khoản mục cũ không xuất hiện trong dữ liệu mới],
    }"""
    info = {"appended": [], "orphans": []}

    new_df = new_df.copy()
    new_df.columns = [str(c) for c in new_df.columns]
    year_cols_new = [c for c in new_df.columns if c != "item"]

    if old_df is None or old_df.empty:
        # Chưa có dữ liệu cũ -> dùng dữ liệu mới làm nền.
        merged = new_df.copy()
    else:
        merged = old_df.copy().reset_index(drop=True)
        merged.columns = [str(c) for c in merged.columns]

        # Đảm bảo các cột năm mới tồn tại trong merged.
        for yc in year_cols_new:
            if yc not in merged.columns:
                merged[yc] = np.nan

        old_keys = _keys_with_occurrence(merged["item"])
        key_to_index = {}
        for i, k in enumerate(old_keys):
            key_to_index.setdefault(k, i)

        matched = set()
        unmatched_rows = []
        new_keys = _keys_with_occurrence(new_df["item"])
        for key, (_, row) in zip(new_keys, new_df.iterrows()):
            old_idx = key_to_index.get(key)
            if old_idx is None:
                unmatched_rows.append(row)
                continue
            matched.add(old_idx)
            for yc in year_cols_new:
                val = row[yc]
                if pd.isna(val):
                    continue  # không đè số cũ bằng giá trị rỗng
                merged.loc[old_idx, yc] = val

        # Khoản mục cũ không còn xuất hiện trong dữ liệu mới.
        orphan_items = [
            str(merged["item"].iloc[i]) for i in range(len(old_keys)) if i not in matched
        ]
        info["orphans"] = orphan_items
        lookup = {o.casefold(): o for o in orphan_items}

        # Không khớp -> thêm nguyên dòng mới xuống CUỐI sheet.
        if unmatched_rows:
            for row in unmatched_rows:
                name = str(row["item"])
                guess = difflib.get_close_matches(name.casefold(), list(lookup), n=1, cutoff=0.6)
                info["appended"].append((name, lookup[guess[0]] if guess else None))
            add_df = pd.DataFrame(unmatched_rows).reindex(columns=merged.columns)
            merged = pd.concat([merged, add_df], ignore_index=True)

    year_cols = _sort_year_columns([c for c in merged.columns if c != "item"])
    merged = merged[["item"] + year_cols]
    return merged, info


def _frames_same(old_df, new_df) -> bool:
    """True nếu merged (new_df) không khác gì dữ liệu cũ (old_df): cùng cột,
    cùng danh sách khoản mục theo thứ tự, giá trị số bằng nhau (có dung sai)."""
    if old_df is None:
        return False
    a = old_df.reset_index(drop=True).copy()
    b = new_df.reset_index(drop=True).copy()
    a.columns = [str(c) for c in a.columns]
    b.columns = [str(c) for c in b.columns]
    if list(a.columns) != list(b.columns) or len(a) != len(b):
        return False
    if not (a["item"].astype(str) == b["item"].astype(str)).all():
        return False
    for c in a.columns:
        if c == "item":
            continue
        x = pd.to_numeric(a[c], errors="coerce").to_numpy(dtype=float)
        y = pd.to_numeric(b[c], errors="coerce").to_numpy(dtype=float)
        if not np.allclose(x, y, rtol=1e-9, atol=1e-6, equal_nan=True):
            return False
    return True


def _report_mismatch(symbol: str, sheet_info: dict) -> None:
    """In cảnh báo khoản mục không khớp tên (hiện thành annotation trên GitHub Actions)."""
    for sheet, info in sheet_info.items():
        appended = info["appended"]
        orphans = info["orphans"]
        if appended:
            print(
                f"::warning::{symbol}/{sheet}: {len(appended)} khoản mục KHÔNG khớp tên với "
                f"dữ liệu cũ -> đã thêm thành dòng mới ở cuối sheet, cần đối chiếu thủ công"
            )
            for name, guess in appended[:MAX_LISTED]:
                hint = f"  (có thể là đổi tên của: {guess})" if guess else ""
                print(f"    + {name}{hint}")
            if len(appended) > MAX_LISTED:
                print(f"    ... và {len(appended) - MAX_LISTED} khoản mục khác")
        if orphans:
            print(f"  [{symbol}/{sheet}] {len(orphans)} khoản mục cũ không có trong dữ liệu mới:")
            for name in orphans[:MAX_LISTED]:
                print(f"    - {name}")
            if len(orphans) > MAX_LISTED:
                print(f"    ... và {len(orphans) - MAX_LISTED} khoản mục khác")


def _write_github_outputs(changed, unchanged, report) -> None:
    """Ghi output cho workflow (changed=...) và tóm tắt hiển thị trên tab Actions."""
    out_path = os.environ.get("GITHUB_OUTPUT")
    if out_path:
        with open(out_path, "a", encoding="utf-8") as f:
            f.write(f"changed={','.join(changed)}\n")

    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not summary_path:
        return

    lines = ["## Kết quả gộp BCTC", ""]
    if changed:
        lines.append(f"- Có cập nhật (sẽ upload Google Drive): **{', '.join(changed)}**")
    if unchanged:
        lines.append(f"- Không có dữ liệu mới: {', '.join(unchanged)}")
    if not changed:
        lines += ["", "**Không có nội dung cập nhật -> bỏ qua upload Google Drive.**"]

    mismatch_lines = []
    for symbol, sheet_info in report.items():
        for sheet, info in sheet_info.items():
            for name, guess in info["appended"][:MAX_LISTED]:
                hint = f" — có thể là đổi tên của `{guess}`" if guess else ""
                mismatch_lines.append(f"- `{symbol}/{sheet}`: `{name}`{hint}")
    if mismatch_lines:
        lines += [
            "",
            "### Khoản mục không khớp tên",
            "Đã thêm thành dòng mới ở cuối sheet, cần đối chiếu thủ công:",
            *mismatch_lines,
        ]

    with open(summary_path, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


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


def process_symbol(symbol: str, report: dict) -> str:
    """Xử lý 1 mã. Trả về 'changed' (đã ghi file mới), 'unchanged' (không có
    dữ liệu mới, không ghi file) hoặc 'skipped' (không xử lý được)."""
    sym_dir = os.path.join(FINANCIALS_DIR, symbol)
    out_path = os.path.join(sym_dir, f"{symbol}.xlsx")

    if not os.path.isdir(sym_dir):
        print(f"  Bỏ qua {symbol}: không tìm thấy thư mục {sym_dir}")
        return "skipped"

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

    output_sheets = {}
    sheet_info = {}
    changed = not os.path.exists(out_path)  # chưa có file -> coi là có thay đổi

    for sheet_name, suffix in STATEMENT_SHEETS.items():
        csv_path = os.path.join(sym_dir, f"{symbol}_{suffix}.csv")
        old_df = existing_sheets.get(sheet_name)
        if not os.path.exists(csv_path):
            print(f"  Bỏ qua sheet '{sheet_name}': không tìm thấy {csv_path}")
            if old_df is not None:
                output_sheets[sheet_name] = old_df  # giữ nguyên sheet cũ, không làm mất
            continue

        new_df = _load_csv(csv_path)
        merged, info = merge_sheet(old_df, new_df)
        output_sheets[sheet_name] = merged
        sheet_info[sheet_name] = info
        if not _frames_same(old_df, merged):
            changed = True

    if not sheet_info:
        print(f"  Không có CSV nào cho {symbol}, bỏ qua.")
        return "skipped"

    report[symbol] = sheet_info
    _report_mismatch(symbol, sheet_info)

    if not changed:
        print(f"  Không có dữ liệu mới cho {symbol} -> giữ nguyên file, không cần upload.")
        return "unchanged"

    backup_path = out_path + ".bak_other_sheets.xlsx"
    if other_sheet_names and os.path.exists(out_path):
        shutil.copyfile(out_path, backup_path)

    ordered_names = [n for n in STATEMENT_SHEETS if n in output_sheets]

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name in ordered_names:
            output_sheets[name].to_excel(writer, sheet_name=name, index=False)
            _format_statement_sheet(writer.sheets[name])

    # Copy nguyên trạng (công thức + định dạng) các sheet khác từ file cũ
    # (vd "chi_so_tai_chinh") sang file vừa ghi ở trên.
    final_order = list(ordered_names)
    if other_sheet_names:
        _copy_other_sheets(out_path, other_sheet_names)
        final_order += other_sheet_names

    print(f"  Đã ghi {out_path}  (sheets: {', '.join(final_order)})")
    return "changed"


def _copy_other_sheets(out_path, sheet_names):
    """Copy NGUYÊN TRẠNG (giá trị, công thức, style, độ rộng cột/dòng, ô
    merge, freeze panes, Conditional Formatting, Comment, Data Validation,
    Hyperlink) các sheet có tên trong `sheet_names` từ BẢN CŨ (backup tạm
    lấy trước khi pandas ghi đè) sang file `out_path` vừa được pandas ghi
    lại (chỉ chứa 3 sheet báo cáo). Đây là các sheet do người dùng tự tạo
    (vd sheet chỉ số tài chính tự nhập công thức) -> không được đụng tới,
    kể cả các phần định dạng nâng cao mà trước đây (bản cũ) từng bị bỏ sót
    khi copy (conditional formatting/comment/data validation/hyperlink)."""
    backup_path = out_path + ".bak_other_sheets.xlsx"
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
    os.remove(backup_path)


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

    symbols = list(dict.fromkeys(symbols))  # bỏ mã trùng, giữ thứ tự

    if not symbols:
        print(f"Không tìm thấy mã nào trong '{FINANCIALS_DIR}/'.")
        sys.exit(1)

    changed_symbols = []
    unchanged_symbols = []
    report = {}
    for symbol in symbols:
        print(f"\n=== Mã {symbol} ===")
        status = process_symbol(symbol, report)
        if status == "changed":
            changed_symbols.append(symbol)
        elif status == "unchanged":
            unchanged_symbols.append(symbol)

    if not changed_symbols and not unchanged_symbols:
        print("\nKhông gộp được dữ liệu cho bất kỳ mã nào.")
        sys.exit(1)

    _write_github_outputs(changed_symbols, unchanged_symbols, report)

    if changed_symbols:
        print(f"\nCó cập nhật: {', '.join(changed_symbols)}")
    else:
        print("\nKhông có nội dung cập nhật ở mã nào -> không cần upload Google Drive.")
    print("\nHoàn tất.")


if __name__ == "__main__":
    main()
