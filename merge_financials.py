"""
Gộp 3 file CSV (balance_sheet, income_statement, cash_flow) của mỗi mã trong
thư mục financials/<MÃ>/ thành 1 file Excel financials/<MÃ>/<MÃ>_financials.xlsx
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

Chạy:
    python merge_financials.py                    # gộp tất cả mã có trong financials/
    python merge_financials.py HPG                 # 1 mã
    python merge_financials.py HPG,TCB,FPT,PNJ      # nhiều mã, cách nhau dấu phẩy
"""

import sys
import os
import pandas as pd

FINANCIALS_DIR = "financials"

# Tên sheet <-> hậu tố tên file CSV, theo đúng thứ tự hiển thị trong Excel.
STATEMENT_SHEETS = {
    "balance_sheet": "balance_sheet",
    "income_statement": "income_statement",
    "cash_flow": "cash_flow",
}

DROP_COLUMNS = ["item_en", "item_id"]


def _load_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="utf-8-sig")
    drop_cols = [c for c in DROP_COLUMNS if c in df.columns]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


def _sort_year_columns(cols):
    """Sắp xếp các cột năm giảm dần (mới nhất trước), cột không phải số năm
    (nếu có) được đẩy xuống cuối, giữ nguyên thứ tự gốc."""

    def key(c):
        try:
            return (0, -int(c))
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


def process_symbol(symbol: str) -> bool:
    sym_dir = os.path.join(FINANCIALS_DIR, symbol)
    out_path = os.path.join(sym_dir, f"{symbol}_financials.xlsx")

    if not os.path.isdir(sym_dir):
        print(f"  Bỏ qua {symbol}: không tìm thấy thư mục {sym_dir}")
        return False

    existing_sheets = {}
    if os.path.exists(out_path):
        try:
            existing_sheets = pd.read_excel(out_path, sheet_name=None, engine="openpyxl")
        except Exception as e:
            print(f"  CẢNH BÁO: không đọc được file Excel cũ {out_path} ({e}) -> tạo mới.")
            existing_sheets = {}

    output_sheets = dict(existing_sheets)  # giữ lại các sheet khác (vd financial_ratios)
    any_written = False

    for sheet_name, suffix in STATEMENT_SHEETS.items():
        csv_path = os.path.join(sym_dir, f"{symbol}_{suffix}.csv")
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

    # Thứ tự sheet: 3 báo cáo gốc trước, các sheet khác (vd financial_ratios) sau.
    ordered_names = [n for n in STATEMENT_SHEETS if n in output_sheets]
    other_names = [n for n in output_sheets if n not in STATEMENT_SHEETS]
    final_order = ordered_names + other_names

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name in final_order:
            output_sheets[name].to_excel(writer, sheet_name=name, index=False)

    print(f"  Đã ghi {out_path}  (sheets: {', '.join(final_order)})")
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

    any_success = False
    for symbol in symbols:
        print(f"\n=== Mã {symbol} ===")
        if process_symbol(symbol):
            any_success = True

    if not any_success:
        print("\nKhông gộp được dữ liệu cho bất kỳ mã nào.")
        sys.exit(1)

    print("\nHoàn tất.")


if __name__ == "__main__":
    main()
