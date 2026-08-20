"""
Lấy BÁO CÁO TÀI CHÍNH ĐẦY ĐỦ theo NĂM (Cân đối kế toán / Kết quả kinh doanh /
Lưu chuyển tiền tệ) của một hoặc nhiều mã chứng khoán từ vnstock (nguồn VCI),
xuất mỗi báo cáo ra 1 file CSV riêng, giữ nguyên toàn bộ khoản mục gốc
(không rút gọn thành vài chỉ số như fetch_hpg_financials.py).

Yêu cầu:
    pip install -r requirements.txt

Chạy:
    python fetch_full_financials.py                # mặc định HPG
    python fetch_full_financials.py HPG             # 1 mã
    python fetch_full_financials.py HPG,TCB,FPT     # nhiều mã, cách nhau bởi dấu phẩy

Kết quả (trong thư mục con financials/<MÃ>/):
    <MÃ>_balance_sheet.csv
    <MÃ>_income_statement.csv
    <MÃ>_cash_flow.csv
"""

import sys
import os
import time
import pandas as pd

DEFAULT_SYMBOLS = ["HPG"]
OUTPUT_DIR = "financials"

# Giữ khoảng cách giữa các lần gọi API để tránh vượt rate limit.
# Community có tối đa 60 requests/phút khi dùng API key.
REQUEST_DELAY_SECONDS = 2.5

STATEMENTS = {
    "balance_sheet": "balance_sheet",
    "income_statement": "income_statement",
    "cash_flow": "cash_flow",
}


# ---------------------------------------------------------------------------
# Vá lỗi UnboundLocalError trong vnstock 4.0.6 (get_hosting_service) — giống
# fetch_hpg_financials.py, giữ đồng nhất giữa 2 script trong repo.
# ---------------------------------------------------------------------------
def _patch_vnstock_hosting_service_bug():
    try:
        from vnstock.core.utils import env as vnstock_env

        _original = vnstock_env.get_hosting_service

        def _safe_get_hosting_service():
            try:
                result = _original()
                return result if result is not None else "Local or Unknown"
            except UnboundLocalError:
                return "Local or Unknown"

        vnstock_env.get_hosting_service = _safe_get_hosting_service
    except Exception:
        pass


_patch_vnstock_hosting_service_bug()


def _register_vnstock_api_key():
    """Đăng ký API key từ biến môi trường, không hard-code key vào repository."""
    api_key = os.getenv("VNSTOCK_API_KEY", "").strip()
    if not api_key:
        print("Không tìm thấy VNSTOCK_API_KEY -> chạy ở chế độ Guest.")
        return

    try:
        from vnstock import register_user
        register_user(api_key=api_key)
        print("Đã xác thực Vnstock bằng API key.")
    except Exception as e:
        print(f"CẢNH BÁO: Không đăng ký được VNSTOCK_API_KEY: {e}")
        print("Tiếp tục chạy; Vnstock có thể sử dụng chế độ Guest.")


def _dedupe_columns(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """Giữ lại cột đầu tiên nếu có tên cột (năm) bị trùng — thấy thực tế ở nguồn VCI."""
    cols = list(df.columns)
    dup_count = len(cols) - len(set(map(str, cols)))
    if dup_count > 0:
        print(
            f"  CẢNH BÁO [{label}]: {dup_count} cột bị trùng tên -> chỉ giữ cột đầu tiên."
        )
        df = df.loc[:, ~df.columns.duplicated(keep="first")]
    return df


def _flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            "_".join([str(c) for c in col if c]).strip() for col in df.columns
        ]
    return df


def _drop_unused_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Bỏ các cột không cần thiết (item_en, item_id) - chỉ giữ tên khoản mục
    tiếng Việt (item)."""
    drop_cols = [c for c in ("item_en", "item_id") if c in df.columns]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


def _fetch_one_statement(finance, symbol: str, method_name: str) -> pd.DataFrame:
    method = getattr(finance, method_name)
    df = method(period="year", lang="en", dropna=False)
    if df is None or df.empty:
        raise RuntimeError(f"{method_name}() trả về rỗng cho mã {symbol}.")
    df = _flatten_columns(df)
    df = _dedupe_columns(df, f"{symbol}.{method_name}()")
    df = _drop_unused_columns(df)
    return df


def fetch_symbol(symbol: str) -> dict:
    """Trả về dict {ten_bao_cao: DataFrame} cho 1 mã. Không dừng toàn bộ nếu
    1 trong 3 báo cáo lỗi -- báo cáo lỗi sẽ bị bỏ qua và in cảnh báo."""
    from vnstock import Finance

    finance = Finance(symbol=symbol, source="VCI")
    results = {}
    for index, (label, method_name) in enumerate(STATEMENTS.items()):
        try:
            print(f"  -> Đang lấy {label} cho {symbol}...")
            results[label] = _fetch_one_statement(finance, symbol, method_name)
        except Exception as e:
            print(f"  LỖI khi lấy {label} cho {symbol}: {e}")
        finally:
            # Không sleep sau request cuối cùng của toàn bộ symbol cũng không
            # gây hại; giữ logic đơn giản và an toàn cho rate limit.
            time.sleep(REQUEST_DELAY_SECONDS)
    return results


def main():
    _register_vnstock_api_key()

    if len(sys.argv) > 1:
        symbols = [s.strip().upper() for s in sys.argv[1].split(",") if s.strip()]
    else:
        symbols = DEFAULT_SYMBOLS

    any_success = False

    for symbol in symbols:
        print(f"\n=== Mã {symbol} ===")
        out_dir = os.path.join(OUTPUT_DIR, symbol)
        os.makedirs(out_dir, exist_ok=True)

        try:
            statements = fetch_symbol(symbol)
        except Exception as e:
            import traceback

            print(
                f"Lỗi khi khởi tạo/gọi API vnstock (VCI) cho {symbol}. Nguyên nhân thường gặp:\n"
                "  - Không có kết nối internet / bị chặn IP tạm thời\n"
                "  - vnstock đã đổi cấu trúc API (kiểm tra bản mới nhất: pip install -U vnstock)\n"
                f"Chi tiết lỗi: {e}"
            )
            traceback.print_exc()
            continue

        if not statements:
            print(f"Không lấy được báo cáo nào cho {symbol}, bỏ qua.")
            continue

        for label, df in statements.items():
            out_path = os.path.join(out_dir, f"{symbol}_{label}.csv")
            df.to_csv(out_path, index=False, encoding="utf-8-sig")
            print(f"  Đã ghi {out_path}  (shape={df.shape})")
            any_success = True

    if not any_success:
        print("\nKhông lấy được dữ liệu cho bất kỳ mã nào.")
        sys.exit(1)

    print(f"\nHoàn tất. Dữ liệu nằm trong thư mục '{OUTPUT_DIR}/<MÃ>/'.")


if __name__ == "__main__":
    main()
