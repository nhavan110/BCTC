"""
Lấy BÁO CÁO TÀI CHÍNH ĐẦY ĐỦ theo NĂM (Cân đối kế toán / Kết quả kinh doanh /
Lưu chuyển tiền tệ) của một hoặc nhiều mã chứng khoán từ vnstock (nguồn VCI),
xuất mỗi báo cáo ra 1 file CSV riêng, giữ nguyên toàn bộ khoản mục gốc.
"""

import sys
import os
import time
import pandas as pd

DEFAULT_SYMBOLS = ["HPG"]
OUTPUT_DIR = "financials"

# Community: 60 requests/phút khi dùng API key.
# 5 giây/request giúp giới hạn tối đa khoảng 12 request/phút.
REQUEST_DELAY_SECONDS = 5
MAX_RETRIES = 3
RATE_LIMIT_WAIT_SECONDS = 65

STATEMENTS = {
    "balance_sheet": "balance_sheet",
    "income_statement": "income_statement",
    "cash_flow": "cash_flow",
}


# ---------------------------------------------------------------------------
# Vá lỗi UnboundLocalError trong vnstock 4.0.6 (get_hosting_service)
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
    """Đăng ký API key từ GitHub Actions Secret, không lưu key trong repository."""
    api_key = os.getenv("VNSTOCK_API_KEY", "").strip()
    if not api_key:
        print("CẢNH BÁO: Không tìm thấy VNSTOCK_API_KEY -> đang chạy Guest.")
        return False

    try:
        from vnstock import register_user
        register_user(api_key=api_key)
        print("Đã xác thực Vnstock bằng API key.")
        return True
    except Exception as e:
        print(f"CẢNH BÁO: Không đăng ký được VNSTOCK_API_KEY: {e}")
        return False


def _dedupe_columns(df: pd.DataFrame, label: str) -> pd.DataFrame:
    cols = list(df.columns)
    dup_count = len(cols) - len(set(map(str, cols)))
    if dup_count > 0:
        print(f"  CẢNH BÁO [{label}]: {dup_count} cột trùng tên -> giữ cột đầu tiên.")
        df = df.loc[:, ~df.columns.duplicated(keep="first")]
    return df


def _flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ["_".join([str(c) for c in col if c]).strip() for col in df.columns]
    return df


def _drop_unused_columns(df: pd.DataFrame) -> pd.DataFrame:
    drop_cols = [c for c in ("item_en", "item_id") if c in df.columns]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


def _fetch_one_statement(finance, symbol: str, method_name: str) -> pd.DataFrame:
    """Gọi một endpoint; nếu gặp rate limit thì chờ rồi tự retry."""
    method = getattr(finance, method_name)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            df = method(period="year", lang="en", dropna=False)
            if df is None or df.empty:
                raise RuntimeError(f"{method_name}() trả về rỗng cho mã {symbol}.")

            df = _flatten_columns(df)
            df = _dedupe_columns(df, f"{symbol}.{method_name}()")
            return _drop_unused_columns(df)

        except Exception as e:
            error_text = str(e).lower()
            is_rate_limit = any(x in error_text for x in (
                "429", "too many requests", "rate limit", "rate_limit", "quota"
            ))

            if is_rate_limit and attempt < MAX_RETRIES:
                print(
                    f"  Rate limit ở {symbol}.{method_name}() "
                    f"(lần {attempt}/{MAX_RETRIES}). Chờ {RATE_LIMIT_WAIT_SECONDS}s..."
                )
                time.sleep(RATE_LIMIT_WAIT_SECONDS)
                continue

            raise


def fetch_symbol(symbol: str) -> dict:
    from vnstock import Finance

    finance = Finance(symbol=symbol, source="VCI")
    results = {}

    for label, method_name in STATEMENTS.items():
        try:
            print(f"  -> Đang lấy {label} cho {symbol}...")
            results[label] = _fetch_one_statement(finance, symbol, method_name)
        except Exception as e:
            print(f"  LỖI khi lấy {label} cho {symbol}: {e}")
        finally:
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
            print(f"Lỗi khi khởi tạo/gọi API vnstock (VCI) cho {symbol}. Chi tiết lỗi: {e}")
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
