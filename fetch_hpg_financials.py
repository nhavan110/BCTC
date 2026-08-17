"""
Lấy chỉ số tài chính theo NĂM của một mã chứng khoán từ vnstock (nguồn VCI)
và xuất ra file CSV: ROE (%), Tăng trưởng LNST (%), Biên lợi nhuận gộp (%).

Yêu cầu:
    pip install -r requirements.txt

Chạy:
    python fetch_hpg_financials.py [MÃ_CK]

Mặc định mã là HPG nếu không truyền tham số.
"""

import sys
import re
import pandas as pd

SYMBOL = sys.argv[1].upper() if len(sys.argv) > 1 else "HPG"
OUTPUT_FILE = f"{SYMBOL}_financial_ratios.csv"

META_COLUMNS = {
    "item", "item_en", "item_id", "unit", "levels", "row_number",
    "ticker", "yearreport", "lengthreport",
}


# ---------------------------------------------------------------------------
# Vá lỗi UnboundLocalError trong vnstock 4.0.6 (get_hosting_service)
# ---------------------------------------------------------------------------
def _patch_vnstock_hosting_service_bug():
    """
    vnstock.core.utils.env.get_hosting_service() không gán giá trị cho biến
    'hosting_service' khi chạy ngoài Colab/Codespace/Replit/Kaggle/HF Spaces
    (vd: GitHub Actions runner thường, máy local), gây crash ngay khi gọi
    API. Patch này chặn lỗi, trả về "Local or Unknown" thay vì raise.
    Có thể xoá khi vnstock phát hành bản vá chính thức.
    """
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


# ---------------------------------------------------------------------------
# Helpers dò cột / dò dòng theo từ khóa (không phụ thuộc format cụ thể)
# ---------------------------------------------------------------------------
def _norm(text) -> str:
    return re.sub(r"[^a-z0-9%]", " ", str(text).lower())


def _is_long_format(df: pd.DataFrame) -> bool:
    """Long format: mỗi DÒNG là 1 chỉ tiêu (có cột item/item_id)."""
    cols_lower = {c.lower() for c in df.columns}
    return "item_id" in cols_lower or "item" in cols_lower


def _period_columns(df: pd.DataFrame) -> list:
    """Các cột không phải metadata -> coi là cột kỳ báo cáo (năm/quý)."""
    return [c for c in df.columns if str(c).lower() not in META_COLUMNS]


def _find_row_long(df: pd.DataFrame, keywords_all: list, exclude_keywords=None):
    """
    Tìm DÒNG đầu tiên mà item_id (hoặc item_en/item) chứa đủ mọi keyword.
    exclude_keywords: nếu dòng chứa bất kỳ từ nào trong đây thì bỏ qua
    (dùng để loại 'minority interest', 'parent company' khi không cần).
    """
    exclude_keywords = exclude_keywords or []
    candidate_cols = [c for c in ["item_id", "item_en", "item"] if c in df.columns]
    for _, row in df.iterrows():
        combined = " ".join(_norm(row[c]) for c in candidate_cols if pd.notna(row.get(c)))
        if all(kw in combined for kw in keywords_all) and not any(
            kw in combined for kw in exclude_keywords
        ):
            return row
    return None


def _find_col_wide(df: pd.DataFrame, keywords_all: list):
    for col in df.columns:
        norm = _norm(col)
        if all(kw in norm for kw in keywords_all):
            return col
    return None


def _dedupe_columns(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """
    Nếu DataFrame có tên cột trùng lặp (đã gặp thực tế với nguồn VCI, ví dụ
    cột '2018' xuất hiện nhiều lần), việc truy cập theo tên cột sẽ trả về
    nhiều giá trị thay vì 1 -> gây lỗi khó hiểu ở phía sau. Hàm này chỉ giữ
    lại lần xuất hiện ĐẦU TIÊN của mỗi tên cột trùng, và cảnh báo rõ ràng.
    """
    cols = list(df.columns)
    dup_count = len(cols) - len(set(map(str, cols)))
    if dup_count > 0:
        print(
            f"CẢNH BÁO: {label} có {dup_count} cột bị trùng tên "
            f"(vd nhiều cột cùng tên năm) -> chỉ giữ lại cột đầu tiên của mỗi tên."
        )
        df = df.loc[:, ~df.columns.duplicated(keep="first")]
    return df


def _dump_diagnostics(label: str, df: pd.DataFrame):
    print(f"\n--- CHẨN ĐOÁN: {label} ---")
    print(f"Shape: {df.shape}")
    cols = list(df.columns)
    print(f"Số cột: {len(cols)} | Số cột duy nhất: {len(set(map(str, cols)))}")
    print("Danh sách cột:")
    for c in cols:
        print(f"  - {c!r}")
    print("5 dòng đầu (toàn bộ, không cắt):")
    with pd.option_context("display.max_columns", None, "display.width", 200):
        print(df.head(5).to_string())
    print("--- HẾT CHẨN ĐOÁN ---\n")


# ---------------------------------------------------------------------------
# Trích xuất 3 chỉ số từ ratio(), fallback tự tính tăng trưởng LNST từ
# income_statement() nếu ratio() không có sẵn cột/dòng tăng trưởng.
# ---------------------------------------------------------------------------
def extract_metrics(finance) -> pd.DataFrame:
    ratio_df = finance.ratio(period="year", lang="en", dropna=False)
    if ratio_df is None or ratio_df.empty:
        raise RuntimeError("ratio() trả về rỗng.")

    if isinstance(ratio_df.columns, pd.MultiIndex):
        ratio_df.columns = [
            "_".join([str(c) for c in col if c]).strip() for col in ratio_df.columns
        ]
    ratio_df = _dedupe_columns(ratio_df, "ratio()")

    long_format = _is_long_format(ratio_df)
    result = pd.DataFrame()

    if long_format:
        periods = _period_columns(ratio_df)
        roe_row = _find_row_long(ratio_df, ["roe"])
        margin_row = _find_row_long(ratio_df, ["gross", "margin"])
        growth_row = _find_row_long(ratio_df, ["profit", "yoy"])
        if growth_row is None:
            growth_row = _find_row_long(ratio_df, ["profit", "growth"])

        if roe_row is None or margin_row is None or periods == []:
            _dump_diagnostics("ratio() - long format nhưng thiếu dòng cần thiết", ratio_df)

        result["Nam"] = periods
        if roe_row is not None:
            result["ROE (%)"] = pd.to_numeric(
                [roe_row[p] for p in periods], errors="coerce"
            )
        if margin_row is not None:
            result["Bien loi nhuan gop (%)"] = pd.to_numeric(
                [margin_row[p] for p in periods], errors="coerce"
            )
        if growth_row is not None:
            result["Tang truong LNST (%)"] = pd.to_numeric(
                [growth_row[p] for p in periods], errors="coerce"
            )
    else:
        year_col = _find_col_wide(ratio_df, ["yearreport"]) or _find_col_wide(
            ratio_df, ["year"]
        )
        roe_col = _find_col_wide(ratio_df, ["roe"])
        margin_col = _find_col_wide(ratio_df, ["gross", "margin"])
        growth_col = _find_col_wide(ratio_df, ["net", "profit", "yoy"]) or _find_col_wide(
            ratio_df, ["net", "profit", "growth"]
        )

        if year_col is None:
            _dump_diagnostics("ratio() - wide format nhưng không tìm thấy cột năm", ratio_df)
            raise RuntimeError("Không xác định được cột năm trong dữ liệu ratio().")

        result["Nam"] = ratio_df[year_col]
        if roe_col:
            result["ROE (%)"] = pd.to_numeric(ratio_df[roe_col], errors="coerce")
        if margin_col:
            result["Bien loi nhuan gop (%)"] = pd.to_numeric(
                ratio_df[margin_col], errors="coerce"
            )
        if growth_col:
            result["Tang truong LNST (%)"] = pd.to_numeric(
                ratio_df[growth_col], errors="coerce"
            )

    # Fallback: nếu chưa có cột tăng trưởng LNST, tự tính từ income_statement()
    if "Tang truong LNST (%)" not in result.columns or result["Tang truong LNST (%)"].isna().all():
        try:
            growth_series = _compute_net_profit_growth(finance)
            if growth_series is not None:
                result = result.merge(growth_series, on="Nam", how="left")
        except Exception as e:
            print(f"(Không tự tính được tăng trưởng LNST từ income_statement: {e})")

    return result


def _compute_net_profit_growth(finance):
    income_df = finance.income_statement(period="year", lang="en", dropna=False)
    if income_df is None or income_df.empty:
        return None

    if isinstance(income_df.columns, pd.MultiIndex):
        income_df.columns = [
            "_".join([str(c) for c in col if c]).strip() for col in income_df.columns
        ]
    income_df = _dedupe_columns(income_df, "income_statement()")

    if _is_long_format(income_df):
        periods = _period_columns(income_df)
        profit_row = _find_row_long(
            income_df,
            ["net", "profit"],
            exclude_keywords=["minority", "parent"],
        )
        if profit_row is None:
            profit_row = _find_row_long(income_df, ["profit", "after", "tax"])

        if profit_row is None or not periods:
            _dump_diagnostics(
                "income_statement() - không tìm thấy dòng lợi nhuận sau thuế", income_df
            )
            return None

        values = pd.to_numeric([profit_row[p] for p in periods], errors="coerce")
        s = pd.Series(values, index=periods).sort_index()
        growth = s.pct_change() * 100
        return pd.DataFrame({"Nam": growth.index, "Tang truong LNST (%)": growth.values})
    else:
        year_col = _find_col_wide(income_df, ["yearreport"]) or _find_col_wide(
            income_df, ["year"]
        )
        profit_col = _find_col_wide(income_df, ["net", "profit"]) or _find_col_wide(
            income_df, ["attributable"]
        )
        if year_col is None or profit_col is None:
            _dump_diagnostics(
                "income_statement() - wide format nhưng thiếu cột cần thiết", income_df
            )
            return None
        tmp = income_df[[year_col, profit_col]].copy()
        tmp.columns = ["Nam", "profit"]
        tmp = tmp.sort_values("Nam")
        tmp["Tang truong LNST (%)"] = pd.to_numeric(tmp["profit"], errors="coerce").pct_change() * 100
        return tmp[["Nam", "Tang truong LNST (%)"]]


# ---------------------------------------------------------------------------
def main():
    try:
        from vnstock import Finance
    except ImportError as e:
        print("Chưa cài vnstock. Chạy: pip install -r requirements.txt")
        raise e

    print(f"Đang lấy báo cáo tài chính theo năm cho mã {SYMBOL} (nguồn: VCI)...")

    try:
        finance = Finance(symbol=SYMBOL, source="VCI")
        result = extract_metrics(finance)
    except Exception as e:
        import traceback

        print(
            "Lỗi khi gọi API vnstock (VCI). Nguyên nhân thường gặp:\n"
            "  - Không có kết nối internet / bị chặn IP tạm thời\n"
            "  - vnstock đã đổi cấu trúc API (kiểm tra bản mới nhất: pip install -U vnstock)\n"
            f"Chi tiết lỗi: {e}\n"
            "\n--- TRACEBACK ĐẦY ĐỦ (để chẩn đoán) ---"
        )
        traceback.print_exc()
        sys.exit(1)

    if result is None or result.empty or "Nam" not in result.columns:
        print(f"Không lấy được dữ liệu cho mã {SYMBOL}. Kiểm tra lại mã chứng khoán.")
        sys.exit(1)

    result = result.drop_duplicates(subset=["Nam"]).sort_values("Nam").reset_index(drop=True)
    result.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

    print(f"\nĐã xuất {len(result)} năm dữ liệu ra file: {OUTPUT_FILE}")
    print(result.to_string(index=False))


if __name__ == "__main__":
    main()
