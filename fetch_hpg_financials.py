"""
Lấy chỉ số tài chính theo NĂM của HPG (Hòa Phát) từ vnstock (nguồn VCI)
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


def _patch_vnstock_hosting_service_bug():
    """
    Vá lỗi UnboundLocalError trong vnstock (bản 4.0.6):
    vnstock.core.utils.env.get_hosting_service() không gán giá trị cho biến
    'hosting_service' khi chạy trên môi trường không phải Colab/Codespace/
    Replit/Kaggle/HF Spaces (ví dụ: GitHub Actions runner thường, máy local
    bình thường), gây crash ngay khi gọi bất kỳ API nào của vnstock.

    Hàm này monkeypatch lại get_hosting_service() để trả về "Local or Unknown"
    thay vì raise lỗi, không ảnh hưởng đến logic lấy dữ liệu.
    Có thể xoá đoạn patch này khi vnstock phát hành bản vá chính thức.
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
        # Nếu cấu trúc nội bộ vnstock đã thay đổi (bản mới hơn đã tự fix bug này),
        # bỏ qua patch, không làm gián đoạn chương trình.
        pass


_patch_vnstock_hosting_service_bug()

SYMBOL = sys.argv[1].upper() if len(sys.argv) > 1 else "HPG"
OUTPUT_FILE = f"{SYMBOL}_financial_ratios.csv"

# Các từ khóa dùng để dò tên cột một cách "mờ" (fuzzy),
# vì vnstock có thể đổi tên cột nhẹ giữa các phiên bản.
COLUMN_KEYWORDS = {
    "ROE (%)": ["roe"],
    "Tang truong LNST (%)": ["net profit", "yoy"],
    "Bien loi nhuan gop (%)": ["gross margin"],
}

YEAR_COLUMN_KEYWORDS = ["yearreport", "year"]


def find_column(columns, keywords):
    """Tìm cột đầu tiên có tên chứa TẤT CẢ keyword (không phân biệt hoa/thường)."""
    for col in columns:
        norm = re.sub(r"[^a-z0-9%]", " ", str(col).lower())
        if all(kw in norm for kw in keywords):
            return col
    return None


def main():
    try:
        from vnstock import Finance
    except ImportError as e:
        print("Chưa cài vnstock. Chạy: pip install -r requirements.txt")
        raise e

    print(f"Đang lấy báo cáo tài chính theo năm cho mã {SYMBOL} (nguồn: VCI)...")

    try:
        finance = Finance(symbol=SYMBOL, source="VCI")
        df = finance.ratio(period="year", lang="en", dropna=False)
    except Exception as e:
        print(
            "Lỗi khi gọi API vnstock (VCI). Nguyên nhân thường gặp:\n"
            "  - Không có kết nối internet / bị chặn IP tạm thời\n"
            "  - vnstock đã đổi cấu trúc API (kiểm tra bản mới nhất: pip install -U vnstock)\n"
            f"Chi tiết lỗi: {e}"
        )
        sys.exit(1)

    if df is None or df.empty:
        print(f"Không lấy được dữ liệu cho mã {SYMBOL}. Kiểm tra lại mã chứng khoán.")
        sys.exit(1)

    # Xử lý MultiIndex columns nếu có (một số phiên bản vnstock trả về 2 tầng cột)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ["_".join([str(c) for c in col if c]).strip() for col in df.columns]

    year_col = find_column(df.columns, ["yearreport"]) or find_column(df.columns, ["year"])
    roe_col = find_column(df.columns, COLUMN_KEYWORDS["ROE (%)"])
    growth_col = find_column(df.columns, COLUMN_KEYWORDS["Tang truong LNST (%)"])
    margin_col = find_column(df.columns, COLUMN_KEYWORDS["Bien loi nhuan gop (%)"])

    missing = [
        name
        for name, col in [
            ("Năm (yearReport)", year_col),
            ("ROE (%)", roe_col),
            ("Tăng trưởng LNST (%)", growth_col),
            ("Biên lợi nhuận gộp (%)", margin_col),
        ]
        if col is None
    ]

    if missing:
        print("CẢNH BÁO: Không tìm thấy các cột sau trong dữ liệu trả về:")
        for m in missing:
            print(f"  - {m}")
        print("\nDanh sách toàn bộ cột hiện có (để bạn tự map lại nếu cần):")
        for c in df.columns:
            print(f"  - {c}")
        if year_col is None:
            sys.exit(1)

    result = pd.DataFrame()
    result["Nam"] = df[year_col]
    if roe_col:
        result["ROE (%)"] = pd.to_numeric(df[roe_col], errors="coerce")
    if growth_col:
        result["Tang truong LNST (%)"] = pd.to_numeric(df[growth_col], errors="coerce")
    if margin_col:
        result["Bien loi nhuan gop (%)"] = pd.to_numeric(df[margin_col], errors="coerce")

    # Loại trùng năm (nếu API trả cả dữ liệu quý lẫn năm), giữ dòng đầu, sort tăng dần theo năm
    result = result.drop_duplicates(subset=["Nam"]).sort_values("Nam").reset_index(drop=True)

    result.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

    print(f"\nĐã xuất {len(result)} năm dữ liệu ra file: {OUTPUT_FILE}")
    print(result.to_string(index=False))


if __name__ == "__main__":
    main()
