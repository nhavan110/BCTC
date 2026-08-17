# BCTC (vnstock)

Hai script lấy dữ liệu tài chính theo **năm** từ [vnstock](https://github.com/thinh-vu/vnstock)
(nguồn dữ liệu VCI - Vietcap):

### 1. `fetch_hpg_financials.py` — chỉ số rút gọn (ratio)

Xuất 3 chỉ số: ROE (%), Biên lợi nhuận gộp (%), Tăng trưởng LNST (%) YoY.

```bash
python fetch_hpg_financials.py        # mặc định mã HPG
python fetch_hpg_financials.py VNM    # hoặc truyền mã khác
```

Kết quả: `<MÃ>_financial_ratios.csv` trong thư mục hiện tại.

### 2. `fetch_full_financials.py` — báo cáo tài chính đầy đủ (BCTC)

Xuất **toàn bộ khoản mục gốc** của 3 báo cáo: Cân đối kế toán (balance sheet),
Kết quả kinh doanh (income statement), Lưu chuyển tiền tệ (cash flow). Không
rút gọn, không tính chỉ số — dữ liệu ở dạng long-format (mỗi dòng 1 khoản
mục, mỗi cột 1 năm) giống hệt cấu trúc trả về từ vnstock.

```bash
python fetch_full_financials.py                    # mặc định mã HPG
python fetch_full_financials.py HPG                 # 1 mã
python fetch_full_financials.py HPG,TCB,FPT,PNJ      # nhiều mã, cách nhau dấu phẩy
```

Kết quả: `financials/<MÃ>/<MÃ>_balance_sheet.csv`,
`financials/<MÃ>/<MÃ>_income_statement.csv`,
`financials/<MÃ>/<MÃ>_cash_flow.csv`.

Nếu 1 mã lỗi (vd bị chặn IP tạm thời), script vẫn tiếp tục chạy các mã còn
lại thay vì dừng toàn bộ. Có sẵn workflow GitHub Actions
`.github/workflows/fetch-full-financials.yml` chạy định kỳ hàng tuần cho cả
rổ 8 mã mặc định (HPG, TCB, FPT, PNJ, MWG, FRT, MBB, TCX).

## Cài đặt

```bash
pip install -r requirements.txt
```

## Lưu ý quan trọng

- vnstock lấy dữ liệu bằng cách gọi API/scrape từ trang của công ty chứng khoán
  Vietcap (VCI), nên máy chạy script **cần có kết nối internet ra ngoài**
  (không hoạt động trong môi trường bị chặn mạng, ví dụ một số sandbox CI có
  network policy hạn chế).
- Nếu chạy trên Google Colab/Kaggle hoặc server cloud và bị chặn IP tạm thời,
  vnstock hỗ trợ tham số `proxy_mode` / `proxy_list` — xem thêm tại
  [tài liệu vnstock](https://vnstocks.com/docs/vnstock/bao-cao-tai-chinh).
- vnstock là thư viện đang phát triển liên tục, tên cột trả về đôi khi thay
  đổi giữa các phiên bản. Script này đã viết cơ chế dò tên cột (fuzzy match)
  để giảm rủi ro vỡ khi có thay đổi nhỏ; nếu vnstock đổi cấu trúc lớn, script
  sẽ in ra danh sách toàn bộ cột hiện có để bạn map lại thủ công.
- Class `Vnstock` (unified, cũ) đã bị deprecate và sẽ EOL vào 31/08/2026 —
  script này dùng trực tiếp class `Finance` (API mới) để tránh cảnh báo/lỗi.
- Luôn nên `pip install -U vnstock` định kỳ để nhận bản vá lỗi API mới nhất.

## Cấu trúc

```
.
├── fetch_hpg_financials.py   # script chính
├── requirements.txt
└── README.md
```
