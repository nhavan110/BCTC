# BCTC (vnstock)

Script lấy dữ liệu tài chính theo **năm** từ [vnstock](https://github.com/thinh-vu/vnstock)
(nguồn dữ liệu VCI - Vietcap), cộng thêm script gộp dữ liệu thành file Excel:

### 1. `fetch_full_financials.py` — báo cáo tài chính đầy đủ (BCTC)

Xuất **toàn bộ khoản mục gốc** của 3 báo cáo: Cân đối kế toán (balance sheet),
Kết quả kinh doanh (income statement), Lưu chuyển tiền tệ (cash flow). Không
rút gọn, không tính chỉ số — dữ liệu ở dạng long-format (mỗi dòng 1 khoản
mục, mỗi cột 1 năm) giống hệt cấu trúc trả về từ vnstock. Cột `item_en` và
`item_id` được loại bỏ, chỉ giữ lại cột `item` (tên khoản mục tiếng Việt).

```bash
python fetch_full_financials.py                    # mặc định mã HPG
python fetch_full_financials.py HPG                 # 1 mã
python fetch_full_financials.py HPG,TCB,FPT,PNJ      # nhiều mã, cách nhau dấu phẩy
```

Kết quả: `financials/<MÃ>/<MÃ>_balance_sheet.csv`,
`financials/<MÃ>/<MÃ>_income_statement.csv`,
`financials/<MÃ>/<MÃ>_cash_flow.csv` — đây chỉ là dữ liệu **tạm**, dùng làm
đầu vào cho `merge_financials.py` (bước 2 bên dưới) rồi bị xoá; repo chỉ giữ
lại file Excel.

Nếu 1 mã lỗi (vd bị chặn IP tạm thời), script vẫn tiếp tục chạy các mã còn
lại thay vì dừng toàn bộ. Có sẵn workflow GitHub Actions
`.github/workflows/fetch-full-financials.yml` (chỉ chạy thủ công qua tab
Actions — không có lịch tự động) cho cả rổ 8 mã mặc định (HPG, TCB, FPT,
PNJ, MWG, FRT, MBB, TCX). Workflow này tự fetch CSV -> merge vào Excel ->
xoá CSV -> commit lại chỉ file `.xlsx`.

### 2. `merge_financials.py` — gộp 3 file CSV thành 1 file Excel theo mã

Với mỗi mã (thư mục con trong `financials/`), gộp 3 file CSV
(`balance_sheet`, `income_statement`, `cash_flow`) thành **1 file Excel**
`financials/<MÃ>/<MÃ>_financials.xlsx` với 3 sheet cùng tên — đây là file
**duy nhất được giữ lại và commit vào repo**; 3 file CSV nguồn chỉ là dữ
liệu tạm và có thể xoá đi sau khi gộp xong (`.gitignore` đã loại `*.csv`
trong `financials/`). Sau này có thể thêm sheet `financial_ratios` (chỉ số
tài chính) — script sẽ giữ nguyên sheet đó nếu đã tồn tại, chỉ cập nhật 3
sheet báo cáo gốc.

```bash
python merge_financials.py                    # gộp tất cả mã có trong financials/
python merge_financials.py HPG                 # 1 mã
python merge_financials.py HPG,TCB,FPT,PNJ      # nhiều mã, cách nhau dấu phẩy
```

**Cách gộp dữ liệu qua thời gian:** dữ liệu mới được đối chiếu với dữ liệu cũ
trong file Excel (nếu đã có) theo cột `item`:
- Nếu tên khoản mục (`item`) **trùng khớp** với dòng đã có → merge các cột
  năm mới vào dòng đó (đè giá trị năm trùng nếu có, thêm cột năm mới nếu
  chưa có) → theo thời gian, các năm cũ được giữ lại, các năm mới được nối
  thêm.
- Nếu **không trùng khớp** (khoản mục mới hoàn toàn, chưa từng thấy trong
  dữ liệu cũ) → **không merge** dòng đó, dữ liệu cũ giữ nguyên không đổi.
- Nếu file Excel chưa tồn tại (lần chạy đầu tiên), toàn bộ dữ liệu CSV hiện
  tại được dùng làm nền ban đầu.

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
├── fetch_full_financials.py   # lấy BCTC đầy đủ từ vnstock -> CSV tạm
├── merge_financials.py        # gộp 3 CSV/mã -> 1 file Excel (3 sheet), merge dữ liệu qua các năm
├── financials/<MÃ>/<MÃ>_financials.xlsx   # file DUY NHẤT được commit, do merge_financials.py tạo
├── requirements.txt
├── .gitignore                 # loại financials/**/*.csv (chỉ là file tạm)
└── README.md
```
