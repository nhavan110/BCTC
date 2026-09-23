# BCTC (vnstock)

Script lấy dữ liệu tài chính theo **năm** và theo **quý** từ [vnstock](https://github.com/thinh-vu/vnstock)
(nguồn dữ liệu VCI - Vietcap), cộng thêm script gộp dữ liệu thành file Excel:

### 1. `fetch_full_financials.py` — báo cáo tài chính đầy đủ (BCTC)

Xuất **toàn bộ khoản mục gốc** của 3 báo cáo: Cân đối kế toán (balance sheet),
Kết quả kinh doanh (income statement), Lưu chuyển tiền tệ (cash flow). Không
rút gọn, không tính chỉ số — dữ liệu ở dạng long-format (mỗi dòng 1 khoản
mục, mỗi cột 1 kỳ báo cáo) giống hệt cấu trúc trả về từ vnstock. Cột `item_en`
và `item_id` được loại bỏ, chỉ giữ lại cột `item` (tên khoản mục tiếng Việt).

Nhận tham số **period** thứ 2 (`year` mặc định, hoặc `quarter`):

```bash
python fetch_full_financials.py                              # mặc định mã HPG, period=year
python fetch_full_financials.py HPG                           # 1 mã, period=year
python fetch_full_financials.py HPG,TCB,FPT,PNJ                # nhiều mã, cách nhau dấu phẩy
python fetch_full_financials.py HPG,TCB,FPT,PNJ quarter         # nhiều mã, theo QUÝ
```

Kết quả (period=year): `financials/<MÃ>/<MÃ>_balance_sheet.csv`,
`financials/<MÃ>/<MÃ>_income_statement.csv`,
`financials/<MÃ>/<MÃ>_cash_flow.csv`.
Kết quả (period=quarter): tương tự nhưng thêm hậu tố `_Q`, vd
`financials/<MÃ>/<MÃ>_Q_balance_sheet.csv` (cột dữ liệu dạng `2024-Q1`,
`2024-Q2`... thay vì `2024`). Đây chỉ là dữ liệu **tạm**, dùng làm đầu vào
cho `merge_financials.py` (bước 2 bên dưới) rồi bị xoá; repo chỉ giữ lại
file Excel.

Nếu 1 mã lỗi (vd bị chặn IP tạm thời), script vẫn tiếp tục chạy các mã còn
lại thay vì dừng toàn bộ. Có sẵn workflow GitHub Actions
`.github/workflows/fetch-full-financials.yml` (chỉ chạy thủ công qua tab
Actions — không có lịch tự động) cho cả rổ 8 mã mặc định (HPG, TCB, FPT,
PNJ, MWG, FRT, MBB, TCX). Workflow này tự fetch CSV -> merge vào Excel ->
xoá CSV -> commit lại chỉ file `.xlsx`.

### 2. `merge_financials.py` — gộp 3 file CSV thành 1 file Excel theo mã

Với mỗi mã (thư mục con trong `financials/`), gộp 3 file CSV
(`balance_sheet`, `income_statement`, `cash_flow`) thành **1 file Excel**
với 3 sheet cùng tên — đây là file **duy nhất được giữ lại và commit vào
repo**; 3 file CSV nguồn chỉ là dữ liệu tạm và có thể xoá đi sau khi gộp
xong (`.gitignore` đã loại `*.csv` trong `financials/`). Sau này có thể
thêm sheet khác (vd `financial_ratios` — chỉ số tài chính) — script sẽ giữ
nguyên sheet đó nếu đã tồn tại, chỉ cập nhật 3 sheet báo cáo gốc.

Cũng nhận tham số **period** thứ 2 giống `fetch_full_financials.py`, quyết
định đọc CSV nào và ghi ra file Excel nào:
- `year` (mặc định): đọc `<MÃ>_<report>.csv` → ghi `financials/<MÃ>/<MÃ>.xlsx`
- `quarter`: đọc `<MÃ>_Q_<report>.csv` → ghi `financials/<MÃ>/<MÃ>_Q.xlsx`
  (file **riêng biệt** với file năm, không đụng tới nhau)

```bash
python merge_financials.py                          # gộp tất cả mã, period=year
python merge_financials.py HPG                       # 1 mã, period=year
python merge_financials.py HPG,TCB,FPT,PNJ            # nhiều mã, period=year
python merge_financials.py HPG,TCB,FPT,PNJ quarter     # nhiều mã, period=quarter -> <MÃ>_Q.xlsx
```

**Cách gộp dữ liệu qua thời gian:** dữ liệu mới được đối chiếu với dữ liệu cũ
trong file Excel (nếu đã có) theo cột `item` — áp dụng như nhau cho cả 2 kỳ:
- Nếu tên khoản mục (`item`) **trùng khớp** với dòng đã có → merge các cột
  kỳ mới vào dòng đó (đè giá trị kỳ trùng nếu có, thêm cột kỳ mới nếu
  chưa có) → theo thời gian, các kỳ cũ được giữ lại, các kỳ mới được nối
  thêm (với quarter, cột được đặt tên `2024-Q1`, `2024-Q2`... và tự sắp
  đúng thứ tự thời gian).
- Nếu **không trùng khớp** (khoản mục mới hoàn toàn, chưa từng thấy trong
  dữ liệu cũ) → **không merge** dòng đó, dữ liệu cũ giữ nguyên không đổi.
- Nếu file Excel chưa tồn tại (lần chạy đầu tiên), toàn bộ dữ liệu CSV hiện
  tại được dùng làm nền ban đầu.

### 3. `gdrive_download.py` / `gdrive_upload.py` — đồng bộ với Google Drive

Các sheet bạn tự tạo thêm trong file Excel (vd sheet chỉ số tài chính với
công thức bạn tự nhập) **không bao giờ bị động tới** bởi `merge_financials.py`
— script chỉ đọc/ghi 3 sheet báo cáo gốc (`balance_sheet`, `income_statement`,
`cash_flow`); mọi sheet khác được copy nguyên trạng (giá trị, công thức,
style, Conditional Formatting, Comment, Data Validation, Hyperlink). Điều
này áp dụng cho cả file năm (`<MÃ>.xlsx`) lẫn file quý (`<MÃ>_Q.xlsx`) —
mỗi file có sheet tự thêm riêng, độc lập với nhau.

Vì file Excel "chuẩn" (đã có sheet bạn tự làm) được lưu trên Google Drive,
2 script này giúp workflow luôn lấy đúng bản mới nhất làm nền trước khi cập
nhật, rồi đẩy kết quả trở lại Drive — đồng bộ **cả file năm lẫn file quý**
trong cùng 1 thư mục con theo mã:

```bash
python gdrive_download.py            # tải toàn bộ <MÃ>.xlsx + <MÃ>_Q.xlsx từ Drive về financials/<MÃ>/
python gdrive_upload.py              # đẩy toàn bộ <MÃ>.xlsx + <MÃ>_Q.xlsx trong financials/ lên lại Drive
```

**Cấu trúc thư mục trên Google Drive** phải khớp với cấu trúc local
(`GDRIVE_FOLDER_ID` là thư mục **gốc**, bên trong là các thư mục con theo
mã; mỗi thư mục con chứa file `<MÃ>.xlsx` (dữ liệu năm) và, nếu bạn muốn
theo dõi thêm theo quý cho mã đó, thêm file `<MÃ>_Q.xlsx` tự tạo với cấu
trúc tương tự — cả 2 file trong cùng 1 thư mục con):

```
<thư mục gốc trên Drive (GDRIVE_FOLDER_ID)>/
├── FPT/
│   └── FPT.xlsx
├── HPG/
│   └── HPG.xlsx
├── MWG/
│   ├── MWG.xlsx
│   └── MWG_Q.xlsx        <- tự tạo thêm, cấu trúc tương tự, dữ liệu theo quý
└── ...
```

`gdrive_download.py` tự dò từng thư mục con để tải đúng file về
`financials/<MÃ>/<MÃ>.xlsx` và `financials/<MÃ>/<MÃ>_Q.xlsx` — mã nào chưa
có file `_Q` trên Drive thì chỉ tải phần dữ liệu năm, không báo lỗi.
`gdrive_upload.py` tự tìm (hoặc tạo mới nếu chưa có) thư mục con cùng tên
mã rồi upload/ghi đè các file `.xlsx` đang có cục bộ (năm và/hoặc quý) vào
đúng thư mục con đó — không cần tự tạo tay thư mục con trên Drive trước.

Cần 2 biến môi trường: `GDRIVE_SA_KEY` (nội dung JSON key của 1 Google
Service Account) và `GDRIVE_FOLDER_ID` (ID thư mục Drive **gốc**). Xem
hướng dẫn thiết lập từng bước ở đầu file `gdrive_utils.py`.

**Thứ tự chạy đầy đủ** (đúng như trong workflow):
`gdrive_download.py` → `fetch_full_financials.py` (year) →
`merge_financials.py` (year) → `fetch_full_financials.py` (quarter) →
`merge_financials.py` (quarter) → `gdrive_upload.py` → xoá CSV tạm →
commit vào Git. Bước quarter có thể tắt bằng input `fetch_quarter: false`
khi chạy workflow thủ công (tab Actions).

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
├── fetch_full_financials.py     # lấy BCTC đầy đủ từ vnstock -> CSV tạm (period: year | quarter)
├── merge_financials.py          # gộp 3 CSV/mã -> 1 file Excel (3 sheet), merge dữ liệu qua thời gian, giữ nguyên các sheet khác (period: year | quarter)
├── gdrive_utils.py               # hàm dùng chung để xác thực Google Drive (Service Account)
├── gdrive_download.py            # tải <MÃ>.xlsx + <MÃ>_Q.xlsx từ Google Drive về (chạy trước merge)
├── gdrive_upload.py              # đẩy <MÃ>.xlsx + <MÃ>_Q.xlsx đã cập nhật lên lại Google Drive (chạy sau merge)
├── financials/<MÃ>/<MÃ>.xlsx     # dữ liệu theo NĂM, do merge_financials.py tạo (period=year)
├── financials/<MÃ>/<MÃ>_Q.xlsx   # dữ liệu theo QUÝ, do merge_financials.py tạo (period=quarter) — tự tạo thêm khi cần
├── requirements.txt
├── .gitignore                   # loại financials/**/*.csv (chỉ là file tạm, gồm cả CSV quý _Q)
└── README.md
```
