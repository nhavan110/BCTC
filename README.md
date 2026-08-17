# HPG Financial Ratios (vnstock)

Script lấy chỉ số tài chính theo **năm** của mã HPG (Hòa Phát) từ [vnstock](https://github.com/thinh-vu/vnstock)
(nguồn dữ liệu VCI - Vietcap) và xuất ra file CSV gồm:

- **ROE (%)**
- **Tăng trưởng Lợi nhuận sau thuế (%)** (YoY)
- **Biên lợi nhuận gộp (%)**

## Cài đặt

```bash
pip install -r requirements.txt
```

## Chạy

```bash
python fetch_hpg_financials.py        # mặc định mã HPG
python fetch_hpg_financials.py VNM    # hoặc truyền mã khác
```

Kết quả xuất ra file `HPG_financial_ratios.csv` trong thư mục hiện tại.

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
