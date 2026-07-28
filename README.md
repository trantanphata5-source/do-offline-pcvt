# D-Office PCVT - Digital Office Offline

Hệ thống quản lý văn bản đến **Công ty Điện lực Vũng Tàu** - phiên bản offline.

## Tính năng

- 📄 Xem danh sách 414 văn bản đến, sắp xếp theo ngày
- 🔍 Tìm kiếm văn bản (số VB, trích yếu, cơ quan ban hành)
- 📋 Xem trước PDF inline
- ✍️ **Chỉ đạo**: Chọn Chủ trì / Phối hợp / Xem để biết cho từng phòng ban
- 🔀 **Chuyển chỉ đạo**: Chuyển văn bản cho lãnh đạo thực hiện
- ℹ️ **Thông tin văn bản**: Xem chi tiết trích yếu, số VB, ngày, cơ quan
- 💾 Lưu trữ dữ liệu trên **Google Sheets** (chia sẻ được cho nhiều người)

## Cấu trúc

```
D-O/
├── app/
│   ├── index.html      # Trang web chính
│   ├── styles.css       # CSS stylesheet
│   ├── app.js           # JavaScript logic
│   └── logo.png         # Logo EVN HCMC
├── documents.json       # Metadata 414 văn bản
├── gas_code.js          # Google Apps Script backend
└── .gitignore
```

## Cài đặt & Chạy

### 1. Clone repo
```bash
git clone https://github.com/YOUR_USERNAME/D-Office-PCVT.git
cd D-Office-PCVT
```

### 2. Chạy local server
```bash
python -m http.server 8080
```

### 3. Mở trình duyệt
```
http://localhost:8080/app/index.html
```

### 4. (Tùy chọn) Kết nối Google Sheets
1. Mở Google Sheet → Extensions → Apps Script
2. Paste nội dung `gas_code.js`
3. Deploy → Web app → Anyone
4. Copy URL → paste vào `GAS_URL` trong `app/app.js` (dòng 13)

## Công nghệ

- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Backend**: Google Apps Script + Google Sheets
- **Font**: Inter (Google Fonts)
- **Icons**: Material Icons Outlined

## Giấy phép

Dự án nội bộ - Công ty Điện lực Vũng Tàu.
