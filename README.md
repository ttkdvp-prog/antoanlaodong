# Hệ thống Theo dõi Tiến độ ATVSLĐ & 5S

Webapp theo dõi tiến độ an toàn vệ sinh lao động và 5S, kết nối trực tiếp với Google Sheets qua Apps Script.

## 🌐 Demo

Xem live tại: [antoanlaodong.vercel.app](https://antoanlaodong.vercel.app)

## 📋 Tính năng

- **Tổng quan Dashboard**: Biểu đồ tiến độ tổng hợp, lịch kiểm tra, ma trận việc trọng yếu
- **Checklist Đơn vị**: Xem và cập nhật checklist theo từng đơn vị trực tiếp lên Google Sheets
- **Tự động đồng bộ**: Cập nhật dữ liệu mỗi 30 giây
- **Dark Mode**: Hỗ trợ chế độ tối/sáng
- **Responsive**: Tương thích mọi thiết bị

## 🚀 Cài đặt

### 1. Triển khai Backend (Apps Script)

1. Mở Google Sheet: https://docs.google.com/spreadsheets/d/1z-D47VH0ALnfbbA-5_yso73g_t7eVPc8FzAj5Qdiyfc/edit
2. Chọn **Tiện ích mở rộng** → **Apps Script**
3. Dán nội dung file `Code.gs` vào editor
4. Vào **Cài đặt dự án** → **Thuộc tính tập lệnh** → Thêm 2 thuộc tính:
   - `API_KEY`: Mã bí mật tự chọn (VD: `vnpt_secret_key_2026`)
   - `SPREADSHEET_ID`: `1z-D47VH0ALnfbbA-5_yso73g_t7eVPc8FzAj5Qdiyfc`
5. **Triển khai** → **Triển khai mới** → **Ứng dụng web**:
   - Thực thi dưới dạng: **Tôi (Me)**
   - Quyền truy cập: **Bất kỳ ai (Anyone)**
6. Copy URL Web App được tạo ra

### 2. Cấu hình Frontend

Mở webapp và nhập:
- **Google Web App URL**: URL đã copy ở bước trên
- **API Key**: Mã bí mật đã đặt ở Script Properties

## 🛠️ Tech Stack

- **Frontend**: HTML, CSS (Tailwind), Vanilla JavaScript
- **Backend**: Google Apps Script (Web App)
- **Database**: Google Sheets
- **Hosting**: Vercel (Static)
- **Charts**: Chart.js
- **Icons**: Lucide Icons

## 📁 Cấu trúc

```
├── index.html          # Webapp chính (SPA)
├── Code.gs             # Apps Script backend
├── appsscript.json     # Cấu hình Apps Script
├── vercel.json         # Cấu hình Vercel deployment
└── README.md           # Hướng dẫn này
```

## 🔐 Bảo mật

- API Key được lưu trong localStorage của trình duyệt
- Apps Script kiểm tra API Key trên mỗi request
- Không có dữ liệu nhạy cảm được lưu trên server

---

Made with ❤️ by VNPT Team
