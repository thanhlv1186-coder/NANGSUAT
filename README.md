# Dashboard Năng Suất Lắp Đặt
**Điện Máy Xanh · dữ liệu tự động từ `NANG SUAT.xlsx`**

## Cấu trúc file
```
├── index.html        # Entry HTML
├── index.jsx         # React entry point
├── Dashboard.jsx     # Component chính
├── excelData.js      # Parser đọc NANG SUAT.xlsx ở runtime
├── data.js           # Dữ liệu dự phòng + logo
├── NANG SUAT.xlsx    # File Excel nguồn, thay file này để cập nhật dashboard
├── package.json
├── vite.config.js    # Phục vụ/copy file Excel vào dev và dist
└── README.md
```

## Cài đặt & chạy
```bash
npm install
npm run dev
```

Khi chạy dev server, dashboard sẽ đọc `NANG SUAT.xlsx` từ thư mục dự án. Nếu thay file mới cùng tên, Vite sẽ reload trang và dashboard cũng tự kiểm tra lại file mỗi 30 giây.

## Build production
```bash
npm run build
```

Build sẽ copy `NANG SUAT.xlsx` vào `dist/`. Khi deploy bản static, chỉ cần thay file `dist/NANG SUAT.xlsx` cùng tên để dashboard đọc số liệu mới.

## Tính năng
- Lọc theo vùng (ĐBSH / ĐTB / Tất cả)
- KPI tổng quan: nhân sự, ML, SPK, TB ML, cảnh báo
- Top 10 năng suất cao nhất
- Cảnh báo NV/CTV ML <= 30
- Bảng chi tiết nhân sự với nhiều bộ lọc
- Donut cơ cấu NV/CTV/ĐT
- Heatmap đơn hàng theo ngày, tự nhận số ngày từ Excel
- So sánh Đối Tác vs NV+CTV
- Phân phối ML theo dải
