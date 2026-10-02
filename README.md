# TOOL TRÍCH XUẤT DANH MỤC DỊCH VỤ VNPT-HIS SANG EXCEL

Tool hỗ trợ nhận đầu vào là **lệnh cURL** từ trình duyệt (khi xem màn hình danh mục dịch vụ trên VNPT-HIS), tự động kết nối bóc tách dữ liệu và xuất ra file Excel chuẩn mẫu import hệ thống.

---

## 📁 Cấu trúc thư mục

```
vnpt_his_excel_tool/
├── run_tool.bat             # File kích hoạt 1-click mở giao diện Web trên Windows
├── server.js                # Web Server cục bộ phục vụ giao diện & API
├── export_excel.js          # Công cụ chạy dòng lệnh (CLI)
├── curl_parser.js           # Module phân tích cURL (hỗ trợ cmd, bash, powershell)
├── his_converter.js         # Module ánh xạ dữ liệu JSON sang 6 Sheet Excel chuẩn
├── curl_input.txt           # File chứa lệnh cURL mẫu (hoặc dán cURL mới vào đây)
├── DMC_DICHVU_output.xlsx   # File Excel đã xuất sẵn từ cURL mẫu của bạn
├── package.json
└── public/
    └── index.html           # Giao diện Web trực quan, có xem trước dữ liệu & tải file
```

---

## 🚀 Hướng dẫn sử dụng

### Cách 1: Sử dụng Giao diện Web (Khuyên dùng - Rất trực quan)
1. Nhấp đúp chuột vào file **`run_tool.bat`** (hoặc mở terminal gõ `npm start`).
2. Trình duyệt sẽ tự động mở trang: `http://localhost:3000`.
3. Dán lệnh cURL mới từ DevTools (hoặc bấm nút **"Dán cURL mẫu"**).
4. Bấm **"Thực hiện Lấy dữ liệu & Xuất Excel"**.
5. Xem trước danh sách dịch vụ trên bảng và bấm nút **"Tải File Excel (.xlsx)"** để tải về máy.

---

### Cách 2: Sử dụng dòng lệnh (CLI)
- Mở PowerShell / CMD tại thư mục `vnpt_his_excel_tool`:
```bash
node export_excel.js
```
- File kết quả sẽ được tạo tại: `DMC_DICHVU_output.xlsx`.
- Để xuất với tên file tùy chỉnh hoặc cURL khác:
```bash
node export_excel.js --input "duong_dan_file_curl.txt" --output "DMC_DICHVU_Thang10.xlsx"
```

---

## 📋 Quy chuẩn dữ liệu Excel được tạo

File Excel được tạo có đầy đủ **6 Sheets**:

1. **`DMC_DICHVU`**: Sheet chính chứa dữ liệu dịch vụ:
   - **Hàng 1**: Hướng dẫn & lưu ý hệ thống.
   - **Hàng 2**: Tiêu đề 50 cột (`KHOA`, `LOAINHOMDICHVU`, `DAXOA`, `CSYTID`, `LOAIAPDUNGGIAMOI`, `MADICHVU`, `TENDICHVU`, `GIABHYT`, `GIADICHVU`, `MADMBYT`, `NHOMKT`...).
   - **Hàng 3**: Hướng dẫn định dạng từng cột (`YYYYMMDD`, bắt buộc, mặc định...).
   - **Hàng 4**: Dòng dữ liệu mẫu (`pttt1`, `3310000`...).
   - **Hàng 5 trở đi**: Dữ liệu trích xuất từ VNPT-HIS.
   - **Các giá trị mặc định**: `KHOA=0`, `LOAINHOMDICHVU=5`, `DAXOA=0`, `CSYTID=1`, `LOAIAPDUNGGIAMOI=1`, `STT_DICHVU=999`.
   - **CHUYENKHOAID**: Luôn bằng `2` (theo phân hệ đơn vị).
   - **TEN_TT37**: Luôn bằng Tên dịch vụ (`TENDICHVU`).
   - **KHOANMUCID (NHOMKT)**: Tự động gán `966` (Phẫu thuật / Thủ thuật) theo đúng danh mục Nhom KT.
   - **Định dạng ngày**: Chuẩn hóa ngày hiệu lực thành `YYYYMMDD` (VD: `20250101`).

2. **`Ma Nhom`**: Bảng mã nhóm (PT, TT...). Tự động nhận diện và cập nhật thêm mã nhóm mới nếu có.
3. **`Nhom BHYT`**: 26 nhóm chuẩn Bộ Y Tế (Khám, Ngày giường, Xét nghiệm, CĐHA, PT, TT...).
4. **`Loai PTTT`**: Danh mục phân loại phẫu thuật, thủ thuật loại 1, 2, 3, đặc biệt...
5. **`Ma Chuyen Khoa`**: Danh mục mã chuyên khoa (Ngoại, Nội, Nhi, RHM, Mắt...).
6. **`Nhom KT`**: Danh mục khoản mục chi phí (962 - 969).
