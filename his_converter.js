/**
 * his_converter.js
 * Chuyển đổi dữ liệu JSON danh mục dịch vụ VNPT-HIS sang file Excel chuẩn mẫu:
 * 1. Tên sheet chính: 'pttt'
 * 2. Format & Style chuẩn mẫu 100%:
 *    - Hàng 1: Chữ ĐỎ, gộp ô (Merge A1:O1), wrap text
 *    - Hàng 2 & 3: Nền VÀNG (#FFFF00), viền đen mỏng, chữ ĐỎ in đậm ở Hàng 2
 *    - Merge cell theo chiều dọc (Row 2 & 3) cho các cột không có chú thích ở Hàng 3
 *    - Hàng 4: Dòng ví dụ mẫu chuẩn
 *    - Hàng 5 trở đi: Dữ liệu thực tế với:
 *        + Cột DENNGAY = '20500101'
 *        + Cột TEN_TT37 = TENDICHVU
 *        + Cột CHUYENKHOAID = 2
 * 3. Danh sách 5 sheet danh mục tiếp theo đúng thứ tự:
 *    - 'Nhom BHYT'
 *    - 'Loai PTTT'
 *    - 'Ma Nhom'
 *    - 'Ma Chuyen Khoa'
 *    - 'Nhom KT'
 */

const ExcelJS = require('exceljs');

// Chuẩn hóa ngày sang định dạng YYYYMMDD
function toYYYYMMDD(dateVal) {
  if (!dateVal) return '';

  if (typeof dateVal === 'number') {
    const d = new Date(dateVal);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}${m}${day}`;
    }
  }

  if (typeof dateVal === 'string') {
    const trimmed = dateVal.trim();
    // DD/MM/YYYY
    const ddmmyyyy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (ddmmyyyy) {
      const day = ddmmyyyy[1].padStart(2, '0');
      const m = ddmmyyyy[2].padStart(2, '0');
      const y = ddmmyyyy[3];
      return `${y}${m}${day}`;
    }
    // YYYY-MM-DD
    const yyyymmdd = trimmed.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (yyyymmdd) {
      const y = yyyymmdd[1];
      const m = yyyymmdd[2].padStart(2, '0');
      const day = yyyymmdd[3].padStart(2, '0');
      return `${y}${m}${day}`;
    }
    if (/^\d{8}$/.test(trimmed)) {
      return trimmed;
    }
  }
  return '';
}

// Bảng tra mã nhóm BHYT theo LOAI_DV (Sheet 'Nhom BHYT')
const LOAI_DV_BHYT_MAP = {
  'KB': 1,
  'GIUONG_NGAY': 2,
  'XN': 3,
  'CDHA': 4,
  'TDCN': 5,
  'PT': 6,
  'DVKT_TYLE': 7,
  'MAU': 8,
  'THUOC': 10,
  'VT': 15,
  'VTYT': 14,
  'VC': 19,
  'GIUONG_NOITRU': 20,
  'DV_KHAC': 21,
  'PHUTHU': 22,
  'GIUONG_LUU': 24,
  'CHEPHAM_MAU': 25,
  'TT': 26
};

// Bảng tra mã nhóm KT (KHOANMUCID trong Sheet 'Nhom KT')
const LOAI_DV_KHOANMUC_MAP = {
  'KB': 963,
  'XN': 964,
  'CDHA': 965,
  'TDCN': 965,
  'PT': 966,
  'TT': 966,
  'THUOC': 967,
  'VT': 967,
  'VTYT': 967,
  'PHUTHU': 968,
  'KHAC': 969
};

const THIN_BORDER = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } }
};

const YELLOW_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFF00' } // Màu vàng rực chuẩn Excel mẫu
};

async function convertHisJsonToWorkbook(rawList) {
  const jsonData = Array.isArray(rawList) ? rawList : [];
  const wb = new ExcelJS.Workbook();
  wb.creator = 'VNPT-HIS Tool';
  wb.lastModifiedBy = 'VNPT-HIS Tool';
  wb.created = new Date();
  wb.modified = new Date();

  // === 1. SHEET CHÍNH: BẮT BUỘC LÀ 'pttt' ===
  const ws = wb.addWorksheet('pttt', {
    views: [{ showGridLines: true }]
  });

  const dmcHeaders = [
    'KHOA', 'LOAINHOMDICHVU', 'DAXOA', 'CSYTID', 'LOAIAPDUNGGIAMOI', 'MADICHVU',
    'TENDICHVU', 'TENDICHVUBHYT', 'TENDICHVUNHANDAN', 'TENDICHVUNUOCNGOAI',
    'MANHOM', 'DONVI', 'GIABHYT', 'NGAYGIACU', 'GIADICHVU', 'NGAYSD_GIAYC',
    'GIANHANDAN', 'NGAYSD_GIAVP', 'GIANUOCNGOAI', 'NGAY_GIA_NUOCNGOAI',
    'NHOM_MABHYT_ID', 'MADMBYT', 'QUYETDINH', 'NGAYCONGBO', 'STTMAU21',
    'MABYTMAU21', 'MATT37', 'TEN_TT37', 'MATT4350', 'CHUYENKHOAID',
    'LOAIPTTT', 'NHOMKT', 'NOTE', 'QDQUYTRINH', 'SLCHUYENGIAO',
    'CSKCBCHUYENGIAO', 'CSKCBGUIDV', 'QUYETDINHDVKT', 'QUYETDINHGIA',
    'TUNGAY', 'DENNGAY', 'DMGUIBH', 'NGAYGUIBH', 'STT_DICHVU',
    'THUTUIN', 'GIABHYTCU', 'GIADICHVUCU', 'GIANHANDANCU', 'GIANUOCNGOAICU', 'SDTHUOC'
  ];

  const dmcGuides = [
    'Mặc định 0', 'Bắt buộc là 5', 'Mặc định 0', 'Mặc định 1', 'Mặc định 1', 'Không được nhập trùng',
    '', '', '', '',
    'Nhập theo mã nhóm. Nếu là nhóm mới thì khai báo mã và tên ở sheet Ma Nhom', '', '',
    '(Bắt buộc nếu có GIABHYT)\nđịnh dạng YYYYMMDD', '',
    '(Bắt buộc nếu có GIADICHVU)\nđịnh dạng YYYYMMDD', '',
    '(Bắt buộc nếu có GIANHANDAN)\nđịnh dạng YYYYMMDD', '',
    '(Bắt buộc nếu có GIANUOCNGOAI)\nđịnh dạng YYYYMMDD',
    'Nhập theo MANHOM của sheet Nhom BHYT', 'vd: 08.0402.0280', '',
    'định dạng YYYYMMDD', '', '', '', '', '',
    'Nhập theo MACHUYENKHOA của sheet Ma Chuyen Khoa',
    'Nhập theo MALOAIPTTT của sheet Loai PTTT',
    'Nhập theo MAKHOANMUC của sheet Nhom KT',
    '', '', '', '', '', '', '',
    'định dạng YYYYMMDD', 'định dạng YYYYMMDD', '',
    '(Bắt buộc nếu DMGUIBH = 1)\nđịnh dạng YYYYMMDD', 'Mặc định 999',
    '', '', '', '', '', ''
  ];

  // Chiều rộng cột
  ws.columns = dmcHeaders.map((header, idx) => {
    let width = 14;
    if (['TENDICHVU', 'TENDICHVUBHYT', 'TENDICHVUNHANDAN', 'TENDICHVUNUOCNGOAI', 'TEN_TT37'].includes(header)) width = 38;
    else if (['MADMBYT', 'MATT37', 'MATT4350', 'MABYTMAU21', 'MADICHVU'].includes(header)) width = 16;
    else if (['KHOA', 'DAXOA', 'CSYTID', 'DONVI', 'MANHOM', 'DMGUIBH'].includes(header)) width = 10;
    else if (header.includes('GIA')) width = 14;
    else if (header.includes('NGAY')) width = 14;
    return { width };
  });

  // HÀNG 1: Ghi chú chữ ĐỎ, Merge A1:O1, WrapText
  const noteText = "Có thể thêm mã nhóm mới vào Sheet 'Ma Nhom' tùy theo đơn vị\n" +
                   "Sheet 'Nhom BHYT' không được sửa\n" +
                   "Bắt đầu nhập dữ liệu từ hàng thứ 5 của sheet này\n" +
                   "Cột nào không rõ thì không nhập\n" +
                   "Khi import dữ liệu cần lưu ý thông tin KHOANMUCID trong DMC_DICHVU";

  ws.getRow(1).height = 80;
  ws.getCell('A1').value = noteText;
  ws.getCell('A1').font = { name: 'Arial', size: 9.5, bold: false, color: { argb: 'FFFF0000' } }; // Chữ đỏ
  ws.getCell('A1').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  ws.mergeCells('A1:O1');

  // HÀNG 2: Tiêu đề cột & HÀNG 3: Hướng dẫn
  const row2 = ws.getRow(2);
  const row3 = ws.getRow(3);
  row2.height = 30;
  row3.height = 55;

  for (let c = 1; c <= 50; c++) {
    const colIdx0 = c - 1;
    const headerTitle = dmcHeaders[colIdx0];
    const guideText = dmcGuides[colIdx0];

    const cellRow2 = row2.getCell(c);
    const cellRow3 = row3.getCell(c);

    cellRow2.fill = YELLOW_FILL;
    cellRow2.border = THIN_BORDER;

    cellRow3.fill = YELLOW_FILL;
    cellRow3.border = THIN_BORDER;

    if (!guideText || guideText.trim() === '') {
      // Cột không có hướng dẫn ở hàng 3 -> MERGE CELL Hàng 2 và Hàng 3 theo chiều dọc
      ws.mergeCells(2, c, 3, c);
      cellRow2.value = headerTitle;
      cellRow2.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FFFF0000' } }; // Đỏ, in đậm
      cellRow2.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    } else {
      // Cột có hướng dẫn -> Hàng 2 là tiêu đề, Hàng 3 là hướng dẫn
      cellRow2.value = headerTitle;
      cellRow2.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FFFF0000' } }; // Đỏ, in đậm
      cellRow2.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

      cellRow3.value = guideText;
      cellRow3.font = { name: 'Arial', size: 8.5, bold: false, italic: (c === 6 || c === 11), color: { argb: 'FF000000' } };
      cellRow3.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
  }

  // HÀNG 4: Dòng ví dụ mẫu chuẩn (giữ nguyên chuẩn file mẫu)
  const sampleValues = [
    0, 5, 0, 1, 1, 'pttt1',
    'Dịch vụ PTTT số 1', 'Dịch vụ PTTT số 1', 'Dịch vụ PTTT số 1', 'Dịch vụ PTTT số 1',
    'PT', 'lần', 3310000, '20260101', 3310000, '20260101', 3310000, '20260101', 3310000, '20260101',
    6, '08.0402.0280', '372/QĐ-SYT', '20141022', '', 'CDHA1', '37.1E03.1494', 'XXXX', '', 1,
    '', '', '', 111, 50, '82xxx', '79xxx', 123, 321, '20251212', '20261231', 1, '20260415', 999,
    '', '', '', '', '', ''
  ];

  const row4 = ws.getRow(4);
  row4.height = 20;
  for (let c = 1; c <= 50; c++) {
    const cell = row4.getCell(c);
    cell.value = sampleValues[c - 1];
    cell.border = THIN_BORDER;
    cell.font = { name: 'Arial', size: 9.5 };
    cell.alignment = { vertical: 'middle', horizontal: typeof sampleValues[c - 1] === 'number' ? 'right' : 'center' };
  }

  // HÀNG 5 TRỞ ĐI: DỮ LIỆU THỰC TẾ
  const distinctGroups = new Map();
  distinctGroups.set('PT', 'Phẫu Thuật');
  distinctGroups.set('TT', 'Thủ Thuật');

  const usedCodes = new Set(['pttt1']);

  for (let i = 0; i < jsonData.length; i++) {
    const item = jsonData[i];
    const rowIndex = 5 + i;
    const row = ws.getRow(rowIndex);
    row.height = 20;

    const loaiDv = (item.LOAI_DV || 'TT').trim().toUpperCase();
    if (!distinctGroups.has(loaiDv)) {
      distinctGroups.set(loaiDv, loaiDv === 'KB' ? 'Khám Bệnh' : (loaiDv === 'XN' ? 'Xét Nghiệm' : (loaiDv === 'CDHA' ? 'Chẩn Đoán Hình Ảnh' : loaiDv)));
    }

    const tenDv = (item.TEN_DV || item.TEN_HIEN_THI || '').trim();
    const tenDvBhyt = (item.TEN_DICH_VU_KBHYT || tenDv).trim();
    const donVi = (item.DVT_DV || '').trim();

    // Giá
    let giaBhyt = 0;
    if (item.GIACOBH_TT22 !== undefined && item.GIACOBH_TT22 !== null && item.GIACOBH_TT22 !== '') {
      giaBhyt = Number(item.GIACOBH_TT22);
    } else if (item.GIABAOHIEM !== undefined && item.GIABAOHIEM !== null && item.GIABAOHIEM !== '') {
      giaBhyt = Number(item.GIABAOHIEM);
    } else if (item.GIA_DV !== undefined && item.GIA_DV !== null && item.GIA_DV !== '') {
      giaBhyt = Number(item.GIA_DV);
    }

    let giaDv = 0;
    if (item.GIAKOBH_TT21 !== undefined && item.GIAKOBH_TT21 !== null && item.GIAKOBH_TT21 !== '') {
      giaDv = Number(item.GIAKOBH_TT21);
    } else if (item.GIA_DV !== undefined && item.GIA_DV !== null && item.GIA_DV !== '') {
      giaDv = Number(item.GIA_DV);
    } else if (item.GIAKHONGBH !== undefined && item.GIAKHONGBH !== null && item.GIAKHONGBH !== '') {
      giaDv = Number(item.GIAKHONGBH);
    }

    const giaNhanDan = giaDv;
    const giaNuocNgoai = giaDv;

    // Ngày hiệu lực
    const effectiveDateStr = toYYYYMMDD(item.NGAY_AP_DUNG_7603) ||
                             toYYYYMMDD(item.NGAY_HL_NHOM9234) ||
                             toYYYYMMDD(item.NGAY_AP_DUNG) ||
                             toYYYYMMDD(item.NGAYAPDUNG_TT15) ||
                             toYYYYMMDD(item.NGAYAPDUNG_TT39) ||
                             '20250101';

    const ngayGiaCu = giaBhyt > 0 ? effectiveDateStr : '';
    const ngaySdGiaYc = giaDv > 0 ? effectiveDateStr : '';
    const ngaySdGiaVp = giaNhanDan > 0 ? effectiveDateStr : '';
    const ngayGiaNuocNgoai = giaNuocNgoai > 0 ? effectiveDateStr : '';

    const nhomMaBhytId = LOAI_DV_BHYT_MAP[loaiDv] || (loaiDv === 'PT' ? 6 : 26);
    const maDmByt = (item.MABAOCAO || item.MABAOCAO_BYT || item.MA_DICHVU_KYTHUAT_DMDC || '').trim();
    const maTt37 = (item.MA_THONGTU_37_DMDC || (item.MA_TT15 ? item.MA_TT15.trim() : '') || '').trim();

    // 1 - TEN_TT37 luôn bằng tên dịch vụ theo yêu cầu
    const tenTt37 = tenDv;
    const maTt4350 = (item.MA_THONGTU_43_50_DMDC || '').trim();

    // 2 - CHUYENKHOAID luôn = 2 theo yêu cầu
    const chuyenKhoaId = 2;

    // MADICHVU luôn bằng MADMBYT theo yêu cầu
    const maDichVu = maDmByt || (item.MANOIBO ? String(item.MANOIBO).trim() : (item.MA_DV !== undefined ? String(item.MA_DV).trim() : `DV_${i + 1}`));

    const quyetDinh = (item.SOQUYETDINH || item.SO_QUYET_DINH_TRUOC_MAP || '').trim();
    const dmGuiBh = 1;
    const ngayGuiBh = effectiveDateStr;

    // 3 - Cột DENNGAY cố định = '20500101' theo yêu cầu
    const denNgay = '20500101';
    const nhomKt = LOAI_DV_KHOANMUC_MAP[loaiDv] || item.KHOANMUCID || 966;

    const rowData = [
      0,                            // 1. KHOA: Mặc định 0
      5,                            // 2. LOAINHOMDICHVU: Bắt buộc là 5
      0,                            // 3. DAXOA: Mặc định 0
      1,                            // 4. CSYTID: Mặc định 1
      1,                            // 5. LOAIAPDUNGGIAMOI: Mặc định 1
      maDichVu,                     // 6. MADICHVU: Không được nhập trùng
      tenDv,                        // 7. TENDICHVU
      tenDvBhyt,                    // 8. TENDICHVUBHYT
      tenDv,                        // 9. TENDICHVUNHANDAN
      tenDv,                        // 10. TENDICHVUNUOCNGOAI
      loaiDv,                       // 11. MANHOM: Nhập theo mã nhóm
      donVi,                        // 12. DONVI
      giaBhyt,                      // 13. GIABHYT
      ngayGiaCu,                    // 14. NGAYGIACU: YYYYMMDD
      giaDv,                        // 15. GIADICHVU
      ngaySdGiaYc,                  // 16. NGAYSD_GIAYC: YYYYMMDD
      giaNhanDan,                   // 17. GIANHANDAN
      ngaySdGiaVp,                  // 18. NGAYSD_GIAVP: YYYYMMDD
      giaNuocNgoai,                 // 19. GIANUOCNGOAI
      ngayGiaNuocNgoai,             // 20. NGAY_GIA_NUOCNGOAI: YYYYMMDD
      nhomMaBhytId,                 // 21. NHOM_MABHYT_ID
      maDmByt,                      // 22. MADMBYT
      quyetDinh,                    // 23. QUYETDINH
      '',                           // 24. NGAYCONGBO
      '',                           // 25. STTMAU21
      maDmByt,                      // 26. MABYTMAU21
      maTt37,                       // 27. MATT37
      tenTt37,                      // 28. TEN_TT37 (luôn bằng tên dịch vụ)
      maTt4350,                     // 29. MATT4350
      chuyenKhoaId,                 // 30. CHUYENKHOAID (luôn = 2)
      '',                           // 31. LOAIPTTT (Cột không rõ thì không nhập)
      nhomKt,                       // 32. NHOMKT (KHOANMUCID = 966)
      item.KYHIEUDV || '',          // 33. NOTE
      '',                           // 34. QDQUYTRINH
      '',                           // 35. SLCHUYENGIAO
      '',                           // 36. CSKCBCHUYENGIAO
      '',                           // 37. CSKCBGUIDV
      '',                           // 38. QUYETDINHDVKT
      '',                           // 39. QUYETDINHGIA
      effectiveDateStr,             // 40. TUNGAY
      denNgay,                      // 41. DENNGAY (luôn = 20500101)
      dmGuiBh,                      // 42. DMGUIBH: 1
      ngayGuiBh,                    // 43. NGAYGUIBH: YYYYMMDD
      999,                          // 44. STT_DICHVU: Mặc định 999
      item.SOLANTHUCHIEN || '',     // 45. THUTUIN
      '',                           // 46. GIABHYTCU
      '',                           // 47. GIADICHVUCU
      '',                           // 48. GIANHANDANCU
      '',                           // 49. GIANUOCNGOAICU
      ''                            // 50. SDTHUOC
    ];

    for (let c = 1; c <= 50; c++) {
      const val = rowData[c - 1];
      const cell = row.getCell(c);
      cell.value = val;
      cell.border = THIN_BORDER;
      cell.font = { name: 'Arial', size: 9.5 };

      // Định dạng canh lề
      if (typeof val === 'number') {
        cell.alignment = { vertical: 'middle', horizontal: [13, 15, 17, 19].includes(c) ? 'right' : 'center' };
      } else if ([7, 8, 9, 10, 28].includes(c)) {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  }

  // === HÀM HỖ TRỢ THÊM SHEET DANH MỤC THAM CHIẾU ===
  function addReferenceSheet(sheetName, headers, rowsData, colWidths = [15, 35]) {
    const refWs = wb.addWorksheet(sheetName, { views: [{ showGridLines: true }] });
    refWs.columns = headers.map((h, idx) => ({ width: colWidths[idx] || 25 }));

    const headRow = refWs.getRow(1);
    headRow.height = 24;
    headers.forEach((h, idx) => {
      const cell = headRow.getCell(idx + 1);
      cell.value = h;
      cell.font = { name: 'Arial', size: 9.5, bold: true };
      cell.border = THIN_BORDER;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    rowsData.forEach((r, rIdx) => {
      const row = refWs.getRow(rIdx + 2);
      row.height = 20;
      r.forEach((val, cIdx) => {
        const cell = row.getCell(cIdx + 1);
        cell.value = val;
        cell.font = { name: 'Arial', size: 9.5 };
        cell.border = THIN_BORDER;
        cell.alignment = { vertical: 'middle', horizontal: typeof val === 'number' ? 'center' : 'left' };
      });
    });
  }

  // === 2. SHEET Nhom BHYT (Thứ tự đúng chuẩn file mẫu) ===
  const nhomBhytRows = [
    [1, 'Khám bệnh'],
    [2, 'Ngày giường bệnh ban ngày'],
    [3, 'Xét nghiệm'],
    [4, 'Chẩn đoán hình ảnh'],
    [5, 'Thăm dò chức năng'],
    [6, 'Phẫu thuật'],
    [7, 'DVKT thanh toán theo tỷ lệ'],
    [8, 'Máu'],
    [9, 'Thuốc, dịch truyền'],
    [10, 'Thuốc'],
    [11, 'Thuốc ngoài danh mục'],
    [12, 'Thuốc điều trị ung thư, chống thải ghép ngoài danh mục'],
    [13, 'Thuốc thanh toán theo tỷ lệ'],
    [14, 'Vật tư y tế'],
    [15, 'Vật tư'],
    [16, 'Vật tư trong danh mục thay thế'],
    [17, 'Vật tư ngoài danh mục'],
    [18, 'VTYT thanh toán theo tỷ lệ'],
    [19, 'Vận chuyển'],
    [20, 'Ngày giường bệnh điều trị nội trú'],
    [21, 'Dịch vụ khác'],
    [22, 'Phụ thu'],
    [23, 'Thuốc vật tư mua ngoài'],
    [24, 'Giường lưu'],
    [25, 'Chế Phẩm máu'],
    [26, 'Thủ thuật']
  ];
  addReferenceSheet('Nhom BHYT', ['MANHOM', 'TENNHOM'], nhomBhytRows, [14, 45]);

  // === 3. SHEET Loai PTTT ===
  const loaiPtttRows = [
    [1, 'Thủ thuật loại 1'],
    [2, 'Thủ thuật loại 2'],
    [3, 'Thủ thuật loại 3'],
    [4, 'Thủ thuật đặc biệt'],
    [5, 'Phẫu thuật loại 1'],
    [6, 'Phẫu thuật loại 2'],
    [7, 'Phẫu thuật loại 3'],
    [8, 'Phẫu thuật đặc biệt'],
    [0, 'Loại khác'],
    [9, 'Hao phí']
  ];
  addReferenceSheet('Loai PTTT', ['MALOAIPTTT', 'TENLOAIPTTT'], loaiPtttRows, [14, 30]);

  // === 4. SHEET Ma Nhom ===
  const maNhomRows = Array.from(distinctGroups.entries()).map(([k, v]) => [k, v]);
  addReferenceSheet('Ma Nhom', ['MANHOM', 'TENNHOM'], maNhomRows, [14, 25]);

  // === 5. SHEET Ma Chuyen Khoa ===
  const maChuyenKhoaRows = [
    [1, 'Hồi sức cấp cứu và Chống độc'],
    [2, 'Nội khoa'],
    [3, 'Nhi khoa'],
    [4, 'Lao (ngoại lao)'],
    [5, 'Da liễu'],
    [6, 'Tâm thần'],
    [7, 'Nội tiết'],
    [8, 'Y học cổ truyền'],
    [9, 'Gây mê hồi sức'],
    [10, 'Ngoại khoa'],
    [11, 'Bỏng'],
    [12, 'Ung bướu'],
    [13, 'Phụ sản'],
    [14, 'Mắt'],
    [15, 'Tai mũi họng'],
    [16, 'Răng hàm mặt'],
    [17, 'Phục hồi chức năng'],
    [18, 'Điện quang'],
    [19, 'Y học hạt nhân'],
    [20, 'Nội soi chẩn đoán, can thiệp'],
    [21, 'Thăm dò chức năng'],
    [22, 'Huyết học - truyền máu'],
    [23, 'Hóa sinh'],
    [24, 'Vi sinh, ký sinh trùng'],
    [25, 'Giải phẫu bệnh và Tế bào bệnh học'],
    [26, 'Vi phẫu'],
    [27, 'Phẫu thuật nội soi'],
    [28, 'Tạo hình- Thẩm mỹ'],
    [102, 'Nội khoa'],
    [105, 'Da liễu'],
    [110, 'Ngoại khoa'],
    [113, 'Phụ sản'],
    [114, 'Mắt'],
    [115, 'Tai mũi họng'],
    [116, 'Răng hàm mặt'],
    [129, 'Cơ Xương Khớp'],
    [130, 'Kết luận'],
    [131, 'Nội tim mạch'],
    [132, 'Nội Tiêu hóa'],
    [133, 'Nội Thận-Tiết niệu'],
    [134, 'Thần kinh'],
    [135, 'Dị ứng'],
    [136, 'Truyền nhiễm'],
    [137, 'Chấn thương chỉnh hình'],
    [138, 'Nội cơ xương khớp'],
    [100, 'Xét nghiệm'],
    [101, 'CĐHA']
  ];
  addReferenceSheet('Ma Chuyen Khoa', ['MACHUYENKHOA', 'TENCHUYENKHOA'], maChuyenKhoaRows, [16, 40]);

  // === 6. SHEET Nhom KT ===
  const nhomKtRows = [
    [962, 'koxacdinh'],
    [963, 'Khám bệnh'],
    [964, 'Xét nghiệm'],
    [965, 'Chẩn đoán hình ảnh & Thăm dò chức năng'],
    [966, 'Phẫu thuật Thủ thuật'],
    [967, 'Tiền thuốc, vật tư'],
    [968, 'Phụ thu'],
    [969, 'Khác']
  ];
  addReferenceSheet('Nhom KT', ['KHOANMUCID', 'TENKHOANMUC'], nhomKtRows, [16, 45]);

  return {
    workbook: wb,
    totalRecords: jsonData.length,
    groups: Array.from(distinctGroups.entries())
  };
}

module.exports = {
  convertHisJsonToWorkbook,
  toYYYYMMDD
};
