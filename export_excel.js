/**
 * export_excel.js
 * CLI Runner: Đọc lệnh cURL, gọi API VNPT-HIS và xuất ra file Excel mẫu
 * Sử dụng:
 *   node export_excel.js
 *   node export_excel.js --input curl_input.txt --output DMC_DICHVU_BacKan.xlsx
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { fetchFromCurl } = require('./curl_parser');
const { convertHisJsonToWorkbook } = require('./his_converter');

async function main() {
  const args = process.argv.slice(2);
  let inputFile = path.join(__dirname, 'curl_input.txt');
  let outputFile = path.join(__dirname, 'DMC_DICHVU_output.xlsx');
  let inlineCurl = null;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--input' || args[i] === '-i') {
      inputFile = path.resolve(args[++i]);
    } else if (args[i] === '--output' || args[i] === '-o') {
      outputFile = path.resolve(args[++i]);
    } else if (args[i] === '--curl' || args[i] === '-c') {
      inlineCurl = args[++i];
    }
  }

  let curlContent = inlineCurl;
  if (!curlContent) {
    if (!fs.existsSync(inputFile)) {
      console.error(`[LỖI] Không tìm thấy file cURL đầu vào tại: ${inputFile}`);
      console.log('Bạn có thể đặt lệnh cURL vào file "curl_input.txt" hoặc dùng tham số: node export_excel.js --curl "curl..."');
      process.exit(1);
    }
    curlContent = fs.readFileSync(inputFile, 'utf-8');
  }

  console.log('====================================================');
  console.log('       TOOL XUẤT DANH MỤC DỊCH VỤ VNPT-HIS RA EXCEL');
  console.log('====================================================');
  console.log(`[1] Đang phân tích lệnh cURL...`);

  try {
    console.log(`[2] Đang kết nối tới máy chủ VNPT-HIS để lấy dữ liệu...`);
    const { reqInfo, data } = await fetchFromCurl(curlContent);

    console.log(` -> URL: ${reqInfo.url.substring(0, 85)}...`);
    console.log(` -> Đã lấy thành công ${data.length} dịch vụ từ hệ thống.`);

    console.log(`[3] Đang ánh xạ dữ liệu & tạo định dạng chuẩn (màu vàng, chữ đỏ, merge cells)...`);
    const { workbook, totalRecords, groups } = await convertHisJsonToWorkbook(data);

    console.log(`[4] Đang ghi file Excel...`);
    await workbook.xlsx.writeFile(outputFile);

    console.log('----------------------------------------------------');
    console.log(` HOÀN THÀNH XUẤT FILE THÀNH CÔNG!`);
    console.log(` Đường dẫn file: ${outputFile}`);
    console.log(` Tổng số dòng dữ liệu: ${totalRecords} dịch vụ (bắt đầu từ dòng 5)`);
    console.log(` Các sheet đã tạo: ${workbook.worksheets.map(w => w.name).join(', ')}`);
    console.log(` Nhóm dịch vụ phát hiện: ${groups.map(g => `${g[0]} (${g[1]})`).join(', ')}`);
    console.log('----------------------------------------------------');

    // In xem trước 3 bản ghi đầu tiên
    console.log('\n Xem trước 3 dịch vụ đầu:');
    data.slice(0, 3).forEach((item, idx) => {
      console.log(`  ${idx + 1}. [Mã: ${item.MA_DV}] ${item.TEN_DV} | Giá BHYT: ${item.GIACOBH_TT22 || item.GIA_DV || 0}đ | Giá VP: ${item.GIAKOBH_TT21 || item.GIA_DV || 0}đ | Mã BYT: ${item.MABAOCAO || 'N/A'}`);
    });
    console.log('\n');
  } catch (err) {
    console.error('\n[LỖI THỰC THI]', err.message);
    process.exit(1);
  }
}

main();
