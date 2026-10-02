/**
 * server.js
 * Local Web Server & API cho Tool Trích xuất Dữ liệu VNPT-HIS sang Excel
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { fetchFromCurl } = require('./curl_parser');
const { convertHisJsonToWorkbook } = require('./his_converter');

const PORT = process.env.PORT || 3050;
const PUBLIC_DIR = path.join(__dirname, 'public');

// Bộ đệm lưu trữ workbook gần nhất cho việc tải về
let lastGeneratedBuffer = null;
let lastGeneratedFilename = 'DMC_DICHVU_output.xlsx';

function serveStatic(res, filePath, contentType) {
  fs.readFile(filePath, (err, content) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('404 Not Found');
    } else {
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(content);
    }
  });
}

const server = http.createServer(async (req, res) => {
  const urlObj = new URL(req.url, `http://${req.headers.host}`);
  const pathname = urlObj.pathname;

  // Cho phép CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // 1. Phục vụ giao diện người dùng
  if (req.method === 'GET' && (pathname === '/' || pathname === '/index.html')) {
    return serveStatic(res, path.join(PUBLIC_DIR, 'index.html'), 'text/html; charset=utf-8');
  }

  // 2. Lấy lệnh cURL mẫu
  if (req.method === 'GET' && pathname === '/api/sample-curl') {
    const sampleCurlPath = path.join(__dirname, 'curl_input.txt');
    let sample = '';
    if (fs.existsSync(sampleCurlPath)) {
      sample = fs.readFileSync(sampleCurlPath, 'utf-8');
    }
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    return res.end(JSON.stringify({ curl: sample }));
  }

  // 3. API xử lý cURL và trích xuất Excel
  if (req.method === 'POST' && pathname === '/api/convert') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', async () => {
      try {
        const payload = JSON.parse(body || '{}');
        const curlCommand = payload.curl;
        if (!curlCommand || !curlCommand.trim()) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          return res.end(JSON.stringify({ error: 'Vui lòng nhập lệnh cURL hợp lệ.' }));
        }

        // Gọi API VNPT-HIS
        const { reqInfo, data } = await fetchFromCurl(curlCommand);

        // Chuyển đổi sang Excel với ExcelJS (đầy đủ format vàng, chữ đỏ, merge cells, tên sheet pttt)
        const { workbook, totalRecords, groups } = await convertHisJsonToWorkbook(data);

        // Tạo Buffer file Excel để cho phép tải về
        const excelBuffer = await workbook.xlsx.writeBuffer();
        lastGeneratedBuffer = excelBuffer;
        const nowStr = new Date().toTimeString().slice(0, 8).replace(/:/g, '');
        lastGeneratedFilename = `pttt_${nowStr}.xlsx`;

        // Ghi luôn 1 bản vào thư mục dự án
        const localSavePath = path.join(__dirname, 'DMC_DICHVU_output.xlsx');
        fs.writeFileSync(localSavePath, excelBuffer);

        // Chuẩn bị dữ liệu xem trước
        const previewRows = data.map((item, idx) => {
          const tenDv = (item.TEN_DV || item.TEN_HIEN_THI || '').trim();
          const maDmByt = (item.MABAOCAO || item.MABAOCAO_BYT || item.MA_DICHVU_KYTHUAT_DMDC || '').trim();
          const maDv = maDmByt || (item.MANOIBO ? String(item.MANOIBO).trim() : (item.MA_DV !== undefined ? String(item.MA_DV).trim() : `DV_${idx + 1}`));
          return {
            stt: idx + 1,
            maDv: maDv, // MADICHVU = MADMBYT
            tenDv: tenDv,
            tenTt37: tenDv, // TEN_TT37 luôn bằng tên dịch vụ
            loaiDv: item.LOAI_DV || 'TT',
            donVi: item.DVT_DV || 'lần',
            giaBhyt: item.GIACOBH_TT22 || item.GIABAOHIEM || item.GIA_DV || 0,
            giaDv: item.GIAKOBH_TT21 || item.GIA_DV || item.GIAKHONGBH || 0,
            maDmByt: maDmByt,
            chuyenKhoa: 2 // Luôn = 2 theo yêu cầu
          };
        });

        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        return res.end(JSON.stringify({
          success: true,
          totalRecords,
          groups,
          sheets: workbook.worksheets.map(w => w.name),
          localSavePath,
          previewRows
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        return res.end(JSON.stringify({
          success: false,
          error: err.message
        }));
      }
    });
    return;
  }

  // 4. API Tải file Excel về máy client
  if (req.method === 'GET' && pathname === '/api/download') {
    if (!lastGeneratedBuffer) {
      const localFile = path.join(__dirname, 'DMC_DICHVU_output.xlsx');
      if (fs.existsSync(localFile)) {
        lastGeneratedBuffer = fs.readFileSync(localFile);
      }
    }

    if (!lastGeneratedBuffer) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end('Chưa có file nào được tạo. Vui lòng thực hiện trích xuất trước.');
    }

    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${lastGeneratedFilename}"`,
      'Cache-Control': 'no-cache, no-store, must-revalidate',
      'Pragma': 'no-cache',
      'Expires': '0',
      'Content-Length': lastGeneratedBuffer.length
    });
    return res.end(lastGeneratedBuffer);
  }

  res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('404 Not Found');
});

server.listen(PORT, () => {
  console.log(`=======================================================`);
  console.log(` Web Server Tool VNPT-HIS đang chạy tại:`);
  console.log(` 👉 http://localhost:${PORT}`);
  console.log(`=======================================================`);
});
