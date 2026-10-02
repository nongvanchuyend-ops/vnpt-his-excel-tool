/**
 * curl_parser.js
 * Module phân tích lệnh cURL (hỗ trợ định dạng Windows cmd, Bash, PowerShell)
 * và thực thi request để lấy dữ liệu JSON.
 */

// Bỏ qua kiểm tra SSL cho các hệ thống bệnh viện / intranet dùng chứng chỉ nội bộ
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

function parseCurl(rawCurl) {
  if (!rawCurl || typeof rawCurl !== 'string') {
    throw new Error('Lệnh cURL không hợp lệ hoặc đang để trống.');
  }

  let cleaned = rawCurl.trim();

  // 1. Chuẩn hóa ký tự ngắt dòng tiếp tục lệnh:
  // Windows cmd kết thúc dòng bằng ^\r?\n
  // Bash kết thúc dòng bằng \\\r?\n
  // PowerShell kết thúc dòng bằng `\r?\n
  cleaned = cleaned.replace(/[\^\\`]\r?\n\s*/g, ' ');

  // 2. Xử lý escape trong Windows CMD:
  // ^%^ hoặc ^% -> %
  cleaned = cleaned.replace(/\^%\^?/g, '%');
  // ^\^" -> " và ^" -> "
  cleaned = cleaned.replace(/\^\\?\^?"/g, '"');
  // ^& -> &, ^! -> !, ^( -> (, ^) -> )
  cleaned = cleaned.replace(/\^([&!|=()<>])/g, '$1');

  // 3. Phân tách tokens (giữ nguyên chuỗi trong ngoặc kép hoặc ngoặc đơn)
  const tokens = [];
  let current = '';
  let inQuote = null;

  for (let i = 0; i < cleaned.length; i++) {
    const char = cleaned[i];

    if (inQuote) {
      if (char === inQuote) {
        if (i > 0 && cleaned[i - 1] === '\\') {
          // Thoát dấu ngoặc
          current = current.slice(0, -1) + char;
        } else {
          inQuote = null;
        }
      } else {
        current += char;
      }
    } else {
      if (char === '"' || char === "'") {
        inQuote = char;
      } else if (/\s/.test(char)) {
        if (current.length > 0) {
          tokens.push(current);
          current = '';
        }
      } else {
        current += char;
      }
    }
  }
  if (current.length > 0) {
    tokens.push(current);
  }

  let url = '';
  let method = 'GET';
  const headers = {};
  const cookies = [];
  let body = null;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token === 'curl' || token === 'curl.exe') continue;

    if (token === '--url') {
      url = tokens[++i];
    } else if (token === '-H' || token === '--header') {
      const headerLine = tokens[++i] || '';
      const colonIdx = headerLine.indexOf(':');
      if (colonIdx > -1) {
        const key = headerLine.slice(0, colonIdx).trim();
        const val = headerLine.slice(colonIdx + 1).trim();
        headers[key] = val;
      }
    } else if (token === '-b' || token === '--cookie') {
      cookies.push(tokens[++i]);
    } else if (token === '-X' || token === '--request') {
      method = (tokens[++i] || 'GET').toUpperCase();
    } else if (token === '-d' || token === '--data' || token === '--data-raw' || token === '--data-binary') {
      method = 'POST';
      body = tokens[++i];
    } else if (!url && !token.startsWith('-')) {
      url = token;
    }
  }

  // Nếu có cookie từ cờ -b, gộp vào Cookie header
  if (cookies.length > 0) {
    const joinedCookies = cookies.join('; ');
    if (headers['Cookie']) {
      headers['Cookie'] = headers['Cookie'] + '; ' + joinedCookies;
    } else {
      headers['Cookie'] = joinedCookies;
    }
  }

  if (!url) {
    throw new Error('Không tìm thấy URL trong lệnh cURL đã cung cấp.');
  }

  // Làm sạch URL nếu còn sót ký tự ^ hoặc ngoặc kép
  url = url.replace(/^\^?"|\^?"$/g, '').trim();

  return { url, method, headers, body };
}

/**
 * Thực thi lệnh cURL và lấy dữ liệu trả về từ server
 */
async function fetchFromCurl(curlCommand) {
  const reqInfo = parseCurl(curlCommand);

  const fetchOptions = {
    method: reqInfo.method,
    headers: reqInfo.headers
  };

  if (reqInfo.body && reqInfo.method !== 'GET' && reqInfo.method !== 'HEAD') {
    fetchOptions.body = reqInfo.body;
  }

  const response = await fetch(reqInfo.url, fetchOptions);

  if (!response.ok) {
    const errText = await response.text().catch(() => '');
    throw new Error(`Yêu cầu thất bại với mã lỗi HTTP ${response.status} (${response.statusText}).\nChi tiết: ${errText.slice(0, 300)}`);
  }

  const responseText = await response.text();
  try {
    const json = JSON.parse(responseText);
    return {
      success: true,
      reqInfo,
      data: Array.isArray(json) ? json : (json.rows || json.data || [json])
    };
  } catch (err) {
    throw new Error(`Dữ liệu máy chủ trả về không phải định dạng JSON (có thể phiên đăng nhập / Cookie đã hết hạn).\nPhản hồi nhận được:\n${responseText.slice(0, 400)}`);
  }
}

module.exports = {
  parseCurl,
  fetchFromCurl
};
