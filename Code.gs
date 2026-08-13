/**
 * Webapp ATVSLĐ & 5S - Apps Script Backend
 * 
 * Cài đặt:
 * 1. Mở Google Sheet: https://docs.google.com/spreadsheets/d/1z-D47VH0ALnfbbA-5_yso73g_t7eVPc8FzAj5Qdiyfc/edit
 * 2. Chọn Tiện ích mở rộng (Extensions) → Apps Script
 * 3. Xóa các mã mặc định và dán toàn bộ nội dung file này vào.
 * 4. Vào Cài đặt dự án (Project Settings - biểu tượng bánh răng) → Thuộc tính tập lệnh (Script Properties) → Thêm 2 thuộc tính:
 *    - API_KEY: [Mã bí mật tự chọn, ví dụ: vnpt_safety_secret_key_2026]
 *    - SPREADSHEET_ID: 1z-D47VH0ALnfbbA-5_yso73g_t7eVPc8FzAj5Qdiyfc
 * 5. Chọn Triển khai (Deploy) → Triển khai mới (New deployment) → Loại triển khai: Ứng dụng web (Web app)
 *    - Thực thi dưới dạng (Execute as): Tôi (Me)
 *    - Quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 * 6. Copy URL của ứng dụng web được sinh ra để cấu hình vào file index.html.
 */

const ENVELOPE_OK = (data, extra) => Object.assign({ ok: true, error: null, data: data }, extra || {});
const ENVELOPE_ERR = (msg) => ({ ok: false, error: String(msg), data: null });

function doGet(e) {
  return handle_(e, e.parameter || {});
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData && e.postData.contents || "{}");
  } catch (err) {}
  const merged = Object.assign({}, e.parameter || {}, body);
  return handle_(e, merged);
}

function handle_(e, p) {
  // CORS configuration
  const origin = e && e.parameter && e.parameter.origin || "*";
  try {
    requireApiKey_(e, p);
    const action = String(p.action || "getOverview");
    let responseData;
    
    switch (action) {
      case "getOverview":
        responseData = getOverview_();
        break;
      case "getChecklist":
        responseData = getChecklist_(p.table);
        break;
      case "updateChecklistRow":
        responseData = updateChecklistRow_(p.table, p.ma, p.data);
        break;
      default:
        return json_(ENVELOPE_ERR("unknown action: " + action), origin);
    }
    return json_(ENVELOPE_OK(responseData), origin);
  } catch (err) {
    return json_(ENVELOPE_ERR(err && err.message || err), origin);
  }
}

function json_(obj, origin) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function requireApiKey_(e, p) {
  const expected = PropertiesService.getScriptProperties().getProperty("API_KEY");
  if (!expected) throw new Error("API_KEY not set in Script Properties");
  const provided = (e && e.parameter && e.parameter.apiKey) || (p && p.apiKey);
  if (provided !== expected) throw new Error("unauthorized: API Key invalid");
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Định dạng ngày giờ thành chuỗi ISO cục bộ
 */
defFormatDate = function(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return val;
};

/**
 * Lấy dữ liệu tổng quan bao gồm lịch kiểm tra, tiến độ tổng hợp, ma trận việc trọng yếu và nhân sự
 */
function getOverview_() {
  const ss = ss_();
  
  // 1. Lịch kiểm tra
  const lichSheet = ss.getSheetByName("Lịch kiểm tra");
  const lichData = [];
  if (lichSheet) {
    const values = lichSheet.getDataRange().getValues();
    // Headers ở dòng 2 (index 1)
    if (values.length > 1) {
      const headers = values[1].map(h => String(h).trim());
      for (let i = 2; i < values.length; i++) {
        const row = values[i];
        if (!row[0] && !row[1]) continue; // bỏ qua dòng trống
        const item = {};
        for (let j = 0; j < headers.length; j++) {
          item[headers[j]] = defFormatDate(row[j]);
        }
        lichData.push(item);
      }
    }
  }
  
  // 2. Tiến độ tổng hợp & Ma trận việc trọng yếu từ sheet "Tổng hợp"
  const tongHopSheet = ss.getSheetByName("Tổng hợp");
  const tienDoData = [];
  const maTranData = [];
  
  if (tongHopSheet) {
    const values = tongHopSheet.getDataRange().getValues();
    
    // Đọc bảng tiến độ tổng hợp: từ dòng 5 (index 4) đến dòng 15 (index 14)
    // Header ở dòng 4 (index 3)
    if (values.length > 3) {
      const headers = values[3].map(h => String(h).trim());
      for (let i = 4; i < 15; i++) {
        if (i >= values.length) break;
        const row = values[i];
        if (!row[1]) continue;
        const item = {};
        for (let j = 0; j < headers.length; j++) {
          item[headers[j]] = defFormatDate(row[j]);
        }
        tienDoData.push(item);
      }
    }
    
    // Đọc bảng ma trận việc trọng yếu: bắt đầu từ dòng 20 (index 19) đến dòng 30 (index 29)
    // Header ở dòng 19 (index 18)
    if (values.length > 18) {
      const headers = values[18].map(h => String(h).trim());
      for (let i = 19; i < 30; i++) {
        if (i >= values.length) break;
        const row = values[i];
        if (!row[0]) continue;
        const item = {};
        for (let j = 0; j < headers.length; j++) {
          item[headers[j]] = row[j];
        }
        maTranData.push(item);
      }
    }
  }
  
  // 3. Danh sách nhân sự
  const nhanSuSheet = ss.getSheetByName("Nhân sự Ban ATVSLĐ-5S");
  const nhanSuData = [];
  if (nhanSuSheet) {
    const values = nhanSuSheet.getDataRange().getValues();
    // Headers ở dòng 4 (index 3)
    if (values.length > 3) {
      const headers = values[3].map(h => String(h).trim());
      for (let i = 4; i < 23; i++) {
        if (i >= values.length) break;
        const row = values[i];
        if (!row[1]) continue;
        const item = {};
        for (let j = 0; j < headers.length; j++) {
          item[headers[j]] = row[j];
        }
        nhanSuData.push(item);
      }
    }
  }
  
  return {
    lichKiemTra: lichData,
    tienDoTongHop: tienDoData,
    maTranTrongYeu: maTranData,
    nhanSu: nhanSuData
  };
}

/**
 * Lấy checklist chi tiết của một Tổ Hạ tầng hoặc Trung tâm
 */
function getChecklist_(sheetName) {
  if (!sheetName) throw new Error("Missing table name");
  const ss = ss_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet not found: " + sheetName);
  
  const values = sh.getDataRange().getValues();
  if (values.length < 13) {
    return { metadata: {}, items: [] };
  }
  
  // Đọc metadata (dòng 1-12)
  const metadata = {
    donVi: sh.getRange("B3").getValue(),
    chuTri: sh.getRange("E3").getValue(),
    tongViec: sh.getRange("I3").getValue(),
    hoanThanh: sh.getRange("I4").getValue(),
    coDuThao: sh.getRange("I5").getValue(),
    dangThucHien: sh.getRange("I6").getValue(),
    chuaThucHien: sh.getRange("I7").getValue(),
    quaHan: sh.getRange("I8").getValue(),
    tyLeSanSang: sh.getRange("I9").getValue(),
    5sHoanThanh: sh.getRange("L3").getValue(),
    xepLoai5s: sh.getRange("L4").getValue(),
    hienTruongScore: sh.getRange("L5").getValue(),
    hoSoScore: sh.getRange("L6").getValue(),
    canhBao: sh.getRange("L7").getValue(),
    ngayKiemTraDau: defFormatDate(sh.getRange("B4").getValue()),
    hanTuRaSoat: defFormatDate(sh.getRange("E4").getValue()),
    hanKhoaChecklist: defFormatDate(sh.getRange("B5").getValue()),
    luuY: sh.getRange("E5").getValue(),
    dieuPhoiVien: sh.getRange("L10").getValue()
  };
  
  // Đọc bảng checklist (Bắt đầu từ dòng 14, header ở dòng 13)
  const headers = values[12].map(h => String(h).trim());
  const items = [];
  
  for (let i = 13; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue; // bỏ qua dòng không có Mã
    const item = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      item[headers[j]] = defFormatDate(val);
    }
    // Lưu số dòng thực tế (1-based index) để cập nhật sau này
    item.__rowNum = i + 1;
    items.push(item);
  }
  
  return {
    metadata: metadata,
    items: items
  };
}

/**
 * Cập nhật một dòng checklist chi tiết
 */
function updateChecklistRow_(sheetName, ma, patch) {
  if (!sheetName) throw new Error("Missing table name");
  if (!ma) throw new Error("Missing item code (Mã)");
  if (!patch) throw new Error("Missing patch data");
  
  const ss = ss_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet not found: " + sheetName);
  
  const values = sh.getDataRange().getValues();
  if (values.length < 14) throw new Error("Sheet has no checklist data");
  
  // Tìm dòng có cột A = ma (bắt đầu từ dòng 14)
  let targetRowIndex = -1;
  for (let i = 13; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(ma).trim()) {
      targetRowIndex = i + 1; // 1-based index của sheet
      break;
    }
  }
  
  if (targetRowIndex === -1) throw new Error("Checklist item code not found: " + ma);
  
  // Cập nhật các cột chỉ định
  // Cột I: Trạng thái (cột thứ 9)
  if (patch.trangThai !== undefined) {
    sh.getRange(targetRowIndex, 9).setValue(patch.trangThai);
  }
  // Cột J: % HT (cột thứ 10)
  if (patch.phanTramHt !== undefined) {
    sh.getRange(targetRowIndex, 10).setValue(Number(patch.phanTramHt));
  }
  // Cột K: Ngày hoàn thành (cột thứ 11)
  if (patch.ngayHoanThanh !== undefined) {
    if (patch.ngayHoanThanh) {
      sh.getRange(targetRowIndex, 11).setValue(new Date(patch.ngayHoanThanh));
    } else {
      sh.getRange(targetRowIndex, 11).clearContent();
    }
  }
  // Cột L: Minh chứng/Đường dẫn (cột thứ 12)
  if (patch.minhChung !== undefined) {
    sh.getRange(targetRowIndex, 12).setValue(patch.minhChung);
  }
  // Cột M: Tồn tại/Ghi chú (cột thứ 13)
  if (patch.ghiChu !== undefined) {
    sh.getRange(targetRowIndex, 13).setValue(patch.ghiChu);
  }
  // Cột P: Điểm tự rà soát (cột thứ 16)
  if (patch.diemTuRaSoat !== undefined) {
    if (patch.diemTuRaSoat === "" || patch.diemTuRaSoat === null) {
      sh.getRange(targetRowIndex, 16).clearContent();
    } else {
      sh.getRange(targetRowIndex, 16).setValue(Number(patch.diemTuRaSoat));
    }
  }
  
  // Trả về dữ liệu chi tiết của sheet sau khi cập nhật để đồng bộ client
  SpreadsheetApp.flush(); // Đảm bảo các công thức được tính toán lại ngay
  return getChecklist_(sheetName);
}
