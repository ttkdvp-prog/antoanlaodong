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
  
  // Đọc metadata từ mảng values[] đã load (nhanh hơn nhiều so với getRange().getValue())
  // B3=values[2][1], E3=values[2][4], I3=values[2][8], L3=values[2][11]...
  const metadata = {
    donVi:            values[2][1],
    chuTri:           values[2][4],
    tongViec:         values[2][8],
    hoanThanh:        values[3][8],
    coDuThao:         values[4][8],
    dangThucHien:     values[5][8],
    chuaThucHien:     values[6][8],
    quaHan:           values[7][8],
    tyLeSanSang:      values[8][8],
    hoanThanh5s:      values[2][11],
    xepLoai5s:        values[3][11],
    hienTruongScore:  values[4][11],
    hoSoScore:        values[5][11],
    canhBao:          values[6][11],
    ngayKiemTraDau:   defFormatDate(values[3][1]),
    hanTuRaSoat:      defFormatDate(values[3][4]),
    hanKhoaChecklist: defFormatDate(values[4][1]),
    luuY:             values[4][4],
    dieuPhoiVien:     values[9][11]
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
  
  // Nếu patch là string JSON (gửi qua GET query param), parse thành object
  if (typeof patch === "string") {
    try { patch = JSON.parse(patch); } catch(e) { throw new Error("Invalid patch data JSON: " + e.message); }
  }
  
  const ss = ss_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet not found: " + sheetName);
  
  const values = sh.getDataRange().getValues();
  if (values.length < 14) throw new Error("Sheet has no checklist data");
  
  // Đọc headers từ dòng 13 (index 12) để tìm đúng vị trí cột
  const headers = values[12].map(h => String(h).trim());
  
  // Tìm chỉ số cột "Mã" trong headers
  const maColIndex = headers.findIndex(h => h === "Mã" || h === "Ma" || h === "MÃ");
  if (maColIndex === -1) throw new Error("Không tìm thấy cột 'Mã' trong header dòng 13");
  
  // Tìm dòng có giá trị cột Mã = ma (bắt đầu từ dòng 14, index 13)
  let targetRowIndex = -1;
  for (let i = 13; i < values.length; i++) {
    if (String(values[i][maColIndex]).trim() === String(ma).trim()) {
      targetRowIndex = i + 1; // 1-based index của sheet
      break;
    }
  }
  
  if (targetRowIndex === -1) throw new Error("Checklist item code not found: " + ma);
  
  // Hàm tiện ích: tìm số cột (1-based) theo tên header
  function colOf(name) {
    const names = Array.isArray(name) ? name : [name];
    for (const n of names) {
      const idx = headers.findIndex(h => h === n);
      if (idx !== -1) return idx + 1; // 1-based
    }
    return -1;
  }
  
  // Cập nhật các cột theo tên header (linh hoạt với mọi cấu trúc sheet)
  const colTrangThai   = colOf(["Trạng thái", "Trạng thái triển khai", "Trang thai"]);
  const colPhanTramHt  = colOf(["% HT", "% Hoàn thành", "% hoan thanh", "Phan tram HT"]);
  const colNgayHt      = colOf(["Ngày hoàn thành", "Ngay hoan thanh"]);
  const colMinhChung   = colOf(["Minh chứng/Đường dẫn", "Minh chung", "Đường dẫn"]);
  const colGhiChu      = colOf(["Tồn tại/Ghi chú", "Ghi chú", "Ghi chu", "Ton tai"]);
  const colDiem        = colOf(["Điểm tự rà soát", "Diem tu ra soat", "Điểm rà soát"]);
  
  if (patch.trangThai !== undefined && colTrangThai > 0) {
    sh.getRange(targetRowIndex, colTrangThai).setValue(patch.trangThai);
  }
  if (patch.phanTramHt !== undefined && colPhanTramHt > 0) {
    sh.getRange(targetRowIndex, colPhanTramHt).setValue(Number(patch.phanTramHt));
  }
  if (patch.ngayHoanThanh !== undefined && colNgayHt > 0) {
    if (patch.ngayHoanThanh) {
      sh.getRange(targetRowIndex, colNgayHt).setValue(new Date(patch.ngayHoanThanh));
    } else {
      sh.getRange(targetRowIndex, colNgayHt).clearContent();
    }
  }
  if (patch.minhChung !== undefined && colMinhChung > 0) {
    sh.getRange(targetRowIndex, colMinhChung).setValue(patch.minhChung);
  }
  if (patch.ghiChu !== undefined && colGhiChu > 0) {
    sh.getRange(targetRowIndex, colGhiChu).setValue(patch.ghiChu);
  }
  if (patch.diemTuRaSoat !== undefined && colDiem > 0) {
    if (patch.diemTuRaSoat === "" || patch.diemTuRaSoat === null) {
      sh.getRange(targetRowIndex, colDiem).clearContent();
    } else {
      sh.getRange(targetRowIndex, colDiem).setValue(Number(patch.diemTuRaSoat));
    }
  }
  
  // Đảm bảo công thức được tính toán lại trước khi trả dữ liệu
  SpreadsheetApp.flush();
  return getChecklist_(sheetName);
}
