/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT - HỆ THỐNG XÁC THỰC LICENSE KEY & QUẢN LÝ AFFILIATE QL SOCIAL
 * ==============================================================================
 * HƯỚNG DẪN CÀI ĐẶT NHANH (Chỉ mất 2 phút):
 * 1. Mở file Google Sheet của bạn (chứa danh sách License Key).
 * 2. Trên thanh menu, chọn: Tiện ích mở rộng (Extensions) > Apps Script.
 * 3. Xóa hết mã cũ trong file Code.gs, dán toàn bộ đoạn mã này vào.
 * 4. Bấm nút "Triển khai" (Deploy) ở góc trên bên phải > Chọn "Triển khai mới" (New deployment).
 * 5. Chọn loại: "Ứng dụng web" (Web App).
 *    - Mô tả: "QL Social Studio API"
 *    - Thực thi dưới dạng (Execute as): "Tôi" (Me)
 *    - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone) - RẤT QUAN TRỌNG!
 * 6. Bấm "Triển khai" (Deploy) > Cấp quyền truy cập (Authorize access).
 * 7. Sao chép "URL Ứng dụng web" (dạng https://script.google.com/macros/s/.../exec) và dán vào Landing Page!
 * ==============================================================================
 */

// Tên các trang tính (Sheet tabs) trong file Google Sheet của bạn
const SHEET_CONFIG = {
  LICENSE_TAB_INDEX: 0, // Mặc định tìm ở tab đầu tiên (hoặc đặt tên cụ thể bên dưới)
  AFFILIATE_TAB_NAME: "DangKy_Affiliate", // Tự động tạo tab này để lưu người đăng ký Affiliate
  ORDERS_TAB_NAME: "DonHang_DatMua"       // Tự động tạo tab này để lưu đơn đặt hàng
};

/**
 * Xử lý yêu cầu GET (Kiểm tra License Key, lấy danh sách...)
 */
function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || "check_license";

    // 1. ACTION: KIỂM TRA LICENSE KEY CÓ TỒN TẠI KHÔNG
    if (action === "check_license") {
      const key = (params.key || "").trim().toUpperCase();
      if (!key || key.length < 3) {
        return createJsonResponse({ valid: false, message: "Vui lòng nhập License Key hợp lệ" });
      }

      const checkResult = verifyLicenseKeyInSheet(key);
      return createJsonResponse(checkResult);
    }

    // 2. ACTION: LƯU ĐĂNG KÝ AFFILIATE (Dự phòng GET nếu form gửi GET)
    if (action === "register_affiliate") {
      const result = saveAffiliateToSheet(params);
      return createJsonResponse(result);
    }

    // 3. ACTION: TEST KẾT NỐI
    if (action === "ping") {
      return createJsonResponse({ success: true, message: "Google Apps Script API đang hoạt động tốt!" });
    }

    return createJsonResponse({ success: false, message: "Hành động không hợp lệ" });
  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Xử lý yêu cầu POST (Lưu đăng ký Affiliate, Lưu Đơn hàng)
 */
function doPost(e) {
  try {
    let data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    const action = data.action || e.parameter?.action || "register_affiliate";

    // 1. ACTION: LƯU THÔNG TIN ĐĂNG KÝ AFFILIATE MỚI
    if (action === "register_affiliate") {
      const result = saveAffiliateToSheet(data);
      return createJsonResponse(result);
    }

    // 2. ACTION: LƯU ĐƠN ĐẶT HÀNG MỚI
    if (action === "create_order") {
      const result = saveOrderToSheet(data);
      return createJsonResponse(result);
    }

    return createJsonResponse({ success: false, message: "Hành động POST không hợp lệ" });
  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Hàm tìm kiếm License Key trong Google Sheet (Bảo mật tối đa, chỉ so khớp chính xác)
 */
function verifyLicenseKeyInSheet(targetKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Lấy sheet đầu tiên chứa License
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return { valid: false, message: "Chưa có dữ liệu License Key trong Sheet" };
  }

  const cleanTarget = targetKey.replace(/[^A-Za-z0-9_-]/g, "").toUpperCase();

  // Quét qua bảng tính: Chỉ so khớp CHÍNH XÁC ô dữ liệu (Tránh quét nhầm email/sđt/thông tin khác)
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cellVal = String(data[r][c] || "").trim().toUpperCase();
      if (cellVal && cellVal === cleanTarget) {
        // Tìm thấy License Key khớp chính xác 100%!
        return {
          valid: true,
          message: "Mã License Key chính xác! Đã áp dụng giảm 100.000đ."
        };
      }
    }
  }

  return {
    valid: false,
    message: "Mã License Key không tồn tại trên hệ thống!"
  };
}

/**
 * Hàm lưu người đăng ký Affiliate vào Tab riêng trong Google Sheet
 */
function saveAffiliateToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG.AFFILIATE_TAB_NAME);

  // Nếu chưa có tab này, tự động tạo mới với tiêu đề chuẩn
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG.AFFILIATE_TAB_NAME);
    const headers = [
      "Thời Gian Đăng Ký",
      "Họ và Tên",
      "Số Điện Thoại (Zalo)",
      "Số Tài Khoản (STK)",
      "Ngân Hàng",
      "Mã Affiliate Ref",
      "Link Giới Thiệu",
      "Trạng Thái",
      "Số Đơn Giới Thiệu",
      "Tổng Hoa Hồng (VNĐ)"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#10b981").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");
  const name = String(data.name || "").trim().toUpperCase();
  const phone = String(data.phone || "").trim();
  const stk = String(data.stk || "").trim();
  const bank = String(data.bank || "").trim().toUpperCase();
  const affCode = String(data.affCode || `STK_${stk}_${bank}`).trim();
  const affUrl = String(data.affUrl || "").trim();

  // Kiểm tra xem STK hoặc SĐT này đã đăng ký trước đó chưa (nếu có thì cập nhật)
  const existingData = sheet.getDataRange().getValues();
  let existingRow = -1;
  for (let i = 1; i < existingData.length; i++) {
    const rowPhone = String(existingData[i][2]).trim();
    const rowStk = String(existingData[i][3]).trim();
    if ((phone && rowPhone === phone) || (stk && rowStk === stk)) {
      existingRow = i + 1;
      break;
    }
  }

  if (existingRow > 0) {
    // Cập nhật dòng hiện có
    sheet.getRange(existingRow, 1, 1, 7).setValues([[nowStr, name, phone, stk, bank, affCode, affUrl]]);
    return {
      success: true,
      message: "Đã cập nhật thông tin Affiliate thành công!",
      affCode: affCode
    };
  } else {
    // Thêm dòng mới
    sheet.appendRow([nowStr, name, "'" + phone, "'" + stk, bank, affCode, affUrl, "Đang Hoạt Động", 0, 0]);
    return {
      success: true,
      message: "Đã lưu đăng ký Affiliate mới vào Google Sheet thành công!",
      affCode: affCode
    };
  }
}

/**
 * Hàm lưu đơn đặt hàng vào Tab DonHang trong Google Sheet
 */
function saveOrderToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG.ORDERS_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG.ORDERS_TAB_NAME);
    const headers = [
      "Thời Gian",
      "Họ và Tên Khách",
      "Số Điện Thoại",
      "Email",
      "Số Tiền (VNĐ)",
      "Mã Giảm Giá / Ref",
      "Loại Đơn Hàng",
      "Nội Dung Chuyển Khoản",
      "Trạng Thái Thanh Toán"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#3b82f6").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");
  const name = String(data.name || "").trim();
  const phone = String(data.phone || "").trim();
  const email = String(data.email || "").trim();
  const amount = Number(data.amount || 1099000);
  const refCode = String(data.refCode || "").trim();
  const orderType = String(data.orderType || "Mua trực tiếp").trim();
  const memo = String(data.memo || "").trim();

  sheet.appendRow([
    nowStr,
    name,
    "'" + phone,
    email,
    amount,
    refCode,
    orderType,
    memo,
    "Chờ Chuyển Khoản VietQR"
  ]);

  return { success: true, message: "Đã lưu đơn hàng vào Google Sheet!" };
}

/**
 * Helper tạo phản hồi JSON chuẩn có hỗ trợ CORS
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
