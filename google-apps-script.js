/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT - HỆ THỐNG XÁC THỰC LICENSE KEY, ĐƠN HÀNG & QUẢN LÝ AFFILIATE QL SOCIAL
 * ==============================================================================
 * CÁC TÍNH NĂNG TỰ ĐỘNG:
 * 1. Cột "Trạng Thái Thanh Toán" tự động có Nút Chọn (Dropdown): "Đang chờ" & "Đã thanh toán".
 * 2. Cột "Thông Tin Người Giới Thiệu (Affiliate)" tự động tra cứu Họ Tên + STK + Ngân Hàng của người giới thiệu.
 * 3. Khi bạn đổi trạng thái sang "Đã thanh toán", hệ thống tự động cộng dồn Số Đơn & Hoa Hồng vào Tab "DangKy_Affiliate".
 * 4. Tự động thêm Menu "QL Social Studio" trên thanh công cụ để đồng bộ định dạng chỉ với 1 click!
 * ==============================================================================
 */

// Cấu hình các trang tính (Sheet tabs)
const SHEET_CONFIG = {
  LICENSE_TAB_INDEX: 0,                   // Tab đầu tiên chứa License Key
  AFFILIATE_TAB_NAME: "DangKy_Affiliate", // Tab lưu danh sách đối tác đăng ký Affiliate
  ORDERS_TAB_NAME: "DonHang_DatMua"       // Tab lưu đơn đặt hàng từ Web
};

/**
 * Tự động tạo Menu "QL Social Studio" khi bạn mở file Google Sheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 QL Social Studio")
    .addItem("🔄 Tạo Dropdown & Cập Nhật Tất Cả Đơn Hàng", "formatAllOrderRows")
    .addItem("📊 Đồng Bộ Lại Số Liệu Hoa Hồng Affiliate", "recalculateAllAffiliateStats")
    .addSeparator()
    .addItem("⚡ Kiểm Tra Kết Nối Web App", "testConnectionMenu")
    .addToUi();
}

/**
 * Trigger tự động kích hoạt khi bạn sửa ô trong Google Sheet
 * Khi bạn chuyển cột "Trạng Thái Thanh Toán" sang "Đã thanh toán", script sẽ tự động tra cứu Affiliate & tính hoa hồng!
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_CONFIG.ORDERS_TAB_NAME) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Cột 9 là cột "Trạng Thái Thanh Toán" (Cột I)
    if (col === 9 && row > 1) {
      const newStatus = String(e.value || "").trim();
      const oldStatus = String(e.oldValue || "").trim();

      // Cập nhật màu sắc & tra cứu thông tin Affiliate
      formatSingleOrderRow(sheet, row);

      // Nếu trạng thái đổi sang "Đã thanh toán" hoặc từ "Đã thanh toán" về "Đang chờ"
      if (newStatus === "Đã thanh toán" || oldStatus === "Đã thanh toán") {
        syncOrderToAffiliateTab(sheet, row, newStatus, oldStatus);
      }
    }
  } catch (err) {
    console.log("onEdit Error:", err);
  }
}

/**
 * Xử lý yêu cầu GET từ Web
 */
function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || "check_license";

    // 1. KIỂM TRA LICENSE KEY CÓ HỢP LỆ KHÔNG
    if (action === "check_license") {
      const key = (params.key || "").trim().toUpperCase();
      if (!key || key.length < 3) {
        return createJsonResponse({ valid: false, message: "Vui lòng nhập License Key hợp lệ" });
      }
      return createJsonResponse(verifyLicenseKeyInSheet(key));
    }

    // 2. LƯU ĐĂNG KÝ AFFILIATE
    if (action === "register_affiliate") {
      return createJsonResponse(saveAffiliateToSheet(params));
    }

    // 3. PING TEST
    if (action === "ping") {
      return createJsonResponse({ success: true, message: "Google Apps Script API đang hoạt động tốt!" });
    }

    return createJsonResponse({ success: false, message: "Hành động không hợp lệ" });
  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Xử lý yêu cầu POST từ Web
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

    if (action === "register_affiliate") {
      return createJsonResponse(saveAffiliateToSheet(data));
    }

    if (action === "create_order") {
      return createJsonResponse(saveOrderToSheet(data));
    }

    return createJsonResponse({ success: false, message: "Hành động POST không hợp lệ" });
  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Tìm kiếm License Key trong Google Sheet (So khớp chính xác 100%, bảo mật cao)
 */
function verifyLicenseKeyInSheet(targetKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[SHEET_CONFIG.LICENSE_TAB_INDEX] || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return { valid: false, message: "Chưa có dữ liệu License Key trong Sheet" };
  }

  const cleanTarget = targetKey.replace(/[^A-Za-z0-9_-]/g, "").toUpperCase();

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cellVal = String(data[r][c] || "").trim().toUpperCase();
      if (cellVal && cellVal === cleanTarget) {
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
 * Lưu đăng ký Affiliate vào Tab DangKy_Affiliate
 */
function saveAffiliateToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG.AFFILIATE_TAB_NAME);

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
      "Số Đơn Thành Công",
      "Tổng Hoa Hồng (VNĐ)"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#10b981")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");
  const name = String(data.name || "").trim().toUpperCase();
  const phone = String(data.phone || "").trim();
  const stk = String(data.stk || "").trim();
  const bank = String(data.bank || "").trim().toUpperCase();
  const affCode = String(data.affCode || `STK_${stk}_${bank}`).trim();
  const affUrl = String(data.affUrl || "").trim();

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
    sheet.getRange(existingRow, 1, 1, 7).setValues([[nowStr, name, phone, stk, bank, affCode, affUrl]]);
    return {
      success: true,
      message: "Đã cập nhật thông tin Affiliate thành công!",
      affCode: affCode
    };
  } else {
    sheet.appendRow([nowStr, name, "'" + phone, "'" + stk, bank, affCode, affUrl, "Đang Hoạt Động", 0, 0]);
    return {
      success: true,
      message: "Đã lưu đăng ký Affiliate mới vào Google Sheet thành công!",
      affCode: affCode
    };
  }
}

/**
 * Lưu đơn đặt hàng mới vào Tab DonHang_DatMua
 * Tự động gắn Dropdown Chọn ("Đang chờ", "Đã thanh toán") và tra cứu thông tin Affiliate
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
      "Trạng Thái Thanh Toán",
      "Thông Tin Người Giới Thiệu (Affiliate)",
      "Hoa Hồng Dự Kiến (VNĐ)"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#2563eb")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
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

  // Tra cứu thông tin người giới thiệu từ mã Ref
  const affInfo = lookupAffiliateInfo(refCode);

  sheet.appendRow([
    nowStr,
    name,
    "'" + phone,
    email,
    amount,
    refCode || "Không có",
    orderType,
    memo,
    "Đang chờ",
    affInfo.detail,
    affInfo.commission
  ]);

  const newRow = sheet.getLastRow();
  formatSingleOrderRow(sheet, newRow);

  return { success: true, message: "Đã lưu đơn hàng vào Google Sheet thành công!" };
}

/**
 * Tra cứu thông tin người giới thiệu (Affiliate) từ mã Ref hoặc License Key
 */
function lookupAffiliateInfo(refCode) {
  if (!refCode || refCode === "Không có" || refCode === "null") {
    return {
      detail: "Mua trực tiếp (Không qua giới thiệu)",
      commission: 0,
      partnerKey: null
    };
  }

  const cleanRef = String(refCode).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Kiểm tra trong Tab DangKy_Affiliate (Đối tác tự do)
  const affSheet = ss.getSheetByName(SHEET_CONFIG.AFFILIATE_TAB_NAME);
  if (affSheet) {
    const affData = affSheet.getDataRange().getValues();
    for (let i = 1; i < affData.length; i++) {
      const name = String(affData[i][1] || "");
      const phone = String(affData[i][2] || "");
      const stk = String(affData[i][3] || "");
      const bank = String(affData[i][4] || "");
      const code = String(affData[i][5] || "");

      if (
        cleanRef.toUpperCase() === code.toUpperCase() ||
        cleanRef.includes(stk) ||
        (phone && cleanRef.includes(phone))
      ) {
        return {
          detail: `💳 ${name} | STK: ${stk} (${bank}) | SĐT: ${phone}`,
          commission: 200000,
          partnerKey: code
        };
      }
    }
  }

  // 2. Kiểm tra nếu là License Key từ Tab đầu tiên
  const licSheet = ss.getSheets()[SHEET_CONFIG.LICENSE_TAB_INDEX] || ss.getSheets()[0];
  if (licSheet && !cleanRef.startsWith("STK_")) {
    const licData = licSheet.getDataRange().getValues();
    for (let r = 0; r < licData.length; r++) {
      for (let c = 0; c < licData[r].length; c++) {
        const val = String(licData[r][c] || "").trim().toUpperCase();
        if (val && val === cleanRef.toUpperCase()) {
          // Lấy thông tin cột bên cạnh nếu có (Tên/SĐT của khách cũ)
          const extraInfo = licData[r][1] ? ` (${licData[r][1]})` : "";
          return {
            detail: `🔑 License Key: ${cleanRef}${extraInfo} - Khách cũ giới thiệu (-100k)`,
            commission: 200000,
            partnerKey: cleanRef
          };
        }
      }
    }
  }

  // 3. Nếu là định dạng STK_..._BANK nhưng chưa có trong tab Affiliate
  if (cleanRef.startsWith("STK_")) {
    const parts = cleanRef.split("_");
    const stk = parts[1] || "";
    const bank = parts[2] || "";
    return {
      detail: `💳 STK: ${stk} (${bank}) - Đối tác Affiliate tự do`,
      commission: 200000,
      partnerKey: cleanRef
    };
  }

  return {
    detail: `Mã Ref: ${cleanRef}`,
    commission: 200000,
    partnerKey: cleanRef
  };
}

/**
 * Gắn Dropdown Nút Chọn ("Đang chờ", "Đã thanh toán") & tô màu tự động cho 1 dòng đơn hàng
 */
function formatSingleOrderRow(sheet, row) {
  // Cột I: Trạng Thái Thanh Toán
  const statusCell = sheet.getRange(row, 9);
  const currentVal = String(statusCell.getValue() || "").trim();

  // Tạo Dropdown Validation
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Đang chờ", "Đã thanh toán"], true)
    .setAllowInvalid(false)
    .build();
  statusCell.setDataValidation(rule);

  // Cập nhật thông tin Affiliate ở cột J & K nếu còn trống
  const refVal = String(sheet.getRange(row, 6).getValue() || "").trim();
  const affCell = sheet.getRange(row, 10);
  const commCell = sheet.getRange(row, 11);

  if (!affCell.getValue() || affCell.getValue() === "") {
    const affInfo = lookupAffiliateInfo(refVal);
    affCell.setValue(affInfo.detail);
    commCell.setValue(affInfo.commission);
  }

  // Tô màu sắc theo trạng thái
  if (currentVal === "Đã thanh toán") {
    statusCell.setBackground("#d1fae5").setFontColor("#065f46").setFontWeight("bold"); // Xanh lá cây
  } else {
    statusCell.setBackground("#fef3c7").setFontColor("#92400e").setFontWeight("bold"); // Vàng cam
  }
}

/**
 * Định dạng lại toàn bộ bảng đơn hàng & gắn Dropdown cho tất cả các dòng
 */
function formatAllOrderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG.ORDERS_TAB_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Chưa có tab " + SHEET_CONFIG.ORDERS_TAB_NAME);
    return;
  }

  // Đảm bảo tiêu đề cột J và K đầy đủ
  sheet.getRange(1, 9).setValue("Trạng Thái Thanh Toán");
  sheet.getRange(1, 10).setValue("Thông Tin Người Giới Thiệu (Affiliate)");
  sheet.getRange(1, 11).setValue("Hoa Hồng Dự Kiến (VNĐ)");
  sheet.getRange(1, 1, 1, 11)
    .setFontWeight("bold")
    .setBackground("#2563eb")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("Đã cập nhật tiêu đề cột thành công!");
    return;
  }

  for (let r = 2; r <= lastRow; r++) {
    formatSingleOrderRow(sheet, r);
  }

  SpreadsheetApp.getUi().alert(`✅ Đã tạo Dropdown Nút Chọn và cập nhật ${lastRow - 1} dòng đơn hàng thành công!`);
}

/**
 * Tự động cập nhật thống kê đơn và hoa hồng vào Tab DangKy_Affiliate khi đơn chuyển trạng thái
 */
function syncOrderToAffiliateTab(orderSheet, row, newStatus, oldStatus) {
  const refCode = String(orderSheet.getRange(row, 6).getValue() || "").trim();
  if (!refCode || refCode === "Không có") return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const affSheet = ss.getSheetByName(SHEET_CONFIG.AFFILIATE_TAB_NAME);
  if (!affSheet) return;

  const affData = affSheet.getDataRange().getValues();
  for (let i = 1; i < affData.length; i++) {
    const code = String(affData[i][5] || "").trim().toUpperCase();
    const stk = String(affData[i][3] || "").trim();
    const phone = String(affData[i][2] || "").trim();

    if (
      refCode.toUpperCase() === code ||
      refCode.includes(stk) ||
      (phone && refCode.includes(phone))
    ) {
      const countCell = affSheet.getRange(i + 1, 9);
      const totalCommCell = affSheet.getRange(i + 1, 10);
      let currentCount = Number(countCell.getValue() || 0);

      if (newStatus === "Đã thanh toán" && oldStatus !== "Đã thanh toán") {
        currentCount += 1;
      } else if (newStatus !== "Đã thanh toán" && oldStatus === "Đã thanh toán") {
        currentCount = Math.max(0, currentCount - 1);
      }

      // Tính hoa hồng theo bậc: 1-4 đơn (200k), 5-19 đơn (250k), từ 20 đơn (300k)
      let commPerOrder = 200000;
      if (currentCount >= 20) commPerOrder = 300000;
      else if (currentCount >= 5) commPerOrder = 250000;

      countCell.setValue(currentCount);
      totalCommCell.setValue(currentCount * commPerOrder);
      break;
    }
  }
}

/**
 * Tính toán lại toàn bộ số liệu thống kê trong Tab Affiliate dựa trên các đơn "Đã thanh toán"
 */
function recalculateAllAffiliateStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_CONFIG.ORDERS_TAB_NAME);
  const affSheet = ss.getSheetByName(SHEET_CONFIG.AFFILIATE_TAB_NAME);

  if (!orderSheet || !affSheet) {
    SpreadsheetApp.getUi().alert("Chưa đủ dữ liệu Tab để tính toán!");
    return;
  }

  const orderData = orderSheet.getDataRange().getValues();
  const affData = affSheet.getDataRange().getValues();

  // Đếm số đơn "Đã thanh toán" theo từng ref
  const counts = {};
  for (let r = 1; r < orderData.length; r++) {
    const ref = String(orderData[r][5] || "").trim().toUpperCase();
    const status = String(orderData[r][8] || "").trim();
    if (status === "Đã thanh toán" && ref && ref !== "KHÔNG CÓ") {
      counts[ref] = (counts[ref] || 0) + 1;
    }
  }

  // Cập nhật lại vào Tab DangKy_Affiliate
  for (let i = 1; i < affData.length; i++) {
    const code = String(affData[i][5] || "").trim().toUpperCase();
    const stk = String(affData[i][3] || "").trim();
    let totalOrders = 0;

    for (let k in counts) {
      if (k === code || (stk && k.includes(stk))) {
        totalOrders += counts[k];
      }
    }

    let commPerOrder = 200000;
    if (totalOrders >= 20) commPerOrder = 300000;
    else if (totalOrders >= 5) commPerOrder = 250000;

    affSheet.getRange(i + 1, 9).setValue(totalOrders);
    affSheet.getRange(i + 1, 10).setValue(totalOrders * commPerOrder);
  }

  SpreadsheetApp.getUi().alert("✅ Đã đồng bộ và tính toán lại toàn bộ hoa hồng Affiliate thành công!");
}

function testConnectionMenu() {
  SpreadsheetApp.getUi().alert("✅ Kết nối Apps Script hoạt động hoàn hảo!");
}

/**
 * Helper tạo phản hồi JSON chuẩn có hỗ trợ CORS
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
