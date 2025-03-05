function doGet(e) {
  var params = e.parameter;
  var LinkUrl = params.LinkUrl;
  var GiftType = params.GiftType;

  if (!LinkUrl || !GiftType) {
    return ContentService.createTextOutput("Thiếu tham số LinkUrl hoặc GiftType").setMimeType(ContentService.MimeType.TEXT);
  }

  var message = addLinkToSheet(LinkUrl, GiftType);

  return ContentService.createTextOutput(message).setMimeType(ContentService.MimeType.TEXT);
}

function initializeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Data";
  var sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    return;
  }
  // Kiểm tra nếu sheet chưa tồn tại thì tạo mới
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Xác định các tiêu đề cột cần thêm, thêm STT vào đầu tiên
  var headers = ["STT", "Gift 100", "Gift 500", "Gift 1000", "Check Lại Link"];
  var totalRows = 1000;
  
  // Ghi tiêu đề vào hàng đầu tiên
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // Đổi màu nền thành xanh lá cây nhạt
  var color = "#ccffcc"; // Mã màu xanh lá cây nhạt
  headerRange.setBackground(color);

  // Kẻ đậm viền cho các tiêu đề
  var fullRange = sheet.getRange(1, 1, totalRows, headers.length);
  fullRange.setBorder(true, true, true, true, true, true);

  // Cố định hàng tiêu đề để không di chuyển khi cuộn
  sheet.setFrozenRows(1);

  // Đặt độ rộng cột: STT giữ nguyên, các cột còn lại gấp đôi (200px)
  sheet.setColumnWidth(1, 100); // Cột STT giữ kích thước mặc định (100px)
  for (var i = 2; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 200); // Các cột còn lại gấp đôi kích thước mặc định
  }
  return sheet;
}

function addLinkToSheet(url, category) {
  var lock = LockService.getScriptLock();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Data";
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = initializeSheet(); // Tạo sheet nếu chưa có
    }

    // Bản đồ chuyển đổi category thành số cột
    var categoryMap = {
      "100": 2,   // Cột B - Gift 100
      "500": 3,   // Cột C - Gift 500
      "1000": 4,  // Cột D - Gift 1000
      "recheck": 5 // Cột E - Check Lại Link
    };

    var column = categoryMap[category.toString()];
    if (!column) {
      return "Giá trị không hợp lệ: " + category;
    }

    // Loại bỏ khoảng trắng thừa trong URL
    url = url.trim();

    var lastRow = sheet.getLastRow();
    var data = [];
    if (lastRow >= 2) {
      // Lấy dữ liệu từ hàng 2 đến cuối, từ cột A đến E
      data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    }

    var newRow = lastRow + 1;
    sheet.getRange(newRow, 1).setValue(newRow - 1); // STT
    sheet.getRange(newRow, column).setValue(url);
    sheet.getRange(newRow, column).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    return `Thêm mới thành công! URL đã được thêm vào cột ${category}`;

  } finally {
    lock.releaseLock(); // Giải phóng lock sau khi hoàn tất
  }
}

