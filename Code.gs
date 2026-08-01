/**
 * Hair Salon Booking - Google Apps Script backend.
 *
 * Với script gắn trực tiếp vào Google Sheet, để trống SPREADSHEET_ID.
 * Với standalone script, điền ID của Google Sheet vào hằng số bên dưới.
 */
const SPREADSHEET_ID = '';

const CONFIG = Object.freeze({
  sheets: {
    services: 'Services',
    stylists: 'Stylists',
    appointments: 'Appointments'
  },
  headers: {
    Services: ['ID', 'Name', 'Price', 'Duration_Minutes', 'Status'],
    Stylists: ['ID', 'Name', 'Phone', 'Status'],
    Appointments: [
      'Booking_ID', 'Customer_Name', 'Customer_Phone', 'Service_IDs',
      'Stylist_ID', 'Booking_Date', 'Booking_Time', 'Total_Price',
      'Status', 'Created_At'
    ]
  },
  openMinute: 8 * 60,
  closeMinute: 20 * 60,
  slotStep: 30,
  blockingStatuses: ['Pending', 'Confirmed', 'Completed']
});

/** Phục vụ giao diện Web App. */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Đặt lịch Hair Salon')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Tự khởi tạo dữ liệu khi chủ sở hữu mở Spreadsheet lần đầu. */
function onOpen(e) {
  setupSalonSheets();
  SpreadsheetApp.getUi()
    .createMenu('Hair Salon')
    .addItem('Khởi tạo / kiểm tra dữ liệu', 'setupSalonSheets')
    .addToUi();
}

/**
 * Chạy hàm này một lần từ Apps Script editor để tạo các sheet và dữ liệu mẫu.
 * Hàm an toàn khi chạy lại: không xóa hoặc ghi đè dữ liệu đang có.
 */
function setupSalonSheets() {
  const ss = getSpreadsheet_();

  Object.keys(CONFIG.headers).forEach(function (sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    const headers = CONFIG.headers[sheetName];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#17202a')
        .setFontColor('#ffffff');
    } else {
      const actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
      if (actual.join('|') !== headers.join('|')) {
        throw new Error('Header của sheet "' + sheetName + '" không đúng cấu trúc yêu cầu.');
      }
    }
  });

  seedIfEmpty_(ss.getSheetByName(CONFIG.sheets.services), [
    ['DV001', 'Cắt tóc nam', 120000, 45, 'Active'],
    ['DV002', 'Gội đầu thư giãn', 80000, 30, 'Active'],
    ['DV003', 'Uốn tóc', 450000, 120, 'Active'],
    ['DV004', 'Nhuộm tóc', 550000, 120, 'Active'],
    ['DV005', 'Tạo kiểu', 150000, 45, 'Active']
  ]);
  seedIfEmpty_(ss.getSheetByName(CONFIG.sheets.stylists), [
    ['ST001', 'Minh Anh', '0900000001', 'Active'],
    ['ST002', 'Quốc Bảo', '0900000002', 'Active'],
    ['ST003', 'Thanh Hà', '0900000003', 'Active']
  ]);

  const appointments = ss.getSheetByName(CONFIG.sheets.appointments);
  appointments.getRange('F:F').setNumberFormat('@');
  appointments.getRange('G:G').setNumberFormat('@');
  appointments.getRange('H:H').setNumberFormat('#,##0');
  appointments.getRange('J:J').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  return { success: true, message: 'Đã khởi tạo dữ liệu Hair Salon.' };
}

/** Lấy dữ liệu dịch vụ và stylist đang hoạt động. */
function getInitialData() {
  try {
    ensureSheets_();
    const ss = getSpreadsheet_();
    const services = rowsToObjects_(ss.getSheetByName(CONFIG.sheets.services))
      .filter(function (row) { return normalizeStatus_(row.Status) === 'active'; })
      .map(function (row) {
        return {
          id: cleanString_(row.ID),
          name: cleanString_(row.Name),
          price: toPositiveNumber_(row.Price, 'Giá dịch vụ'),
          duration: toPositiveInteger_(row.Duration_Minutes, 'Thời lượng dịch vụ')
        };
      });
    const stylists = rowsToObjects_(ss.getSheetByName(CONFIG.sheets.stylists))
      .filter(function (row) { return normalizeStatus_(row.Status) === 'active'; })
      .map(function (row) {
        return { id: cleanString_(row.ID), name: cleanString_(row.Name) };
      });

    return {
      success: true,
      services: services,
      stylists: stylists,
      businessHours: { open: '08:00', close: '20:00', stepMinutes: CONFIG.slotStep },
      today: Utilities.formatDate(new Date(), getTimeZone_(), 'yyyy-MM-dd')
    };
  } catch (error) {
    console.error(error);
    throw new Error(safeErrorMessage_(error));
  }
}

/**
 * Trả về các khoảng thời gian đã bận. stylistId = "ANY" sẽ trả lịch của mọi thợ.
 */
function checkAvailability(date, stylistId) {
  try {
    ensureSheets_();
    const bookingDate = validateDate_(date);
    const requestedStylist = cleanString_(stylistId || 'ANY');
    const stylists = getActiveStylists_();
    const stylistIds = requestedStylist === 'ANY'
      ? stylists.map(function (item) { return item.id; })
      : [requestedStylist];

    if (requestedStylist !== 'ANY' && !stylists.some(function (item) { return item.id === requestedStylist; })) {
      throw new Error('Stylist không tồn tại hoặc đang ngừng hoạt động.');
    }

    const busyByStylist = getBusyIntervals_(bookingDate, stylistIds);
    return {
      success: true,
      date: bookingDate,
      stylistId: requestedStylist,
      busyByStylist: busyByStylist,
      openMinute: CONFIG.openMinute,
      closeMinute: CONFIG.closeMinute,
      stepMinutes: CONFIG.slotStep
    };
  } catch (error) {
    console.error(error);
    throw new Error(safeErrorMessage_(error));
  }
}

/** Tạo lịch hẹn, tính giá tại server và chống đặt trùng lịch. */
function createBooking(bookingData) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) throw new Error('Hệ thống đang bận. Vui lòng thử lại sau vài giây.');
    ensureSheets_();

    const data = bookingData || {};
    const customerName = cleanString_(data.customerName);
    const customerPhone = normalizePhone_(data.customerPhone);
    const bookingDate = validateDate_(data.bookingDate);
    const bookingTime = validateTime_(data.bookingTime);
    const requestedStylist = cleanString_(data.stylistId || 'ANY');
    const serviceIds = uniqueStrings_(Array.isArray(data.serviceIds) ? data.serviceIds : []);

    if (customerName.length < 2 || customerName.length > 100) {
      throw new Error('Họ tên phải có từ 2 đến 100 ký tự.');
    }
    if (!/^(?:\+84|0)\d{9,10}$/.test(customerPhone)) {
      throw new Error('Số điện thoại không hợp lệ.');
    }
    if (!serviceIds.length) throw new Error('Vui lòng chọn ít nhất một dịch vụ.');

    const serviceMap = getActiveServiceMap_();
    const selectedServices = serviceIds.map(function (id) {
      if (!serviceMap[id]) throw new Error('Dịch vụ "' + id + '" không tồn tại hoặc đã ngừng phục vụ.');
      return serviceMap[id];
    });
    const totalPrice = selectedServices.reduce(function (sum, item) { return sum + item.price; }, 0);
    const totalDuration = selectedServices.reduce(function (sum, item) { return sum + item.duration; }, 0);
    const startMinute = timeToMinute_(bookingTime);
    const now = new Date();
    const today = Utilities.formatDate(now, getTimeZone_(), 'yyyy-MM-dd');
    const currentMinute = Number(Utilities.formatDate(now, getTimeZone_(), 'H')) * 60 +
      Number(Utilities.formatDate(now, getTimeZone_(), 'm'));

    if (bookingDate === today && startMinute <= currentMinute) {
      throw new Error('Không thể đặt một khung giờ đã qua.');
    }

    if (startMinute < CONFIG.openMinute || startMinute + totalDuration > CONFIG.closeMinute) {
      throw new Error('Khung giờ nằm ngoài giờ hoạt động hoặc không đủ thời gian thực hiện dịch vụ.');
    }
    if ((startMinute - CONFIG.openMinute) % CONFIG.slotStep !== 0) {
      throw new Error('Khung giờ không hợp lệ.');
    }

    const activeStylists = getActiveStylists_();
    let candidateIds;
    if (requestedStylist === 'ANY') {
      candidateIds = activeStylists.map(function (item) { return item.id; });
    } else {
      if (!activeStylists.some(function (item) { return item.id === requestedStylist; })) {
        throw new Error('Stylist không tồn tại hoặc đang ngừng hoạt động.');
      }
      candidateIds = [requestedStylist];
    }
    if (!candidateIds.length) throw new Error('Hiện chưa có stylist hoạt động.');

    const busyByStylist = getBusyIntervals_(bookingDate, candidateIds);
    const availableIds = candidateIds.filter(function (id) {
      return !hasOverlap_(startMinute, startMinute + totalDuration, busyByStylist[id] || []);
    });
    if (!availableIds.length) {
      throw new Error('Khung giờ vừa được người khác đặt. Vui lòng chọn giờ khác.');
    }

    const assignedStylistId = requestedStylist === 'ANY'
      ? availableIds[Math.floor(Math.random() * availableIds.length)]
      : availableIds[0];
    const bookingId = 'BK-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
    const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.appointments);
    sheet.appendRow([
      bookingId,
      escapeFormula_(customerName),
      "'" + customerPhone,
      serviceIds.join(','),
      assignedStylistId,
      bookingDate,
      bookingTime,
      totalPrice,
      'Pending',
      new Date()
    ]);

    return {
      success: true,
      message: 'Đặt lịch thành công! Salon sẽ sớm liên hệ xác nhận.',
      bookingId: bookingId,
      stylistId: assignedStylistId,
      totalPrice: totalPrice,
      totalDuration: totalDuration
    };
  } catch (error) {
    console.error(error);
    return { success: false, message: safeErrorMessage_(error) };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Không tìm thấy Google Sheet. Hãy gắn script vào Sheet hoặc cấu hình SPREADSHEET_ID.');
  return ss;
}

function ensureSheets_() {
  const ss = getSpreadsheet_();
  const missing = Object.keys(CONFIG.headers).some(function (name) { return !ss.getSheetByName(name); });
  if (missing) setupSalonSheets();
}

function seedIfEmpty_(sheet, rows) {
  if (sheet.getLastRow() === 1 && rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function rowsToObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(String);
  return values.filter(function (row) {
    return row.some(function (cell) { return cell !== ''; });
  }).map(function (row) {
    return headers.reduce(function (obj, header, index) {
      obj[header] = row[index];
      return obj;
    }, {});
  });
}

function getActiveStylists_() {
  return rowsToObjects_(getSpreadsheet_().getSheetByName(CONFIG.sheets.stylists))
    .filter(function (row) { return normalizeStatus_(row.Status) === 'active'; })
    .map(function (row) { return { id: cleanString_(row.ID), name: cleanString_(row.Name) }; })
    .filter(function (row) { return row.id; });
}

function getActiveServiceMap_() {
  return getServiceMap_(true);
}

function getServiceMap_(activeOnly) {
  return rowsToObjects_(getSpreadsheet_().getSheetByName(CONFIG.sheets.services))
    .filter(function (row) { return !activeOnly || normalizeStatus_(row.Status) === 'active'; })
    .reduce(function (map, row) {
      const id = cleanString_(row.ID);
      if (id) {
        map[id] = {
          id: id,
          name: cleanString_(row.Name),
          price: toPositiveNumber_(row.Price, 'Giá dịch vụ'),
          duration: toPositiveInteger_(row.Duration_Minutes, 'Thời lượng dịch vụ')
        };
      }
      return map;
    }, {});
}

function getBusyIntervals_(date, stylistIds) {
  const result = {};
  stylistIds.forEach(function (id) { result[id] = []; });
  const idSet = stylistIds.reduce(function (set, id) { set[id] = true; return set; }, {});
  // Lịch cũ vẫn phải giữ đúng thời lượng khi dịch vụ đã chuyển sang Inactive.
  const serviceMap = getServiceMap_(false);
  const rows = rowsToObjects_(getSpreadsheet_().getSheetByName(CONFIG.sheets.appointments));

  rows.forEach(function (row) {
    const stylistId = cleanString_(row.Stylist_ID);
    const rowDate = formatSheetDate_(row.Booking_Date);
    const status = cleanString_(row.Status);
    if (!idSet[stylistId] || rowDate !== date || CONFIG.blockingStatuses.indexOf(status) === -1) return;

    const start = timeToMinute_(formatSheetTime_(row.Booking_Time));
    const duration = cleanString_(row.Service_IDs).split(',').reduce(function (sum, id) {
      const service = serviceMap[cleanString_(id)];
      return sum + (service ? service.duration : CONFIG.slotStep);
    }, 0) || CONFIG.slotStep;
    result[stylistId].push({ start: start, end: start + duration });
  });
  return result;
}

function hasOverlap_(start, end, intervals) {
  return intervals.some(function (item) { return start < item.end && end > item.start; });
}

function validateDate_(value) {
  const text = cleanString_(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error('Ngày đặt lịch không hợp lệ.');
  const parts = text.split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (date.getFullYear() !== parts[0] || date.getMonth() !== parts[1] - 1 || date.getDate() !== parts[2]) {
    throw new Error('Ngày đặt lịch không tồn tại.');
  }
  const today = Utilities.formatDate(new Date(), getTimeZone_(), 'yyyy-MM-dd');
  if (text < today) throw new Error('Không thể đặt lịch trong quá khứ.');
  return text;
}

function validateTime_(value) {
  const text = cleanString_(value);
  if (!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(text)) throw new Error('Giờ đặt lịch không hợp lệ.');
  return text;
}

function timeToMinute_(time) {
  const parts = validateTime_(time).split(':').map(Number);
  return parts[0] * 60 + parts[1];
}

function formatSheetDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, getTimeZone_(), 'yyyy-MM-dd');
  }
  return cleanString_(value).replace(/^'/, '');
}

function formatSheetTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, getTimeZone_(), 'HH:mm');
  }
  return cleanString_(value).replace(/^'/, '').slice(0, 5);
}

function getTimeZone_() {
  return getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
}

function normalizePhone_(value) {
  return cleanString_(value).replace(/[\s.()-]/g, '');
}

function escapeFormula_(value) {
  const text = cleanString_(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function normalizeStatus_(value) { return cleanString_(value).toLowerCase(); }
function cleanString_(value) { return value == null ? '' : String(value).trim(); }
function uniqueStrings_(values) {
  return values.map(cleanString_).filter(Boolean).filter(function (value, index, array) {
    return array.indexOf(value) === index;
  });
}
function toPositiveNumber_(value, label) {
  const number = Number(value);
  if (!isFinite(number) || number < 0) throw new Error(label + ' không hợp lệ.');
  return number;
}
function toPositiveInteger_(value, label) {
  const number = Number(value);
  if (!Number.isInteger(number) || number <= 0) throw new Error(label + ' không hợp lệ.');
  return number;
}
function safeErrorMessage_(error) {
  return error && error.message ? error.message : 'Đã xảy ra lỗi không xác định. Vui lòng thử lại.';
}

/*
 * DEPLOY WEB APP:
 * 1. Mở Extensions > Apps Script từ Google Sheet và thêm Code.gs + Index.html.
 * 2. Chạy setupSalonSheets() một lần, cấp các quyền được yêu cầu.
 * 3. Chọn Deploy > New deployment > Web app.
 * 4. Execute as: Me; Who has access: Anyone.
 * 5. Deploy, cấp quyền và dùng URL /exec được cung cấp.
 */
