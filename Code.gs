// ============================================================
//  巨石玩理 · 學費管理系統 · Google Apps Script
// ============================================================
//
//  【三步驟完成設定】
//
//  步驟一：開啟你的 Google 試算表
//    → https://docs.google.com/spreadsheets/d/1hD8yESUhpKkGmTz9qzMMN-KEXY9iyut4eniHRCIRoe0
//
//  步驟二：貼上這段程式碼並執行一次
//    試算表上方選單 → 擴充功能 → Apps Script
//    把這整份內容貼上 → 儲存
//    先執行 initializeSheets()（按▶或上方執行選單）
//    再執行 importHistoricalData()（119 筆學生資料會自動匯入）
//
//  步驟三：部署為網頁應用程式
//    右上角「部署」→ 新增部署作業
//    類型：網頁應用程式
//    執行身分：我（你的帳號）
//    誰可以存取：所有人
//    → 複製產生的網址，貼到 register.html 和 admin.html 的 SCRIPT_URL 位置
//
// ============================================================

// 你的 Google 試算表 ID（已填好，不需要修改）
const SHEET_ID    = '1hD8yESUhpKkGmTz9qzMMN-KEXY9iyut4eniHRCIRoe0';
const ADMIN_EMAIL = 'paco578578@gmail.com';
const STUDENT_SHEET      = '學生名單';
const PAYMENT_SHEET      = '繳費紀錄';
const SCORE_SHEET        = '成績紀錄';
const SCHOLARSHIP_SHEET  = '獎學金紀錄';
const REVIEW_SHEET       = '複習班名單';

// ============================================================
//  接收學生報名（POST）
// ============================================================

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.form_type === 'review') {
      var rIds = saveReviewStudents(data);
      sendReviewNotificationEmail(data, rIds);
      output.setContent(JSON.stringify({ status: 'success', ids: rIds }));
    } else {
      var id = saveStudent(data);
      sendNotificationEmail(data, id);
      output.setContent(JSON.stringify({ status: 'success', id: id }));
    }
  } catch(err) {
    Logger.log('doPost error: ' + err.toString());
    output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
  }
  return output;
}

// ============================================================
//  管理介面 API（GET）
// ============================================================

function doGet(e) {
  var action = e.parameter.action || '';
  var token  = e.parameter.token  || '';
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  if (token !== getAdminToken()) {
    output.setContent(JSON.stringify({ status: 'error', message: '未授權' }));
    return output;
  }

  var result;
  try {
    if (action === 'getStudents') {
      result = { status: 'success', data: getAllStudents() };

    } else if (action === 'updatePayment') {
      updatePaymentStatus(parseInt(e.parameter.student_id), e.parameter.payment_status, e.parameter.amount || '');
      result = { status: 'success' };

    } else if (action === 'getStudentHistory') {
      result = { status: 'success', data: getStudentPaymentHistory(parseInt(e.parameter.student_id)) };

    } else if (action === 'recordPayment') {
      recordStudentPayment(parseInt(e.parameter.student_id), e.parameter.payment_date || '', e.parameter.amount || '');
      result = { status: 'success' };

    } else if (action === 'updatePeriod') {
      updateStudentPeriod(parseInt(e.parameter.student_id), e.parameter.course_plan, e.parameter.start_date, e.parameter.end_date);
      result = { status: 'success' };

    } else if (action === 'markInactive') {
      markStudentInactive(parseInt(e.parameter.student_id));
      result = { status: 'success' };

    } else if (action === 'addScore') {
      var sid   = parseInt(e.parameter.student_id);
      var exam  = e.parameter.exam_name || '';
      var cur   = parseFloat(e.parameter.current_score) || 0;
      var prev  = e.parameter.prev_score !== undefined && e.parameter.prev_score !== '' ? parseFloat(e.parameter.prev_score) : null;
      result = { status: 'success', data: addStudentScore(sid, exam, cur, prev) };

    } else if (action === 'getScoreHistory') {
      result = { status: 'success', data: getStudentScoreHistory(parseInt(e.parameter.student_id)) };

    } else if (action === 'getScholarshipBalance') {
      result = { status: 'success', data: getStudentScholarshipBalance(parseInt(e.parameter.student_id)) };

    } else if (action === 'useScholarship') {
      var sid    = parseInt(e.parameter.student_id);
      var amt    = parseFloat(e.parameter.amount) || 0;
      var type   = e.parameter.type || '折抵';
      var remark = e.parameter.remark || '';
      useStudentScholarship(sid, amt, type, remark);
      result = { status: 'success' };

    } else if (action === 'importScores') {
      var exam   = e.parameter.exam_name || '段考';
      var scores = JSON.parse(e.parameter.scores || '[]');
      result = { status: 'success', data: importStudentScoresBatch(scores, exam) };

    } else if (action === 'getReviewStudents') {
      result = { status: 'success', data: getAllReviewStudents() };

    } else if (action === 'recordReviewPayment') {
      recordReviewStudentPayment(parseInt(e.parameter.student_id), e.parameter.payment_date || '', e.parameter.amount || '');
      result = { status: 'success' };

    } else if (action === 'markReviewInactive') {
      markReviewStudentInactive(parseInt(e.parameter.student_id));
      result = { status: 'success' };

    } else {
      result = { status: 'error', message: '未知 action' };
    }
  } catch(err) {
    Logger.log('doGet error: ' + err.toString());
    result = { status: 'error', message: err.toString() };
  }

  // 支援 JSONP（callback 參數），解決瀏覽器 CORS 限制
  var callback = e.parameter.callback || '';
  if (callback) {
    output = ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')');
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    output.setContent(JSON.stringify(result));
  }

  return output;
}

// ============================================================
//  寫入新學生（來自報名表單）
// ============================================================

function saveStudent(data) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STUDENT_SHEET) || createStudentSheet(ss);
  var id    = getNextId(sheet);

  sheet.appendRow([
    id,
    data.submitted_at || new Date().toISOString(),
    data.tuition_year || '116',
    data.student_name || '',
    data.school       || '',
    data.grade        || '',
    data.phone        || '',
    data.line_id      || '',
    data.address      || '',
    data.course_type  || '',
    extractFee(data.course_type),
    '', '', '待確認', data.notes || ''
  ]);

  var lastRow = sheet.getLastRow();
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, 15).setBackground('#f9f9f9');
  }
  return id;
}

// ============================================================
//  取得所有學生（供管理介面）
// ============================================================

function getAllStudents() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STUDENT_SHEET);
  if (!sheet) return [];

  var tz      = Session.getScriptTimeZone();
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
    // Sheets 會把 "2026-06" 自動轉成 Date，這裡強制轉回 YYYY-MM 字串
    ['課程開始', '課程截止'].forEach(function(col) {
      if (obj[col] instanceof Date) {
        obj[col] = Utilities.formatDate(obj[col], tz, 'yyyy-MM');
      }
    });
    result.push(obj);
  }
  return result;
}

// ============================================================
//  更新繳費狀態
// ============================================================

function updatePaymentStatus(studentId, status, amount) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STUDENT_SHEET);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      sheet.getRange(i + 1, 14).setValue(status);
      if (amount) sheet.getRange(i + 1, 11).setValue(amount);
      logPayment(studentId, data[i][2], data[i][3], status, amount);
      break;
    }
  }
}

// ============================================================
//  取得學生繳費歷史紀錄
// ============================================================

function getStudentPaymentHistory(studentId) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(PAYMENT_SHEET);
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
      result.push(obj);
    }
  }
  return result.reverse();
}

// ============================================================
//  記錄繳費（新版：含繳費日期）
// ============================================================

function recordStudentPayment(studentId, paymentDate, amount) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STUDENT_SHEET);
  if (!sheet) return;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var nameIdx = headers.indexOf('學生姓名');
  var yearIdx = headers.indexOf('學費年度');
  var feeIdx  = headers.indexOf('月費金額');
  var stIdx   = headers.indexOf('繳費狀態');
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      sheet.getRange(i + 1, stIdx + 1).setValue('已繳');
      var paySheet = ss.getSheetByName(PAYMENT_SHEET) || createPaymentSheet(ss);
      var date     = paymentDate || new Date().toISOString().slice(0, 10);
      var fee      = amount || data[i][feeIdx] || '';
      paySheet.appendRow([new Date().toISOString(), studentId, data[i][yearIdx], data[i][nameIdx], date, fee, '已繳']);
      break;
    }
  }
}

// ============================================================
//  更新學期（課程方案、起訖日期）
// ============================================================

function updateStudentPeriod(studentId, coursePlan, startDate, endDate) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STUDENT_SHEET);
  if (!sheet) return;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var planIdx  = headers.indexOf('課程方案');
  var feeIdx   = headers.indexOf('月費金額');
  var startIdx = headers.indexOf('課程開始');
  var endIdx   = headers.indexOf('課程截止');
  var stIdx    = headers.indexOf('繳費狀態');
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      if (coursePlan) { sheet.getRange(i+1, planIdx+1).setValue(coursePlan); }
      if (startDate)  { sheet.getRange(i+1, startIdx+1).setValue(startDate); }
      if (endDate)    { sheet.getRange(i+1, endIdx+1).setValue(endDate); }
      sheet.getRange(i+1, stIdx+1).setValue('待確認');
      var fee = extractFee(coursePlan);
      if (fee) sheet.getRange(i+1, feeIdx+1).setValue(fee);
      break;
    }
  }
}

// ============================================================
//  標記不續班
// ============================================================

function markStudentInactive(studentId) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STUDENT_SHEET);
  if (!sheet) return;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var stIdx   = headers.indexOf('繳費狀態');
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      sheet.getRange(i+1, stIdx+1).setValue('不續班');
      break;
    }
  }
}

// ============================================================
//  Email 通知
// ============================================================

function sendNotificationEmail(data, id) {
  var body = [
    '✅ 新學生報名通知',
    '─────────────',
    '編號：#' + id,
    '學生：' + data.student_name,
    '學校：' + data.school + '　' + (data.grade || ''),
    '電話：' + data.phone,
    'LINE：' + (data.line_id || ''),
    '方案：' + data.course_type,
    '地址：' + data.address,
    '備註：' + (data.notes || '無'),
    '─────────────',
    '時間：' + data.submitted_at,
    '年度：' + (data.tuition_year || '116'),
  ].join('\n');

  MailApp.sendEmail(
    ADMIN_EMAIL,
    '【巨石玩理】新學生報名 #' + id + '：' + data.student_name,
    body
  );
}

// ============================================================
//  建立工作表
// ============================================================

function createStudentSheet(ss) {
  var sheet   = ss.insertSheet(STUDENT_SHEET);
  var headers = ['id','報名時間','學費年度','學生姓名','就讀學校','年級','家長電話','LINE ID','地址','課程方案','月費金額','課程開始','課程截止','繳費狀態','備註'];
  var range   = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground('#1c1c1e');
  range.setFontColor('white');
  range.setFontWeight('bold');
  range.setFontSize(12);
  sheet.setFrozenRows(1);
  [40,160,70,80,130,80,160,150,250,110,80,80,80,80,160].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  return sheet;
}

function createPaymentSheet(ss) {
  var sheet   = ss.insertSheet(PAYMENT_SHEET);
  var headers = ['紀錄時間','學生id','學費年度','學生姓名','繳費月份','金額','狀態'];
  var range   = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground('#1c1c1e');
  range.setFontColor('white');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

// ============================================================
//  獎學金計算邏輯
// ============================================================

function calcScholarship(current, prev) {
  var candidates = [];
  if (current === 100) candidates.push(500);
  if (current >= 90)   candidates.push(300);
  if (prev !== null && prev !== undefined && !isNaN(prev) && (current - prev) >= 20) {
    candidates.push(200);
  }
  return candidates.length > 0 ? Math.max.apply(null, candidates) : 0;
}

// ============================================================
//  新增成績紀錄
// ============================================================

function addStudentScore(studentId, examName, currentScore, prevScore) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SCORE_SHEET) || createScoreSheet(ss);

  // 取學生姓名
  var sSheet  = ss.getSheetByName(STUDENT_SHEET);
  var sData   = sSheet ? sSheet.getDataRange().getValues() : [];
  var sName   = '';
  for (var i = 1; i < sData.length; i++) {
    if (sData[i][0] === studentId) { sName = sData[i][3]; break; }
  }

  var improvement = (prevScore !== null && prevScore !== undefined && !isNaN(prevScore)) ? (currentScore - prevScore) : '';
  var earned      = calcScholarship(currentScore, prevScore);
  var id          = getNextIdBySheet(sheet);

  sheet.appendRow([id, new Date().toISOString(), studentId, sName, examName,
                   prevScore !== null && prevScore !== undefined ? prevScore : '',
                   currentScore, improvement, earned]);

  return { scholarship_earned: earned, improvement: improvement };
}

// ============================================================
//  取得成績歷史
// ============================================================

function getStudentScoreHistory(studentId) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SCORE_SHEET);
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === studentId) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
      result.push(obj);
    }
  }
  return result.reverse();
}

// ============================================================
//  取得獎學金餘額
// ============================================================

function getStudentScholarshipBalance(studentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // 累計獲得
  var scoreSheet = ss.getSheetByName(SCORE_SHEET);
  var earned = 0;
  if (scoreSheet) {
    var sd = scoreSheet.getDataRange().getValues();
    for (var i = 1; i < sd.length; i++) {
      if (sd[i][2] === studentId) earned += (parseFloat(sd[i][8]) || 0);
    }
  }

  // 累計使用
  var usedSheet = ss.getSheetByName(SCHOLARSHIP_SHEET);
  var used = 0;
  if (usedSheet) {
    var ud = usedSheet.getDataRange().getValues();
    for (var i = 1; i < ud.length; i++) {
      if (ud[i][2] === studentId) used += (parseFloat(ud[i][5]) || 0);
    }
  }

  return { balance: earned - used, earned: earned, used: used };
}

// ============================================================
//  使用獎學金（折抵或領現）
// ============================================================

function useStudentScholarship(studentId, amount, type, remark) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SCHOLARSHIP_SHEET) || createScholarshipSheet(ss);

  // 取學生姓名
  var sSheet = ss.getSheetByName(STUDENT_SHEET);
  var sData  = sSheet ? sSheet.getDataRange().getValues() : [];
  var sName  = '';
  for (var i = 1; i < sData.length; i++) {
    if (sData[i][0] === studentId) { sName = sData[i][3]; break; }
  }

  var id = getNextIdBySheet(sheet);
  sheet.appendRow([id, new Date().toISOString(), studentId, sName, type, amount, remark || '']);
}

// ============================================================
//  批次匯入成績
// ============================================================

function importStudentScoresBatch(scoresArray, examName) {
  // scoresArray: [{student_id, student_name, current_score, prev_score}]
  var results = [];
  for (var i = 0; i < scoresArray.length; i++) {
    var item    = scoresArray[i];
    var sid     = parseInt(item.student_id);
    var cur     = parseFloat(item.current_score);
    var prev    = (item.prev_score !== '' && item.prev_score !== null && item.prev_score !== undefined) ? parseFloat(item.prev_score) : null;
    if (isNaN(sid) || isNaN(cur)) continue;
    var r = addStudentScore(sid, examName, cur, prev);
    results.push({ student_id: sid, student_name: item.student_name, result: r });
  }
  return results;
}

// ============================================================
//  建立成績紀錄工作表
// ============================================================

function createScoreSheet(ss) {
  var sheet   = ss.insertSheet(SCORE_SHEET);
  var headers = ['id','記錄時間','學生id','學生姓名','考試名稱','前次成績','當次成績','進步分數','獲得獎學金'];
  var range   = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground('#1c1c1e');
  range.setFontColor('white');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
  [40,160,70,80,120,80,80,80,100].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  return sheet;
}

// ============================================================
//  建立獎學金紀錄工作表
// ============================================================

function createScholarshipSheet(ss) {
  var sheet   = ss.insertSheet(SCHOLARSHIP_SHEET);
  var headers = ['id','記錄時間','學生id','學生姓名','類型','金額','備註'];
  var range   = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground('#1c1c1e');
  range.setFontColor('white');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
  [40,160,70,80,60,70,200].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  return sheet;
}

function logPayment(studentId, year, name, status, amount) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(PAYMENT_SHEET) || createPaymentSheet(ss);
  var now   = new Date();
  var month = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
  sheet.appendRow([new Date().toISOString(), studentId, year, name, month, amount||'', status]);
}

// ============================================================
//  工具函式
// ============================================================

function getNextId(sheet) {
  return getNextIdBySheet(sheet);
}

function getNextIdBySheet(sheet) {
  var last = sheet.getLastRow();
  if (last <= 1) return 1;
  var ids = sheet.getRange(2, 1, last-1, 1).getValues().flat().filter(Number);
  return ids.length > 0 ? Math.max.apply(null, ids) + 1 : 1;
}

function extractFee(courseType) {
  if (!courseType) return '';
  var m = courseType.match(/(\d+)$/);
  return m ? parseInt(m[1]) : '';
}

function getAdminToken() {
  var t = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  return t || 'paco2024';
}

// 想修改後台密碼時，改這裡的字串再執行一次
function setAdminToken() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', 'paco2024');
  Logger.log('密碼已設定');
}

// ============================================================
//  複習班相關函式
// ============================================================

function getReviewFee(classType, identity) {
  if (identity === '班內生') return 2500;
  return classType === '生物複習班' ? 3900 : 3600;
}

function saveReviewStudents(data) {
  var ids = [];
  if (data.bio_class === 'on' || data.bio_class === true || data.bio_class === '是') {
    ids.push(saveOneReviewStudent(data, '生物複習班'));
  }
  if (data.phy_class === 'on' || data.phy_class === true || data.phy_class === '是') {
    ids.push(saveOneReviewStudent(data, '理化複習班'));
  }
  return ids;
}

function saveOneReviewStudent(data, classType) {
  var ss     = SpreadsheetApp.openById(SHEET_ID);
  var sheet  = ss.getSheetByName(REVIEW_SHEET) || createReviewSheet(ss);
  var id     = getNextIdBySheet(sheet);
  var identity = (data.is_member === '是') ? '班內生' : '班外生';
  var fee    = getReviewFee(classType, identity);
  sheet.appendRow([
    id,
    data.submitted_at || new Date().toISOString(),
    data.tuition_year || '116',
    data.student_name || '',
    data.school       || '',
    data.grade        || '',
    data.contact      || '',
    data.address      || '',
    classType,
    identity,
    fee,
    '2026-07',
    '2026-08',
    '待確認',
    '',
    data.notes || ''
  ]);
  return id;
}

function getAllReviewStudents() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(REVIEW_SHEET);
  if (!sheet) return [];
  var tz      = Session.getScriptTimeZone();
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
    ['課程開始', '課程截止'].forEach(function(col) {
      if (obj[col] instanceof Date) {
        obj[col] = Utilities.formatDate(obj[col], tz, 'yyyy-MM');
      }
    });
    result.push(obj);
  }
  return result;
}

function recordReviewStudentPayment(studentId, paymentDate, amount) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(REVIEW_SHEET);
  if (!sheet) return;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var stIdx   = headers.indexOf('繳費狀態');
  var dtIdx   = headers.indexOf('繳費日期');
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      sheet.getRange(i + 1, stIdx + 1).setValue('已繳');
      if (dtIdx >= 0) sheet.getRange(i + 1, dtIdx + 1).setValue(paymentDate || new Date().toISOString().slice(0, 10));
      break;
    }
  }
}

function markReviewStudentInactive(studentId) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(REVIEW_SHEET);
  if (!sheet) return;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var stIdx   = headers.indexOf('繳費狀態');
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      sheet.getRange(i + 1, stIdx + 1).setValue('不續班');
      break;
    }
  }
}

function createReviewSheet(ss) {
  var sheet   = ss.insertSheet(REVIEW_SHEET);
  var headers = ['id','報名時間','學費年度','學生姓名','就讀學校','年級','聯絡方式','地址','班別','身份','費用','課程開始','課程截止','繳費狀態','繳費日期','備註'];
  var range   = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground('#1c1c1e');
  range.setFontColor('white');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
  [40,160,70,80,130,70,160,250,100,70,70,80,80,80,80,160].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  return sheet;
}

function sendReviewNotificationEmail(data, ids) {
  var classes = [];
  if (data.bio_class === 'on' || data.bio_class === '是') classes.push('生物複習班');
  if (data.phy_class === 'on' || data.phy_class === '是') classes.push('理化複習班');
  var body = [
    '✅ 複習班報名通知',
    '─────────────',
    '編號：#' + ids.join(', '),
    '學生：' + data.student_name,
    '學校：' + data.school + '　' + (data.grade || ''),
    '聯絡：' + data.contact,
    '身份：' + (data.is_member === '是' ? '班內生' : '班外生'),
    '報名班別：' + (classes.join('、') || '無'),
    '地址：' + data.address,
    '備註：' + (data.notes || '無'),
    '─────────────',
    '時間：' + data.submitted_at,
  ].join('\n');
  MailApp.sendEmail(ADMIN_EMAIL, '【巨石玩理】複習班報名 ' + data.student_name, body);
}

// ============================================================
//  步驟二：先執行這個函式，建立試算表結構
// ============================================================

function initializeSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  if (!ss.getSheetByName(STUDENT_SHEET))     createStudentSheet(ss);
  if (!ss.getSheetByName(PAYMENT_SHEET))     createPaymentSheet(ss);
  if (!ss.getSheetByName(SCORE_SHEET))       createScoreSheet(ss);
  if (!ss.getSheetByName(SCHOLARSHIP_SHEET)) createScholarshipSheet(ss);
  if (!ss.getSheetByName(REVIEW_SHEET))      createReviewSheet(ss);
  Logger.log('✅ 工作表建立完成（含成績紀錄、獎學金紀錄、複習班名單）');
}

// ============================================================
//  步驟二：再執行這個函式，匯入 119 筆歷史學生資料
// ============================================================

function importHistoricalData() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STUDENT_SHEET);

  if (!sheet) {
    Logger.log('❌ 請先執行 initializeSheets()');
    return;
  }

  if (sheet.getLastRow() > 1) {
    Logger.log('⚠️ 工作表已有資料，請確認後再執行（避免重複匯入）');
    return;
  }

  var rows = [
[1,"2026-03-02","116","張暻奕","照南國中","","0921370721","","苗栗縣竹南鎮博愛街81號","四個月9000",9000,"2026-03","2026-06","待確認",""],
[2,"2026-03-12","116","洪千喨","薇閣中學","","0933005114","","台北市北投區復興四路110號2樓","四個月9000",9000,"2026-03","2026-06","待確認",""],
[3,"2026-03-03","116","陳韋潔","台北私立復興中學","","0928844229（line ID)","","台北市松山區南京東路五段250巷18弄11-2號1樓","四個月9000",9000,"2026-03","2026-06","待確認",""],
[4,"2026-03-12","116","黃敬翔","東大附中","","0975566538","","台中市清水區海濱路183-1號","四個月9000",9000,"2026-03","2026-06","待確認",""],
[5,"2026-03-02","116","何書佑","台南市麻豆國中","","0981885826，kikia0929","","台南市麻豆區油車里新生北路82巷34號","四個月9000",9000,"2026-03","2026-06","待確認",""],
[6,"2026-03-02","116","李宛霏","聖功女中","","0970540145","","台南市安南區中安街一段146巷92弄3號","四個月9000",9000,"2026-03","2026-06","待確認",""],
[7,"2026-03-02","116","陳妤柔","社頭國中","","0988221112","","彰化縣田中鎮大社里社斗路二段53號","四個月9000",9000,"2026-03","2026-06","待確認",""],
[8,"2026-03-02","116","蔡宜伊","鳳鳴國中","","0937047522","","鳳鳴一街195號6樓","四個月9000",9000,"2026-03","2026-06","待確認",""],
[9,"2026-03-02","116","蘇昱瑋","中山國中","","0937311266/yci0831","","高雄市前鎮區復興四路12號A棟3F-3","四個月9000",9000,"2026-03","2026-06","待確認",""],
[10,"2026-03-03","116","廖于溱","新北市板橋區海山國中","","ID：0921619233","","10058台北市中正區忠孝東路一段1號","四個月9000",9000,"2026-03","2026-06","待確認",""],
[11,"2026-03-04","116","郭品妍","鹽埕國中","","0929051519 & karena-chi(line id)","","karena.chi@gmail.com","四個月9000",9000,"2026-03","2026-06","待確認",""],
[12,"2026-03-03","116","蔡嫙","大有國中","","0922661756","","桃園市桃園區民有十五街26號8樓","四個月9000",9000,"2026-03","2026-06","待確認",""],
[13,"2026-03-04","116","宮靖洋","五福國中","","0989610824","","高雄市三民區凱旋一路143號四樓","四個月9000",9000,"2026-03","2026-06","待確認",""],
[14,"2026-03-02","116","吳文樂","七賢國中","","0921251695&lineid:0921251695","","高雄市鼓山區美術南三路303號29樓","四個月9500",9500,"2026-03","2026-06","待確認",""],
[15,"2026-03-03","116","蔡珽安","湖口-新湖國中","","0953811513 & vmpnadia (邱小薰）","","新竹縣新豐鄉中崙村5鄰202-16號","四個月9500",9500,"2026-03","2026-06","待確認",""],
[16,"2026-03-03","116","游秉叡","二重國中","","0988236222，dorislee6988","","新竹縣竹東鎮中興路二段152巷39弄12號3樓","四個月9500",9500,"2026-03","2026-06","待確認",""],
[17,"2026-03-30","116","林宣喬","草屯國中","","0908620808、line: ifunlife","","南投縣草屯鎮御富路430號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[18,"2026-03-04","116","品沂","烏眉國中","","0916-207709","","苗栗縣通霄鎮通南里中山路36號","四個月9500",9500,"2026-03","2026-06","待確認",""],
[19,"2026-03-04","116","張望宣","和平高中國中部","","0939699778","","新北市新店區中央六街27號三樓","四個月9500",9500,"2026-03","2026-06","待確認",""],
[20,"2026-03-06","116","郭芯語","花壇國中","","Line:yan810728","","彰化縣大村鄉山腳路144巷38號","四個月9500",9500,"2026-03","2026-06","待確認",""],
[21,"2026-03-03","116","張家緁","民生國中","","0910145493 & 珮珮","","台北市內湖區金湖路403之一號1樓","四個月9500",9500,"2026-03","2026-06","待確認",""],
[22,"2026-02-24","116","旻吟","瀛海國中","","0927190770","","台南市安定區安加里安定295-4號","六個月14400",14400,"2026-03","2026-08","待確認",""],
[23,"2026-03-02","116","蔡承翰","福山國中","","0920816856","","高雄市左營區榮總路420號","四個月9000",9000,"2026-03","2026-06","待確認",""],
[24,"2026-03-02","116","黃少甫","台中市私立弘文中學","","0955808326","","台中市北屯區軍南二街27號","四個月9500",9500,"2026-03","2026-06","待確認",""],
[25,"2026-05-11","116","張祐綸","鹽行國中","","0987085095 szuyen0715","","台南市歸仁區大埔路一段109-3號","六個月14400",14400,"2026-05","2026-01","待確認",""],
[26,"2026-05-06","116","鄭羽彤","社頭國中","","0934063518 & regine_h","","彰化縣社頭鄉張厝二巷456號","六個月14400",14400,"2026-05","2026-01","待確認",""],
[27,"2026-04-07","116","翌桀","中正國中","","0976228979","","屏東市公民街15號4樓之一","六個月14400",14400,"2026-04","2026-09","待確認",""],
[28,"2026-03-31","116","王宜閔","民德國中","","0970685865","","嘉義縣民雄鄉西安村進興街132巷16號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[29,"2026-03-31","116","林恩守","彰化市立彰安國中","","0931643331","","50567彰化縣鹿港鎮彰頂路54號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[30,"2026-04-02","116","黃丞煒","清泉國中","","0919854646、namikuo1104","","台中市清水區民有路278巷15-2號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[31,"2026-01-08","116","張詠淩","弘文中學","","0916728729","","台中市豐原區富陽路61之2號八樓之二","六個月14400",14400,"2026-01","2026-06","待確認",""],
[32,"2026-01-24","116","廖品涵","文興高中（國中部）","","0920529636","","彰化縣社頭鄉廣福村社石路785號5樓之2","六個月14400",14400,"2026-01","2026-07","待確認",""],
[33,"2026-01-09","116","張桀瑞","照南國中","","0982331076/0982331076","","苗栗縣竹南鎮大營路14巷31號","六個月14400",14400,"2026-01","2026-06","待確認",""],
[34,"2026-01-09","116","陳品伃","育英國中","","0937277357&Line ID:vivian668","","台中市東區建功街171之8號","六個月14400",14400,"2026-01","2026-06","待確認",""],
[35,"2025-09-29","116","楊濬澤","安平國中","","0918309872","","台南市安平區健康三街165巷32號","六個月14400",14400,"2026-02","2026-07","待確認",""],
[36,"2026-03-02","116","吳芷妘","二重國中","","0911103653&irene0911103653","","新竹縣竹東鎮光明路126巷65弄30號10樓","六個月14400",14400,"2026-03","2026-08","待確認",""],
[37,"2026-03-09","116","沈雨葇","和平","","0931385931/sugertseng1207","","新北市永和區永貞路324號4樓","六個月14400",14400,"2026-03","2026-08","待確認",""],
[38,"2026-03-04","116","李羿璇","康橋","","0963033229","","新竹市東區光華街27巷15弄4號2樓","六個月14400",14400,"2026-03","2026-08","待確認",""],
[39,"2026-04-01","116","張馨予","草屯國中","","0958398727","","南投縣名間鄉南雲路13巷29號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[40,"2026-04-03","116","林子程","華興","","CHAN-SHU-LI","","新北市蘆洲區成功路72巷1號2樓","四個月10000",10000,"2026-04","2026-07","待確認",""],
[41,"2026-04-01","116","莊詠媗","馬公國中","","0952930617   pin333555","","880澎湖縣馬公市樹德路36號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[42,"2026-04-01","116","李力衡","龍華國中","","0986731127","","807高雄市三民區察哈爾一街106號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[43,"2026-04-27","116","王紫暄","草屯國中","","0910242048","","南投縣草屯鎮上林里草溪路212之29號","四個月10000",10000,"2026-04","2026-07","待確認",""],
[44,"2026-04-18","116","湯苡禎","明仁國中","","0910151247(沒有line ID )","","苗栗縣苗栗市大同路133號8樓（病理科）","四個月10000",10000,"2026-04","2026-07","待確認",""],
[45,"2026-04-28","116","黃榆棓","正興國中","","0935489317","","高雄市三民區大豐二路135巷10弄8號","六個月14400",14400,"2026-04","2026-09","待確認",""],
[46,"2025-12-20","116","李承翰","精誠中學","","0902193868","","彰化縣和美鎮鐵勢路170號","六個月14400",14400,"2025-12","2025-06","待確認",""],
[47,"2026-05-06","116","田有瀚","林口崇林國中","","0965321355","","桃園市龜山區長庚醫護新村332號1樓","四個月10000",10000,"2026-05","2026-08","待確認",""],
[48,"2025-12-23","116","涂天淨","大華國民中學","","0933411633&binbyun","","台中市大雅區民族街147北二巷17號","六個月14400",14400,"2026-05","2026-01","待確認",""],
[49,"2026-05-04","116","李芮綺","新莊國中","","0920341663/lala23576","","新北市泰山區壽山路134巷7號","四個月10000",10000,"2026-05","2026-08","待確認",""],
[50,"2026-01-20","116","張彥勛","彰化精誠中學","","0931-509619&otnv4460","","彰化縣北斗鎮文苑東路69號","六個月14400",14400,"2026-02","2026-07","待確認",""],
[51,"2026-01-24","116","陳品妍","薇閣中學","","0919068667 / amelie068","","247新北市蘆洲區長興路103號9樓","六個月14400",14400,"2026-02","2026-07","待確認",""],
[52,"2026-01-25","116","王霈璇","鳳西國中","","0952952830，marta0819","","高雄市鳳山區中山西路109號5樓","四個月10000",10000,"2026-02","2026-05","待確認",""],
[53,"2026-01-26","116","張育銓","南榮國中","","0930397808","","932屏東縣新園鄉新東村平和路338-1號","六個月14400",14400,"2026-02","2026-07","待確認",""],
[54,"2026-01-26","116","洪毓茹","安和國中","","0987638229 & @marcella0131","","（407）台中市西屯區台灣大道四段859號8樓","四個月10000",10000,"2026-02","2026-05","待確認",""],
[55,"2026-01-26","116","林品睿","磐石國中","","0926153808","","新竹市香山區牛埔南路142巷33弄41號4樓","六個月14400",14400,"2026-02","2026-07","待確認",""],
[56,"2026-01-27","116","陳彥碩","林口康橋","","0916706170","","桃園市蘆竹區長興路3段229巷60號","四個月10000",10000,"2026-02","2026-05","待確認",""],
[57,"2026-01-27","116","林辰澐","蘭雅國中","","0912285581 @bw015432","","台北市士林區中山北路六段35巷8號10F","六個月14400",14400,"2026-02","2026-07","待確認",""],
[58,"2026-01-27","116","陳筠媗","台中大業國中","","0988293220&joanna32212","","台中市西屯區大墩路895號2樓","四個月10000",10000,"2026-02","2026-05","待確認",""],
[59,"2026-01-27","116","胡芫瑞","福和國中","","0913559677＆joanne7955","","新北市新店區安康路二段341巷35號4樓","六個月14400",14400,"2026-02","2026-07","待確認",""],
[60,"2026-01-28","116","林彥承","明湖國中","","0919220310","","台北市內湖區民權東路六段21巷33號6樓","四個月10000",10000,"2026-02","2026-05","待確認",""],
[61,"2026-01-28","116","魏品妍","博愛國中（竹北市）","","0937935450／joy0216","","新竹縣竹北市中正西路555巷20號","四個月10000",10000,"2026-02","2026-05","待確認",""],
[62,"2026-01-28","116","吳承翰","介壽國中","","0928570085","","台北市松山區光復北路２３０巷２５號１１樓","四個月10000",10000,"2026-02","2026-05","待確認",""],
[63,"2026-01-28","116","方靖雁","瑞祥國中","","0938922639&chienchi0919","","高雄市鼓山區蓬萊路30號","六個月14400",14400,"2026-02","2026-07","待確認",""],
[64,"2026-01-29","116","楊栯昕","後甲國中","","0921246225","","台南市東區德光街15巷39號5樓","六個月14400",14400,"2026-02","2026-07","待確認",""],
[65,"2026-01-29","116","賴璟鈞","鹿寮國中","","0976-398-352  line:2729881","","台中市清水區中央路99之77號","六個月14400",14400,"2026-02","2026-07","待確認",""],
[66,"2026-01-30","116","許豫柔","金城國中","","0937392807","","金門縣金城鎮民權路226巷4弄21號1樓","六個月14400",14400,"2026-02","2026-07","待確認",""],
[67,"2026-01-31","116","翁翊浚","復華國中部","","0931862325/0912369039","","高雄市前金區中正四路117號10樓之1","六個月14400",14400,"2026-02","2026-07","待確認",""],
[68,"2026-05-23","116","歐以晴","天母國中","","0910807903&  angel0903ou","","台北市北投區同德街17-3號2 F","六個月14400",14400,"2026-07","2026-12","待確認",""],
[69,"2026-02-01","116","陳毅","國風國中","","0911276552","","花蓮縣新城鄉佳林村佳林6-6號","四個月10000",10000,"2026-02","2026-05","待確認",""],
[70,"2026-02-03","116","曾士倫","三民國中(台北)","","0920-015-663 & Line ID:540ivy","","台北市內湖區民權東路六段180巷42弄18號1F","六個月14400",14400,"2026-02","2026-07","待確認",""],
[71,"2026-02-03","116","林妍蕎","陽明國中","","0988352238","","高雄市三民區昌富街69號","四個月10000",10000,"2026-02","2026-05","待確認",""],
[72,"2026-05-23","116","林筠宜","新園國中","","0988352238","","屏東縣新園鄉田洋村仙吉路797之8號","六個月14400",14400,"2026-05","2026-11","待確認",""],
[73,"2026-02-05","116","呂曜丞","壽山國中","","0937321561&937321561","","高雄市鼓山區麗雄街101號","六個月14400",14400,"2026-03","2026-08","待確認",""],
[74,"2026-02-14","116","黃映慈","育賢國中","","0927205033&ping6995","","新竹市東區明湖路243巷115號3樓","四個月10000",10000,"2026-02","2026-06","待確認",""],
[75,"2026-02-23","116","許祐睿","馬公國中","","0988557510","","澎湖縣馬公市石泉里32-52號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[76,"2026-02-23","116","王涵緹","新竹市 培英國中","","0982016538 & LeoWang","","新竹市力行一路1號 3A1-3 (台灣彩光)","六個月14400",14400,"2026-03","2026-08","待確認",""],
[77,"2026-05-23","116","賴彥廷","積穗國中","","0906275970","","新北市中和區民利街140號二樓","六個月14400",14400,"2026-06","2026-11","待確認",""],
[78,"2026-02-23","116","林廷叡","宜蘭市復興國中","","0911219056","","宜蘭縣員山鄉賢好路23巷33弄9號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[79,"2026-02-23","116","邱于禎","屏榮中學","","0935951789","","屏東市民生路79-16號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[80,"2026-02-23","116","劉祐丞","福營國中","","0918819834","","242新北市新莊區建安街76巷3弄5號7樓","六個月14400",14400,"2026-03","2026-08","待確認",""],
[81,"2026-02-25","116","蔡沛安","文府國中","","0968335856/chia_yi_yen","","高雄市左營區文學路193號13樓","四個月10000",10000,"2026-03","2026-06","待確認",""],
[82,"2026-02-26","116","朱亮暻","崇光中學","","0928179099","","新北市新店區建國路108號10樓之1","六個月14400",14400,"2026-03","2026-08","待確認",""],
[83,"2026-02-26","116","黃士洺","桃園慈文","","0953196236 (已加line）","","33754 桃園市大園區三和路28巷8號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[84,"2026-02-26","116","林婷瑀","台南私立長榮中學國中部","","0980909666／jean-0401","","台南市南區明興路1255號","六個月14400",14400,"2026-03","2026-08","待確認",""],
[85,"2026-02-26","116","傅楷荃","臺中市立鹿寮國民中學","","092172633  line : tendy.fu","","台中市南屯區龍富十路100號10樓之1","六個月14400",14400,"2026-03","2026-08","待確認",""],
[86,"2026-02-26","116","洪鏗翰","蘭雅國中","","0960503800 / agogomm","","110台北市信義區松隆路102號6樓","四個月10000",10000,"2026-03","2026-06","待確認",""],
[87,"2026-02-27","116","翁靖喬","康橋國中新竹校區","","0920191055","","新竹縣竹北市六家七路228巷26號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[88,"2026-02-28","116","廖苡彤","彰化成功國中","","0922285563","","彰化縣溪湖鎮興農街159號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[89,"2026-03-01","116","廖品貽","崇德國中","","0985660781","","台中市北屯區崇德二路一段157巷7號5樓之三","四個月10000",10000,"2026-03","2026-06","待確認",""],
[90,"2026-03-01","116","石宥婕","景美國中","","0928603306","","新北市新店區700巷52號6樓","四個月10000",10000,"2026-03","2026-06","待確認",""],
[91,"2026-03-05","116","林立豈","大業國中","","0925118228","","台中市西屯區四川路87巷57號4樓之1","四個月10000",10000,"2026-03","2026-06","待確認",""],
[92,"2026-03-06","116","王映辰","居仁國中","","0952033125  & sisi5151","","406台中市北屯區安順東六街27號2樓之三","四個月10000",10000,"2026-03","2026-06","待確認",""],
[93,"2026-03-09","116","蘇浚瑞","三峽國中","","0933693836","","新北市三峽區復興路110巷2弄10號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[94,"2026-03-12","116","凃宣伶","鳳翔國中","","0915060938","","高雄市鳳山區祥和街52巷11號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[95,"2026-03-14","116","洪椽鈞","道明國中","","0931843068&coco023188","","高雄市鳳山區北明街85號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[96,"2026-03-15","116","陳佳佑","新興國中","","0989048585&蕙嘉","","台北市文山區景華街121巷8號4樓","四個月10000",10000,"2026-03","2026-06","待確認",""],
[97,"2026-03-17","116","蘇鈺琁","國昌國中","","0910710245/ssy1216","","高雄市楠梓區智昌街606號","四個月10000",10000,"2026-03","2026-06","待確認",""],
[98,"2026-03-26","116","方嘉寶","至善國中","","0963360839/tiffanylin","","台中市北屯區中清路2段640號14樓","四個月10000",10000,"2026-04","2026-07","待確認",""],
[99,"2026-03-26","116","張右新","福豐國中","","0911863862&tine5243","","33458桃園市八德區東勇街490巷75弄70號","四個月10000",10000,"2026-04","2026-07","待確認",""],
[100,"2026-05-23","116","黃玨淏","彰化陽明國中","","0988169996/@：g660123","","彰化縣線西鄉溝內路116-8號","六個月14400",14400,"2026-07","2026-12","待確認",""],
[101,"2026-03-26","116","柏瑜","左營國中","","0933657693","","高雄市左營區文瑞路23號10樓","四個月10000",10000,"2026-04","2026-07","待確認",""],
[102,"2026-03-26","116","蘇淮媗","臺南市東區後甲國中","","0937341192/Ti","","臺南市永康區中華路22號10樓之3","四個月10000",10000,"2026-04","2026-07","待確認",""],
[103,"2026-03-26","116","侯信宇","後甲國中","","0987379212","","台南市東區東門路三段253號11樓","四個月10000",10000,"2026-04","2026-07","待確認",""],
[104,"2026-03-26","116","薛同家","台東縣關山國中","","ID跟電話0938060885","","台東縣鹿野鄉光榮路380號","四個月10000",10000,"2026-04","2026-07","待確認",""],
[105,"2026-03-26","116","陳睿豐","大義國中","","電話：0933658920","","高雄市岡山區岡山南路22號","六個月14400",14400,"2026-04","2026-07","待確認",""],
[106,"2026-03-27","116","宋允皓","石牌國中","","0915139396 & kuku7483","","台北市士林區社中街335號7樓","四個月10000",10000,"2026-04","2026-07","待確認",""],
[107,"2026-03-30","116","陳宥綾","民生國中","","0960105833 ID：0960105833","","嘉義市民生南路306號","四個月10000",10000,"2026-04","2026-07","待確認",""],
[108,"2026-03-30","116","郁綺","五權國中","","0928-553109/cintypu","","台中市西區大墩十街61號8樓之1","四個月10000",10000,"2026-04","2026-07","待確認",""],
[109,"2026-03-31","116","黃子睿","弘文中學","","0921390621","","台中市豐原區中陽路10號","四個月10000",10000,"2026-04","2026-07","待確認",""],
[110,"2026-04-04","116","謝品丞","苗栗明仁國中","","0926-573-900、emily24888","","苗栗縣苗栗市恭敬里9鄰恭敬86號","四個月10800",10800,"2026-04","2026-07","待確認",""],
[111,"2026-04-05","116","葉思妤","台中市立北新國中","","0937550970/0937550970","","台中市潭子區弘智二街128號2樓之1","四個月10800",10800,"2026-04","2026-07","待確認",""],
[112,"2026-04-09","116","蘇景璇","竹林國中","","line ID:0932193262","","新北市中和區中正路803號4樓之3","六個月15600",15600,"2026-04","2026-07","待確認",""],
[113,"2026-04-09","116","陳椲力","台南市立鹽行國中","","0912176521","","台南市永康區鹽行路146巷20弄7-15號","四個月10800",10800,"2026-04","2026-07","待確認",""],
[114,"2026-04-30","116","許晏誠","七賢國中","","0972798071/whhsu227chanchichiug","","高雄市鼓山區光榮里建榮路51巷1號","六個月15600",15600,"2026-05","2026-10","待確認",""],
[115,"2026-04-30","116","方姵勛","竹北東興國中","","0911879028","","竹北市嘉豐二街一段21號12樓","四個月10800",10800,"2026-05","2026-08","待確認",""],
[116,"2026-05-06","116","李昀杰","民權國中","","0921624351、同電話號碼","","台北市大同區酒泉街165巷1號三樓","四個月10800",10800,"2026-05","2026-08","待確認",""],
[117,"2026-05-06","116","陳柔妡","曉明女中","","0920038656(ID:aven0531)","","台中市西屯區市政北一路269號5樓之8","四個月10800",10800,"2026-05","2026-08","待確認",""],
[118,"2026-05-07","116","許光爵","中興國中","","0917784128（line同）","","南投市彰南路三段880巷26弄39號","四個月10800",10800,"2026-05","2026-08","待確認",""],
[119,"2026-05-09","116","柯宜珊","德光中學","","0912703625","","台南市北區西門路三段41號幼兒園吳妃恂老師收","四個月10800",10800,"2026-05","2026-08","待確認",""]
  ];

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // 偶數列底色
  for (var i = 0; i < rows.length; i++) {
    if ((i + 2) % 2 === 0) {
      sheet.getRange(i + 2, 1, 1, rows[0].length).setBackground('#f9f9f9');
    }
  }

  Logger.log('✅ 匯入完成，共 ' + rows.length + ' 筆學生資料');
}
