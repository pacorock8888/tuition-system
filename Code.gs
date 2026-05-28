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
//    學生歷史名單不在程式碼裡，請直接維護 Google 試算表
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
const ADMIN_EMAIL = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || '';
const STUDENT_SHEET      = '學生名單';
const PAYMENT_SHEET      = '繳費紀錄';
const SCORE_SHEET        = '成績紀錄';
const SCHOLARSHIP_SHEET  = '獎學金紀錄';

// ============================================================
//  接收學生報名（POST）
// ============================================================

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    var data = JSON.parse(e.postData.contents);
    var id   = saveStudent(data);
    sendNotificationEmail(data, id);
    output.setContent(JSON.stringify({ status: 'success', id: id }));
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

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
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
  return PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
}

// 想修改後台密碼時，改這裡的字串再執行一次
function setAdminToken() {
  throw new Error('管理密碼只可寫入 Apps Script 屬性，不可寫進程式碼');
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
  Logger.log('✅ 工作表建立完成（含成績紀錄、獎學金紀錄）');
}

// ============================================================
//  學生名單已改由 Google 試算表維護，此函式不再內嵌資料
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

  var rows = []; // student data lives in the Google Sheet

  if (!rows.length) {
    Logger.log('學生資料不在程式碼中，請維護 Google 試算表');
    return;
  }


  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // 偶數列底色
  for (var i = 0; i < rows.length; i++) {
    if ((i + 2) % 2 === 0) {
      sheet.getRange(i + 2, 1, 1, rows[0].length).setBackground('#f9f9f9');
    }
  }

  Logger.log('✅ 匯入完成，共 ' + rows.length + ' 筆學生資料');
}
