/**
 * Automated Weekly Report for ESTA / KETA / UK
 * Runs every Tuesday at 0:00 JST, writes previous Mon-Sun data to Google Spreadsheet.
 *
 * Trigger setup (one-time):
 *   ScriptApp.newTrigger('runWeeklyReport')
 *     .timeBased()
 *     .onWeekDay(ScriptApp.WeekDay.TUESDAY)
 *     .atHour(0)
 *     .create();
 *
 * Script Properties required:
 *   ESTA_URL, ESTA_KEY, KETA_URL, KETA_KEY, UKETA_URL, UKETA_KEY
 */

var SPREADSHEET_ID = '1QrcAwoTvaxKX6HREzzHxjDzcw443iFz4PKMLYYN7Z6Q';

var SERVICE_CONFIG = {
  ESTA: {
    urlProp: 'ESTA_URL',
    keyProp: 'ESTA_KEY',
    sheetName: 'ESTA',
    color: '#4472C4'
  },
  KETA: {
    urlProp: 'KETA_URL',
    keyProp: 'KETA_KEY',
    sheetName: 'KETA',
    color: '#BF8F00'
  },
  UK: {
    urlProp: 'UKETA_URL',
    keyProp: 'UKETA_KEY',
    sheetName: 'UK',
    color: '#548235'
  }
};

var STAFF_DISPLAY_NAMES = {
  '裕子': '岩澤裕子',
  '有沙': '木村有沙',
  'ひかり': '石川ひかり',
  '者': '管理者'
};

var PROCESSING_STAFF_LIST = ['岩澤裕子', '木村有沙', '石川ひかり', 'くるみ'];

var DAY_LABELS = ['月', '火', '水', '木', '金', '土', '日'];

var RATE_BG = '#FFF2CC';

// ============================================================
// Entry points
// ============================================================

/**
 * Main entry — processes previous Mon-Sun (called on Tuesday).
 */
function runWeeklyReport() {
  var now = new Date();
  // Shift to JST
  var jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  // Go back to last Monday (jstNow is Tuesday, so -1 day = Monday)
  var monday = new Date(jstNow);
  monday.setDate(monday.getDate() - (monday.getUTCDay() === 0 ? 6 : monday.getUTCDay() - 1));
  // If today is Tuesday, last Monday is 8 days ago
  // Actually let's just go back to previous Monday properly
  var dayOfWeek = jstNow.getUTCDay(); // 0=Sun,1=Mon,...
  var daysBack = dayOfWeek === 0 ? 6 : dayOfWeek - 1; // days since last Monday
  daysBack += 7; // we want PREVIOUS week's Monday
  monday = new Date(jstNow);
  monday.setUTCDate(monday.getUTCDate() - daysBack);
  monday.setUTCHours(0, 0, 0, 0);

  var startStr = Utilities.formatDate(monday, 'UTC', 'yyyy-MM-dd');
  _processWeek(startStr);
}

/**
 * Manual entry — takes the Monday date string 'yyyy-MM-dd'.
 */
function runWeeklyReportManual(startDate) {
  _processWeek(startDate);
}

// ============================================================
// Core
// ============================================================

function _processWeek(mondayStr) {
  var parts = mondayStr.split('-');
  var mondayJST = new Date(Date.UTC(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10) - 1,
    parseInt(parts[2], 10),
    0, 0, 0
  ));
  // mondayJST represents midnight JST of that Monday.
  // In UTC that is mondayJST - 9h
  var startUTC = new Date(mondayJST.getTime() - 9 * 60 * 60 * 1000);
  var endUTC = new Date(startUTC.getTime() + 7 * 24 * 60 * 60 * 1000);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var props = PropertiesService.getScriptProperties();

  var services = ['ESTA', 'KETA', 'UK'];
  for (var s = 0; s < services.length; s++) {
    var svcKey = services[s];
    var cfg = SERVICE_CONFIG[svcKey];
    var baseUrl = props.getProperty(cfg.urlProp);
    var apiKey = props.getProperty(cfg.keyProp);

    Logger.log('Processing ' + svcKey + ' ...');
    _processService(ss, cfg, baseUrl, apiKey, startUTC, endUTC, mondayJST);
  }

  Logger.log('Weekly report complete.');
}

function _processService(ss, cfg, baseUrl, apiKey, startUTC, endUTC, mondayJST) {
  var sheet = _getOrCreateSheet(ss, cfg.sheetName);
  sheet.clear();
  sheet.clearFormats();

  // ---- Fetch journal details for the week ----
  var journals = _fetchJournalsForWeek(baseUrl, apiKey, startUTC, endUTC);
  // journals: array of {journal_id, issue_id, staff_id, created_at, old_value, new_value}

  // ---- Fetch staff map ----
  var staffMap = _fetchStaffMap(baseUrl, apiKey);
  // staffMap: id -> display name

  // ---- Fetch issues to exclude ----
  var excludedIssueIds = _fetchExcludedIssueIds(baseUrl, apiKey);

  // ---- Filter out excluded issues ----
  journals = journals.filter(function(j) {
    return excludedIssueIds.indexOf(j.issue_id) === -1;
  });

  // ---- Classify events ----
  var payments = []; // {dt_jst, staff, issue_id, journal entries...}
  var refunds = [];
  var processings = []; // for processing time calc

  for (var i = 0; i < journals.length; i++) {
    var j = journals[i];
    var dtUTC = new Date(j.created_at);
    var dtJST = new Date(dtUTC.getTime() + 9 * 60 * 60 * 1000);
    var staffName = _resolveStaffName(j.staff_id, staffMap);

    if (j.old_value === '申込完了' && j.new_value === '申込決済完了') {
      payments.push({
        dt_jst: dtJST,
        staff: staffName,
        issue_id: j.issue_id,
        hour: dtJST.getUTCHours()
      });
    }
    if (j.new_value === '返金依頼-手数料のみ') {
      refunds.push({
        dt_jst: dtJST,
        staff: staffName,
        issue_id: j.issue_id,
        hour: dtJST.getUTCHours()
      });
    }
    if (j.old_value === '申込決済完了' && j.new_value === '手動申請結果待ち') {
      processings.push({
        dt_jst: dtJST,
        staff: staffName,
        issue_id: j.issue_id,
        hour: dtJST.getUTCHours()
      });
    }
  }

  // ---- Build per-day data for cancel rate table ----
  var cancelData = _buildCancelData(payments, refunds, mondayJST);

  // ---- Build processing time data ----
  // Need to find pairs: last 申込決済完了 -> 手動申請結果待ち for each issue
  var processingTimeData = _buildProcessingTimeData(baseUrl, apiKey, processings, startUTC, endUTC, staffMap, excludedIssueIds);

  // ---- Write Table 1: Cancel Rate ----
  var svcName = cfg.sheetName;
  var nextRow = _writeCancelRateTable(sheet, cancelData, cfg.color, mondayJST, svcName);

  // ---- Write Table 2: Staff Processing Time ----
  nextRow += 2; // gap
  _writeProcessingTimeTable(sheet, processingTimeData, cfg.color, mondayJST, nextRow, svcName);

  // Auto-resize columns
  var lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    for (var c = 1; c <= lastCol; c++) {
      sheet.autoResizeColumn(c);
    }
  }
}

// ============================================================
// Data fetching
// ============================================================

function _supabaseGet(baseUrl, apiKey, path) {
  var url = baseUrl + '/rest/v1/' + path;
  var separator = url.indexOf('?') === -1 ? '?' : '&';
  // Pagination
  var allRows = [];
  var offset = 0;
  var limit = 1000;

  while (true) {
    var pagedUrl = url + separator + 'limit=' + limit + '&offset=' + offset;
    var options = {
      method: 'get',
      headers: {
        'apikey': apiKey,
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json',
        'Prefer': 'count=exact'
      },
      muteHttpExceptions: true
    };
    var resp = UrlFetchApp.fetch(pagedUrl, options);
    var code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('Supabase error ' + code + ': ' + resp.getContentText().substring(0, 500));
      break;
    }
    var rows = JSON.parse(resp.getContentText());
    if (!rows || rows.length === 0) break;
    allRows = allRows.concat(rows);
    if (rows.length < limit) break;
    offset += limit;
  }
  return allRows;
}

function _fetchJournalsForWeek(baseUrl, apiKey, startUTC, endUTC) {
  // Fetch journal_details with field_name=status_code, joined with journals
  var startISO = startUTC.toISOString();
  var endISO = endUTC.toISOString();

  var path = 'journal_details?select=id,old_value,new_value,journal_id,journals!inner(id,created_at,issue_id,staff_id)'
    + '&field_name=eq.status_code'
    + '&journals.created_at=gte.' + startISO
    + '&journals.created_at=lt.' + endISO;

  var rows = _supabaseGet(baseUrl, apiKey, path);

  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var j = r.journals;
    result.push({
      journal_id: j.id,
      issue_id: j.issue_id,
      staff_id: j.staff_id,
      created_at: j.created_at,
      old_value: r.old_value,
      new_value: r.new_value
    });
  }
  return result;
}

function _fetchStaffMap(baseUrl, apiKey) {
  var rows = _supabaseGet(baseUrl, apiKey, 'staffs?select=id,firstname');
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var name = rows[i].firstname || '';
    var display = STAFF_DISPLAY_NAMES[name] || name;
    map[rows[i].id] = display;
  }
  return map;
}

function _fetchExcludedIssueIds(baseUrl, apiKey) {
  var rows = _supabaseGet(baseUrl, apiKey,
    'issues?select=id&status_code=in.(other,entry_before_payment)');
  var ids = [];
  for (var i = 0; i < rows.length; i++) {
    ids.push(rows[i].id);
  }
  return ids;
}

function _resolveStaffName(staffId, staffMap) {
  return staffMap[staffId] || staffId || '不明';
}

// ============================================================
// Cancel rate data
// ============================================================

function _buildCancelData(payments, refunds, mondayJST) {
  // 7 days: Mon=0 .. Sun=6
  var data = {
    totalPayments: [0, 0, 0, 0, 0, 0, 0],
    whPayments: [0, 0, 0, 0, 0, 0, 0],   // working hours (hour >= 9)
    whRefunds: [0, 0, 0, 0, 0, 0, 0],
    ohPayments: [0, 0, 0, 0, 0, 0, 0],    // outside hours
    ohRefunds: [0, 0, 0, 0, 0, 0, 0]
  };

  for (var i = 0; i < payments.length; i++) {
    var p = payments[i];
    var dayIdx = _dayIndex(p.dt_jst, mondayJST);
    if (dayIdx < 0 || dayIdx > 6) continue;
    data.totalPayments[dayIdx]++;
    if (p.hour >= 9) {
      data.whPayments[dayIdx]++;
    } else {
      data.ohPayments[dayIdx]++;
    }
  }

  for (var i = 0; i < refunds.length; i++) {
    var r = refunds[i];
    var dayIdx = _dayIndex(r.dt_jst, mondayJST);
    if (dayIdx < 0 || dayIdx > 6) continue;
    if (r.hour >= 9) {
      data.whRefunds[dayIdx]++;
    } else {
      data.ohRefunds[dayIdx]++;
    }
  }

  return data;
}

function _dayIndex(dtJST, mondayJST) {
  var diff = dtJST.getTime() - mondayJST.getTime();
  return Math.floor(diff / (24 * 60 * 60 * 1000));
}

// ============================================================
// Processing time data
// ============================================================

function _buildProcessingTimeData(baseUrl, apiKey, processings, startUTC, endUTC, staffMap, excludedIssueIds) {
  // For each processing event (申込決済完了 -> 手動申請結果待ち), we need the payment time.
  // Collect unique issue_ids from processings
  var issueIds = [];
  for (var i = 0; i < processings.length; i++) {
    if (issueIds.indexOf(processings[i].issue_id) === -1) {
      issueIds.push(processings[i].issue_id);
    }
  }

  // Fetch all 申込決済完了 events for these issues to find the last payment datetime
  var paymentMap = {}; // issue_id -> [{dt_utc}]
  if (issueIds.length > 0) {
    // Fetch in batches to avoid URL length limits
    var batchSize = 50;
    for (var b = 0; b < issueIds.length; b += batchSize) {
      var batch = issueIds.slice(b, b + batchSize);
      var idsParam = '(' + batch.join(',') + ')';
      var path = 'journal_details?select=id,old_value,new_value,journal_id,journals!inner(id,created_at,issue_id,staff_id)'
        + '&field_name=eq.status_code'
        + '&new_value=eq.申込決済完了'
        + '&journals.issue_id=in.' + idsParam;
      var rows = _supabaseGet(baseUrl, apiKey, path);
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var issId = r.journals.issue_id;
        if (!paymentMap[issId]) paymentMap[issId] = [];
        paymentMap[issId].push({
          dt_utc: new Date(r.journals.created_at)
        });
      }
    }
  }

  // For each processing event, find the last 申込決済完了 before it, compute time diff
  var results = []; // {staff, dt_jst, minutes, issue_id}

  for (var i = 0; i < processings.length; i++) {
    var proc = processings[i];
    var procDtUTC = new Date(proc.dt_jst.getTime() - 9 * 60 * 60 * 1000); // back to UTC
    var payEvents = paymentMap[proc.issue_id] || [];

    // Find the last payment event before (or at) this processing event
    var lastPayDt = null;
    for (var p = 0; p < payEvents.length; p++) {
      var pdt = payEvents[p].dt_utc;
      if (pdt.getTime() <= procDtUTC.getTime()) {
        if (!lastPayDt || pdt.getTime() > lastPayDt.getTime()) {
          lastPayDt = pdt;
        }
      }
    }

    if (!lastPayDt) continue;

    var payDtJST = new Date(lastPayDt.getTime() + 9 * 60 * 60 * 1000);
    // Only include if payment was during working hours (hour >= 9)
    if (payDtJST.getUTCHours() < 9) continue;

    var minutes = (procDtUTC.getTime() - lastPayDt.getTime()) / (60 * 1000);

    results.push({
      staff: proc.staff,
      dt_jst: proc.dt_jst,
      minutes: minutes,
      issue_id: proc.issue_id
    });
  }

  return results;
}

// ============================================================
// Write Cancel Rate Table
// ============================================================

function _writeCancelRateTable(sheet, data, serviceColor, mondayJST, serviceName) {
  var startRow = 1;

  // Title row
  var endDateJST = new Date(mondayJST.getTime() + 6 * 24 * 60 * 60 * 1000);
  var titleStartMM = ('0' + (mondayJST.getUTCMonth() + 1)).slice(-2);
  var titleStartDD = ('0' + mondayJST.getUTCDate()).slice(-2);
  var titleEndMM = ('0' + (endDateJST.getUTCMonth() + 1)).slice(-2);
  var titleEndDD = ('0' + endDateJST.getUTCDate()).slice(-2);
  var title = serviceName + ' キャンセル率（' + titleStartMM + '/' + titleStartDD + '〜' + titleEndMM + '/' + titleEndDD + '）';
  sheet.getRange(startRow, 1).setValue(title).setFontWeight('bold').setFontColor(serviceColor);
  startRow++;

  // Generate date strings for header
  var dateHeaders = [];
  for (var d = 0; d < 7; d++) {
    var dt = new Date(mondayJST.getTime() + d * 24 * 60 * 60 * 1000);
    var mm = ('0' + (dt.getUTCMonth() + 1)).slice(-2);
    var dd = ('0' + dt.getUTCDate()).slice(-2);
    dateHeaders.push(mm + '/' + dd + '(' + DAY_LABELS[d] + ')');
  }

  // Header row
  var headerRow = [''].concat(dateHeaders).concat(['合計', '率平均']);
  sheet.getRange(startRow, 1, 1, headerRow.length).setValues([headerRow]);
  sheet.getRange(startRow, 1, 1, headerRow.length)
    .setBackground(serviceColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  // Helper to compute rate
  function rate(cancels, payments) {
    if (payments === 0) return 0;
    return cancels / payments * 100;
  }

  function fmtRate(r) {
    return r.toFixed(1) + '%';
  }

  function avgNonZeroRates(arr) {
    var vals = [];
    for (var i = 0; i < arr.length; i++) {
      if (arr[i] !== 0) vals.push(arr[i]);
    }
    if (vals.length === 0) return 0;
    var sum = 0;
    for (var i = 0; i < vals.length; i++) sum += vals[i];
    return sum / vals.length;
  }

  function sumArr(arr) {
    var s = 0;
    for (var i = 0; i < arr.length; i++) s += arr[i];
    return s;
  }

  // Compute rates
  var whRates = [];
  var ohRates = [];
  for (var d = 0; d < 7; d++) {
    whRates.push(rate(data.whRefunds[d], data.whPayments[d]));
    ohRates.push(rate(data.ohRefunds[d], data.ohPayments[d]));
  }

  // Rows
  var rows = [
    ['決済数'].concat(data.totalPayments).concat([sumArr(data.totalPayments), '']),
    ['時間内:決済数'].concat(data.whPayments).concat([sumArr(data.whPayments), '']),
    ['　キャンセル'].concat(data.whRefunds).concat([sumArr(data.whRefunds), '']),
    ['　キャンセル率'].concat(whRates.map(fmtRate)).concat(['', fmtRate(avgNonZeroRates(whRates))]),
    ['時間外:決済数'].concat(data.ohPayments).concat([sumArr(data.ohPayments), '']),
    ['　キャンセル'].concat(data.ohRefunds).concat([sumArr(data.ohRefunds), '']),
    ['　キャンセル率'].concat(ohRates.map(fmtRate)).concat(['', fmtRate(avgNonZeroRates(ohRates))])
  ];

  var dataStartRow = startRow + 1;
  sheet.getRange(dataStartRow, 1, rows.length, rows[0].length).setValues(rows);

  // Yellow background for rate rows (rows 4 and 7, i.e. indices 3 and 6)
  var rateRowIndices = [3, 6]; // 0-based within rows
  for (var ri = 0; ri < rateRowIndices.length; ri++) {
    var rowNum = dataStartRow + rateRowIndices[ri];
    sheet.getRange(rowNum, 2, 1, headerRow.length - 1).setBackground(RATE_BG);
  }

  // Borders for the whole table
  var totalRows = 1 + rows.length;
  var totalCols = headerRow.length;
  sheet.getRange(startRow, 1, totalRows, totalCols)
    .setBorder(true, true, true, true, true, true);

  return startRow + totalRows;
}

// ============================================================
// Write Processing Time Table
// ============================================================

function _writeProcessingTimeTable(sheet, processingData, serviceColor, mondayJST, startRow, serviceName) {
  // Title row
  sheet.getRange(startRow, 1).setValue(serviceName + ' スタッフ別処理時間（就業時間内）')
    .setFontWeight('bold').setFontColor(serviceColor);
  startRow++;

  // Header
  var headers = ['担当者', '日付', '件数', '5分以内', '5分以内率', '5〜10分', '5〜10分率',
                 '11〜30分', '11〜30分率', '30分〜', '30分〜率'];
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(startRow, 1, 1, headers.length)
    .setBackground(serviceColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  // Group processing data by staff and day
  var staffDayData = {}; // staffName -> dayIdx -> [minutes, ...]
  for (var i = 0; i < processingData.length; i++) {
    var pd = processingData[i];
    var staff = pd.staff;
    var dayIdx = _dayIndex(pd.dt_jst, mondayJST);
    if (dayIdx < 0 || dayIdx > 6) continue;

    if (!staffDayData[staff]) staffDayData[staff] = {};
    if (!staffDayData[staff][dayIdx]) staffDayData[staff][dayIdx] = [];
    staffDayData[staff][dayIdx].push(pd.minutes);
  }

  var currentRow = startRow + 1;
  var rateCols = [5, 7, 9, 11]; // 1-based columns for rate cells

  for (var si = 0; si < PROCESSING_STAFF_LIST.length; si++) {
    var staffName = PROCESSING_STAFF_LIST[si];
    var dayData = staffDayData[staffName] || {};
    var staffFirstRow = currentRow;
    var hasData = false;

    var dailyBuckets = []; // for average calculation: [{count, b5, b10, b30, b30p, r5, r10, r30, r30p}]

    for (var d = 0; d < 7; d++) {
      var mins = dayData[d];
      if (!mins || mins.length === 0) continue;
      hasData = true;

      var count = mins.length;
      var b5 = 0, b10 = 0, b30 = 0, b30p = 0;
      for (var m = 0; m < mins.length; m++) {
        if (mins[m] <= 5) b5++;
        else if (mins[m] <= 10) b10++;
        else if (mins[m] <= 30) b30++;
        else b30p++;
      }

      var r5 = count > 0 ? (b5 / count * 100) : 0;
      var r10 = count > 0 ? (b10 / count * 100) : 0;
      var r30 = count > 0 ? (b30 / count * 100) : 0;
      var r30p = count > 0 ? (b30p / count * 100) : 0;

      dailyBuckets.push({ count: count, b5: b5, b10: b10, b30: b30, b30p: b30p, r5: r5, r10: r10, r30: r30, r30p: r30p });

      // Date label
      var dt = new Date(mondayJST.getTime() + d * 24 * 60 * 60 * 1000);
      var mm = ('0' + (dt.getUTCMonth() + 1)).slice(-2);
      var dd = ('0' + dt.getUTCDate()).slice(-2);
      var dateLabel = mm + '-' + dd + ' (' + DAY_LABELS[d] + ')';

      var nameCell = (currentRow === staffFirstRow) ? staffName : '';
      var row = [nameCell, dateLabel, count,
                 b5, r5.toFixed(1) + '%',
                 b10, r10.toFixed(1) + '%',
                 b30, r30.toFixed(1) + '%',
                 b30p, r30p.toFixed(1) + '%'];
      sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);

      // Yellow bg on rate cells
      for (var rc = 0; rc < rateCols.length; rc++) {
        sheet.getRange(currentRow, rateCols[rc]).setBackground(RATE_BG);
      }

      currentRow++;
    }

    if (!hasData) continue;

    // Average row
    var totalCount = 0, totalB5 = 0, totalB10 = 0, totalB30 = 0, totalB30p = 0;
    var rateArrays = { r5: [], r10: [], r30: [], r30p: [] };

    for (var db = 0; db < dailyBuckets.length; db++) {
      var bucket = dailyBuckets[db];
      totalCount += bucket.count;
      totalB5 += bucket.b5;
      totalB10 += bucket.b10;
      totalB30 += bucket.b30;
      totalB30p += bucket.b30p;
      if (bucket.r5 > 0) rateArrays.r5.push(bucket.r5);
      if (bucket.r10 > 0) rateArrays.r10.push(bucket.r10);
      if (bucket.r30 > 0) rateArrays.r30.push(bucket.r30);
      if (bucket.r30p > 0) rateArrays.r30p.push(bucket.r30p);
    }

    function avgArr(arr) {
      if (arr.length === 0) return 0;
      var s = 0;
      for (var i = 0; i < arr.length; i++) s += arr[i];
      return s / arr.length;
    }

    var avgRow = ['', '平均', totalCount,
                  totalB5, avgArr(rateArrays.r5).toFixed(1) + '%',
                  totalB10, avgArr(rateArrays.r10).toFixed(1) + '%',
                  totalB30, avgArr(rateArrays.r30).toFixed(1) + '%',
                  totalB30p, avgArr(rateArrays.r30p).toFixed(1) + '%'];
    sheet.getRange(currentRow, 1, 1, avgRow.length).setValues([avgRow]);
    sheet.getRange(currentRow, 1, 1, avgRow.length).setFontWeight('bold');

    // Yellow bg on rate cells for average row
    for (var rc = 0; rc < rateCols.length; rc++) {
      sheet.getRange(currentRow, rateCols[rc]).setBackground(RATE_BG);
    }

    currentRow++;
  }

  // Borders for the whole processing time table
  if (currentRow > startRow + 1) {
    sheet.getRange(startRow, 1, currentRow - startRow, headers.length)
      .setBorder(true, true, true, true, true, true);
  }

  return currentRow;
}

// ============================================================
// Helpers
// ============================================================

function _getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}
