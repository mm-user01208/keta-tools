/**
 * K-ETA Ticket Processing Time Daily Report
 *
 * Google Apps Script that queries Supabase for K-ETA ticket processing data,
 * calculates processing times per staff, and writes aggregated results
 * to a Google Spreadsheet.
 *
 * Sheets:
 *   - 日報 (Daily Report): aggregated stats per staff per day
 *   - 詳細 (Details): individual ticket data
 *
 * Entry points:
 *   - runDailyReport()        — triggered daily, processes previous day
 *   - runManualReport(dateStr) — manual run for a specific JST date (e.g. '2026-04-12')
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

var CONFIG = {
  SUPABASE_URL: 'https://bmuvklukfntqmopblhox.supabase.co',
  SUPABASE_API_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJtdXZrbHVrZm50cW1vcGJsaG94Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTczMzI4NTgyNCwiZXhwIjoyMDQ4ODYxODI0fQ.5ZRquHJnInh58sLHABGIAgRR87MJRyoc6t16jv7H0Q8',
  SPREADSHEET_ID: '1ZDyvqYbyTEuYUjlUojtIYPUh3A2f9xQ1vFI4TrfOATw',
  PAGE_SIZE: 1000,
  JST_OFFSET_MS: 9 * 60 * 60 * 1000  // +9 hours in milliseconds
};

// ---------------------------------------------------------------------------
// Supabase helpers
// ---------------------------------------------------------------------------

/**
 * Fetch data from the Supabase REST API. Handles pagination automatically.
 * @param {string} endpoint - REST path including query params (e.g. '/rest/v1/staffs?select=id,firstname')
 * @return {Array<Object>} parsed JSON rows
 */
function fetchFromSupabase(endpoint) {
  var allRows = [];
  var offset = 0;
  var hasMore = true;

  while (hasMore) {
    var separator = endpoint.indexOf('?') === -1 ? '?' : '&';
    var url = CONFIG.SUPABASE_URL + endpoint + separator + 'limit=' + CONFIG.PAGE_SIZE + '&offset=' + offset;

    var options = {
      method: 'get',
      headers: {
        'apikey': CONFIG.SUPABASE_API_KEY,
        'Authorization': 'Bearer ' + CONFIG.SUPABASE_API_KEY,
        'Accept': 'application/json',
        'Prefer': 'count=exact'
      },
      muteHttpExceptions: true
    };

    Logger.log('Fetching: ' + url);
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code < 200 || code >= 300) {
      Logger.log('Supabase error (' + code + '): ' + response.getContentText());
      throw new Error('Supabase request failed with status ' + code + ': ' + response.getContentText());
    }

    var rows = JSON.parse(response.getContentText());
    if (!Array.isArray(rows)) {
      throw new Error('Unexpected Supabase response: ' + response.getContentText());
    }

    allRows = allRows.concat(rows);

    if (rows.length < CONFIG.PAGE_SIZE) {
      hasMore = false;
    } else {
      offset += CONFIG.PAGE_SIZE;
    }
  }

  Logger.log('Fetched ' + allRows.length + ' rows from ' + endpoint.split('?')[0]);
  return allRows;
}

// ---------------------------------------------------------------------------
// Staff mapping
// ---------------------------------------------------------------------------

/**
 * Fetch staff id -> firstname mapping from Supabase.
 * @return {Object} map of staff uuid to firstname
 */
function getStaffMap() {
  var rows = fetchFromSupabase('/rest/v1/staffs?select=id,firstname');
  var map = {};
  rows.forEach(function (row) {
    map[row.id] = row.firstname || '(不明)';
  });
  Logger.log('Staff map loaded: ' + Object.keys(map).length + ' entries');
  return map;
}

// ---------------------------------------------------------------------------
// Date helpers
// ---------------------------------------------------------------------------

/**
 * Return a JST date string (yyyy-MM-dd) for "yesterday" relative to now.
 * @return {string}
 */
function getYesterdayJST() {
  var now = new Date();
  var jstNow = new Date(now.getTime() + CONFIG.JST_OFFSET_MS);
  var yesterday = new Date(jstNow.getTime() - 24 * 60 * 60 * 1000);
  return Utilities.formatDate(yesterday, 'UTC', 'yyyy-MM-dd');
}

/**
 * Convert a UTC ISO timestamp string to a JST Date object.
 * @param {string} utcStr
 * @return {Date}
 */
function utcToJST(utcStr) {
  var d = new Date(utcStr);
  return new Date(d.getTime() + CONFIG.JST_OFFSET_MS);
}

/**
 * Format a Date object as JST display string.
 * @param {Date} d
 * @return {string}
 */
function formatJST(d) {
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd HH:mm:ss');
}

// ---------------------------------------------------------------------------
// Data retrieval
// ---------------------------------------------------------------------------

/**
 * Fetch all status-change journal_details whose journal.created_at falls
 * within the given JST date (converted to UTC range).
 *
 * @param {string} dateStr - JST date in yyyy-MM-dd format
 * @return {Array<Object>} journal_detail rows with embedded journal
 */
function getTargetTickets(dateStr) {
  // JST day boundaries → UTC
  var startUTC = dateStr + 'T00:00:00+09:00';
  var endDate = new Date(new Date(dateStr + 'T00:00:00+09:00').getTime() + 24 * 60 * 60 * 1000);
  var endUTC = Utilities.formatDate(endDate, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var startUTCFormatted = Utilities.formatDate(new Date(startUTC), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");

  Logger.log('Date range (UTC): ' + startUTCFormatted + ' to ' + endUTC);

  // Fetch journal_details for status_code changes within the date range
  var endpoint = '/rest/v1/journal_details?field_name=eq.status_code' +
    '&select=journal_id,old_value,new_value,journals!inner(id,created_at,issue_id,staff_id)' +
    '&journals.created_at=gte.' + startUTCFormatted +
    '&journals.created_at=lt.' + endUTC;

  var rows = fetchFromSupabase(endpoint);
  Logger.log('Total status changes for ' + dateStr + ': ' + rows.length);
  return rows;
}

/**
 * Fetch ALL status-change history for a specific issue (ticket).
 * Used to find the 申込決済完了 transition timestamps.
 *
 * @param {number} issueId
 * @return {Array<Object>} journal_detail rows with embedded journal, ordered by created_at
 */
function getIssueHistory(issueId) {
  var endpoint = '/rest/v1/journal_details?field_name=eq.status_code' +
    '&select=journal_id,old_value,new_value,journals!inner(id,created_at,issue_id,staff_id)' +
    '&journals.issue_id=eq.' + issueId +
    '&order=journals(created_at).asc';

  return fetchFromSupabase(endpoint);
}

// ---------------------------------------------------------------------------
// Processing time calculation
// ---------------------------------------------------------------------------

/**
 * Calculate processing times for target tickets.
 *
 * Target: tickets whose status changed to 「手動申請結果待ち」 on the given day.
 * (Any ticket that has ever been set to 手動申請結果待ち, regardless of current status.)
 *
 * Processing time (正味作業時間): time between the LAST transition TO 「申込決済完了」
 * and the transition FROM 「申込決済完了」 TO 「手動申請結果待ち」.
 * This excludes wait time during intermediate statuses (情報修正, 申請停止中, etc.).
 *
 * @param {Array<Object>} dayChanges - all status changes for the target day
 * @param {Object} staffMap - staff id → name
 * @return {Array<Object>} per-ticket result objects
 */
function calculateProcessingTimes(dayChanges, staffMap) {
  // Step 1: Find target tickets — those with a status change to
  //         手動申請結果待ち on the given day.
  var targetIssues = {};  // issueId → { staffId, ... }

  dayChanges.forEach(function (row) {
    if (row.new_value === '手動申請結果待ち') {
      var issueId = row.journals.issue_id;
      var staffId = row.journals.staff_id;
      if (!targetIssues[issueId]) {
        targetIssues[issueId] = {
          issueId: issueId,
          staffId: staffId,
          staffName: staffMap[staffId] || '(不明)',
          targetStatus: row.new_value
        };
      }
    }
  });

  var issueIds = Object.keys(targetIssues);
  Logger.log('Target tickets: ' + issueIds.length);

  if (issueIds.length === 0) {
    return [];
  }

  // Step 2: For each target ticket, fetch full history and compute processing time.
  var results = [];

  issueIds.forEach(function (issueId) {
    var info = targetIssues[issueId];
    var history = getIssueHistory(issueId);

    // Sort by journal created_at ascending
    history.sort(function (a, b) {
      return new Date(a.journals.created_at) - new Date(b.journals.created_at);
    });

    // Find the transition FROM 申込決済完了 TO 手動申請結果待ち
    // and the LAST transition TO 申込決済完了 that occurred before it.
    // This handles cases where intermediate statuses (情報修正, 申請停止中)
    // occur between the first 申込決済完了 and 手動申請結果待ち.
    var paymentEntry = null;      // last transition TO 申込決済完了 before application
    var applicationEntry = null;  // transition FROM 申込決済完了 TO 手動申請結果待ち
    var isAnomaly = false;
    var hadIntermediateStatus = false;

    // First pass: find the transition TO 手動申請結果待ち from 申込決済完了
    for (var i = 0; i < history.length; i++) {
      var h = history[i];
      if (h.old_value === '申込決済完了' && h.new_value === '手動申請結果待ち') {
        applicationEntry = h;
        break;
      }
    }

    if (applicationEntry) {
      // Find the LAST transition TO 申込決済完了 that occurred before the application entry
      var appTime = new Date(applicationEntry.journals.created_at).getTime();
      for (var i = 0; i < history.length; i++) {
        var h = history[i];
        if (h.new_value === '申込決済完了' && new Date(h.journals.created_at).getTime() < appTime) {
          paymentEntry = h;  // keep updating — we want the LAST one before application
        }
      }

      // Check if there were intermediate statuses between first 申込決済完了 and 手動申請結果待ち
      var firstPayment = null;
      for (var i = 0; i < history.length; i++) {
        if (history[i].new_value === '申込決済完了') {
          firstPayment = history[i];
          break;
        }
      }
      if (firstPayment && paymentEntry && firstPayment.journal_id !== paymentEntry.journal_id) {
        hadIntermediateStatus = true;
      }
    } else {
      // No direct 申込決済完了 → 手動申請結果待ち found
      // Check if there's any transition FROM 申込決済完了 (anomaly)
      for (var i = 0; i < history.length; i++) {
        var h = history[i];
        if (h.new_value === '申込決済完了') {
          paymentEntry = h;
        }
        if (h.old_value === '申込決済完了' && h.new_value !== '手動申請結果待ち') {
          isAnomaly = true;
          applicationEntry = h;
          break;
        }
      }
    }

    var result = {
      issueId: issueId,
      staffId: info.staffId,
      staffName: info.staffName,
      paymentTime: null,
      applicationTime: null,
      processingMinutes: null,
      isAnomaly: isAnomaly,
      hadIntermediateStatus: hadIntermediateStatus,
      hasWarning: false,
      warningReasons: []
    };

    if (paymentEntry) {
      result.paymentTime = paymentEntry.journals.created_at;
    }
    if (applicationEntry) {
      result.applicationTime = applicationEntry.journals.created_at;
    }

    // Calculate processing time (only for non-anomaly tickets)
    if (!isAnomaly && paymentEntry && applicationEntry) {
      var paymentDate = new Date(paymentEntry.journals.created_at);
      var applicationDate = new Date(applicationEntry.journals.created_at);
      var diffMs = applicationDate.getTime() - paymentDate.getTime();
      var diffMinutes = Math.round(diffMs / 60000 * 100) / 100;  // 2 decimal places
      result.processingMinutes = diffMinutes;

      // Warning checks
      if (diffMinutes < 0) {
        result.hasWarning = true;
        result.warningReasons.push('処理時間がマイナス');
      }
      if (diffMinutes > 720) {
        result.hasWarning = true;
        result.warningReasons.push('処理時間が12時間超過');
      }
    }

    // History order anomaly check: payment should come before application
    if (paymentEntry && applicationEntry) {
      var pIdx = -1, aIdx = -1;
      for (var j = 0; j < history.length; j++) {
        if (history[j].journal_id === paymentEntry.journal_id) pIdx = j;
        if (history[j].journal_id === applicationEntry.journal_id) aIdx = j;
      }
      if (pIdx >= 0 && aIdx >= 0 && pIdx > aIdx) {
        result.hasWarning = true;
        result.warningReasons.push('履歴の順序異常');
      }
    }

    // Missing data warnings
    if (!paymentEntry) {
      result.hasWarning = true;
      result.warningReasons.push('申込決済完了への遷移が見つからない');
    }
    if (!applicationEntry) {
      result.hasWarning = true;
      result.warningReasons.push('申込決済完了からの遷移が見つからない');
    }

    results.push(result);
  });

  Logger.log('Processed ' + results.length + ' tickets (' +
    results.filter(function (r) { return r.isAnomaly; }).length + ' anomalies, ' +
    results.filter(function (r) { return r.hasWarning; }).length + ' warnings)');

  return results;
}

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

/**
 * Aggregate ticket results by staff.
 *
 * @param {Array<Object>} ticketResults
 * @return {Array<Object>} per-staff summary objects, plus an overall "全体" entry
 */
function aggregateByStaff(ticketResults) {
  if (ticketResults.length === 0) {
    return [];
  }

  var staffGroups = {};

  ticketResults.forEach(function (t) {
    var name = t.staffName;
    if (!staffGroups[name]) {
      staffGroups[name] = {
        staffName: name,
        count: 0,
        times: [],
        anomalyCount: 0,
        warningCount: 0
      };
    }
    var g = staffGroups[name];
    g.count++;
    if (t.isAnomaly) g.anomalyCount++;
    if (t.hasWarning) g.warningCount++;
    if (t.processingMinutes !== null && !t.isAnomaly) {
      g.times.push(t.processingMinutes);
    }
  });

  var summaries = [];
  var allTimes = [];

  Object.keys(staffGroups).sort().forEach(function (name) {
    var g = staffGroups[name];
    var fastest = g.times.length > 0 ? Math.min.apply(null, g.times) : null;
    var slowest = g.times.length > 0 ? Math.max.apply(null, g.times) : null;
    var average = g.times.length > 0
      ? Math.round(g.times.reduce(function (a, b) { return a + b; }, 0) / g.times.length * 100) / 100
      : null;

    allTimes = allTimes.concat(g.times);

    summaries.push({
      staffName: name,
      count: g.count,
      fastest: fastest,
      slowest: slowest,
      average: average,
      anomalyCount: g.anomalyCount,
      warningCount: g.warningCount
    });
  });

  // Overall summary row
  var totalCount = ticketResults.length;
  var totalAnomaly = ticketResults.filter(function (t) { return t.isAnomaly; }).length;
  var totalWarning = ticketResults.filter(function (t) { return t.hasWarning; }).length;
  var totalFastest = allTimes.length > 0 ? Math.min.apply(null, allTimes) : null;
  var totalSlowest = allTimes.length > 0 ? Math.max.apply(null, allTimes) : null;
  var totalAverage = allTimes.length > 0
    ? Math.round(allTimes.reduce(function (a, b) { return a + b; }, 0) / allTimes.length * 100) / 100
    : null;

  summaries.push({
    staffName: '全体',
    count: totalCount,
    fastest: totalFastest,
    slowest: totalSlowest,
    average: totalAverage,
    anomalyCount: totalAnomaly,
    warningCount: totalWarning
  });

  return summaries;
}

// ---------------------------------------------------------------------------
// Spreadsheet output
// ---------------------------------------------------------------------------

/**
 * Write aggregated results and detail rows to the spreadsheet.
 *
 * @param {string} dateStr - JST date string (yyyy-MM-dd)
 * @param {Array<Object>} aggregated - per-staff summaries from aggregateByStaff
 * @param {Array<Object>} details - per-ticket results from calculateProcessingTimes
 */
function writeToSpreadsheet(dateStr, aggregated, details) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // --- Sheet 1: 日報 ---
  var dailySheet = ss.getSheetByName('日報');
  if (!dailySheet) {
    dailySheet = ss.insertSheet('日報');
    dailySheet.appendRow(['日付', '担当者', '件数', '最速/分', '最遅/分', '平均/分', '異例', '警告']);
    Logger.log('Created sheet: 日報');
  }

  aggregated.forEach(function (row) {
    dailySheet.appendRow([
      dateStr,
      row.staffName,
      row.count,
      row.fastest !== null ? row.fastest : '',
      row.slowest !== null ? row.slowest : '',
      row.average !== null ? row.average : '',
      row.anomalyCount,
      row.warningCount
    ]);
  });

  Logger.log('Wrote ' + aggregated.length + ' rows to 日報');

  // --- Sheet 2: 詳細 ---
  var detailSheet = ss.getSheetByName('詳細');
  if (!detailSheet) {
    detailSheet = ss.insertSheet('詳細');
    detailSheet.appendRow([
      '日付', 'チケットID', '担当者', '決済時刻/JST', '申請時刻/JST',
      '処理時間/分', '中断あり', '異例', '警告', '警告理由'
    ]);
    Logger.log('Created sheet: 詳細');
  }

  details.forEach(function (t) {
    var paymentJST = t.paymentTime ? formatJST(utcToJST(t.paymentTime)) : '';
    var applicationJST = t.applicationTime ? formatJST(utcToJST(t.applicationTime)) : '';

    detailSheet.appendRow([
      dateStr,
      t.issueId,
      t.staffName,
      paymentJST,
      applicationJST,
      t.processingMinutes !== null ? t.processingMinutes : '',
      t.hadIntermediateStatus ? '○' : '',
      t.isAnomaly ? '○' : '',
      t.hasWarning ? '○' : '',
      t.warningReasons.join(', ')
    ]);
  });

  Logger.log('Wrote ' + details.length + ' rows to 詳細');
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/**
 * Main entry point — designed to be called by a daily time-driven trigger.
 * Processes the previous JST day's data.
 */
function runDailyReport() {
  var dateStr = getYesterdayJST();
  Logger.log('=== Daily Report for ' + dateStr + ' ===');
  _processDate(dateStr);
}

/**
 * Manual entry point for testing or backfilling.
 * @param {string} dateStr - JST date in yyyy-MM-dd format (e.g. '2026-04-12')
 */
function runManualReport(dateStr) {
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    throw new Error('Invalid date format. Use yyyy-MM-dd (e.g. "2026-04-12")');
  }
  Logger.log('=== Manual Report for ' + dateStr + ' ===');
  _processDate(dateStr);
}

/**
 * Internal: run the full pipeline for a given JST date.
 * @param {string} dateStr
 */
function _processDate(dateStr) {
  try {
    // 1. Load staff mapping
    var staffMap = getStaffMap();

    // 2. Get all status changes for the target day
    var dayChanges = getTargetTickets(dateStr);

    if (dayChanges.length === 0) {
      Logger.log('No status changes found for ' + dateStr + '. Nothing to write.');
      return;
    }

    // 3. Calculate processing times per ticket
    var ticketResults = calculateProcessingTimes(dayChanges, staffMap);

    if (ticketResults.length === 0) {
      Logger.log('No target tickets found for ' + dateStr + '. Nothing to write.');
      return;
    }

    // 4. Aggregate by staff
    var aggregated = aggregateByStaff(ticketResults);

    // 5. Write to spreadsheet
    writeToSpreadsheet(dateStr, aggregated, ticketResults);

    Logger.log('=== Report complete for ' + dateStr + ' ===');
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    throw e;
  }
}
