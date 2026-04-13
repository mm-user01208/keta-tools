/**
 * Stripe Decision Dashboard — Google Apps Script
 *
 * Fetches Stripe transaction data from Supabase and writes
 * dashboard views to a Google Spreadsheet.
 *
 * Sheets:
 *   - 日次サマリー: Daily summary (succeeded, matched, unmatched, held, 3DS breakdown)
 *   - 全チケット: All transactions with status, 3DS type, reconciliation
 *   - 未突合・保留: Only unmatched and held payments (alerts)
 *
 * Entry points:
 *   - updateDashboard()  — main, called by 30-min trigger
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

var CONFIG = {
  SUPABASE_URL: 'https://bmuvklukfntqmopblhox.supabase.co',
  SUPABASE_API_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJtdXZrbHVrZm50cW1vcGJsaG94Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTczMzI4NTgyNCwiZXhwIjoyMDQ4ODYxODI0fQ.5ZRquHJnInh58sLHABGIAgRR87MJRyoc6t16jv7H0Q8',
  SPREADSHEET_ID: '1-r7UdFKOLJ_2BZ3WVRYkzOuIj52nlTLF5cIqyejBqZM',
  JST_OFFSET_MS: 9 * 60 * 60 * 1000
};

// ---------------------------------------------------------------------------
// Supabase query helper (views)
// ---------------------------------------------------------------------------

function queryView(viewName, queryParams) {
  var url = CONFIG.SUPABASE_URL + '/rest/v1/' + viewName;
  if (queryParams) {
    url += '?' + queryParams;
  }
  var options = {
    method: 'get',
    headers: {
      'apikey': CONFIG.SUPABASE_API_KEY,
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_API_KEY,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code < 200 || code >= 300) {
    Logger.log('Query error (' + code + '): ' + response.getContentText());
    throw new Error('Supabase query ' + viewName + ' failed: ' + code);
  }

  return JSON.parse(response.getContentText());
}

// ---------------------------------------------------------------------------
// Date/time helpers
// ---------------------------------------------------------------------------

function toJST(isoStr) {
  if (!isoStr) return '';
  var d = new Date(isoStr);
  var jst = new Date(d.getTime() + CONFIG.JST_OFFSET_MS);
  return Utilities.formatDate(jst, 'UTC', 'yyyy-MM-dd HH:mm:ss');
}

function classify3DS(result) {
  if (!result || result === '' || result === 'authenticated') {
    return 'フリクションレス';
  }
  if (result === 'challenge') {
    return 'チャレンジ';
  }
  return 'その他 (' + result + ')';
}

function classifyStatus(status) {
  switch (status) {
    case 'succeeded': return '決済完了';
    case 'requires_payment_method': return '決済保留（支払方法待ち）';
    case 'requires_action': return '決済保留（認証待ち）';
    case 'requires_capture': return '決済保留（キャプチャ待ち）';
    case 'processing': return '処理中';
    case 'canceled': return 'キャンセル';
    default: return status || '不明';
  }
}

function classifyReconciliation(status) {
  switch (status) {
    case 'matched': return '突合済み';
    case 'unmatched': return '未突合';
    case 'pending': return '確認中';
    default: return status || '未処理';
  }
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function writeSheet(sheet, headers, rows) {
  sheet.clearContents();
  sheet.appendRow(headers);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // ヘッダー行の書式設定
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // 列幅自動調整
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

function updateDashboard() {
  Logger.log('=== Stripe Dashboard Update ===');

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // --- 日次サマリー ---
  Logger.log('Fetching daily summary...');
  var dailySummary = queryView('stripe_daily_summary', 'order=report_date.desc');

  var summarySheet = getOrCreateSheet(ss, 'Stripe日次サマリー');
  var summaryHeaders = [
    '日付', '決済完了', '突合済み', '未突合', '確認中', '保留',
    'フリクションレス', 'チャレンジ', 'その他3DS'
  ];
  var summaryRows = dailySummary.map(function (r) {
    return [
      r.report_date,
      r.total_succeeded || 0,
      r.matched || 0,
      r.unmatched || 0,
      r.pending || 0,
      r.held || 0,
      r.frictionless || 0,
      r.challenge || 0,
      r.three_ds_other || 0
    ];
  });
  writeSheet(summarySheet, summaryHeaders, summaryRows);
  Logger.log('Daily summary: ' + summaryRows.length + ' days');

  // --- 全チケット ---
  Logger.log('Fetching all transactions...');
  var allTx = queryView('stripe_dashboard', 'order=stripe_created_at.desc');

  var allSheet = getOrCreateSheet(ss, 'Stripe全チケット');
  var allHeaders = [
    '決済日時(JST)', 'PaymentIntent ID', '金額', '通貨',
    'ステータス', '突合状況', '3DSフロー',
    'カードブランド', 'カード末尾4桁', 'カード国',
    'メールアドレス', '突合日時(JST)', '同期日時(JST)'
  ];
  var allRows = allTx.map(function (t) {
    return [
      toJST(t.stripe_created_at),
      t.payment_intent_id,
      t.amount ? t.amount / 100 : 0,
      t.currency ? t.currency.toUpperCase() : '',
      classifyStatus(t.status),
      classifyReconciliation(t.reconciliation_status),
      classify3DS(t.three_d_secure_result),
      t.card_brand || '',
      t.card_last4 || '',
      t.card_country || '',
      t.customer_email || '',
      toJST(t.matched_at),
      toJST(t.synced_at)
    ];
  });
  writeSheet(allSheet, allHeaders, allRows);
  Logger.log('All transactions: ' + allRows.length + ' rows');

  // --- 未突合・保留 ---
  var alertTx = allTx.filter(function (t) {
    return t.reconciliation_status === 'unmatched' ||
           t.reconciliation_status === 'pending' ||
           ['requires_payment_method', 'requires_action', 'requires_capture', 'processing'].indexOf(t.status) !== -1;
  });

  var alertSheet = getOrCreateSheet(ss, 'Stripe未突合・保留');
  var alertHeaders = [
    '決済日時(JST)', 'PaymentIntent ID', '金額', '通貨',
    'ステータス', '突合状況', '3DSフロー',
    'カードブランド', 'カード末尾4桁', 'カード国'
  ];
  var alertRows = alertTx.map(function (t) {
    return [
      toJST(t.stripe_created_at),
      t.payment_intent_id,
      t.amount ? t.amount / 100 : 0,
      t.currency ? t.currency.toUpperCase() : '',
      classifyStatus(t.status),
      classifyReconciliation(t.reconciliation_status),
      classify3DS(t.three_d_secure_result),
      t.card_brand || '',
      t.card_last4 || '',
      t.card_country || ''
    ];
  });
  writeSheet(alertSheet, alertHeaders, alertRows);

  // 未突合・保留があれば背景色で警告
  if (alertRows.length > 0) {
    for (var i = 0; i < alertRows.length; i++) {
      var rowRange = alertSheet.getRange(i + 2, 1, 1, alertHeaders.length);
      var status = alertRows[i][4];
      if (status.indexOf('保留') !== -1) {
        rowRange.setBackground('#fff3cd'); // 黄色 — 保留
      } else if (alertRows[i][5] === '未突合') {
        rowRange.setBackground('#f8d7da'); // 赤 — 未突合
      }
    }
  }

  Logger.log('Alerts: ' + alertRows.length + ' rows');

  // 最終更新時刻を記録
  var now = new Date();
  var jstNow = new Date(now.getTime() + CONFIG.JST_OFFSET_MS);
  var updateStr = Utilities.formatDate(jstNow, 'UTC', 'yyyy-MM-dd HH:mm:ss') + ' JST';

  summarySheet.getRange(summaryRows.length + 3, 1).setValue('最終更新: ' + updateStr);
  allSheet.getRange(allRows.length + 3, 1).setValue('最終更新: ' + updateStr);
  alertSheet.getRange(alertRows.length + 3, 1).setValue('最終更新: ' + updateStr);

  Logger.log('=== Dashboard update complete at ' + updateStr + ' ===');
}
