/**
 * Google Store 注文メール → スプレッドシート追記
 * 設定値
 */
const ORDER_LABEL_NAME = 'ご注文完了';     // 注文完了メール用ラベル名
const ORDER_CANCELED_LABEL_NAME = '注文キャンセル'; // 注文キャンセルメール用ラベル名
const ORDER_ON_HOLD_LABEL_NAME = '注文保留'; // 注文保留メール用ラベル名
const DELIVERY_CHANGED_LABEL_NAME = 'お届け予定日変更'; // お届け予定日変更メール用ラベル名
const SHIPPED_LABEL_NAME = '発送完了';     // 発送完了メール用ラベル名
const PAYMENT_ERROR_LABEL_NAME = 'ペイメントエラー'; // ペイメントエラーメール用ラベル名
const PAYMENT_RECOVERY_LABEL_NAME = 'ペイメント復帰'; // ペイメント復帰メール用ラベル名
const PROCESSED_LABEL_NAME = '注文抽出済'; // 処理済みラベル名
const SHEET_NAME = 'Tracking';            // 書き込み先シート名
const ACCOUNT_SHEET_NAME = 'アカウント管理'; // アカウント管理シート名
const TRADE_IN_CONFIRM_LABEL_NAME = '下取り本人確認'; // 下取り本人確認メール用ラベル名
const TRADE_IN_CONFIRMED_LABEL_NAME = '下取り本人確認完了'; // 下取り本人確認完了メール用ラベル名
const TRADE_IN_SHIPPED_LABEL_NAME = '下取り発送'; // 下取り発送メール用ラベル名
const TRADE_IN_CANCEL_LABEL_NAME = '下取りキャンセル'; // 下取りキャンセルメール用ラベル名（同じGS注文番号の本人確認メールを削除するため）
const TRACKING2_SHEET_NAME = 'tracking2';   // 下取り本人確認用シート名

/** 1時間おき実行時：読み取るメールの範囲（何時間前まで）。1=直近1時間、2=直近2時間 */
const LOOKBACK_HOURS = 2;

/** 下取りキャンセル時の本人確認メール削除：何日前まで参照するか（4=直近4日） */
const TRADE_IN_CANCEL_DELETE_LOOKBACK_DAYS = 4;

/** 下取り発送メールの検索範囲（日）。直近2時間だと追跡番号更新が一度も走らないため長めに */
const TRADE_IN_SHIPPED_LOOKBACK_DAYS = 14;

/** テスト用：下取り関連の処理を上から何件に制限するか。3=上から3件のみ。0=制限なし */
const TRADE_IN_TEST_LIMIT = 3;

/**
 * Webアプリとしてデプロイして使用
 * - tracking=番号&tradeInId=xxx → tracking2 の D列・L列を更新
 */
function doGet(e) {
  var result = { ok: false, message: '' };
  try {
    var params = e && e.parameter ? e.parameter : {};
    var action = (params.action || '').toString().trim();
    var tradeInId = (params.tradeInId || '').toString().trim().replace(/-/g, '');

    if (action === 'getPending' && tradeInId) {
      var props = PropertiesService.getScriptProperties();
      var key = 'PENDING_' + tradeInId;
      var shortUrl = props.getProperty(key);
      if (shortUrl) props.deleteProperty(key);
      return ContentService.createTextOutput(JSON.stringify({ shortUrl: shortUrl || '' })).setMimeType(ContentService.MimeType.JSON);
    }

    var tracking = (params.tracking || '').toString().trim();
    if (!tracking || !tradeInId) {
      result.message = 'tracking と tradeInId を指定してください';
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      result.message = 'SPREADSHEET_ID が未設定です';
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(TRACKING2_SHEET_NAME);
    if (!sheet) {
      result.message = 'tracking2 シートがありません';
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    var row = findRowByOrderId_(sheet, tradeInId);
    if (!row) {
      result.message = '下取りIDに一致する行が見つかりません: ' + tradeInId;
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    sheet.getRange(row, 4).setValue(tracking);
    sheet.getRange(row, 12).setValue('発送');
    result.ok = true;
    result.message = '更新しました row=' + row;
  } catch (err) {
    result.message = String(err);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * メイン処理
 * 時間主導トリガーから呼び出すことを想定
 */
function processOrderEmails() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    Logger.log('SPREADSHEET_ID が ScriptProperties に設定されていません');
    return;
  }

  try {
    runProcessOrderEmails_(props, spreadsheetId);
  } catch (e) {
    var msg = (e && e.message) ? String(e.message) : String(e);
    if (msg.indexOf('Service invoked too many times') >= 0 || msg.indexOf('gmail') >= 0) {
      Logger.log('Gmailの1日あたりの呼び出し制限に達しました。明日までお待ちください。');
    }
    throw e;
  }
}

function runProcessOrderEmails_(props, spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // 注文番号(A列)は先頭ゼロを保持するため常にテキスト扱いにする
  sheet.getRange('A:A').setNumberFormat('@');

  const processedMessageIds = loadProcessedMessageIds_(props);
  const existingOrderIds = loadExistingOrderIds_(sheet);

  const orderResult = processNewOrderMails_(sheet, existingOrderIds, processedMessageIds);
  const canceledResult = processStatusMails_(sheet, processedMessageIds, ORDER_CANCELED_LABEL_NAME, 'キャンセル');
  const holdResult = processStatusMails_(sheet, processedMessageIds, ORDER_ON_HOLD_LABEL_NAME, '保留');
  const changedResult = processDeliveryChangedMails_(sheet, processedMessageIds);
  const shippedResult = processShippedMails_(sheet, processedMessageIds);
  const cancelResult = applyCancelByDeadline_(sheet);
  const accountSyncResult = syncAccountSheetFromShipped_(ss, sheet, props);
  const paymentErrorResult = processPaymentErrorMails_(ss, processedMessageIds);
  const paymentRecoveryResult = processPaymentRecoveryMails_(ss, processedMessageIds);
  const tradeInCancelDeleteResult = deleteTradeInConfirmMailsWhenCancelled_(ss);
  const tradeInResult = processTradeInConfirmMails_(ss, processedMessageIds);
  const tradeInConfirmedResult = processTradeInConfirmedMails_(ss, processedMessageIds);
  const tradeInShippedResult = processTradeInShippedMails_(ss, processedMessageIds);

  saveProcessedMessageIds_(props, processedMessageIds);
  Logger.log(
    '注文取込: 追加=' + orderResult.added +
    ', 重複注文=' + orderResult.skippedDuplicateOrder +
    ', 既処理messageId=' + orderResult.skippedAlreadyProcessed +
    ', 解析失敗=' + orderResult.parseFailed
  );
  Logger.log(
    '注文キャンセル更新: 更新=' + canceledResult.updated +
    ', 行未検出=' + canceledResult.noMatchedRow +
    ', 既処理messageId=' + canceledResult.skippedAlreadyProcessed +
    ', 解析失敗=' + canceledResult.parseFailed
  );
  Logger.log(
    '注文保留更新: 更新=' + holdResult.updated +
    ', 行未検出=' + holdResult.noMatchedRow +
    ', 既処理messageId=' + holdResult.skippedAlreadyProcessed +
    ', 解析失敗=' + holdResult.parseFailed
  );
  Logger.log(
    '予定日変更: 更新=' + changedResult.updated +
    ', 行未検出=' + changedResult.noMatchedRow +
    ', 既処理messageId=' + changedResult.skippedAlreadyProcessed +
    ', 解析失敗=' + changedResult.parseFailed
  );
  Logger.log(
    '発送更新: 更新=' + shippedResult.updated +
    ', 行未検出=' + shippedResult.noMatchedRow +
    ', 既処理messageId=' + shippedResult.skippedAlreadyProcessed +
    ', 解析失敗=' + shippedResult.parseFailed
  );
  Logger.log('キャンセル更新: ' + cancelResult.cancelled + '件');
  Logger.log('アカウント管理更新: 反映注文数=' + accountSyncResult.applied + ', 対象なし=' + accountSyncResult.skipped);
  Logger.log(
    'ペイメントエラー更新: 更新=' + paymentErrorResult.updated +
    ', 既処理messageId=' + paymentErrorResult.skippedAlreadyProcessed +
    ', 解析失敗=' + paymentErrorResult.parseFailed
  );
  Logger.log(
    'ペイメント復帰更新: 更新=' + paymentRecoveryResult.updated +
    ', 既処理messageId=' + paymentRecoveryResult.skippedAlreadyProcessed +
    ', 解析失敗=' + paymentRecoveryResult.parseFailed
  );
  Logger.log('下取りキャンセル対応: メール削除=' + tradeInCancelDeleteResult.deleted + '件, スプシ行削除=' + tradeInCancelDeleteResult.rowsDeleted + '件');
  Logger.log(
    '下取り本人確認(tracking2): 追加=' + tradeInResult.added +
    ', 重複ID=' + tradeInResult.skippedDuplicate +
    ', 既処理messageId=' + tradeInResult.skippedAlreadyProcessed +
    ', 解析失敗=' + tradeInResult.parseFailed
  );
  Logger.log(
    '下取り本人確認完了(tracking2): 更新=' + tradeInConfirmedResult.updated +
    ', 行未検出=' + tradeInConfirmedResult.noMatchedRow +
    ', 既処理messageId=' + tradeInConfirmedResult.skippedAlreadyProcessed +
    ', 解析失敗=' + tradeInConfirmedResult.parseFailed
  );
  Logger.log(
    '下取り発送(tracking2): 更新=' + tradeInShippedResult.updated +
    ', 行未検出=' + tradeInShippedResult.noMatchedRow +
    ', 既処理messageId=' + tradeInShippedResult.skippedAlreadyProcessed +
    ', 解析失敗=' + tradeInShippedResult.parseFailed
  );
}

/**
 * 調査用（GmailApp）：「ご注文完了」ラベルの全スレッドと「注文抽出済」の有無をログに出す
 * 実行後は 表示 > ログ で確認
 */
function debugListOrderLabelThreads() {
  var orderLabel = GmailApp.getUserLabelByName(ORDER_LABEL_NAME);
  if (!orderLabel) {
    Logger.log('ラベルが存在しません: ' + ORDER_LABEL_NAME);
    return;
  }
  var threads = orderLabel.getThreads(0, 100);
  Logger.log('=== 「' + ORDER_LABEL_NAME + '」ラベル スレッド一覧（取得数: ' + threads.length + '）===');
  var processedCount = 0;
  for (var i = 0; i < threads.length; i++) {
    var th = threads[i];
    var labels = th.getLabels();
    var hasProcessed = false;
    for (var l = 0; l < labels.length; l++) {
      if (labels[l].getName() === PROCESSED_LABEL_NAME) {
        hasProcessed = true;
        processedCount++;
        break;
      }
    }
    var msgs = th.getMessages();
    var firstSubject = msgs.length > 0 ? msgs[0].getSubject() : '(なし)';
    var firstDate = msgs.length > 0 ? msgs[0].getDate() : null;
    Logger.log(
      (i + 1) + '. 注文抽出済=' + (hasProcessed ? 'あり' : 'なし') +
      ' | メール数=' + msgs.length +
      ' | 件名=' + firstSubject.substring(0, 50) +
      (firstDate ? ' | 日時=' + firstDate : '')
    );
  }
  Logger.log('--- 集計: 合計=' + threads.length + ', 注文抽出済あり=' + processedCount + ', 未処理(次回対象)=' + (threads.length - processedCount) + ' ---');
}

/**
 * 調査用（Gmail API）：ラベル別のスレッド数をAPIで取得
 * 初回は「Gmail API」を有効化する必要あり（拡張サービス）
 * 実行後は 表示 > ログ で確認
 */
function debugGmailApiLabelCount() {
  try {
    var labelList = Gmail.Users.Labels.list('me').labels;
    var orderLabelId = null;
    var processedLabelId = null;
    for (var i = 0; i < labelList.length; i++) {
      var L = labelList[i];
      if (L.name === ORDER_LABEL_NAME) {
        orderLabelId = L.id;
      }
      if (L.name === PROCESSED_LABEL_NAME) {
        processedLabelId = L.id;
      }
    }
    if (!orderLabelId) {
      Logger.log('Gmail API: ラベル「' + ORDER_LABEL_NAME + '」が見つかりません');
      return;
    }
    Logger.log('=== Gmail API 調査 ===');
    Logger.log('ラベル「' + ORDER_LABEL_NAME + '」 ID=' + orderLabelId);
    if (processedLabelId) {
      Logger.log('ラベル「' + PROCESSED_LABEL_NAME + '」 ID=' + processedLabelId);
    }
    var total = 0;
    var pageToken = null;
    do {
      var opt = { labelIds: [orderLabelId], maxResults: 100 };
      if (pageToken) opt.pageToken = pageToken;
      var res = Gmail.Users.Threads.list('me', opt);
      var threads = res.threads || [];
      total += threads.length;
      for (var j = 0; j < threads.length; j++) {
        Logger.log('  スレッドID: ' + threads[j].id);
      }
      pageToken = res.nextPageToken || null;
    } while (pageToken);
    Logger.log('--- Gmail API で取得した「' + ORDER_LABEL_NAME + '」のスレッド総数: ' + total + ' ---');
  } catch (e) {
    Logger.log('Gmail API エラー（拡張サービスが有効か確認してください）: ' + e.toString());
  }
}

/**
 * 再テスト用：「ご注文完了」が付いているスレッドから「注文抽出済」ラベルを外す
 * スプレッドシートを空にして同じメールで再度テストしたいときに1回だけ実行
 */
function removeProcessedLabelForTest() {
  const orderLabel = GmailApp.getUserLabelByName(ORDER_LABEL_NAME);
  if (!orderLabel) {
    Logger.log('ラベルが存在しません: ' + ORDER_LABEL_NAME);
    return;
  }
  const processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL_NAME);
  if (!processedLabel) {
    Logger.log('ラベルが存在しません: ' + PROCESSED_LABEL_NAME);
    return;
  }
  const threads = orderLabel.getThreads(0, 100);
  var count = 0;
  for (var i = 0; i < threads.length; i++) {
    const th = threads[i];
    const labels = th.getLabels();
    for (var j = 0; j < labels.length; j++) {
      if (labels[j].getName() === PROCESSED_LABEL_NAME) {
        processedLabel.removeFromThread(th);
        count++;
        break;
      }
    }
  }
  Logger.log('「注文抽出済」を外したスレッド数: ' + count);
}

/**
 * 時間主導トリガーを 5分おきで作成
 * 最初に 1回だけ手動実行してください
 */
function createTimeTrigger() {
  ScriptApp.newTrigger('processOrderEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function processNewOrderMails_(sheet, existingOrderIds, processedMessageIds) {
  const orderLabel = GmailApp.getUserLabelByName(ORDER_LABEL_NAME);
  if (!orderLabel) {
    Logger.log('注文用ラベルが存在しません: ' + ORDER_LABEL_NAME);
    return { added: 0, skippedDuplicateOrder: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  const threads = getThreadsByLabelNewerThan_(orderLabel, LOOKBACK_HOURS);
  const rowsToAppend = [];
  var skippedAlreadyProcessed = 0;
  var skippedDuplicateOrder = 0;
  var parseFailed = 0;

  for (var i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      const message = messages[j];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        const parsed = parseOrderMail_(message);
        if (!parsed || !parsed.orderId) {
          parseFailed++;
          continue;
        }

        if (existingOrderIds.has(parsed.orderId)) {
          skippedDuplicateOrder++;
          processedMessageIds.add(message.getId());
          continue;
        }

        const row = [
          stringifyOrderId_(parsed.orderId),  // A: 注文番号（数字のみ・先頭ゼロ保持）
          parsed.accountEmail || '',          // B: アカウント名（メールアドレス）
          parsed.orderDate || '',             // C: 注文日
          parsed.paymentMethod || '',         // D: 決済方法
          parsed.productName || '',           // E: 商品名
          parsed.idNumber || '',              // F: ID番号
          parsed.subtotal || 0,               // G: 小計
          parsed.tax || 0,                    // H: 消費税
          parsed.usedStorePoints || 0,        // I: 使用ストアP
          parsed.couponDiscount || 0,         // J: クーポン割引
          parsed.paidAmount || 0,             // K: 支払金額
          parsed.bonusStorePoints || 0,       // L: 特典ストアP
          parsed.trackingNumber || '',        // M: 追跡番号
          parsed.deliverySchedule || '',      // N: お届け予定日時
          '未発送'                            // O: ステータス
        ];

        rowsToAppend.push(row);
        existingOrderIds.add(parsed.orderId);
        processedMessageIds.add(message.getId());
      } catch (e) {
        Logger.log('注文メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }
  return {
    added: rowsToAppend.length,
    skippedDuplicateOrder: skippedDuplicateOrder,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

function processStatusMails_(sheet, processedMessageIds, labelName, targetStatus) {
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    Logger.log('ラベルが存在しません: ' + labelName);
    return { updated: 0, noMatchedRow: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  const threads = getThreadsByLabelNewerThan_(label, LOOKBACK_HOURS);
  var updated = 0;
  var noMatchedRow = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;

  for (var i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      const message = messages[j];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        const parsed = parseOrderIdFromAnyMail_(message);
        if (!parsed || !parsed.orderId) {
          parseFailed++;
          continue;
        }

        const row = findRowByOrderId_(sheet, parsed.orderId);
        if (!row) {
          noMatchedRow++;
          continue;
        }

        sheet.getRange(row, 15).setValue(targetStatus); // O列
        processedMessageIds.add(message.getId());
        updated++;
      } catch (e) {
        Logger.log(labelName + ' メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  return {
    updated: updated,
    noMatchedRow: noMatchedRow,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

function processDeliveryChangedMails_(sheet, processedMessageIds) {
  const changedLabel = GmailApp.getUserLabelByName(DELIVERY_CHANGED_LABEL_NAME);
  if (!changedLabel) {
    Logger.log('予定日変更ラベルが存在しません: ' + DELIVERY_CHANGED_LABEL_NAME);
    return { updated: 0, noMatchedRow: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  const threads = getThreadsByLabelNewerThan_(changedLabel, LOOKBACK_HOURS);
  var updated = 0;
  var noMatchedRow = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;

  for (var i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      const message = messages[j];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        const changed = parseDeliveryChangedMail_(message);
        if (!changed || !changed.orderId) {
          parseFailed++;
          continue;
        }

        const row = findRowByOrderId_(sheet, changed.orderId);
        if (!row) {
          // 注文行が未作成の可能性があるので、次回再試行できるよう未処理のままにする
          noMatchedRow++;
          continue;
        }

        const current = sheet.getRange(row, 1, 1, 15).getValues()[0];
        const oldSchedule = String(current[13] || '').trim(); // N
        const newSchedule = String(changed.deliverySchedule || '').trim();
        if (newSchedule && newSchedule !== oldSchedule) {
          current[13] = newSchedule;
          current[14] = '変更'; // O
          sheet.getRange(row, 1, 1, 15).setValues([current]);
          sheet.getRange(row, 14).setBackground('#f4cccc'); // N列を赤
          updated++;
        } else {
          // 変更がない場合は更新せず処理済みにする
          if (newSchedule) {
            current[13] = newSchedule;
            sheet.getRange(row, 14).setValue(newSchedule);
          }
        }
        processedMessageIds.add(message.getId());
      } catch (e) {
        Logger.log('予定日変更メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  return {
    updated: updated,
    noMatchedRow: noMatchedRow,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

function processShippedMails_(sheet, processedMessageIds) {
  const shippedLabel = GmailApp.getUserLabelByName(SHIPPED_LABEL_NAME);
  if (!shippedLabel) {
    Logger.log('発送ラベルが存在しません: ' + SHIPPED_LABEL_NAME);
    return { updated: 0, noMatchedRow: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  const threads = getThreadsByLabelNewerThan_(shippedLabel, LOOKBACK_HOURS);
  var updated = 0;
  var noMatchedRow = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;

  for (var i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      const message = messages[j];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        const shipped = parseShippedMail_(message);
        if (!shipped || !shipped.orderId) {
          parseFailed++;
          continue;
        }

        const row = findRowByOrderId_(sheet, shipped.orderId);
        if (!row) {
          // 注文メールより先に発送メールが来るケースを考慮して未処理のままにする
          noMatchedRow++;
          continue;
        }

        const current = sheet.getRange(row, 1, 1, 15).getValues()[0];
        if (shipped.idNumber) current[5] = shipped.idNumber;           // F
        if (shipped.trackingNumber) current[12] = shipped.trackingNumber; // M
        if (shipped.deliverySchedule) current[13] = shipped.deliverySchedule; // N
        current[14] = '発送完了'; // O

        sheet.getRange(row, 1, 1, 15).setValues([current]);
        processedMessageIds.add(message.getId());
        updated++;
      } catch (e) {
        Logger.log('発送メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  return {
    updated: updated,
    noMatchedRow: noMatchedRow,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

function applyCancelByDeadline_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return { cancelled: 0 };

  const values = sheet.getRange(1, 1, lastRow, 15).getValues();
  const now = new Date();
  var cancelled = 0;

  for (var i = 0; i < values.length; i++) {
    const row = values[i];
    const orderId = String(row[0] || '').trim(); // A
    const orderDate = String(row[2] || '').trim(); // C
    const deliverySchedule = String(row[13] || '').trim(); // N
    const status = String(row[14] || '').trim(); // O

    if (!orderId) continue;
    if (status === '発送完了' || status === 'キャンセル') continue;
    if (!deliverySchedule) continue;

    const deadline = getCancelDeadline_(deliverySchedule, orderDate);
    if (!deadline) continue;

    if (now.getTime() > deadline.getTime()) {
      row[14] = 'キャンセル';
      cancelled++;
    }
  }

  if (cancelled > 0) {
    sheet.getRange(1, 1, values.length, 15).setValues(values);
  }
  return { cancelled: cancelled };
}

function syncAccountSheetFromShipped_(ss, trackingSheet, props) {
  let accountSheet = ss.getSheetByName(ACCOUNT_SHEET_NAME);
  if (!accountSheet) {
    accountSheet = ss.insertSheet(ACCOUNT_SHEET_NAME);
  }
  accountSheet.getRange('A:A').setNumberFormat('@');

  const syncedOrderIds = loadSyncedAccountOrderIds_(props);

  const trackingLastRow = trackingSheet.getLastRow();
  if (trackingLastRow < 1) {
    return { applied: 0, skipped: 0 };
  }
  const trackingValues = trackingSheet.getRange(1, 1, trackingLastRow, 15).getValues();

  const accountLastRow = accountSheet.getLastRow();
  const accountRows = accountLastRow > 0 ? accountSheet.getRange(1, 1, accountLastRow, 2).getValues() : [];
  const accountRowIndexMap = {};
  for (var i = 0; i < accountRows.length; i++) {
    const account = String(accountRows[i][0] || '').trim();
    if (account) accountRowIndexMap[account] = i;
  }

  var applied = 0;
  var skipped = 0;

  for (var r = 0; r < trackingValues.length; r++) {
    const row = trackingValues[r];
    const orderId = String(row[0] || '').trim(); // A
    const accountName = String(row[1] || '').trim(); // B
    const usedStorePoint = Number(row[8] || 0); // I
    const status = String(row[14] || '').trim(); // O

    if (!orderId || status !== '発送完了') {
      continue;
    }
    if (syncedOrderIds.has(orderId)) {
      skipped++;
      continue;
    }
    if (!accountName) {
      skipped++;
      continue;
    }

    let idx = accountRowIndexMap[accountName];
    if (idx === undefined) {
      idx = accountRows.length;
      accountRowIndexMap[accountName] = idx;
      accountRows.push([accountName, 0]);
    }
    const current = Number(accountRows[idx][1] || 0);
    accountRows[idx][1] = current + usedStorePoint;

    syncedOrderIds.add(orderId);
    applied++;
  }

  if (accountRows.length > 0) {
    accountSheet.getRange(1, 1, accountRows.length, 2).setValues(accountRows);
  }
  saveSyncedAccountOrderIds_(props, syncedOrderIds);
  return { applied: applied, skipped: skipped };
}

function processPaymentErrorMails_(ss, processedMessageIds) {
  const errorLabel = GmailApp.getUserLabelByName(PAYMENT_ERROR_LABEL_NAME);
  if (!errorLabel) {
    Logger.log('ペイメントエラーラベルが存在しません: ' + PAYMENT_ERROR_LABEL_NAME);
    return { updated: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  let accountSheet = ss.getSheetByName(ACCOUNT_SHEET_NAME);
  if (!accountSheet) {
    accountSheet = ss.insertSheet(ACCOUNT_SHEET_NAME);
  }
  accountSheet.getRange('A:A').setNumberFormat('@');

  const threads = getThreadsByLabelNewerThan_(errorLabel, LOOKBACK_HOURS);
  const lastRow = accountSheet.getLastRow();
  const accountRows = lastRow > 0 ? accountSheet.getRange(1, 1, lastRow, 4).getValues() : [];
  const rowIndexByAccount = {};
  for (var i = 0; i < accountRows.length; i++) {
    const account = String(accountRows[i][0] || '').trim();
    if (account) rowIndexByAccount[account] = i;
  }

  var updated = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;
  var dirty = false;

  for (var t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      const message = messages[m];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        const account = String(extractAccountEmail_(message) || '').trim();
        if (!account) {
          parseFailed++;
          continue;
        }

        let idx = rowIndexByAccount[account];
        if (idx === undefined) {
          idx = accountRows.length;
          rowIndexByAccount[account] = idx;
          accountRows.push([account, 0, '', 'エラー']);
          updated++;
          dirty = true;
        } else if (String(accountRows[idx][3] || '') !== 'エラー') {
          accountRows[idx][3] = 'エラー'; // D列
          updated++;
          dirty = true;
        }

        processedMessageIds.add(message.getId());
      } catch (e) {
        Logger.log('ペイメントエラーメール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  if (dirty && accountRows.length > 0) {
    accountSheet.getRange(1, 1, accountRows.length, 4).setValues(accountRows);
  }
  return {
    updated: updated,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

/**
 * ペイメント復帰ラベル: To のアカウントに対応するアカウント管理の D列を「エラー」→「-」に変更
 */
function processPaymentRecoveryMails_(ss, processedMessageIds) {
  const recoveryLabel = GmailApp.getUserLabelByName(PAYMENT_RECOVERY_LABEL_NAME);
  if (!recoveryLabel) {
    Logger.log('ペイメント復帰ラベルが存在しません: ' + PAYMENT_RECOVERY_LABEL_NAME);
    return { updated: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  var accountSheet = ss.getSheetByName(ACCOUNT_SHEET_NAME);
  if (!accountSheet) {
    return { updated: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  const threads = getThreadsByLabelNewerThan_(recoveryLabel, LOOKBACK_HOURS);
  var updated = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;
  var dirty = false;
  const lastRow = accountSheet.getLastRow();
  const accountRows = lastRow > 0 ? accountSheet.getRange(1, 1, lastRow, 4).getValues() : [];
  const rowIndexByAccount = {};
  for (var i = 0; i < accountRows.length; i++) {
    var account = String(accountRows[i][0] || '').trim();
    if (account) rowIndexByAccount[account] = i;
  }

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        var account = String(extractAccountEmail_(message) || '').trim();
        if (!account) {
          parseFailed++;
          continue;
        }

        var idx = rowIndexByAccount[account];
        if (idx === undefined) {
          parseFailed++;
          continue;
        }
        if (String(accountRows[idx][3] || '') === 'エラー') {
          accountRows[idx][3] = '-';
          updated++;
          dirty = true;
        }
        processedMessageIds.add(message.getId());
      } catch (e) {
        Logger.log('ペイメント復帰メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  if (dirty && accountRows.length > 0) {
    accountSheet.getRange(1, 1, accountRows.length, 4).setValues(accountRows);
  }
  return {
    updated: updated,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

/**
 * 流れ: 本人確認メール検索 → 下取りキャンセルメール検索。注文番号が一致したら tracking2 の行と本人確認メールを削除
 */
function deleteTradeInConfirmMailsWhenCancelled_(ss) {
  var confirmLabel = GmailApp.getUserLabelByName(TRADE_IN_CONFIRM_LABEL_NAME);
  var cancelLabel = GmailApp.getUserLabelByName(TRADE_IN_CANCEL_LABEL_NAME);
  if (!confirmLabel || !cancelLabel) return { deleted: 0, rowsDeleted: 0 };

  // 1. 下取りキャンセルラベル：直近N日分から注文番号(GS)を収集
  var cancelOrderIds = new Set();
  var cancelThreads = getThreadsByLabelNewerThanDays_(cancelLabel, TRADE_IN_CANCEL_DELETE_LOOKBACK_DAYS);
  for (var t = 0; t < cancelThreads.length; t++) {
    var messages = cancelThreads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var parsed = parseOrderIdFromAnyMail_(messages[m]);
      if (parsed && parsed.orderId) cancelOrderIds.add(parsed.orderId);
    }
  }
  if (cancelOrderIds.size === 0) return { deleted: 0, rowsDeleted: 0 };

  // 2. 本人確認ラベル：直近N日分を検索し、注文番号がキャンセル側と一致するメールとスプシ行を削除（テスト時はメール3件まで）
  var sheet = ss ? ss.getSheetByName(TRACKING2_SHEET_NAME) : null;
  var toDelete = []; // { row: number, message: GmailMessage }
  var confirmThreads = getThreadsByLabelNewerThanDays_(confirmLabel, TRADE_IN_CANCEL_DELETE_LOOKBACK_DAYS);
  var deleteLimit = TRADE_IN_TEST_LIMIT > 0 ? TRADE_IN_TEST_LIMIT : 999999;
  for (var t = 0; t < confirmThreads.length && toDelete.length < deleteLimit; t++) {
    var messages = confirmThreads[t].getMessages();
    for (var m = 0; m < messages.length && toDelete.length < deleteLimit; m++) {
      var message = messages[m];
      try {
        var orderParsed = parseOrderIdFromAnyMail_(message);
        if (!orderParsed || !orderParsed.orderId || !cancelOrderIds.has(orderParsed.orderId)) continue;
        var confirmParsed = parseTradeInConfirmMail_(message);
        var tradeInId = confirmParsed && confirmParsed.tradeInId ? confirmParsed.tradeInId : '';
        var row = (sheet && tradeInId) ? findRowByOrderId_(sheet, tradeInId) : 0;
        toDelete.push({ row: row, message: message });
      } catch (e) {
        Logger.log('下取り本人確認マッチングエラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  // 下の行から削除してインデックスずれを防ぐ
  toDelete.sort(function (a, b) { return (b.row || 0) - (a.row || 0); });
  var rowsDeleted = 0;
  for (var i = 0; i < toDelete.length; i++) {
    var item = toDelete[i];
    if (sheet && item.row > 0) {
      try {
        sheet.deleteRow(item.row);
        rowsDeleted++;
      } catch (e) {
        Logger.log('tracking2行削除エラー row=' + item.row + ': ' + e);
      }
    }
    try {
      item.message.moveToTrash();
    } catch (e) {
      Logger.log('本人確認メール削除エラー: ' + e);
    }
  }
  return { deleted: toDelete.length, rowsDeleted: rowsDeleted };
}

/**
 * 下取り本人確認ラベル: tracking2 に A=下取りID(ハイフンなし), B=To, C=開始日, E=デバイス名, F=支払い方法, G=見積もり額 を追記
 */
function processTradeInConfirmMails_(ss, processedMessageIds) {
  const label = GmailApp.getUserLabelByName(TRADE_IN_CONFIRM_LABEL_NAME);
  if (!label) {
    Logger.log('下取り本人確認ラベルが存在しません: ' + TRADE_IN_CONFIRM_LABEL_NAME);
    return { added: 0, skippedDuplicate: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  let sheet = ss.getSheetByName(TRACKING2_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TRACKING2_SHEET_NAME);
  }
  sheet.getRange('A:A').setNumberFormat('@');

  const existingIds = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 1) {
    const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var id = String(colA[i][0] || '').trim();
      if (id) existingIds.add(id);
    }
  }

  var threads = getThreadsByLabelNewerThan_(label, LOOKBACK_HOURS);
  const rowsToAppend = [];
  var skippedDuplicate = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;
  var addLimit = TRADE_IN_TEST_LIMIT > 0 ? TRADE_IN_TEST_LIMIT : 999999;

  for (var t = 0; t < threads.length && rowsToAppend.length < addLimit; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length && rowsToAppend.length < addLimit; m++) {
      var message = messages[m];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        var parsed = parseTradeInConfirmMail_(message);
        if (!parsed || !parsed.tradeInId) {
          parseFailed++;
          continue;
        }

        if (existingIds.has(parsed.tradeInId)) {
          skippedDuplicate++;
          processedMessageIds.add(message.getId());
          continue;
        }

        rowsToAppend.push([
          parsed.tradeInId,                    // A: 下取りID（ハイフンなし）
          parsed.accountEmail || '',          // B: To のアカウント名
          parsed.startDate || '',             // C: 開始日（2026年3月9日のみ）
          '',                                 // D
          parsed.deviceName || '',            // E: デバイス名
          parsed.paymentMethod || '',         // F: 支払い方法（メアド含む全文）
          parsed.estimateAmount || 0,         // G: 見積もり額（数値のみ）
          '', '', '', '',                     // H〜K
          '未完了'                             // L: ステータス（メール受信直後）
        ]);
        existingIds.add(parsed.tradeInId);
        processedMessageIds.add(message.getId());
      } catch (e) {
        Logger.log('下取り本人確認メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  if (rowsToAppend.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, 12).setValues(rowsToAppend);
  }

  return {
    added: rowsToAppend.length,
    skippedDuplicate: skippedDuplicate,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

function parseTradeInConfirmMail_(message) {
  var content = getMessageContent_(message);
  var text = content.text;
  var html = content.html;
  var haystack = (text || '') + '\n' + (html || '');

  if (!haystack) return null;

  var accountEmail = extractAccountEmail_(message);

  // 下取り ID: 5251-2046-0910 → A列はハイフンなし 525120460910
  var idRaw = matchGroup_(haystack, /下取り\s*ID[\s\S]{0,80}?(\d{4}-\d{4}-\d{4})/i) ||
              matchGroup_(haystack, /(\d{4}-\d{4}-\d{4})/);
  var tradeInId = idRaw ? idRaw.replace(/-/g, '') : '';

  // 下取り手続きの開始日 → 2026年3月9日 のみ
  var startDateLine = extractValueAfterLabel_(text, ['下取り手続きの開始日']) ||
                      matchGroup_(haystack, /下取り手続きの開始日[\s\S]{0,120}?(\d{4}年\d{1,2}月\d{1,2}日)/i);
  var startDate = matchGroup_(String(startDateLine || ''), /(\d{4}年\d{1,2}月\d{1,2}日)/);

  // お支払い方法 → 全文（PayPal: xxx@gmail.com など）
  var paymentMethod = extractValueAfterLabel_(text, ['お支払い方法', 'お支払方法']) ||
                      matchGroup_(haystack, /お支払い方法[\s\S]{0,400}?eds__body[^>]*>([^<]+)</);

  // 申告されたデバイス
  var deviceName = extractValueAfterLabel_(text, ['申告されたデバイス']) ||
                   matchGroup_(haystack, /申告されたデバイス[\s\S]{0,400}?eds__body[^>]*>([^<]+)</);

  // 見積もり額 → 数値のみ
  var estimateStr = extractValueAfterLabel_(text, ['見積もり額']) ||
                    matchGroup_(haystack, /見積もり額[\s\S]{0,120}?[￥¥]?\s*([\d,]+)/i);
  var estimateAmount = parseNumber_(estimateStr);

  return {
    tradeInId: tradeInId,
    accountEmail: accountEmail,
    startDate: startDate,
    deviceName: deviceName ? deviceName.trim() : '',
    paymentMethod: paymentMethod ? paymentMethod.trim() : '',
    estimateAmount: estimateAmount
  };
}

/**
 * 下取り本人確認完了ラベル: 同じ下取りIDの tracking2 の行の L列 を「完了」に更新
 */
function processTradeInConfirmedMails_(ss, processedMessageIds) {
  var label = GmailApp.getUserLabelByName(TRADE_IN_CONFIRMED_LABEL_NAME);
  if (!label) {
    Logger.log('下取り本人確認完了ラベルが存在しません: ' + TRADE_IN_CONFIRMED_LABEL_NAME);
    return { updated: 0, noMatchedRow: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  var sheet = ss.getSheetByName(TRACKING2_SHEET_NAME);
  if (!sheet) {
    return { updated: 0, noMatchedRow: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  var threads = getThreadsByLabelNewerThan_(label, LOOKBACK_HOURS);
  var updated = 0;
  var noMatchedRow = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;
  var updateLimit = TRADE_IN_TEST_LIMIT > 0 ? TRADE_IN_TEST_LIMIT : 999999;

  for (var t = 0; t < threads.length && updated < updateLimit; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length && updated < updateLimit; m++) {
      var message = messages[m];
      try {
        if (processedMessageIds.has(message.getId())) {
          skippedAlreadyProcessed++;
          continue;
        }

        var tradeInId = parseTradeInIdFromConfirmedMail_(message);
        if (!tradeInId) {
          parseFailed++;
          continue;
        }

        var row = findRowByOrderId_(sheet, tradeInId);
        if (!row) {
          noMatchedRow++;
          processedMessageIds.add(message.getId());
          continue;
        }

        sheet.getRange(row, 12).setValue('完了');
        processedMessageIds.add(message.getId());
        updated++;
      } catch (e) {
        Logger.log('下取り本人確認完了メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  return {
    updated: updated,
    noMatchedRow: noMatchedRow,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

function parseTradeInIdFromConfirmedMail_(message) {
  var content = getMessageContent_(message);
  var html = content.html;
  var text = content.text;
  var haystack = (html || '') + '\n' + (text || '');
  var idRaw = matchGroup_(haystack, /下取り\s*ID[\s\S]{0,120}?(\d{4}-\d{4}-\d{4})/i) ||
              matchGroup_(haystack, /(\d{4}-\d{4}-\d{4})/);
  return idRaw ? idRaw.replace(/-/g, '') : '';
}

/**
 * 下取り発送ラベル: 同一下取りIDの tracking2 の D列に追跡番号（67始まり）、L列を「発送」に更新
 * 追跡番号はメール内の短縮URL(c.gle等)をリダイレクト追跡して日本郵便URLの reqCodeNo1 から取得
 */
function processTradeInShippedMails_(ss, processedMessageIds) {
  var label = GmailApp.getUserLabelByName(TRADE_IN_SHIPPED_LABEL_NAME);
  if (!label) {
    Logger.log('下取り発送ラベルが存在しません: ' + TRADE_IN_SHIPPED_LABEL_NAME);
    return { updated: 0, noMatchedRow: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  var sheet = ss.getSheetByName(TRACKING2_SHEET_NAME);
  if (!sheet) {
    return { updated: 0, noMatchedRow: 0, skippedAlreadyProcessed: 0, parseFailed: 0 };
  }

  var threads = getThreadsByLabelNewerThanDays_(label, TRADE_IN_SHIPPED_LOOKBACK_DAYS);
  var shippedLimit = TRADE_IN_TEST_LIMIT > 0 ? TRADE_IN_TEST_LIMIT : 999999;
  Logger.log('[下取り発送] 直近' + TRADE_IN_SHIPPED_LOOKBACK_DAYS + '日のスレッド数=' + threads.length + (TRADE_IN_TEST_LIMIT > 0 ? '（テスト用にメール' + TRADE_IN_TEST_LIMIT + '件まで）' : ''));
  var updated = 0;
  var noMatchedRow = 0;
  var skippedAlreadyProcessed = 0;
  var parseFailed = 0;

  for (var t = 0; t < threads.length && updated < shippedLimit; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length && updated < shippedLimit; m++) {
      var message = messages[m];
      try {
        var parsed = parseTradeInShippedMail_(message);
        if (!parsed || !parsed.tradeInId) {
          Logger.log('[下取り発送] 解析失敗: tradeInId=' + (parsed ? parsed.tradeInId : 'null'));
          parseFailed++;
          continue;
        }

        var row = findRowByOrderId_(sheet, parsed.tradeInId);
        if (!row) {
          noMatchedRow++;
          processedMessageIds.add(message.getId());
          continue;
        }

        var trackingNumber = parsed.trackingNumber || '';
        if (!trackingNumber && parsed.trackingLink) {
          var props = PropertiesService.getScriptProperties();
          if (props.getProperty('GITHUB_TOKEN') && props.getProperty('GITHUB_REPO')) {
            Logger.log('[下取り発送] GitHub に任せます tradeInId=' + parsed.tradeInId);
            triggerGitHubResolve_(parsed.trackingLink, parsed.tradeInId);
          } else {
            Logger.log('[下取り発送] 短縮URLを解決開始: ' + parsed.trackingLink.substring(0, 50) + '...');
            trackingNumber = resolveTrackingNumberFromShortUrl_(parsed.trackingLink);
            Logger.log('[下取り発送] 解決結果: 追跡番号=' + (trackingNumber || '取得できず'));
          }
        }
        var needRetryLater = parsed.trackingLink && !trackingNumber;
        if (processedMessageIds.has(message.getId()) && !needRetryLater) {
          skippedAlreadyProcessed++;
          continue;
        }
        if (processedMessageIds.has(message.getId()) && needRetryLater) {
          Logger.log('[下取り発送] 再試行（前回追跡番号未取得） tradeInId=' + parsed.tradeInId);
        }

        Logger.log('[下取り発送] 下取りID=' + parsed.tradeInId + ', リンク=' + (parsed.trackingLink ? parsed.trackingLink.substring(0, 60) + '...' : 'なし') + ', 本文の番号=' + (parsed.trackingNumber || 'なし'));

        sheet.getRange(row, 4).setValue(trackingNumber);
        sheet.getRange(row, 12).setValue('発送');
        if (!needRetryLater) processedMessageIds.add(message.getId());
        updated++;
      } catch (e) {
        Logger.log('下取り発送メール処理エラー: ' + e + ' messageId=' + message.getId());
      }
    }
  }

  return {
    updated: updated,
    noMatchedRow: noMatchedRow,
    skippedAlreadyProcessed: skippedAlreadyProcessed,
    parseFailed: parseFailed
  };
}

function parseTradeInShippedMail_(message) {
  var content = getMessageContent_(message);
  var html = content.html;
  var text = content.text;
  var haystack = (html || '') + '\n' + (text || '');
  haystack = haystack.replace(/=\s*\r?\n/g, '');
  haystack = haystack.replace(/\r?\n/g, ' ');
  var idRaw = matchGroup_(haystack, /下取り\s*ID[\s\S]{0,120}?(\d{4}-\d{4}-\d{4})/i) ||
              matchGroup_(haystack, /(\d{4}-\d{4}-\d{4})/);
  var tradeInId = idRaw ? idRaw.replace(/-/g, '') : '';
  var trackingNumber = '';
  var directJapanPost = haystack.match(/trackings\.post\.japanpost\.jp[^\s"']*reqCodeNo1=(\d+)/);
  if (directJapanPost) trackingNumber = directJapanPost[1];
  var trackingLink = '';
  // HTMLの<a href="...">から c.gle を最優先で抽出（途中で切れないよう最長URLを採用）
  if (html) {
    var cgleFromHtml = [];
    var reA = /<a[^>]+href\s*=\s*["'](https?:\/\/c\.gle\/[^"']+)["'][^>]*>/ig;
    var ma;
    while ((ma = reA.exec(html)) !== null) {
      cgleFromHtml.push(ma[1]);
    }
    if (cgleFromHtml.length > 0) {
      cgleFromHtml.sort(function (a, b) { return b.length - a.length; });
      trackingLink = cgleFromHtml[0];
    }
  }
  if (haystack.indexOf('trackings.post.japanpost') >= 0) {
    var jpMatch = haystack.match(/reqCodeNo1=(\d{12,14})/);
    if (jpMatch) trackingNumber = jpMatch[1];
  }
  var hrefBlock = haystack.match(/href\s*=\s*["']?(https?:\/\/c\.gle\/[^"']+)["']?/i) ||
                  haystack.match(/href\s*=\s*3D\s*["']?(https?:\/\/c\.gle\/[^"'\s>]+)/i);
  if (!trackingLink && hrefBlock) {
    trackingLink = hrefBlock[1].replace(/&amp;/g, '&').replace(/\s/g, '').replace(/["']+$/, '').trim();
  }
  if (!trackingLink && haystack.indexOf('c.gle') >= 0) {
    var allCgle = haystack.match(/https?:\/\/c\.gle\/[^\s"'<>]+/g);
    if (allCgle) {
      allCgle.sort(function (a, b) { return b.length - a.length; });
      trackingLink = allCgle[0].replace(/&amp;/g, '&').replace(/["']+$/, '').trim();
    }
  }
  if (!trackingLink) {
    var hrefMatch = haystack.match(/荷物の追跡[\s\S]{0,800}?href\s*=\s*["']?(https?:\/\/[^\s"'>]+)/i) ||
                    haystack.match(/荷物の追跡[\s\S]{0,800}?href\s*=\s*3D\s*["']?(https?:\/\/[^\s"'>]+)/i);
    if (hrefMatch) trackingLink = hrefMatch[1].replace(/&amp;/g, '&').replace(/["']+$/, '').replace(/\s/g, '').trim();
  }
  if (trackingLink) {
    trackingLink = trackingLink
      .replace(/["'>].*$/, '')
      .replace(/(?:style|target|class|rel|aria-[a-z-]+)=.*$/i, '')
      .replace(/&amp;/g, '&')
      .trim();
    var cleanCgle = trackingLink.match(/https?:\/\/c\.gle\/[A-Za-z0-9._~!$&'()*+,;=:@%\/?-]+/i);
    if (cleanCgle) trackingLink = cleanCgle[0];
  }
  if (!trackingLink && haystack.indexOf('japanpost') >= 0) {
    var anyJp = haystack.match(/https?:\/\/[^\s"']*trackings?\.post\.japanpost[^\s"']*reqCodeNo1=\d+/);
    if (anyJp) {
      var num = anyJp[0].match(/reqCodeNo1=(\d+)/);
      if (num) trackingNumber = num[1];
    }
  }
  Logger.log(
    '[下取り発送parse] tradeInId=' + tradeInId +
    ', linkLen=' + (trackingLink ? trackingLink.length : 0) +
    ', linkTail=' + (trackingLink ? trackingLink.slice(-30) : '') +
    ', number=' + (trackingNumber || '')
  );
  return { tradeInId: tradeInId, trackingLink: trackingLink, trackingNumber: trackingNumber };
}

/**
 * 短縮URL(c.gle等)をリダイレクト追跡し、日本郵便のURLまたはレスポンス本文から reqCodeNo1（追跡番号）を取得
 */
function resolveTrackingNumberFromShortUrl_(shortUrl) {
  if (!shortUrl || shortUrl.length < 10) {
    Logger.log('[追跡URL] 短縮URLが空または短い');
    return '';
  }
  var url = shortUrl.replace(/&amp;/g, '&').trim();
  var maxRedirects = 10;
  for (var i = 0; i < maxRedirects; i++) {
    try {
      Logger.log('[追跡URL] fetch i=' + i + ' url=' + url.substring(0, 70) + (url.length > 70 ? '...' : ''));
      var resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: false,
        validateHttpsCertificates: false,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; rv:109.0) Gecko/20100101 Firefox/115.0' }
      });
      var code = resp.getResponseCode();
      Logger.log('[追跡URL] 応答 code=' + code);
      if (code >= 200 && code < 300) {
        var match = url.match(/reqCodeNo1=(\d+)/);
        if (match) {
          Logger.log('[追跡URL] URLから取得: ' + match[1]);
          return match[1];
        }
        var body = resp.getContentText();
        if (body) {
          var m = body.match(/reqCodeNo1=(\d+)/);
          if (m) {
            Logger.log('[追跡URL] 本文から取得: ' + m[1]);
            return m[1];
          }
          var m67 = body.match(/jp&reqCodeNo1=(\d+)/);
          if (m67) {
            Logger.log('[追跡URL] 本文(jp&)から取得: ' + m67[1]);
            return m67[1];
          }
        }
        Logger.log('[追跡URL] 200だがreqCodeNo1なし');
        break;
      }
      if (code !== 301 && code !== 302 && code !== 307 && code !== 308) {
        Logger.log('[追跡URL] リダイレクト以外で終了 code=' + code);
        break;
      }
      var loc = resp.getHeaders()['Location'];
      if (Array.isArray(loc)) loc = loc[0];
      if (!loc) {
        Logger.log('[追跡URL] Locationヘッダなし');
        break;
      }
      loc = String(loc).replace(/\s/g, '');
      Logger.log('[追跡URL] Location=' + loc.substring(0, 80) + (loc.length > 80 ? '...' : ''));
      if (loc.indexOf('http') === 0 || loc.indexOf('https') === 0) {
        url = loc;
      } else {
        var base = url.replace(/\?.*$/, '').replace(/#.*$/, '').replace(/\/[^/]*$/, '');
        url = base + (loc.indexOf('/') === 0 ? loc : '/' + loc);
      }
    } catch (e) {
      Logger.log('[追跡URL] エラー: ' + e + ' url=' + url.substring(0, 60));
      break;
    }
  }
  Logger.log('[追跡URL] リダイレクト追跡で未取得、followRedirects=trueで再試行');
  try {
    var resp2 = UrlFetchApp.fetch(shortUrl.replace(/&amp;/g, '&').trim(), {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: false,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; rv:109.0) Gecko/20100101 Firefox/115.0' }
    });
    var body2 = resp2.getContentText();
    Logger.log('[追跡URL] 再試行 本文長=' + (body2 ? body2.length : 0) + ' reqCodeNo1含む=' + (body2 && body2.indexOf('reqCodeNo1=') >= 0));
    if (body2 && body2.indexOf('reqCodeNo1=') >= 0) {
      var m2 = body2.match(/reqCodeNo1=(\d{12,14})/);
      if (m2) {
        Logger.log('[追跡URL] 再試行で取得: ' + m2[1]);
        return m2[1];
      }
    }
    if (body2 && body2.indexOf('japanpost') >= 0) {
      var m3 = body2.match(/reqCodeNo1=(\d+)/);
      if (m3) {
        Logger.log('[追跡URL] 再試行(japanpost)で取得: ' + m3[1]);
        return m3[1];
      }
    }
  } catch (e2) {
    Logger.log('[追跡URL] 再試行エラー: ' + e2);
  }
  Logger.log('[追跡URL] 取得できず');
  return '';
}

/**
 * 下取り発送メールの短縮URLを GitHub Actions へ直接渡して起動する
 * URLは base64 で渡し、長いリンクでも JSON で欠けにくくする
 */
function triggerGitHubResolve_(shortUrl, tradeInId) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  var repo = props.getProperty('GITHUB_REPO');
  var workflowId = props.getProperty('GITHUB_WORKFLOW_ID');
  if (!token || !repo || !workflowId) {
    Logger.log('[GitHub] 未設定のためスキップ: GITHUB_TOKEN, GITHUB_REPO, GITHUB_WORKFLOW_ID');
    return;
  }
  if (!tradeInId || !shortUrl) return;
  var shortUrlB64 = Utilities.base64Encode(shortUrl);
  var url = 'https://api.github.com/repos/' + repo + '/actions/workflows/' + encodeURIComponent(workflowId) + '/dispatches';
  var payload = JSON.stringify({
    ref: 'main',
    inputs: {
      trade_in_id: tradeInId,
      short_url_b64: shortUrlB64
    }
  });
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
      Logger.log('[GitHub] ワークフロー起動しました tradeInId=' + tradeInId);
    } else {
      Logger.log('[GitHub] 起動失敗: ' + resp.getResponseCode() + ' ' + resp.getContentText());
    }
  } catch (e) {
    Logger.log('[GitHub] エラー: ' + e);
  }
}

/**
 * 指定ラベルの「直近 hours 時間以内」のスレッドだけ取得（1時間おき実行用・最小限の読み取り）
 */
function getThreadsByLabelNewerThan_(label, hours) {
  if (!label || hours < 1) return [];
  var name = (label.getName() || '').replace(/"/g, '\\"');
  var query = 'label:"' + name + '" newer_than:' + hours + 'h';
  return GmailApp.search(query);
}

/**
 * 指定ラベルの「直近 days 日以内」のスレッドを取得（下取りキャンセル時の削除用など）
 */
function getThreadsByLabelNewerThanDays_(label, days) {
  if (!label || days < 1) return [];
  var name = (label.getName() || '').replace(/"/g, '\\"');
  var query = 'label:"' + name + '" newer_than:' + days + 'd';
  return GmailApp.search(query);
}

/**
 * ラベルのスレッドをページングで全件取得する（デバッグ等で使用）
 */
function getAllThreadsByLabel_(label, maxThreads) {
  const out = [];
  const pageSize = 100;
  var start = 0;
  while (start < maxThreads) {
    const batch = label.getThreads(start, Math.min(pageSize, maxThreads - start));
    if (!batch || batch.length === 0) break;
    for (var i = 0; i < batch.length; i++) {
      out.push(batch[i]);
    }
    if (batch.length < pageSize) break;
    start += batch.length;
  }
  return out;
}

function findRowByOrderId_(sheet, orderId) {
  const target = String(orderId || '').trim();
  if (!target) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 0;
  const ids = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0] || '').trim() === target) {
      return i + 1;
    }
  }
  return 0;
}

function getCancelDeadline_(deliverySchedule, orderDateStr) {
  const schedule = String(deliverySchedule || '');
  const m = schedule.match(/([0-9]{1,2})月([0-9]{1,2})日/);
  if (!m) return null;

  var year = extractYearFromOrderDate_(orderDateStr);
  if (!year) year = new Date().getFullYear();
  const month = Number(m[1]);
  const day = Number(m[2]);
  if (!month || !day) return null;

  const deliveryDate = new Date(year, month - 1, day, 0, 0, 0, 0);
  const orderDate = parseOrderDate_(orderDateStr);
  if (orderDate && deliveryDate.getTime() < orderDate.getTime() - 24 * 60 * 60 * 1000) {
    deliveryDate.setFullYear(year + 1);
  }

  const deadline = new Date(deliveryDate.getTime());
  deadline.setDate(deadline.getDate() - 1);
  deadline.setHours(23, 50, 0, 0);
  return deadline;
}

function extractYearFromOrderDate_(s) {
  const m = String(s || '').match(/([0-9]{4})年/);
  return m ? Number(m[1]) : 0;
}

function parseOrderDate_(s) {
  const m = String(s || '').match(/([0-9]{4})年([0-9]{1,2})月([0-9]{1,2})日/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
}

function loadProcessedMessageIds_(props) {
  const raw = props.getProperty('PROCESSED_MESSAGE_IDS');
  if (!raw) return new Set();
  try {
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return new Set();
    return new Set(arr);
  } catch (e) {
    Logger.log('PROCESSED_MESSAGE_IDS の読み込みに失敗: ' + e);
    return new Set();
  }
}

function saveProcessedMessageIds_(props, setObj) {
  // Script Properties サイズ対策: 直近5000件だけ保持
  const arr = Array.from(setObj);
  const limited = arr.length > 5000 ? arr.slice(arr.length - 5000) : arr;
  props.setProperty('PROCESSED_MESSAGE_IDS', JSON.stringify(limited));
}

function loadSyncedAccountOrderIds_(props) {
  const raw = props.getProperty('ACCOUNT_SYNCED_ORDER_IDS');
  if (!raw) return new Set();
  try {
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return new Set();
    return new Set(arr);
  } catch (e) {
    Logger.log('ACCOUNT_SYNCED_ORDER_IDS の読み込みに失敗: ' + e);
    return new Set();
  }
}

function saveSyncedAccountOrderIds_(props, setObj) {
  const arr = Array.from(setObj);
  const limited = arr.length > 10000 ? arr.slice(arr.length - 10000) : arr;
  props.setProperty('ACCOUNT_SYNCED_ORDER_IDS', JSON.stringify(limited));
}

/**
 * 再テスト用：処理済みmessageIdの記録をクリア
 */
function resetProcessedMessageIdsForTest() {
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_MESSAGE_IDS');
  Logger.log('PROCESSED_MESSAGE_IDS をクリアしました');
}

function resetAccountSyncForTest() {
  PropertiesService.getScriptProperties().deleteProperty('ACCOUNT_SYNCED_ORDER_IDS');
  Logger.log('ACCOUNT_SYNCED_ORDER_IDS をクリアしました');
}

/**
 * Tracking シートから既存の注文番号セットを読み込む
 */
function loadExistingOrderIds_(sheet) {
  const lastRow = sheet.getLastRow();
  const set = new Set();

  if (lastRow < 1) {
    return set;
  }

  const values = sheet.getRange(1, 1, lastRow, 1).getValues(); // A列
  for (var i = 0; i < values.length; i++) {
    const id = String(values[i][0] || '').trim();
    if (id) {
      set.add(id);
    }
  }
  return set;
}

/**
 * メールが指定ラベルを持っているか
 */
function messageHasLabel_(message, label) {
  const labels = message.getThread().getLabels();
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].getName() === label.getName()) {
      return true;
    }
  }
  return false;
}

/**
 * ラベルを取得（なければ作成）
 */
function getOrCreateLabel_(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

/**
 * メール1通から必要情報を正規表現で抽出
 */
function parseOrderMail_(message) {
  const content = getMessageContent_(message);
  const text = content.text;
  const html = content.html;
  const haystack = (text || '') + '\n' + (html || '');

  if (!text && !html) {
    return null;
  }

  const accountEmail = extractAccountEmail_(message);

  const orderNumberRaw = matchGroup_(haystack, /(GS\.[0-9\-]+)/); // 例: GS.0310-8892-8323
  // 数字だけに整形
  const orderId = orderNumberRaw ? orderNumberRaw.replace(/[^\d]/g, '') : '';

  const orderDateLine = (
    extractValueAfterLabel_(text, ['ご注文日', '注文日', 'ご注文日時']) ||
    matchGroup_(text, /([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日)\s*[0-9]{1,2}:[0-9]{2}/) ||
    matchGroup_(haystack, /([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日)\s*[0-9]{1,2}:[0-9]{2}/)
  );
  const orderDate = matchGroup_(String(orderDateLine || ''), /([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日)/);

  const paymentRaw = extractValueAfterLabel_(text, ['お支払い方法', 'お支払方法']);
  const paymentMethod = normalizePaymentMethod_(paymentRaw);

  const productName = (
    extractProductNameFromHtml_(html) ||
    extractProductNameFromText_(text) ||
    ''
  );

  const idNumber = (
    matchGroup_(haystack, /ID\s*番号[:：]?\s*(?:<br[^>]*>\s*)?(\d{8,})/i) ||
    matchGroup_(text, /ID\s*番号[:：]?\s*(\d{8,})/) ||
    matchGroup_(text, /(?:注文\s*ID|注文ID)[:：]\s*([A-Za-z0-9\-]+)/) ||
    ''
  );

  const subtotal = parseNumber_(
    extractValueAfterLabel_(text, ['小計']) ||
    matchGroup_(text, /小計\s*￥?\s*([-\d,]+)/)
  );

  const tax = parseNumber_(
    extractValueAfterLabel_(text, ['消費税', '税金']) ||
    matchGroup_(text, /(?:消費税|税金)\s*￥?\s*([-\d,]+)/)
  );

  const usedStorePoints = parseNumber_(
    extractValueAfterLabel_(text, [
      '使用ストアP',
      '使用ストアポイント',
      '使用ストアクレジット',
      'Google ストア クレジット',
      'Google ストアクレジット',
      'ストア クレジット',
      'Google ストア ポイント',
      'Google ストアポイント'
    ]) ||
    matchGroup_(text, /Google\s*ストア\s*(?:クレジット|ポイント)\s*-?￥?\s*([-\d,]+)/)
  );

  const couponDiscount = parseNumber_(
    extractValueAfterLabel_(text, ['クーポン割引', 'クーポン', '割引']) ||
    matchGroup_(text, /(?:クーポン割引|クーポン|割引)\s*-?￥?\s*([-\d,]+)/)
  );

  const paidAmount = parseNumber_(
    extractValueAfterLabel_(text, ['お支払い総額', 'お支払い金額', 'ご請求額', '合計']) ||
    matchGroup_(text, /(?:お支払い総額|お支払い金額|ご請求額|合計)\s*￥?\s*([-\d,]+)/)
  );

  const bonusStorePoints = parseNumber_(
    matchGroup_(html, /Google\s*ストア\s*ポイント\s*プロモーション[\s\S]{0,300}?item-value[^>]*>\s*(?:&yen;|￥)?\s*([-\d,]+)/i) ||
    extractValueAfterLabel_(text, ['特典ストアP', '獲得予定ストアポイント', 'Google ストアポイント プロモーション', 'Google ストア ポイント プロモーション', 'Google ストアポイント', 'Google ストア ポイント']) ||
    matchGroup_(text, /Google\s*ストア\s*ポイント\s*プロモーション[\s\S]{0,120}?￥?\s*([-\d,]+)/) ||
    matchGroup_(text, /Google\s*ストア\s*ポイント[\s\S]{0,80}?￥?\s*([\d,]+)/)
  );

  const trackingNumber = (
    matchGroup_(html, /itemprop\s*=\s*["']trackingNumber["'][^>]*content\s*=\s*["']([^"']+)["']/i) ||
    matchGroup_(html, /https?:\/\/store\.google\.com\/track\/(\d+)/i) ||
    matchGroup_(text, /(?:追跡番号|お荷物番号|お問い合わせ伝票番号|配送業者のお問い合わせ番号)[:：]?\s*([A-Za-z0-9\-]+)/) ||
    ''
  );
  const deliverySchedule = normalizeDeliverySchedule_(
    matchGroup_(html, /お届け予定日時[:：]?\s*<\/span>\s*([^<\n]+)/i) ||
    extractValueAfterLabel_(text, ['お届け予定日時', 'お届け予定日']) ||
    matchGroup_(text, /([0-9]{1,2}月[0-9]{1,2}日\s*\([^)]+\)\s*[0-9]{1,2}:[0-9]{2}\s*-\s*[0-9]{1,2}:[0-9]{2})/)
  );

  return {
    orderId: orderId,
    accountEmail: accountEmail,
    orderDate: orderDate,
    paymentMethod: paymentMethod,
    productName: productName,
    idNumber: idNumber,
    subtotal: subtotal,
    tax: tax,
    usedStorePoints: usedStorePoints,
    couponDiscount: couponDiscount,
    paidAmount: paidAmount,
    bonusStorePoints: bonusStorePoints,
    trackingNumber: trackingNumber,
    deliverySchedule: deliverySchedule
  };
}

function parseShippedMail_(message) {
  const content = getMessageContent_(message);
  const text = content.text;
  const html = content.html;
  const haystack = (text || '') + '\n' + (html || '');

  const orderRaw = (
    matchGroup_(html, /itemprop\s*=\s*["']orderNumber["'][^>]*content\s*=\s*["'](GS\.[0-9\-]+)["']/i) ||
    matchGroup_(haystack, /(GS\.[0-9\-]+)/)
  );
  const orderId = orderRaw ? orderRaw.replace(/[^\d]/g, '') : '';

  const trackingNumber = (
    matchGroup_(html, /itemprop\s*=\s*["']trackingNumber["'][^>]*content\s*=\s*["']([^"']+)["']/i) ||
    matchGroup_(html, /https?:\/\/store\.google\.com\/track\/(\d+)/i) ||
    ''
  );

  const idNumber = (
    matchGroup_(haystack, /ID\s*番号[:：]?\s*(?:<br[^>]*>\s*)?([A-Za-z0-9\-]{8,})/i) ||
    matchGroup_(text, /ID\s*番号[:：]?\s*([A-Za-z0-9\-]{8,})/) ||
    ''
  );

  const deliverySchedule = normalizeDeliverySchedule_(
    matchGroup_(html, /お届け予定日時[:：]?\s*<\/span>\s*([^<\n]+)/i) ||
    extractValueAfterLabel_(text, ['お届け予定日時', 'お届け予定日']) ||
    matchGroup_(text, /([0-9]{1,2}月[0-9]{1,2}日\s*\([^)]+\)\s*[0-9]{1,2}:[0-9]{2}\s*-\s*[0-9]{1,2}:[0-9]{2})/)
  );

  return {
    orderId: orderId,
    trackingNumber: trackingNumber,
    idNumber: idNumber,
    deliverySchedule: deliverySchedule
  };
}

function parseDeliveryChangedMail_(message) {
  const content = getMessageContent_(message);
  const text = content.text;
  const html = content.html;
  const haystack = (text || '') + '\n' + (html || '');

  const orderRaw = (
    matchGroup_(html, /注文番号[:：]?\s*(GS\.[0-9\-]+)/i) ||
    matchGroup_(haystack, /(GS\.[0-9\-]+)/)
  );
  const orderId = orderRaw ? orderRaw.replace(/[^\d]/g, '') : '';

  const deliverySchedule = normalizeDeliverySchedule_(
    matchGroup_(html, /お届け予定日(?:時)?[:：]?\s*<\/span>\s*([^<\n]+)/i) ||
    extractValueAfterLabel_(text, ['お届け予定日時', 'お届け予定日']) ||
    matchGroup_(text, /([0-9]{1,2}月[0-9]{1,2}日(?:\s*\([^)]+\))?(?:\s*[0-9]{1,2}:[0-9]{2}\s*-\s*[0-9]{1,2}:[0-9]{2})?)/)
  );

  return {
    orderId: orderId,
    deliverySchedule: deliverySchedule
  };
}

function parseOrderIdFromAnyMail_(message) {
  const content = getMessageContent_(message);
  const text = content.text;
  const html = content.html;
  const haystack = (text || '') + '\n' + (html || '');
  const orderRaw = (
    matchGroup_(html, /注文番号[:：]?\s*(GS\.[0-9\-]+)/i) ||
    matchGroup_(html, /itemprop\s*=\s*["']orderNumber["'][^>]*content\s*=\s*["'](GS\.[0-9\-]+)["']/i) ||
    matchGroup_(haystack, /(GS\.[0-9\-]+)/)
  );
  return {
    orderId: orderRaw ? orderRaw.replace(/[^\d]/g, '') : ''
  };
}

/**
 * メール本文（text/plain 優先）と html を返す
 */
function getMessageContent_(message) {
  const plainRaw = message.getPlainBody() || '';
  const htmlRaw = message.getBody() || '';

  const plain = maybeDecodeQuotedPrintable_(plainRaw);
  const html = maybeDecodeQuotedPrintable_(htmlRaw);

  const text = normalizeText_(plain && plain.trim() ? plain : stripHtml_(html));
  return { text: text, html: html };
}

/**
 * HTML からテキストをざっくり抽出
 */
function stripHtml_(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&yen;/g, '￥')
    .replace(/&amp;/g, '&')
    .replace(/\r\n/g, '\n');
}

function normalizeText_(text) {
  return String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * quoted-printable っぽい文字列を（必要なら）UTF-8として復元
 * Gmailの「元のメッセージ」由来の =E3=81=... のような断片に対応
 */
function maybeDecodeQuotedPrintable_(s) {
  const input = String(s || '');
  if (!input) return '';
  // 典型的な quoted-printable の断片が少ない場合はそのまま
  if (!/=[0-9A-F]{2}/i.test(input) || (input.match(/=[0-9A-F]{2}/gi) || []).length < 5) {
    return input;
  }
  try {
    return decodeQuotedPrintableUtf8_(input);
  } catch (e) {
    return input;
  }
}

function decodeQuotedPrintableUtf8_(qp) {
  // soft line break: =\r\n / =\n を除去
  const s = String(qp).replace(/=\r?\n/g, '');
  const bytes = [];

  for (var i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '=' && i + 2 < s.length && /[0-9A-F]{2}/i.test(s.substr(i + 1, 2))) {
      bytes.push(parseInt(s.substr(i + 1, 2), 16));
      i += 2;
      continue;
    }
    const code = s.charCodeAt(i);
    // ASCIIはそのまま、非ASCIIはUTF-16のまま入る可能性があるのでUTF-8化はしない（混在対策）
    if (code <= 0xFF) {
      bytes.push(code);
    } else {
      // 既にデコード済みっぽい場合はそのまま返す
      throw new Error('already decoded');
    }
  }
  return Utilities.newBlob(bytes).getDataAsString('UTF-8');
}

function extractValueAfterLabel_(text, labelCandidates) {
  const t = String(text || '');
  if (!t) return '';
  for (var i = 0; i < labelCandidates.length; i++) {
    const label = labelCandidates[i];
    const idx = t.indexOf(label);
    if (idx < 0) continue;

    let after = t.slice(idx + label.length);
    after = after.replace(/^[\s:：]+/, '');

    // 同じ行に値があるケース
    const sameLine = after.split('\n')[0].trim();
    if (sameLine) return sameLine;

    // 次行以降の最初の非空行
    const lines = after.split('\n');
    for (var j = 1; j < lines.length; j++) {
      const line = String(lines[j] || '').trim();
      if (!line) continue;
      if (line === label) continue;
      return line;
    }
  }
  return '';
}

function normalizePaymentMethod_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  if (/paypal/i.test(s)) {
    // 例: "PayPal: shosuke0824@gmail.com" はメール末尾まで保持
    return s;
  }
  const card = normalizeCardPaymentText_(s);
  if (card) return card;
  return s;
}

function normalizeCardPaymentText_(s) {
  const cleaned = String(s || '')
    .replace(/[•●・]/g, ' ')
    .replace(/\*{2,}/g, ' ')
    .replace(/[xX]{2,}/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const m = cleaned.match(/(visa|mastercard|master|jcb|amex|american\s*express)\s*(\d{4})/i);
  if (!m) return '';
  var brand = m[1].toLowerCase();
  if (brand === 'master') brand = 'mastercard';
  if (brand === 'american express') brand = 'amex';
  return brand + ' ' + m[2];
}

function normalizeDeliverySchedule_(s) {
  const value = String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const m = value.match(/([0-9]{1,2}月[0-9]{1,2}日(?:\s*\([^)]+\))?(?:\s*[0-9]{1,2}:[0-9]{2}\s*-\s*[0-9]{1,2}:[0-9]{2})?)/);
  return m && m[1] ? m[1].replace(/\s+/g, ' ').trim() : value;
}

function extractProductNameFromHtml_(html) {
  const h = String(html || '');
  if (!h) return '';
  const m = h.match(/text--heading[^>]*>\s*([^<]+)\s*<\/p>/i);
  if (m && m[1]) return normalizeText_(m[1]);
  return '';
}

function extractProductNameFromText_(text) {
  const t = String(text || '');
  if (!t) return '';
  const labeled = extractValueAfterLabel_(t, ['商品名', '商品', 'アイテム']);
  if (labeled) return labeled;

  const idx = t.indexOf('ID 番号');
  if (idx > 0) {
    const before = t.slice(0, idx);
    const lines = before.split('\n').map(function (x) { return String(x || '').trim(); }).filter(function (x) { return !!x; });
    for (var i = lines.length - 1; i >= 0; i--) {
      const line = lines[i];
      if (line.length < 5) continue;
      if (/^(ID|注文番号|ご注文日|お支払い|小計|合計|消費税|送料)/.test(line)) continue;
      return line;
    }
  }
  return '';
}

/**
 * 正規表現の第1キャプチャを返すユーティリティ
 */
function matchGroup_(text, regex) {
  const m = text.match(regex);
  return m && m[1] ? m[1].trim() : '';
}

/**
 * 金額などの数値文字列を数値に変換（見つからなければ 0）
 */
function parseNumber_(str) {
  if (!str) return 0;
  const cleaned = String(str).replace(/[^\d]/g, '');
  return cleaned ? Number(cleaned) : 0;
}

function stringifyOrderId_(orderId) {
  if (!orderId) return '';
  // setValues時に数値化されないよう明示的に文字列で返す
  return String(orderId);
}

/**
 * アカウントメールアドレスをヘッダから抽出
 * 優先順位: Delivered-To > To > From
 */
function extractAccountEmail_(message) {
  const deliveredTo = message.getHeader('Delivered-To');
  const to = message.getTo();
  const from = message.getFrom();

  return (
    extractEmailFromHeader_(deliveredTo) ||
    extractEmailFromHeader_(to) ||
    extractEmailFromHeader_(from) ||
    ''
  );
}

/**
 * ヘッダ文字列からメールアドレスだけを取り出す
 */
function extractEmailFromHeader_(headerValue) {
  if (!headerValue) return '';
  const m = headerValue.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[A-Za-z]{2,})/);
  return m ? m[1] : '';
}

