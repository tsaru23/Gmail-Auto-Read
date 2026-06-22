/**
 * Gmail不要メール自動既読処理のメインエントリーポイント
 * このスクリプトは時間主導型トリガーによって定期的に実行されることを想定しています。
 */
function autoMarkAsRead() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    console.error('エラー: スクリプトプロパティ「SPREADSHEET_ID」が設定されていません。');
    return;
  }
  
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    console.error('エラー: 指定されたスプレッドシートを開くことができませんでした。IDを確認してください。', e.toString());
    return;
  }
  
  // フィルタールールの読み込み
  const rules = loadFilterRules(spreadsheet);
  if (rules.length === 0) {
    console.log('有効なフィルタールールが見つかりませんでした。処理を終了します。');
    return;
  }
  
  console.log(`${rules.length} 件のフィルタールールを読み込みました。`);
  
  // 未読メールの検索（1回につき最大100件処理）
  // 効率化のため「is:unread」で未読スレッドのみを対象とする
  const threads = GmailApp.search('is:unread', 0, 100);
  console.log(`未読スレッドが ${threads.length} 件見つかりました。解析を開始します。`);
  
  let processedCount = 0;
  
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    
    // スレッド内の最新のメッセージ（あるいはすべてのメッセージ）をチェック
    // ここではスレッド全体が既読になるため、スレッド内のいずれかのメッセージがルールに合致するか確認
    let shouldMarkRead = false;
    let matchingRule = null;
    let matchedMessageInfo = null;
    
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      
      // メッセージが未読の場合のみチェック
      if (message.isUnread()) {
        const from = message.getFrom();
        const subject = message.getSubject();
        const body = message.getPlainBody();
        
        // ルール判定
        for (const rule of rules) {
          if (isMatch(rule, from, subject, body)) {
            shouldMarkRead = true;
            matchingRule = rule;
            matchedMessageInfo = {
              from: from,
              subject: subject,
              id: message.getId()
            };
            break;
          }
        }
      }
      if (shouldMarkRead) break;
    }
    
    // ルールにマッチした場合、スレッドを既読にする
    if (shouldMarkRead && matchedMessageInfo) {
      thread.markRead();
      processedCount++;
      
      const logMessage = `既読化成功: [ルール: ${matchingRule.type} = ${matchingRule.pattern}] 件名: "${matchedMessageInfo.subject}" (From: ${matchedMessageInfo.from})`;
      console.log(logMessage);
      
      // スレッドにログシートがある場合は記録
      logToSpreadsheet(spreadsheet, matchedMessageInfo.from, matchedMessageInfo.subject, matchingRule);
    }
  }
  
  console.log(`処理完了: ${processedCount} 件のスレッドを自動既読にしました。`);
}

/**
 * スプレッドシートからフィルタールールを読み込む
 * シート名: "FilterRules"
 * カラム構成: [Type (From/Subject/Body), Pattern, MatchType (Exact/Contains/RegExp), Description, Active (TRUE/FALSE)]
 */
function loadFilterRules(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('FilterRules');
  if (!sheet) {
    console.warn('警告: 「FilterRules」シートが見つかりません。新規作成します。');
    createDefaultRulesSheet(spreadsheet);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const rules = [];
  
  // 1行目はヘッダーなので2行目から読み込み
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = String(row[0]).trim();
    const pattern = String(row[1]).trim();
    const matchType = String(row[2]).trim();
    const description = String(row[3]).trim();
    const isActive = row[4]; // Boolean値
    
    // 必須入力項目のチェックおよび有効フラグの確認
    if (type && pattern && matchType && isActive === true) {
      rules.push({
        type: type, // 'From', 'Subject', 'Body'
        pattern: pattern,
        matchType: matchType, // 'Exact', 'Contains', 'RegExp'
        description: description
      });
    }
  }
  
  return rules;
}

/**
 * メッセージがルールに合致するか判定する
 */
function isMatch(rule, from, subject, body) {
  let targetValue = '';
  
  // 判定対象の選択
  switch (rule.type.toLowerCase()) {
    case 'from':
      targetValue = from;
      break;
    case 'subject':
      targetValue = subject;
      break;
    case 'body':
      targetValue = body;
      break;
    default:
      return false;
  }
  
  // 一致方式による判定
  switch (rule.matchType.toLowerCase()) {
    case 'exact':
      return targetValue === rule.pattern;
      
    case 'contains':
      return targetValue.indexOf(rule.pattern) !== -1;
      
    case 'regexp':
      try {
        // パブリック公開時の予期せぬエラー防止のため、正規表現エラーをキャッチ
        const regex = new RegExp(rule.pattern, 'i'); // 大文字小文字を区別しない
        return regex.test(targetValue);
      } catch (e) {
        console.error(`正規表現の解析エラー: ルールパターン "${rule.pattern}"`, e.toString());
        return false;
      }
      
    default:
      return false;
  }
}

/**
 * 処理結果をスプレッドシートのログシートに記録する
 * シート名: "Logs"
 */
function logToSpreadsheet(spreadsheet, from, subject, rule) {
  let sheet = spreadsheet.getSheetByName('Logs');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Logs');
    sheet.appendRow(['日時', '送信元', '件名', 'マッチしたルール種別', 'マッチしたパターン', 'ルールの説明']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  const timestamp = new Date();
  sheet.appendRow([
    timestamp,
    from,
    subject,
    rule.type,
    rule.pattern,
    rule.description
  ]);
  
  // ログが増えすぎないように過去データの上限設定（例: 最大1000行）を設けると親切
  const maxRows = 1000;
  const currentRows = sheet.getLastRow();
  if (currentRows > maxRows) {
    // ヘッダー行（1行目）を残して、2行目から超過分を削除
    const deleteCount = currentRows - maxRows;
    sheet.deleteRows(2, deleteCount);
  }
}

/**
 * デフォルトのルールシートを作成するヘルパー関数
 */
function createDefaultRulesSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('FilterRules');
  sheet.appendRow(['Type (From/Subject/Body)', 'Pattern', 'MatchType (Exact/Contains/RegExp)', 'Description', 'Active (TRUE/FALSE)']);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e0e0e0');
  
  // サンプルのルールを追加（無効状態で追加）
  sheet.appendRow(['From', 'noreply@newsletter.com', 'Exact', 'メルマガ配信元（サンプル）', false]);
  sheet.appendRow(['Subject', '【広告】', 'Contains', '広告メール（サンプル）', false]);
  sheet.appendRow(['Body', '登録解除はこちら', 'Contains', '本文に登録解除があるもの（サンプル）', false]);
  
  // 列幅を調整
  sheet.autoResizeColumns(1, 5);
}

/**
 * 解析用にGmailの直近の未読メールをスプレッドシートに書き出す関数
 * シート名: "TempEmails"
 * ※解析後はこのシートを削除するか、データを消去してください。
 */
function exportEmailsForAnalysis() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    console.error('エラー: スクリプトプロパティ「SPREADSHEET_ID」が設定されていません。');
    return;
  }
  
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    console.error('エラー: 指定されたスプレッドシートを開くことができませんでした。IDを確認してください。', e.toString());
    return;
  }
  
  let sheet = spreadsheet.getSheetByName('TempEmails');
  if (sheet) {
    sheet.clear(); // 既存のデータをクリア
  } else {
    sheet = spreadsheet.insertSheet('TempEmails');
  }
  
  // ヘッダーの設定
  sheet.appendRow(['日時', '送信元 (From)', '件名 (Subject)', '本文プレビュー (Body Preview)']);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d9ead3');
  
  // 直近の未読スレッド50件を取得
  const threads = GmailApp.search('is:unread', 0, 50);
  console.log(`未読メール ${threads.length} 件をスプレッドシートにエクスポートします。`);
  
  const data = [];
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    
    // スレッド内の未読メッセージを走査
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      if (message.isUnread()) {
        const date = message.getDate();
        const from = message.getFrom();
        const subject = message.getSubject();
        // 本文を300文字に切り詰めてプレビューとする
        const body = message.getPlainBody().substring(0, 300) + '...';
        
        data.push([date, from, subject, body]);
      }
    }
  }
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 4).setValues(data);
    sheet.autoResizeColumns(1, 4);
    console.log(`エクスポートが完了しました。「TempEmails」シートを確認してください。計 ${data.length} 件`);
  } else {
    console.log('未読メールが見つかりませんでした。');
  }
}

