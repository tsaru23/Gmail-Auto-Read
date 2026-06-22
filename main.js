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
  const threads = GmailApp.search('is:unread', 0, 100);
  console.log(`未読スレッドが ${threads.length} 件見つかりました。解析を開始します。`);
  
  let processedCount = 0;
  
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    
    let shouldMarkRead = false;
    let matchingRule = null;
    let matchedMessageInfo = null;
    
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      
      if (message.isUnread()) {
        const from = message.getFrom();
        const subject = message.getSubject();
        const body = message.getPlainBody();
        
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
      
      const logMessage = `既読化成功: [ルール: ${matchingRule.displayType} = ${matchingRule.pattern}] 件名: "${matchedMessageInfo.subject}" (From: ${matchedMessageInfo.from})`;
      console.log(logMessage);
      
      // ログ記録
      logToSpreadsheet(spreadsheet, matchedMessageInfo.from, matchedMessageInfo.subject, matchingRule);
    }
  }
  
  console.log(`処理完了: ${processedCount} 件のスレッドを自動既読にしました。`);
}

/**
 * スプレッドシートからフィルタールールを読み込む
 * シート名: "FilterRules"
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
  
  // マッピング定義（日本語入力値を内部コードにマッピング、後方互換性のため旧英語入力も許容）
  const typeMap = {
    '送信元(From)': 'From', '送信元': 'From', 'From': 'From', 'from': 'From',
    '件名(Subject)': 'Subject', '件名': 'Subject', 'Subject': 'Subject', 'subject': 'Subject',
    '本文(Body)': 'Body', '本文': 'Body', 'Body': 'Body', 'body': 'Body'
  };
  
  const matchTypeMap = {
    '完全一致(Exact)': 'Exact', '完全一致': 'Exact', 'Exact': 'Exact', 'exact': 'Exact',
    '部分一致(Contains)': 'Contains', '部分一致': 'Contains', 'Contains': 'Contains', 'contains': 'Contains',
    '正規表現(RegExp)': 'RegExp', '正規表現': 'RegExp', 'RegExp': 'RegExp', 'regexp': 'RegExp'
  };
  
  // 1行目はヘッダーなので2行目から読み込み
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawType = String(row[0]).trim();
    const pattern = String(row[1]).trim();
    const rawMatchType = String(row[2]).trim();
    const description = String(row[3]).trim();
    const isActive = row[4]; // チェックボックス(Boolean)
    
    const type = typeMap[rawType];
    const matchType = matchTypeMap[rawMatchType];
    
    // 必須入力項目のチェックおよび有効フラグの確認
    if (type && pattern && matchType && isActive === true) {
      rules.push({
        type: type, // 'From', 'Subject', 'Body'
        displayType: rawType, // ログ表示用のオリジナル値
        pattern: pattern,
        matchType: matchType, // 'Exact', 'Contains', 'RegExp'
        displayMatchType: rawMatchType,
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
  
  switch (rule.type) {
    case 'From':
      targetValue = from;
      break;
    case 'Subject':
      targetValue = subject;
      break;
    case 'Body':
      targetValue = body;
      break;
    default:
      return false;
  }
  
  switch (rule.matchType) {
    case 'Exact':
      return targetValue === rule.pattern;
      
    case 'Contains':
      return targetValue.indexOf(rule.pattern) !== -1;
      
    case 'RegExp':
      try {
        const regex = new RegExp(rule.pattern, 'i');
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
    rule.displayType,
    rule.pattern,
    rule.description
  ]);
  
  // ログ上限設定（最大1000行）
  const maxRows = 1000;
  const currentRows = sheet.getLastRow();
  if (currentRows > maxRows) {
    const deleteCount = currentRows - maxRows;
    sheet.deleteRows(2, deleteCount);
  }
}

/**
 * デフォルトのルールシートおよびヘルプシートを作成する
 */
function createDefaultRulesSheet(spreadsheet) {
  // 直接実行された場合など、引数がない場合はスクリプトプロパティから開く
  if (!spreadsheet) {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      console.error('エラー: スクリプトプロパティ「SPREADSHEET_ID」が設定されていません。');
      return;
    }
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      console.error('エラー: 指定されたスプレッドシートを開くことができませんでした。', e.toString());
      return;
    }
  }

  let sheet = spreadsheet.getSheetByName('FilterRules');
  const isNew = !sheet;
  
  if (isNew) {
    sheet = spreadsheet.insertSheet('FilterRules');
  } else {
    sheet.clear();
  }
  
  // ヘッダーの作成
  sheet.appendRow(['種別 (Type)', 'キーワード (Pattern)', '一致方法 (MatchType)', '説明 (Description)', '有効にする (Active)']);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#d9e1f2').setHorizontalAlignment('center');
  
  // データの入力規則（プルダウン設定）
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['送信元(From)', '件名(Subject)', '本文(Body)'], true)
    .setAllowInvalid(false)
    .setHelpText('種別を選択してください。')
    .build();
  sheet.getRange('A2:A100').setDataValidation(typeRule);
  
  const matchTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['完全一致(Exact)', '部分一致(Contains)', '正規表現(RegExp)'], true)
    .setAllowInvalid(false)
    .setHelpText('一致方法を選択してください。')
    .build();
  sheet.getRange('C2:C100').setDataValidation(matchTypeRule);
  
  // チェックボックスをE列（有効フラグ）に挿入
  sheet.getRange('E2:E100').insertCheckboxes();
  
  // サンプルの設定
  sheet.getRange('A2').setValue('送信元(From)');
  sheet.getRange('B2').setValue('info@nikki.ne.jp');
  sheet.getRange('C2').setValue('完全一致(Exact)');
  sheet.getRange('D2').setValue('みんなのキャンパスのメルマガを既読にする');
  sheet.getRange('E2').setValue(true); // 最初から有効化しておく
  
  sheet.getRange('A3').setValue('件名(Subject)');
  sheet.getRange('B3').setValue('(起業|クラブオフ|イベント|キャンペーン|おしゃべり|クイズ)');
  sheet.getRange('C3').setValue('正規表現(RegExp)');
  sheet.getRange('D3').setValue('大学からの不要なイベント・広告メール');
  sheet.getRange('E3').setValue(true); // 最初から有効化しておく

  sheet.getRange('A4').setValue('送信元(From)');
  sheet.getRange('B4').setValue('.*@mail\\.axol\\.jp');
  sheet.getRange('C4').setValue('正規表現(RegExp)');
  sheet.getRange('D4').setValue('就活関係の自動配信メール（一括既読化）');
  sheet.getRange('E4').setValue(false); // 安全のため最初は無効化
  
  sheet.autoResizeColumns(1, 5);
  
  // ヘルプシートの作成
  createHelpSheet(spreadsheet);
}

/**
 * 使い方ヘルプシートを作成する
 */
function createHelpSheet(spreadsheet) {
  let helpSheet = spreadsheet.getSheetByName('使い方ヘルプ');
  if (helpSheet) {
    spreadsheet.deleteSheet(helpSheet); // 常に最新の状態に更新
  }
  helpSheet = spreadsheet.insertSheet('使い方ヘルプ');
  
  helpSheet.appendRow(['★ Gmail自動既読システム 使い方ヘルプ ★']);
  helpSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1f4e78');
  
  helpSheet.appendRow(['']);
  helpSheet.appendRow(['このシステムは、「FilterRules」シートで設定したルールに従って、Gmailの未読メールを自動で既読にします。']);
  helpSheet.appendRow(['AIに頼らなくても、以下のルール設定例を参考に「FilterRules」に自分で記入して管理することができます。']);
  helpSheet.appendRow(['']);
  
  helpSheet.appendRow(['▼ 設定項目の説明']);
  helpSheet.getRange(6, 1).setFontWeight('bold').setFontSize(12);
  
  helpSheet.appendRow(['項目名', '設定方法', '説明']);
  helpSheet.getRange(7, 1, 1, 3).setFontWeight('bold').setBackground('#d9e1f2');
  helpSheet.appendRow(['種別', '「送信元(From)」「件名(Subject)」「本文(Body)」から選択', 'メールのどの部分を検査するかを決定します。']);
  helpSheet.appendRow(['キーワード', 'メールアドレス、件名のキーワード、または本文のテキストを入力', '既読にしたい対象のテキストを入力します。']);
  helpSheet.appendRow(['一致方法', '「完全一致(Exact)」「部分一致(Contains)」「正規表現(RegExp)」から選択', 'キーワードがどのように一致するかを決定します。通常は「部分一致」が一番簡単でおすすめです。']);
  helpSheet.appendRow(['説明', '自由にテキストを入力', 'どのような目的で作成したルールかをメモとして残すことができます。']);
  helpSheet.appendRow(['有効にする', 'チェックボックスをオン/オフ', 'チェックを入れるとこのルールが動作し、チェックを外すと動作を一時停止します。']);
  
  helpSheet.appendRow(['']);
  helpSheet.appendRow(['▼ よく使われるフィルタ設定の具体例']);
  helpSheet.getRange(15, 1).setFontWeight('bold').setFontSize(12);
  
  helpSheet.appendRow(['目的', '種別', 'キーワード', '一致方法', '説明']);
  helpSheet.getRange(16, 1, 1, 5).setFontWeight('bold').setBackground('#f2f2f2');
  helpSheet.appendRow(['特定のメルマガを完全に既読化', '送信元(From)', 'info@nikki.ne.jp', '完全一致(Exact)', 'みんなのキャンパスなどの特定のアドレスを既読化']);
  helpSheet.appendRow(['件名に特定の文字があるものを既読化', '件名(Subject)', '【広告】', '部分一致(Contains)', '件名に【広告】とつくものをすべて既読化']);
  helpSheet.appendRow(['本文に「登録解除」があるものを既読化', '本文(Body)', '配信停止はこちら', '部分一致(Contains)', '本文に退会や配信停止の案内があるものを既読化']);
  helpSheet.appendRow(['特定のドメインを一括既読化', '送信元(From)', '.*@mail\\.axol\\.jp', '正規表現(RegExp)', 'axol.jp ドメインの就活メールを一括既読化（※上級者向け）']);
  helpSheet.appendRow(['大学の不要なイベントメールのみを既読化', '件名(Subject)', '(起業|キャンペーン|イベント|クイズ)', '正規表現(RegExp)', '大学の重要連絡（休講・補講）は残しつつ、不要なイベントのみを既読化（※上級者向け）']);
  
  helpSheet.getRange('A7:C12').setBorder(true, true, true, true, true, true);
  helpSheet.getRange('A16:E21').setBorder(true, true, true, true, true, true);
  helpSheet.autoResizeColumns(1, 5);
}

/**
 * 解析用にGmailの直近の未読メールをスプレッドシートに書き出す関数
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
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet('TempEmails');
  }
  
  sheet.appendRow(['日時', '送信元 (From)', '件名 (Subject)', '本文プレビュー (Body Preview)']);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d9ead3');
  
  const threads = GmailApp.search('is:unread', 0, 50);
  console.log(`未読メール ${threads.length} 件をスプレッドシートにエクスポートします。`);
  
  const data = [];
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      if (message.isUnread()) {
        const date = message.getDate();
        const from = message.getFrom();
        const subject = message.getSubject();
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
