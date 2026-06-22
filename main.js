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
  
  // ヘッダーのスタイル調整
  const headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange.setFontWeight('bold')
             .setBackground('#d9e1f2')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 28);
  
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
  sheet.getRange('B2').setValue('newsletter@example.com');
  sheet.getRange('C2').setValue('完全一致(Exact)');
  sheet.getRange('D2').setValue('特定のメルマガを完全に既読化');
  sheet.getRange('E2').setValue(true); // 最初から有効化しておく
  
  sheet.getRange('A3').setValue('件名(Subject)');
  sheet.getRange('B3').setValue('(イベント|キャンペーン|アンケート|広告)');
  sheet.getRange('C3').setValue('正規表現(RegExp)');
  sheet.getRange('D3').setValue('特定のキーワードを含む不要なメールのみを既読化');
  sheet.getRange('E3').setValue(true); // 最初から有効化しておく

  sheet.getRange('A4').setValue('送信元(From)');
  sheet.getRange('B4').setValue('.*@spam\\.example\\.net');
  sheet.getRange('C4').setValue('正規表現(RegExp)');
  sheet.getRange('D4').setValue('特定のドメインからの自動配信メールを一括既読化');
  sheet.getRange('E4').setValue(false); // 安全のため最初は無効化
  
  // スタイル適用（中央揃えなど）
  sheet.getRange('A2:A100').setHorizontalAlignment('center');
  sheet.getRange('C2:C100').setHorizontalAlignment('center');
  sheet.getRange('E2:E100').setHorizontalAlignment('center');
  
  // 列幅を明示的に指定して、プルダウン矢印が文字と重ならないようにする
  sheet.setColumnWidth(1, 150); // 種別
  sheet.setColumnWidth(2, 280); // キーワード
  sheet.setColumnWidth(3, 150); // 一致方法
  sheet.setColumnWidth(4, 280); // 説明
  sheet.setColumnWidth(5, 120); // 有効にする
  
  // ヘルプシートの作成
  createHelpSheet(spreadsheet);
}

/**
 * 使い方ヘルプシートを作成する
 */
function createHelpSheet(spreadsheet) {
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

  let helpSheet = spreadsheet.getSheetByName('使い方ヘルプ');
  if (helpSheet) {
    spreadsheet.deleteSheet(helpSheet); // 常に最新の状態に更新
  }
  helpSheet = spreadsheet.insertSheet('使い方ヘルプ');
  
  // タイトル行（セルの結合をして、A列幅に影響しないようにする）
  helpSheet.getRange('A1:E1').merge().setValue('★ Gmail自動既読システム 使い方ヘルプ ★')
           .setFontSize(16).setFontWeight('bold').setFontColor('#1f4e78')
           .setVerticalAlignment('middle').setHorizontalAlignment('left');
  helpSheet.setRowHeight(1, 35);
  
  // 説明文行（結合と折り返し）
  helpSheet.getRange('A2:E2').merge().setValue('');
  
  const descText = 'このシステムは、「FilterRules」シートで設定したルールに従って、Gmailの未読メールを自動で既読にします。\n' +
                   'AIに頼らなくても、以下のルール設定例を参考に「FilterRules」に自分で記入して管理することができます。';
  helpSheet.getRange('A3:E4').merge().setValue(descText)
           .setWrap(true).setVerticalAlignment('top').setFontSize(10).setFontColor('#333333');
  helpSheet.setRowHeight(3, 20);
  helpSheet.setRowHeight(4, 20);
  
  // セクション1: 設定項目の説明
  helpSheet.getRange('A6').setValue('▼ 設定項目の説明').setFontWeight('bold').setFontSize(12).setFontColor('#1f4e78');
  
  const headers1 = ['項目名', '設定方法', '説明'];
  for (let i = 0; i < headers1.length; i++) {
    helpSheet.getRange(7, i + 1).setValue(headers1[i]);
  }
  helpSheet.getRange('A7:C7').setFontWeight('bold').setBackground('#d9e1f2').setHorizontalAlignment('center');
  
  const items1 = [
    ['種別', '「送信元(From)」「件名(Subject)」「本文(Body)」から選択', 'メールのどの部分を検査するかを決定します。'],
    ['キーワード', 'メールアドレス、件名のキーワード、または本文のテキストを入力', '既読にしたい対象のテキストを入力します。'],
    ['一致方法', '「完全一致(Exact)」「部分一致(Contains)」「正規表現(RegExp)」から選択', 'キーワードがどのように一致するかを決定します。通常は「部分一致」が一番簡単でおすすめです。'],
    ['説明', '自由にテキストを入力', 'どのような目的で作成したルールかをメモとして残すことができます。'],
    ['有効にする', 'チェックボックスをオン/オフ', 'チェックを入れるとこのルールが動作し、チェックを外すと動作を一時停止します。']
  ];
  
  for (let r = 0; r < items1.length; r++) {
    for (let c = 0; c < items1[r].length; c++) {
      helpSheet.getRange(8 + r, c + 1).setValue(items1[r][c]);
    }
  }
  
  // 設定項目の説明テーブルに枠線を引く＆テキスト折り返し
  const descTableRange = helpSheet.getRange('A7:C12');
  descTableRange.setBorder(true, true, true, true, true, true, '#ccc', SpreadsheetApp.BorderStyle.SOLID);
  descTableRange.setWrap(true).setVerticalAlignment('middle');
  
  // セクション2: フィルタ設定の具体例
  helpSheet.getRange('A14').setValue('▼ よく使われるフィルタ設定の具体例').setFontWeight('bold').setFontSize(12).setFontColor('#1f4e78');
  
  const headers2 = ['目的', '種別', 'キーワード', '一致方法', '説明'];
  for (let i = 0; i < headers2.length; i++) {
    helpSheet.getRange(15, i + 1).setValue(headers2[i]);
  }
  helpSheet.getRange('A15:E15').setFontWeight('bold').setBackground('#f2f2f2').setHorizontalAlignment('center');
  
  const items2 = [
    ['特定のメルマガを完全に既読化', '送信元(From)', 'newsletter@example.com', '完全一致(Exact)', '特定のメルマガからのメールを完全に既読化'],
    ['件名に特定の文字があるものを既読化', '件名(Subject)', '【広告】', '部分一致(Contains)', '件名に【広告】とつくものをすべて既読化'],
    ['本文に特定の文字があるものを既読化', '本文(Body)', '配信停止はこちら', '部分一致(Contains)', '本文に退会や配信停止の案内があるものを既読化'],
    ['特定のドメインを一括既読化', '送信元(From)', '.*@spam\\.example\\.net', '正規表現(RegExp)', '特定のドメインからのメールを一括既読化（※上級者向け）'],
    ['特定のキーワードを含むメールのみを既読化', '件名(Subject)', '(イベント|キャンペーン|アンケート|広告)', '正規表現(RegExp)', '重要な連絡は残しつつ、特定のキーワードを含む不要なメールのみを既読化（※上級者向け）']
  ];
  
  for (let r = 0; r < items2.length; r++) {
    for (let c = 0; c < items2[r].length; c++) {
      helpSheet.getRange(16 + r, c + 1).setValue(items2[r][c]);
    }
  }
  
  // 具体例テーブルに枠線を引く＆テキスト折り返し
  const exampleTableRange = helpSheet.getRange('A15:E20');
  exampleTableRange.setBorder(true, true, true, true, true, true, '#ccc', SpreadsheetApp.BorderStyle.SOLID);
  exampleTableRange.setWrap(true).setVerticalAlignment('middle');
  
  // 表の各行の高さを少し広げて見やすくする
  for (let row = 7; row <= 12; row++) {
    helpSheet.setRowHeight(row, 24);
  }
  for (let row = 15; row <= 20; row++) {
    helpSheet.setRowHeight(row, 24);
  }
  
  // 列幅を明示的に指定して折り返しがきれいに収まるようにする
  helpSheet.setColumnWidth(1, 160); // 項目名 / 目的
  helpSheet.setColumnWidth(2, 140); // 設定方法 / 種別
  helpSheet.setColumnWidth(3, 260); // 説明 / キーワード
  helpSheet.setColumnWidth(4, 140); // 一致方法
  helpSheet.setColumnWidth(5, 260); // 説明
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
