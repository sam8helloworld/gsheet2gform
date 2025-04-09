/**
 * Googleフォームの回答送信時に実行されるトリガー関数
 * 
 * @param {GoogleAppsScript.Events.FormsOnSubmit} e - フォーム送信時のイベントオブジェクト
 */
function onSubmit(e) {
  const email = e.response.getRespondentEmail();
  const itemResponses = e.response.getItemResponses();
  let responseMap = {};
  for (const itemResponse of itemResponses) {
    const item = itemResponse.getItem();
    // 設問のタイトルと値を取得する
    responseMap[item.getTitle()] = itemResponse.getResponse();
  }
  // メール文面シートを見て設問の条件すべてが合う行を抽出する
  const row = retrieveRow(responseMap);
  // 該当なし
  if (Object.keys(row).length === 0) {
    // 抽出した行のtemplateと特定の列の値を取得
    const template = row["自動返信する文章"];

    // メールの設定
    let subject = "【診断結果】あなたが選ぶべき光回線はこれです"; //件名
    let body = 'htmlメールが表示できませんでした'; //body.htmlが表示できない時の予備
    //使用するhtmlファイルを指定
    var htmlTemplate = HtmlService.createTemplateFromFile("該当なし");
    const form = FormApp.getActiveForm();
    htmlTemplate.formPublishedUrl = form.getPublishedUrl();
    html = htmlTemplate.evaluate().getContent();

    let options = {
      // from: '~~~~~~~~~~', //送り元（必要に応じてカスタム）
      // cc: '~~~~~~~~~, ~~~~~~~~', //CCの設定（必要に応じてカスタム）
      htmlBody: html
    }

    ///メール送信
    MailApp.sendEmail(
      email,
      subject, 
      body,
      options
    );
  } else {
    // 抽出した行のtemplateと特定の列の値を取得
    const template = row["自動返信する文章"];

    // メールの設定
    let subject = "【診断結果】あなたが選ぶべき光回線はこれです"; //件名
    let body = 'htmlメールが表示できませんでした'; //body.htmlが表示できない時の予備
    //使用するhtmlファイルを指定
    var htmlTemplate = HtmlService.createTemplate(template);

    // `data` プロパティを手動で定義
    htmlTemplate.data = {};
    Object.keys(responseMap).map(key => {
      // 日本語のキーを持つオブジェクトを設定
      htmlTemplate.data[key] = responseMap[key];
    });
    


    /*********************************
     * 
     * 設問にはないが自動返信文章に挿入したい列がある場合はここに以下のフォーマットで追加
     * htmlTemplate.data["列名"] = row["列名"];
     * 
     *********************************/
    htmlTemplate.data["回線"] = row["回線"];


    /*********************************
    * 
    * X投稿文を変更したい時
    * htmlTemplate.data["x投稿文"] = `文章`;
    * 
    *********************************/
    htmlTemplate.data["x投稿文"] = `${responseMap["絶対に避けたい後悔ポイント"]}な私にぴったりの光回線は${row["回線"]}でした！
by回線診断ツール
#回線診断
#光回線診断
#診断結果
URL`;


    htmlTemplate.data["住まいのエリア（都道府県）"] = responseMap["住まいのエリア（都道府県）"];
    htmlTemplate.data["市区町村を選択"] = responseMap["市区町村を選択"];
    html = htmlTemplate.evaluate().getContent();

    let options = {
      // from: '~~~~~~~~~~', //送り元（必要に応じてカスタム）
      // cc: '~~~~~~~~~, ~~~~~~~~', //CCの設定（必要に応じてカスタム）
      htmlBody: html
    }

    ///メール送信
    MailApp.sendEmail(
      email,
      subject, 
      body,
      options
    );
  }
}

/**
 * トラブルシューティング用のデバッグ関数
 * 開発時はフォームに回答しなくてこの関数を実行してメールを送信する
 */
function debug() {
  const responseMap = {
    "居住タイプ": "一戸建て",
    "都道府県": "東京都",
    "スマホキャリア": "ドコモ",
    "絶対に避けたい後悔ポイント": "あとで“もっと安いサービスあったじゃん”って知るのが嫌だ"
  };

  // メール文面シートを見て設問の条件すべてが合う行を抽出する
  const row = retrieveRow(responseMap);
  // 抽出した行のtemplateと特定の列の値を取得
  const template = row["自動返信する文章"];
  
  // メールの設定
  let subject = "アンケート回答の御礼"; //件名
  let body = 'htmlメールが表示できませんでした'; //body.htmlが表示できない時の予備
  //使用するhtmlファイルを指定
  var htmlTemplate = HtmlService.createTemplate(template);
  // `data` プロパティを手動で定義
  htmlTemplate.data = {};
  Object.keys(responseMap).map(key => {
    // 日本語のキーを持つオブジェクトを設定
    htmlTemplate.data[key] = responseMap[key];
  });
  
  htmlTemplate.data["回線"] = row["回線"];
  htmlTemplate.data["住まいのエリア（都道府県）"] = row["住まいのエリア（都道府県）"];
  htmlTemplate.data["市区町村を選択"] = responseMap["市区町村を選択"];
  html = htmlTemplate.evaluate().getContent();

  let options = {
    // from: '~~~~~~~~~~', //送り元（必要に応じてカスタム）
    // cc: '~~~~~~~~~, ~~~~~~~~', //CCの設定（必要に応じてカスタム）
    htmlBody: html
  }

  ///メール送信
  GmailApp.sendEmail(
  "babi634@neko2.net",
   subject, 
   body,
   options
   );
}

/**
 * 回答のオブジェクトを受け取りスプレッドシートから該当する行を返す
 * @param {Object.<string, string>} responseMap - キーと値のペアを持つオブジェクト
 * @return {Object.<string, string>} - 処理後のオブジェクト
 */
function retrieveRow(responseMap) {
  const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID"))
  const dataSheet = sheet.getSheetByName("メール文面テンプレート"); // 設問があるシート名を指定
  const data = dataSheet.getDataRange().getValues();

  const headers = data[0]; // ヘッダー行を取得
  
  const filteredRows = [];
  
  // 各キーがどの列に対応しているかをマッピング
  const columnIndices = {};
  for (let i = 0; i < headers.length; i++) {
    if (responseMap.hasOwnProperty(headers[i])) {
      columnIndices[headers[i]] = i;
    }
  }

  // 2行目以降のデータを検索
  for (let i = 1; i < data.length; i++) {
    let match = true;

    for (const key in responseMap) {
      const colIndex = columnIndices[key];
      if (colIndex === undefined) continue;

      const cellValue = data[i][colIndex].toString(); // セルの値を文字列に変換
      const searchValue = responseMap[key];

      // カンマ区切りの値がある場合、部分一致を確認
      if (!cellValue.split(",").map(v => v.trim()).includes(searchValue)) {
        match = false;
        break;
      }
    }

    if (match) {
      // ヘッダーをキーとしてオブジェクト化
      const rowObject = {};
      for (let j = 0; j < headers.length; j++) {
        rowObject[headers[j]] = data[i][j];
      }
      filteredRows.push(rowObject);
      break;
    }
  }

  return filteredRows[0] ?? {};
}

