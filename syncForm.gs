/**
 * この関数を実行することでGoogleフォームにスプレッドシートの設問が同期されます
 */
function syncForm() {
  // Googleフォームの先頭に都道府県の選択肢を追加
  const form = FormApp.openById(PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_ID"))

  form.getItems().forEach(item => form.deleteItem(item)); // 既存の設問を削除
  const prefectures = getPrefectures();
  const prefectureForm = addPrefectureForm(form, prefectures);
  const sections = addMunicipalitiesForm(form, prefectures)


  // --- 都道府県の選択肢に「市区町村選択ページへのジャンプ」を設定 ---
  let prefectureChoices = prefectures.map(pref => 
      prefectureForm.createChoice(pref, sections[pref])
  );
  prefectureForm.setChoices(prefectureChoices);

  // --- すべての市区町村選択ページから「回答完了」へジャンプ ---
  Object.values(sections).forEach(section => {
    section.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  });
}

/**
 * Googleフォームに都道府県を選択する設問を同期します
 * @param {GoogleAppsScript.Forms.Form} form - formインスタンス
 * @param {string[]} prefectures - 都道府県のリスト
 * @return {GoogleAppsScript.Forms.ListItem} 追加された都道府県のリストアイテム
 */
function addPrefectureForm(form, prefectures) {
  const prefectureItem = form.addListItem()
      .setTitle("住まいのエリア（都道府県）")
      .setChoiceValues(prefectures)
      .setRequired(true);
    
  return prefectureItem;
}

/**
 * スプレッドシートから都道府県一覧を取得します
 * @return {string[]} 都道府県のリスト
 */
function getPrefectures() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("都道府県"); // シート名を設定
  const data = sheet.getDataRange().getValues(); // シートのデータを取得
  // 都道府県リスト（1行目）
  const prefectures = data[0].filter(pref => pref !== ""); // 空白セルを除外
  return prefectures;
}

/**
 * Googleフォームの都道府県毎のセクションを作成し、そのセクションに市区町村の設問とその他の全ての設問を同期する
 * @param {GoogleAppsScript.Forms.Form} form - formインスタンス
 * @param {string[]} prefectures - 都道府県のリスト
 * @return Object.<string, GoogleAppsScript.Forms.PageBreakItem> セクション
 */
function addMunicipalitiesForm(form, prefectures) {
  const regionMap = getPrefectureMap()
  let sections = {};
  // --- 都道府県ごとのページ（セクション）を作成 ---
  prefectures.forEach(pref => {
    let section = form.addPageBreakItem().setTitle(`${pref}の市区町村を選択`);
    let cityItem = form.addListItem()
      .setTitle("市区町村を選択")
      .setChoiceValues(regionMap[pref])
      .setRequired(true);

    addQuestionForm(form);
    sections[pref] = section;
  });
  return sections
}

/**
 * GoogleフォームにGoogleスプレッドシートの設問を1つずつ同期する
 * @param {GoogleAppsScript.Forms.Form} form - formインスタンス
 */
function addQuestionForm(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = sheet.getSheetByName("設問"); // 設問があるシート名を指定
  const data = dataSheet.getDataRange().getValues(); 

  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));
  const transposedData = transpose(data);

  // スプレッドシートの各行に対応する設問を作成
  transposedData.forEach((row, index) => {
    const questionTitle = row[0]; // 設問（1列目）
    const questionType = row[1]; // 設問タイプ（2列目）
    const mandatory = row[2]; // 必須かどうか（3列目）

    if(questionType == "記述式(短文)") {
      const item = form.addTextItem();
      item.setTitle(questionTitle).setRequired(mandatory == "必須" ? true : false);
    } else if(questionType == "段落") {
      const item = form.addParagraphTextItem();
      item.setTitle(questionTitle).setRequired(mandatory == "必須" ? true : false);
    } else if(questionType == "ラジオボタン") {
      const choices = row.slice(3).filter(cell => cell !== ""); // 回答（3列目以降）
      const item = form.addMultipleChoiceItem();
      item.setTitle(questionTitle).setChoiceValues(choices).setRequired(mandatory == "必須" ? true : false);
    } else if(questionType == "チェックボックス") {
      const choices = row.slice(3).filter(cell => cell !== ""); // 回答（3列目以降）
      const item = form.addCheckboxItem();
      item.setTitle(questionTitle).setChoiceValues(choices).setRequired(mandatory == "必須" ? true : false);
    } else if(questionType == "プルダウン") {
      const choices = row.slice(3).filter(cell => cell !== ""); // 回答（3列目以降）
      const item = form.addListItem();
      item.setTitle(questionTitle).setChoiceValues(choices).setRequired(mandatory == "必須" ? true : false);
    } else {
      Logger.log("設問タイプが定義されていません。")
    }
  });
}

/**
 * スプレッドシートから都道府県毎の市区町村一覧の全データを取得する
 * @return {'都道府県':string[]} { 都道府県: [市区町村] } のマップ
 */
function getPrefectureMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("都道府県"); // シート名を設定
  const data = sheet.getDataRange().getValues(); // シートのデータを取得

  // 都道府県リスト（1行目）
  const prefectures = data[0].filter(pref => pref !== ""); // 空白セルを除外

  // 市区町村リスト（都道府県ごと）
  let regionMap = {}; // { 都道府県: [市区町村] } のマップ
  for (let col = 0; col < prefectures.length; col++) {
    let cities = [];
    for (let row = 1; row < data.length; row++) {
      if (data[row][col]) {
        cities.push(data[row][col]);
      }
    }
    regionMap[prefectures[col]] = cities;
  }

  return regionMap;
}