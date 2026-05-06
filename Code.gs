// ============================================================
// 西町GP前店 在庫管理システム - Google Apps Script
// ============================================================
//
// 【シート構成】
//   商品マスタ   … 全商品の基本情報（発注先・単価など）
//   仕入れログ   … 仕入れフォームから自動記録
//   棚卸_YYYYMM  … 月次在庫シート（GASが自動生成）
//
// 【2つのフォーム】
//   仕入れフォーム … 納品のたびに入力（都度）
//   棚卸フォーム   … 月末に実棚数を商品ごとに入力→現在庫へ自動反映
//
// 【月次シートの計算ロジック】
//   現在庫 … 棚卸フォームから直接入力（実際に数えた数）
//   使用数 … 繰越在庫 + 当月仕入れ - 現在庫（自動計算）
// ============================================================

// ============================================================
// 定数定義
// ============================================================

// ★ スプレッドシートのIDをここに貼り付けてください
// URLの https://docs.google.com/spreadsheets/d/【ここ】/edit の部分
const SPREADSHEET_ID = '17ryGvEfavGDPgOq2Coobhjob3O3NDoYkDAVWSW3cQOs';

function _ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const SHEET = {
  MASTER:   '商品マスタ',
  PURCHASE: '仕入れログ',
};

// 商品マスタの列インデックス（1始まり）
const MASTER_COL = {
  SUPPLIER: 1,  // 発注先
  CATEGORY: 2,  // カテゴリ
  TYPE:     3,  // 商品種別
  NAME:     4,  // 商品名
  SPEC:     5,  // 規格・サイズ
  LOT:      6,  // 入数/1LOT
  UNIT:     7,  // 単位
  PRICE:    8,  // 単価(税抜)
  TAX_TYPE: 9,  // 消費税区分（標準10%/軽8%）
};

// 仕入れログの列インデックス（1始まり）
const PURCHASE_COL = {
  TIMESTAMP: 1,  // タイムスタンプ
  DATE:      2,  // 納品日
  SLIP_NO:   3,  // 伝票No.
  NAME:      4,  // 商品名
  QTY:       5,  // 数量(LOT)
  PRICE:     6,  // 単価(税抜) ※マスタから自動セット
  AMOUNT:    7,  // 金額(税抜) ※自動計算
  TAX_TYPE:  8,  // 消費税区分 ※マスタから自動セット
  NOTE:      9,  // 備考
};

// 月次棚卸シートの列インデックス（1始まり）
const MONTHLY_COL = {
  SUPPLIER:   1,  // 発注先
  CATEGORY:   2,  // カテゴリ
  TYPE:       3,  // 商品種別
  NAME:       4,  // 商品名
  SPEC:       5,  // 規格・サイズ
  LOT:        6,  // 入数/1LOT
  UNIT:       7,  // 単位
  PRICE:      8,  // 単価(税抜)
  CARRY:      9,  // 繰越在庫（前月末から自動引継ぎ）
  PURCHASE:  10,  // 当月仕入れ（仕入れログからSUMIF）
  STOCK:     11,  // 現在庫（棚卸フォームから入力）← 直接入力
  USED:      12,  // 使用数 = 繰越 + 仕入れ - 現在庫（自動計算）
  AMOUNT_10: 13,  // 在庫金額(税10%)（自動計算）
  AMOUNT_8:  14,  // 在庫金額(税8%)（自動計算）
  NOTE:      15,  // 備考
};

// ============================================================
// 仕入れフォーム送信トリガー（都度入力）
// ============================================================

function onPurchaseFormSubmit(e) {
  const ss           = _ss();
  const purchaseSheet = ss.getSheetByName(SHEET.PURCHASE);
  const masterSheet   = ss.getSheetByName(SHEET.MASTER);

  const response    = e.namedValues;
  const productName = response['商品名'][0];
  const qty         = Number(response['数量(LOT)'][0]);
  const delivDate   = response['納品日'][0];
  const slipNo      = response['伝票No.'][0] || '';
  const note        = response['備考'][0]     || '';

  // マスタから単価・税区分を取得
  const masterData = masterSheet.getDataRange().getValues();
  let price = 0, taxType = '';
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][MASTER_COL.NAME - 1] === productName) {
      price   = masterData[i][MASTER_COL.PRICE    - 1];
      taxType = masterData[i][MASTER_COL.TAX_TYPE - 1];
      break;
    }
  }

  purchaseSheet.appendRow([
    new Date(), delivDate, slipNo,
    productName, qty, price, price * qty, taxType, note,
  ]);

  SpreadsheetApp.flush();
}

// ============================================================
// 棚卸フォーム送信トリガー（月次・現在庫を直接入力）
// ============================================================

function onInventoryFormSubmit(e) {
  const ss       = _ss();
  const response = e.namedValues;

  const productName = response['商品名'][0];
  const stock       = Number(response['現在庫数量'][0]);
  const note        = response['備考'][0] || '';

  // 対象月のシートを特定（フォームに「対象月」がある場合はそちら優先）
  let targetSheet = null;
  if (response['対象月'] && response['対象月'][0]) {
    // 例: "2026年5月" → 棚卸_202605
    const raw = response['対象月'][0];
    const m   = raw.match(/(\d{4})年(\d{1,2})月/);
    if (m) {
      const sheetName = `棚卸_${m[1]}${String(m[2]).padStart(2, '0')}`;
      targetSheet = ss.getSheetByName(sheetName);
    }
  }
  // 対象月が取れない場合は当月シートにフォールバック
  if (!targetSheet) {
    const today = new Date();
    const sheetName = `棚卸_${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}`;
    targetSheet = ss.getSheetByName(sheetName);
  }

  if (!targetSheet) {
    // シートが存在しない場合はエラーログを残して終了
    Logger.log(`棚卸シートが見つかりません: 商品=${productName}`);
    return;
  }

  // 商品名で行を検索して現在庫・備考を更新
  const data = targetSheet.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    if (data[i][MONTHLY_COL.NAME - 1] === productName) {
      targetSheet.getRange(i + 1, MONTHLY_COL.STOCK).setValue(stock);
      if (note) {
        targetSheet.getRange(i + 1, MONTHLY_COL.NOTE).setValue(note);
      }
      break;
    }
  }

  SpreadsheetApp.flush();
}

// ============================================================
// 月次棚卸シート生成
// ============================================================

function createNextMonthSheet() {
  const today     = new Date();
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  _createMonthlySheet(_ss(), nextMonth);
}

function createThisMonthSheet() {
  const today = new Date();
  _createMonthlySheet(_ss(), today);
}

function _createMonthlySheet(ss, targetDate) {
  const year      = targetDate.getFullYear();
  const month     = targetDate.getMonth() + 1;
  const sheetName = `棚卸_${year}${String(month).padStart(2, '0')}`;

  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert(`${sheetName} はすでに存在します。`);
    return;
  }

  const masterSheet = ss.getSheetByName(SHEET.MASTER);
  const masterData  = masterSheet.getDataRange().getValues();

  // 前月在庫マップを作成
  const prevDate      = new Date(year, month - 2, 1);
  const prevSheetName = `棚卸_${prevDate.getFullYear()}${String(prevDate.getMonth() + 1).padStart(2, '0')}`;
  const prevSheet     = ss.getSheetByName(prevSheetName);
  const prevStockMap  = {};
  if (prevSheet) {
    const prevData = prevSheet.getDataRange().getValues();
    for (let i = 2; i < prevData.length; i++) {
      const name  = prevData[i][MONTHLY_COL.NAME  - 1];
      const stock = prevData[i][MONTHLY_COL.STOCK - 1];
      if (name) prevStockMap[name] = (stock === '' || stock === null) ? 0 : Number(stock);
    }
  }

  const newSheet = ss.insertSheet(sheetName);

  // タイトル行
  newSheet.getRange(1, 1).setValue(`西町GP前店　棚卸在庫表　${year}年${month}月`);
  newSheet.getRange(1, 1, 1, 15).merge();

  // ヘッダー行
  newSheet.getRange(2, 1, 1, 15).setValues([[
    '発注先', 'カテゴリ', '商品種別', '商品名', '規格・サイズ',
    '入数/1LOT', '単位', '単価(税抜)',
    '繰越在庫', '当月仕入れ', '現在庫', '使用数',
    '在庫金額(税10%)', '在庫金額(税8%)', '備考',
  ]]);

  // データ行を準備
  const rows = [];
  const taxTypes = [];
  for (let i = 1; i < masterData.length; i++) {
    const row  = masterData[i];
    const name = row[MASTER_COL.NAME - 1];
    if (!name) continue;

    const carry = prevStockMap[name] !== undefined ? prevStockMap[name] : 0;
    taxTypes.push(row[MASTER_COL.TAX_TYPE - 1]);
    rows.push([
      row[MASTER_COL.SUPPLIER - 1],
      row[MASTER_COL.CATEGORY - 1],
      row[MASTER_COL.TYPE     - 1],
      name,
      row[MASTER_COL.SPEC - 1],
      row[MASTER_COL.LOT  - 1],
      row[MASTER_COL.UNIT - 1],
      row[MASTER_COL.PRICE - 1],
      carry,  // 繰越在庫
      '',     // 当月仕入れ（SUMIF数式を後でセット）
      '',     // 現在庫（棚卸フォームから入力）
      '',     // 使用数（自動計算）
      '',     // 在庫金額(税10%)
      '',     // 在庫金額(税8%)
      '',     // 備考
    ]);
  }

  if (rows.length > 0) {
    newSheet.getRange(3, 1, rows.length, 15).setValues(rows);

    // 数式をセット
    for (let r = 0; r < rows.length; r++) {
      const sr = r + 3;  // シート行番号

      // 当月仕入れ: 仕入れログをSUMIF集計
      newSheet.getRange(sr, MONTHLY_COL.PURCHASE).setFormula(
        `=IFERROR(SUMIF(仕入れログ!$D:$D,$D${sr},仕入れログ!$E:$E),0)`
      );

      // 使用数 = 繰越 + 仕入れ - 現在庫（現在庫が未入力なら空白）
      newSheet.getRange(sr, MONTHLY_COL.USED).setFormula(
        `=IF(K${sr}="","",I${sr}+J${sr}-K${sr})`
      );

      // 在庫金額（税区分で振り分け）
      if (taxTypes[r] === '軽8%') {
        newSheet.getRange(sr, MONTHLY_COL.AMOUNT_10).setValue(0);
        newSheet.getRange(sr, MONTHLY_COL.AMOUNT_8).setFormula(
          `=IF(K${sr}="",0,K${sr}*H${sr})`
        );
      } else {
        newSheet.getRange(sr, MONTHLY_COL.AMOUNT_10).setFormula(
          `=IF(K${sr}="",0,K${sr}*H${sr})`
        );
        newSheet.getRange(sr, MONTHLY_COL.AMOUNT_8).setValue(0);
      }
    }

    // 合計行
    const totalRow = rows.length + 3;
    newSheet.getRange(totalRow, MONTHLY_COL.NAME).setValue('合計');
    newSheet.getRange(totalRow, MONTHLY_COL.AMOUNT_10).setFormula(`=SUM(M3:M${totalRow - 1})`);
    newSheet.getRange(totalRow, MONTHLY_COL.AMOUNT_8).setFormula(`=SUM(N3:N${totalRow - 1})`);
  }

  _applyMonthlySheetFormat(newSheet, rows.length);
  SpreadsheetApp.getUi().alert(`${sheetName} を作成しました。`);
}

// ============================================================
// 書式設定
// ============================================================

function _applyMonthlySheetFormat(sheet, dataRows) {
  const lastRow = dataRows + 3;

  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  sheet.getRange(2, 1, 1, 15)
    .setBackground('#4A90D9').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');

  for (let r = 3; r < lastRow; r++) {
    sheet.getRange(r, 1, 1, 15).setBackground(r % 2 === 0 ? '#F0F7FF' : '#FFFFFF');
  }

  // 現在庫列（入力セル）を黄色で強調
  if (dataRows > 0) {
    sheet.getRange(3, MONTHLY_COL.STOCK, dataRows).setBackground('#FFF9C4');
  }

  sheet.getRange(lastRow, 1, 1, 15).setBackground('#D9EAD3').setFontWeight('bold');

  const colWidths = [90, 80, 90, 150, 180, 80, 50, 90, 80, 90, 90, 80, 120, 120, 150];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  if (dataRows > 0) {
    sheet.getRange(3, MONTHLY_COL.PRICE,     dataRows).setNumberFormat('#,##0');
    sheet.getRange(3, MONTHLY_COL.CARRY,     dataRows).setNumberFormat('#,##0.##');
    sheet.getRange(3, MONTHLY_COL.PURCHASE,  dataRows).setNumberFormat('#,##0.##');
    sheet.getRange(3, MONTHLY_COL.STOCK,     dataRows).setNumberFormat('#,##0.##');
    sheet.getRange(3, MONTHLY_COL.USED,      dataRows).setNumberFormat('#,##0.##');
    sheet.getRange(3, MONTHLY_COL.AMOUNT_10, dataRows).setNumberFormat('#,##0');
    sheet.getRange(3, MONTHLY_COL.AMOUNT_8,  dataRows).setNumberFormat('#,##0');
  }

  sheet.getRange(2, 1, dataRows + 2, 15).setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(4);
}

// ============================================================
// フォーム作成（共通ヘルパー）
// ============================================================

function _getProductNames(ss) {
  const masterSheet = ss.getSheetByName(SHEET.MASTER);
  if (!masterSheet) return [];
  const data  = masterSheet.getDataRange().getValues();
  const names = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][MASTER_COL.NAME - 1];
    if (name) names.push(String(name));
  }
  return names;
}

function _openFormById(editUrl) {
  const match = editUrl.match(/\/d\/([a-zA-Z0-9-_]+)\//);
  return match ? FormApp.openById(match[1]) : null;
}

function _updateFormProductList(form, productNames) {
  const items = form.getItems(FormApp.ItemType.LIST);
  for (const item of items) {
    if (item.getTitle() === '商品名') {
      item.asListItem().setChoiceValues(productNames);
      return;
    }
  }
}

// ============================================================
// 仕入れフォーム作成
// ============================================================

function createPurchaseForm() {
  const ss           = _ss();
  const productNames = _getProductNames(ss);
  if (productNames.length === 0) {
    SpreadsheetApp.getUi().alert('商品マスタに商品が登録されていません。');
    return;
  }

  const form = FormApp.create('西町GP前店 仕入れ記録フォーム');
  form.setDescription('商品が納品されたときに入力してください。');
  form.setConfirmationMessage('記録しました。ありがとうございます。');
  form.setCollectEmail(false);

  form.addDateItem().setTitle('納品日').setRequired(true);
  form.addTextItem().setTitle('伝票No.').setRequired(false);
  form.addListItem().setTitle('商品名').setRequired(true).setChoiceValues(productNames);
  form.addTextItem()
    .setTitle('数量(LOT)')
    .setHelpText('例: 1、0.5')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumber().build());
  form.addParagraphTextItem().setTitle('備考').setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === 'onPurchaseFormSubmit')) {
    ScriptApp.newTrigger('onPurchaseFormSubmit').forForm(form).onFormSubmit().create();
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('PURCHASE_FORM_URL',      form.getPublishedUrl());
  props.setProperty('PURCHASE_FORM_EDIT_URL', form.getEditUrl());

  SpreadsheetApp.getUi().alert(
    '仕入れフォームを作成しました！\n\n' +
    '【スタッフ共有用URL】\n' + form.getPublishedUrl() + '\n\n' +
    'このURLをLINEでスタッフに共有してください。'
  );
}

// ============================================================
// 棚卸フォーム作成（月次・現在庫入力用）
// ============================================================

function createInventoryForm() {
  const ss           = _ss();
  const productNames = _getProductNames(ss);
  if (productNames.length === 0) {
    SpreadsheetApp.getUi().alert('商品マスタに商品が登録されていません。');
    return;
  }

  // 対象月の選択肢を生成（今月・来月・先月の3択）
  const today      = new Date();
  const monthLabels = [];
  for (let delta = -1; delta <= 1; delta++) {
    const d = new Date(today.getFullYear(), today.getMonth() + delta, 1);
    monthLabels.push(`${d.getFullYear()}年${d.getMonth() + 1}月`);
  }

  const form = FormApp.create('西町GP前店 棚卸入力フォーム');
  form.setDescription(
    '月末棚卸のときに使います。\n' +
    '商品を一つ選んで、実際に数えた在庫数を入力してください。\n' +
    '全商品ぶん繰り返し入力してください。'
  );
  form.setConfirmationMessage('記録しました。次の商品を入力する場合はフォームに戻ってください。');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(true);  // 入力ミスを修正できるようにする

  // 対象月
  form.addListItem()
    .setTitle('対象月')
    .setRequired(true)
    .setChoiceValues(monthLabels);

  // 商品名
  form.addListItem()
    .setTitle('商品名')
    .setRequired(true)
    .setChoiceValues(productNames);

  // 現在庫数量
  form.addTextItem()
    .setTitle('現在庫数量')
    .setHelpText('実際に数えた在庫数を入力（例: 3、0.5）')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build());

  // 備考
  form.addParagraphTextItem()
    .setTitle('備考')
    .setHelpText('メモがあれば（任意）')
    .setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === 'onInventoryFormSubmit')) {
    ScriptApp.newTrigger('onInventoryFormSubmit').forForm(form).onFormSubmit().create();
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('INVENTORY_FORM_URL',      form.getPublishedUrl());
  props.setProperty('INVENTORY_FORM_EDIT_URL', form.getEditUrl());

  SpreadsheetApp.getUi().alert(
    '棚卸フォームを作成しました！\n\n' +
    '【スタッフ共有用URL】\n' + form.getPublishedUrl() + '\n\n' +
    '月末棚卸のときにスタッフがスマホからこのURLにアクセスして\n' +
    '商品ごとに現在庫数を入力します。'
  );
}

// ============================================================
// フォームの商品リスト同期（マスタ更新後に実行）
// ============================================================

function syncFormProducts() {
  const ss           = _ss();
  const productNames = _getProductNames(ss);
  const props        = PropertiesService.getScriptProperties();
  let updated        = 0;

  const purchaseEditUrl  = props.getProperty('PURCHASE_FORM_EDIT_URL');
  const inventoryEditUrl = props.getProperty('INVENTORY_FORM_EDIT_URL');

  if (purchaseEditUrl) {
    const f = _openFormById(purchaseEditUrl);
    if (f) { _updateFormProductList(f, productNames); updated++; }
  }
  if (inventoryEditUrl) {
    const f = _openFormById(inventoryEditUrl);
    if (f) { _updateFormProductList(f, productNames); updated++; }
  }

  if (updated === 0) {
    SpreadsheetApp.getUi().alert('フォームが見つかりません。先にフォームを作成してください。');
  } else {
    SpreadsheetApp.getUi().alert(`${updated}件のフォームの商品リストを ${productNames.length} 件に更新しました。`);
  }
}

// ============================================================
// 初期セットアップ
// ============================================================

function setup() {
  const ss = _ss();
  _setupMasterSheet(ss);
  _setupPurchaseLogSheet(ss);

  SpreadsheetApp.getUi().alert(
    'セットアップ完了！\n\n次のステップ:\n' +
    '1. 商品マスタ.csv をインポート\n' +
    '2. 「仕入れフォームを作成」を実行\n' +
    '3. 「棚卸フォームを作成」を実行\n' +
    '4. 「今月の棚卸シートを作成」を実行'
  );
}

function _setupMasterSheet(ss) {
  let sheet = ss.getSheetByName(SHEET.MASTER);
  if (!sheet) sheet = ss.insertSheet(SHEET.MASTER);
  sheet.clearContents();
  const headers = ['発注先','カテゴリ','商品種別','商品名','規格・サイズ','入数/1LOT','単位','単価(税抜)','消費税区分(標準10%/軽8%)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  [4, 5, 9].forEach(c => sheet.setColumnWidth(c, c === 4 ? 150 : 180));
}

function _setupPurchaseLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET.PURCHASE);
  if (!sheet) sheet = ss.insertSheet(SHEET.PURCHASE);
  sheet.clearContents();
  const headers = ['タイムスタンプ','納品日','伝票No.','商品名','数量(LOT)','単価(税抜)','金額(税抜)','消費税区分','備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(4, 150);
}

// ============================================================
// 手動再集計
// ============================================================

function refreshCurrentMonth() {
  const today     = new Date();
  const sheetName = `棚卸_${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}`;
  const sheet     = _ss().getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`${sheetName} が見つかりません。`);
    return;
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('再集計しました。');
}

// ============================================================
// GASウェブアプリ（棚卸一括入力画面）
// ============================================================

// ウェブアプリのエントリーポイント
function doGet(e) {
  const action = e.parameter && e.parameter.action;

  // APIリクエスト → JSON返却（認証不要）
  if (action === 'getData') {
    return ContentService
      .createTextOutput(JSON.stringify(getInventoryData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // HTML画面 → ContentServiceで返すことでGoogleの認証をバイパス
  const url  = ScriptApp.getService().getUrl();
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.scriptUrl = url;
  const html = tmpl.evaluate().getContent();
  return ContentService
    .createTextOutput(html)
    .setMimeType(ContentService.MimeType.HTML);
}

// 在庫保存（POSTリクエスト）
function doPost(e) {
  const data   = JSON.parse(e.postData.contents);
  const result = saveInventory(data.entries);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 棚卸シートの商品データをウェブアプリに渡す
// 戻り値: { month, categories: [{ name, products: [{ name, spec, unit, carry, purchase, stock }] }] }
function getInventoryData() {
  const ss    = _ss();
  const today = new Date();

  // 当月シートを探す（なければnullを返す）
  const sheetName    = `棚卸_${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}`;
  const monthlySheet = ss.getSheetByName(sheetName);

  if (!monthlySheet) {
    return { error: `${sheetName} が見つかりません。先にメニューから「今月の棚卸シートを作成」を実行してください。` };
  }

  const data       = monthlySheet.getDataRange().getValues();
  const categoryMap = {};

  for (let i = 2; i < data.length; i++) {
    const row      = data[i];
    const name     = row[MONTHLY_COL.NAME     - 1];
    const category = row[MONTHLY_COL.CATEGORY - 1];
    if (!name || name === '合計') continue;

    const cat = category || 'その他';
    if (!categoryMap[cat]) categoryMap[cat] = [];
    categoryMap[cat].push({
      name:     String(name),
      spec:     String(row[MONTHLY_COL.SPEC     - 1] || ''),
      unit:     String(row[MONTHLY_COL.UNIT     - 1] || ''),
      lot:      String(row[MONTHLY_COL.LOT      - 1] || ''),
      carry:    Number(row[MONTHLY_COL.CARRY    - 1] || 0),
      purchase: Number(row[MONTHLY_COL.PURCHASE - 1] || 0),
      stock:    row[MONTHLY_COL.STOCK - 1] === '' ? null : Number(row[MONTHLY_COL.STOCK - 1]),
    });
  }

  // カテゴリを固定順に並べる
  const order      = ['フード', 'ドリンク', 'お土産', '備品', 'その他'];
  const categories = order
    .filter(c => categoryMap[c])
    .map(c => ({ name: c, products: categoryMap[c] }));

  // 定義外カテゴリがあれば末尾に追加
  Object.keys(categoryMap).forEach(c => {
    if (!order.includes(c)) categories.push({ name: c, products: categoryMap[c] });
  });

  return {
    month:      `${today.getFullYear()}年${today.getMonth() + 1}月`,
    sheetName,
    categories,
  };
}

// ウェブアプリからの一括送信を受け取って棚卸シートに書き込む
// entries: [{ name, stock }]
function saveInventory(entries) {
  const ss    = _ss();
  const today = new Date();
  const sheetName    = `棚卸_${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}`;
  const monthlySheet = ss.getSheetByName(sheetName);

  if (!monthlySheet) {
    return { success: false, message: `${sheetName} が見つかりません。` };
  }

  const data    = monthlySheet.getDataRange().getValues();
  // 商品名→行番号のマップを作成（高速検索）
  const rowMap  = {};
  for (let i = 2; i < data.length; i++) {
    const name = data[i][MONTHLY_COL.NAME - 1];
    if (name) rowMap[String(name)] = i + 1; // 1-indexed
  }

  let updated = 0;
  for (const entry of entries) {
    if (entry.stock === null || entry.stock === undefined || entry.stock === '') continue;
    const rowNum = rowMap[entry.name];
    if (!rowNum) continue;
    monthlySheet.getRange(rowNum, MONTHLY_COL.STOCK).setValue(Number(entry.stock));
    updated++;
  }

  SpreadsheetApp.flush();
  return { success: true, message: `${updated}件の在庫数を更新しました。` };
}

// ============================================================
// カスタムメニュー
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在庫管理')
    .addItem('🆕 初期セットアップ', 'setup')
    .addSeparator()
    .addItem('📋 仕入れフォームを作成', 'createPurchaseForm')
    .addItem('📦 棚卸入力アプリを開く',  'openInventoryApp')
    .addItem('🔁 フォームの商品リストを同期', 'syncFormProducts')
    .addSeparator()
    .addItem('📅 今月の棚卸シートを作成', 'createThisMonthSheet')
    .addItem('📅 翌月の棚卸シートを作成', 'createNextMonthSheet')
    .addItem('🔄 今月の仕入れ合計を再集計', 'refreshCurrentMonth')
    .addToUi();
}

// ウェブアプリのURLをダイアログに表示する
function openInventoryApp() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'ウェブアプリがまだデプロイされていません。\n\n' +
      'Apps Script エディタ →「デプロイ」→「新しいデプロイ」→\n' +
      '種類:ウェブアプリ / アクセス:自分のGoogleアカウント または 全員\n' +
      'でデプロイしてください。'
    );
    return;
  }
  SpreadsheetApp.getUi().alert('棚卸入力アプリのURL:\n\n' + url + '\n\nこのURLをスタッフに共有してください。');
}
