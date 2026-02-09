const CONFIG_SHEET_NAME = '設定';
const AI_SHEET_NAME = 'AIキーワード抽出';

const COLUMN_NAMES = {
  STATUS: 'ステータス',
  SALE_DATE: '販売日',
  SALE_PLACE: '販売場所',
  SALE_PRICE: '販売価格',
  INCOME: '粗利',
  PROFIT: '利益',
  PROFIT_RATE: '利益率'
};

const ANALYSIS_HEADER_ROW = 15;

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const invMenu = ui.createMenu('棚卸')
    .addItem('今月を開始', 'startNewMonth')
    .addItem('今月に新規IDを同期', 'syncCurrentMonthIds')
    .addItem('最新月の理論を前月実地で再計算', 'recalcCurrentTheoryFromPrev');

  ui.createMenu('管理メニュー')
    .addItem('1. 回収完了リストを更新(抽出)', 'generateCompletionList')
    .addItem('2. 配布用リスト作成(チェック行を印刷・CSV用)', 'createBuyerSheet')
    .addItem('3. 売却反映(チェック行を一括処理)', 'processSelectedSales')
    .addSeparator()
    .addItem('ハイブランドソート', 'runHighBrandSort')
    .addSeparator()
    .addItem('マニュアル', 'showManual')
    .addItem('基本設定', 'showBasicSettings')
    .addItem('手数料設定', 'showFeeSettings')
    .addSeparator()
    .addSubMenu(invMenu)
    .addToUi();
}

function onEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== '回収完了') return;

  rc_handleRecoveryCompleteOnEdit_(e);

  if (e.range.getRow() === 4 && e.range.getColumn() === 2) {
    sortByField(sh);
  }
}


function generateCompletionList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('商品管理');
  var analysis = ss.getSheetByName('在庫分析');
  var out = ss.getSheetByName('回収完了');
  var aiSheet = ss.getSheetByName(AI_SHEET_NAME);
  var returnSheet = ss.getSheetByName('返送管理');

  ss.toast('リスト抽出を開始します...', '処理開始', 5);

  var daysThreshold = out.getRange('B1').getValue();
  var rawRateThreshold = out.getRange('B2').getValue();
  var rateThreshold = rawRateThreshold / 100;

  var lastRowAn = analysis.getLastRow();
  var anHeaders = analysis.getRange(ANALYSIS_HEADER_ROW, 1, 1, analysis.getLastColumn()).getValues()[0];
  var colIdIdx = anHeaders.indexOf('仕入れID');
  var colRateIdx = anHeaders.indexOf('回収割合');

  if (colIdIdx === -1 || colRateIdx === -1) {
    Browser.msgBox('エラー：「在庫分析」シートに「仕入れID」または「回収割合」の列が見つかりません。');
    return;
  }

  var anDataRange = analysis.getRange(ANALYSIS_HEADER_ROW + 1, 1, lastRowAn - ANALYSIS_HEADER_ROW, analysis.getLastColumn());
  var anData = anDataRange.getValues();
  var rateMap = {};
  anData.forEach(function (r) {
    var id = r[colIdIdx];
    var val = r[colRateIdx];
    var rate = 0;
    if (typeof val === 'number') {
      rate = val;
    } else if (typeof val === 'string') {
      rate = parseFloat(val.replace('%', '')) / 100;
    }
    if (id) rateMap[id] = rate;
  });

  var aiMap = {};
  if (aiSheet) {
    var aiData = aiSheet.getDataRange().getValues();
    var aiHeaders = aiData.shift();
    var aiIdIdx = aiHeaders.indexOf('管理番号');
    var keywordIndices = [];
    aiHeaders.forEach(function (h, i) {
      if (String(h).match(/キーワード|Keyword/)) keywordIndices.push(i);
    });

    if (aiIdIdx > -1 && keywordIndices.length > 0) {
      aiData.forEach(function (r) {
        var id = String(r[aiIdIdx]).trim();
        if (id) {
          var words = [];
          keywordIndices.forEach(function (idx) {
            var val = r[idx];
            if (val && String(val).trim() !== "") words.push(val);
          });
          aiMap[id] = words.join(' ');
        }
      });
    }
  }

  var boxMap = {};
  if (returnSheet) {
    var rData = returnSheet.getDataRange().getValues();
    for (var i = 1; i < rData.length; i++) {
      var row = rData[i];
      var boxId = row[0];
      var mgmtIdsStr = String(row[2]);

      if (mgmtIdsStr) {
        var ids = mgmtIdsStr.split(/[、,]/);
        ids.forEach(function (id) {
          boxMap[id.trim()] = boxId;
        });
      }
    }
  }

  var mData = main.getDataRange().getValues();
  var headers = mData.shift();
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  if (idx['ステータス'] === undefined) {
    Browser.msgBox('エラー：「商品管理」シートに「ステータス」列が見つかりません。ヘッダ名を確認してください。');
    return;
  }

  var outArr = [];
  var today = new Date();
  var msecPerDay = 24 * 60 * 60 * 1000;

  mData.forEach(function (r) {
    var status = String(r[idx['ステータス']] || '').trim();
    if (status === '売却済み') return;

    var ld = r[idx['出品日']];
    if (!ld || !(ld instanceof Date)) return;
    if (r[idx['販売日']] || r[idx['返品日']] || r[idx['キャンセル日']] || r[idx['廃棄日']]) return;

    var days = Math.floor((today - ld) / msecPerDay);
    if (days < daysThreshold) return;

    var id = r[idx['仕入れID']];
    var rate = rateMap[id] || 0;

    if (rate <= rateThreshold) return;

    var mgmtId = String(r[idx['管理番号']]).trim();
    var aiTitle = aiMap[mgmtId] || "";
    var boxId = boxMap[mgmtId] || "";

    outArr.push([
      false,
      boxId,
      mgmtId,
      r[idx['ブランド']],
      r[idx['メルカリサイズ']],
      r[idx['性別']],
      r[idx['カテゴリ2']],
      aiTitle,
      ld,
      r[idx['使用アカウント']],
      r[idx['仕入れ値']],
      r[idx['納品場所']],
      ""
    ]);
  });

  var maxRows = out.getMaxRows();
  if (maxRows >= 7) {
    var lastCol = out.getLastColumn() || 20;
    out.getRange(7, 1, maxRows - 6, lastCol).clearContent();
    out.getRange(7, 1, maxRows - 6, 1).removeCheckboxes();
  }

  var headerTitles = [
    '確認', '箱ID', '管理番号', 'ブランド', 'サイズ', '性別', 'カテゴリ', 'AIタイトル(KW1-8)', '出品日', 'アカウント', '仕入れ値', '納品場所', '【入力】まとめID'
  ];
  out.getRange(6, 1, 1, headerTitles.length).setValues([headerTitles]).setFontWeight('bold').setBackground('#f3f3f3');

  if (outArr.length > 0) {
    out.getRange(7, 1, outArr.length, outArr[0].length).setValues(outArr);
    out.getRange(7, 1, outArr.length, 1).insertCheckboxes();
    ss.toast(outArr.length + '件抽出完了', '完了', 5);
  } else {
    ss.toast('対象データはありませんでした', '完了', 5);
  }
}


function createBuyerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- モーダルで入力を受け取る ---
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body { font-family: sans-serif; padding: 10px; }' +
    'label { display: block; margin-top: 10px; font-weight: bold; }' +
    'input { width: 100%; padding: 6px; margin-top: 4px; box-sizing: border-box; }' +
    'button { margin-top: 16px; padding: 8px 20px; }' +
    '</style>' +
    '<label>会社名 / 氏名</label>' +
    '<input type="text" id="name" />' +
    '<label>受付番号</label>' +
    '<input type="text" id="number" />' +
    '<br><button onclick="submit()">実行</button>' +
    '<script>' +
    'function submit(){' +
    '  var name = document.getElementById("name").value;' +
    '  var number = document.getElementById("number").value;' +
    '  if(!name || !number){ alert("両方入力してください"); return; }' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); })' +
    '    .createBuyerSheetWithInput(name, number);' +
    '}' +
    '</script>'
  ).setWidth(350).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, '配布用リスト作成');
}

function createBuyerSheetWithInput(inputName, inputNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listSheet = ss.getSheetByName('回収完了');
  var mainSheet = ss.getSheetByName('商品管理');

  var exportSheetName = '配布用リスト';
  var exportSheet = ss.getSheetByName(exportSheetName);

  if (!exportSheet) {
    exportSheet = ss.insertSheet(exportSheetName);
  } else {
    var maxRows = exportSheet.getMaxRows();
    var maxCols = exportSheet.getMaxColumns();

    // --- 1行目はデータ反映セル(A1, B1, E1, I1)だけクリア、2行目以降は全クリア ---
    if (maxRows >= 2) {
      exportSheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
      exportSheet.getRange(2, 1, maxRows - 1, 1).removeCheckboxes();
    }
    // 1行目の対象セルだけクリア
    exportSheet.getRange('A1').clearContent();
    exportSheet.getRange('B1').clearContent();
    exportSheet.getRange('E1').clearContent();
    exportSheet.getRange('I1').clearContent();
  }

  var lastRow = listSheet.getLastRow();
  if (lastRow < 7) {
    ss.toast('リストが空です', 'エラー');
    return;
  }

  var listData = listSheet.getRange(7, 1, lastRow - 6, 13).getValues();
  var targetRows = listData.filter(function (r) { return r[0] === true; });

  if (targetRows.length === 0) {
    ss.toast('「回収完了」シートでチェックされた項目がありません。', '処理中断');
    return;
  }

  var summaryId = targetRows[0][12] || "";

  var mData = mainSheet.getDataRange().getValues();
  var headers = mData.shift();
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var headerRow = ['確認', '箱ID', '管理番号(照合用)', 'ブランド', 'AIタイトル候補', 'アイテム', 'サイズ', '状態', '傷汚れ詳細', '採寸情報', '即出品用説明文（コピペ用）'];
  var exportData = [headerRow];

  targetRows.forEach(function (listRow) {
    var boxId = listRow[1];
    var targetId = listRow[2];
    var aiTitle = listRow[7];

    if (targetId == "") return;

    var row = mData.find(function (r) { return r[idx['管理番号']] === targetId; });
    if (!row) return;

    var condition = row[idx['状態']] || '目立った傷や汚れなし';
    var damageDetail = row[idx['傷汚れ詳細']] || '';
    var brand = row[idx['ブランド']] || '';
    var size = row[idx['メルカリサイズ']] || '';
    var item = row[idx['カテゴリ2']] || '古着';
    if (!aiTitle) aiTitle = "";

    var length = row[idx['着丈']] || '-';
    var width = row[idx['身幅']] || '-';
    var shoulder = row[idx['肩幅']] || '-';
    var sleeve = row[idx['袖丈']] || '-';
    var waist = row[idx['ウエスト']];
    var rise = row[idx['股上']];
    var inseam = row[idx['股下']];

    var measurementText = "";
    if (length != '-' || width != '-') {
      measurementText += "着丈: " + length + " / 身幅: " + width + " / 肩幅: " + shoulder + " / 袖丈: " + sleeve + "\n";
    }
    if (waist) {
      measurementText += "ウエスト: " + waist + " / 股上: " + rise + " / 股下: " + inseam;
    }
    measurementText = measurementText.trim();

    var description =
      "【管理番号】\n" +
      "【ブランド】" + brand + "\n" +
      "【サイズ】" + size + "\n" +
      "【状態】" + condition + "\n";

    if (damageDetail !== "") {
      description += "【状態詳細】\n" + damageDetail + "\n";
    }

    description += "【実寸(cm)】\n" + measurementText + "\n" +
      "\n※素人採寸のため多少の誤差はご了承ください。";

    exportData.push([false, boxId, targetId, brand, aiTitle, item, size, condition, damageDetail, measurementText, description]);
  });

  // --- 1行目：A1にまとめID、B1に値、E1に会社名/氏名、I1に受付番号 ---
  exportSheet.getRange('A1').setValue('まとめID');
  exportSheet.getRange('B1').setValue(summaryId);
  exportSheet.getRange('E1').setValue(inputName);
  exportSheet.getRange('I1').setValue(inputNumber);

  // --- 2行目以降にデータ ---
  exportSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);

  if (exportData.length > 1) {
    exportSheet.getRange(3, 1, exportData.length - 1, 1).insertCheckboxes();
  }

  ss.toast(targetRows.length + '件を配布用リストへ反映しました', '成功', 3);
}

function processSelectedSales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('回収完了');
  var main = ss.getSheetByName('商品管理');

  var lastRow = sh.getLastRow();
  if (lastRow < 7) {
    ss.toast('データがありません', '終了');
    return;
  }

  ss.toast('ステータス＋まとめID反映と削除を開始します...', '処理中', 30);

  var headerRow = main.getRange(1, 1, 1, main.getLastColumn()).getValues()[0];
  var colMap = {};
  headerRow.forEach(function (name, i) {
    if (name) colMap[String(name).trim()] = i + 1;
  });

  var statusCol = colMap[COLUMN_NAMES.STATUS];
  if (!statusCol) {
    Browser.msgBox('エラー：ステータス列が見つかりません。「★列とシートの診断」を実行してください。');
    return;
  }

  var summaryCol = colMap['まとめID'] || 67;

  var idCol = colMap['管理番号'];
  if (!idCol) {
    Browser.msgBox('エラー：管理番号列が見つかりません。');
    return;
  }

  var mainLastRow = main.getLastRow();
  if (mainLastRow < 2) {
    ss.toast('商品管理にデータがありません', '終了', 5);
    return;
  }

  var mainIds = main.getRange(2, idCol, mainLastRow - 1, 1).getValues().flat();
  var idToRowMap = {};
  mainIds.forEach(function (id, index) {
    var k = String(id).trim();
    if (k !== '') idToRowMap[k] = index + 2;
  });

  var values = sh.getRange(7, 1, lastRow - 6, 17).getValues();

  var rowsToDelete = [];
  var uniqueRowSet = {};
  var statusRows = [];
  var summaryMap = {};

  for (var i = 0; i < values.length; i++) {
    var rowData = values[i];

    var isChecked = rowData[0] === true;
    if (!isChecked) continue;

    var id = String(rowData[2] || '').trim();
    if (id === '') continue;

    var tgtRow = idToRowMap[id];
    if (!tgtRow) continue;

    rowsToDelete.push(i + 7);

    if (!uniqueRowSet[tgtRow]) {
      uniqueRowSet[tgtRow] = true;
      statusRows.push(tgtRow);
    }

    var summaryId = rowData[12];
    if (summaryId !== '' && summaryId != null) {
      summaryMap[tgtRow] = summaryId;
    }
  }

  if (rowsToDelete.length === 0) {
    ss.toast('処理対象がありませんでした', '完了', 3);
    return;
  }

  var statusA1s = [];
  var statusColLetter = colNumToLetter_(statusCol);
  for (var a = 0; a < statusRows.length; a++) {
    statusA1s.push(statusColLetter + statusRows[a]);
  }
  main.getRangeList(statusA1s).setValue('売却済み');

  var summaryKeys = Object.keys(summaryMap);
  if (summaryKeys.length > 0) {
    var minRow = null;
    var maxRow = null;
    for (var b = 0; b < summaryKeys.length; b++) {
      var r = Number(summaryKeys[b]);
      if (minRow === null || r < minRow) minRow = r;
      if (maxRow === null || r > maxRow) maxRow = r;
    }

    var height = maxRow - minRow + 1;
    var rng = main.getRange(minRow, summaryCol, height, 1);
    var cur = rng.getValues();

    for (var c = 0; c < summaryKeys.length; c++) {
      var rr = Number(summaryKeys[c]);
      cur[rr - minRow][0] = summaryMap[rr];
    }

    rng.setValues(cur);
  }

  SpreadsheetApp.flush();

  rowsToDelete.sort(function (x, y) { return y - x; }).forEach(function (r) {
    sh.deleteRow(r);
  });

  ss.toast(rowsToDelete.length + '件を処理しました（売却済み＋まとめID反映＋回収完了から削除）', '処理完了', 5);

  // _colToA1Letter_ は Utils.gs の colNumToLetter_ に統合済み
}

function debugCheckColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('商品管理');
  var aiSheet = ss.getSheetByName(AI_SHEET_NAME);
  var analysis = ss.getSheetByName('在庫分析');

  var msg = '【診断レポート】\n\n';

  if (aiSheet) {
    msg += '✅ AIシート「' + AI_SHEET_NAME + '」発見\n';
  } else {
    msg += '❌ AIシート「' + AI_SHEET_NAME + '」が見つかりません\n';
  }

  if (analysis) {
    msg += '✅ 在庫分析シート発見\n';
    var h = analysis.getRange(15, 1, 1, analysis.getLastColumn()).getValues()[0];
    var colRateIdx = h.indexOf('回収割合');
    msg += '  - 回収割合: ' + (colRateIdx > -1 ? (colRateIdx + 1) + '列目' : '❌見つかりません(15行目を確認してください)') + '\n';
  } else {
    msg += '❌ 在庫分析シートが見つかりません\n';
  }

  msg += '\n✅ 商品管理シート列確認:\n';
  var headerRow = main.getRange(1, 1, 1, main.getLastColumn()).getValues()[0];
  var map = {};
  headerRow.forEach(function (n, i) { if (n) map[n.toString().trim()] = i + 1; });

  for (var k in COLUMN_NAMES) {
    var name = COLUMN_NAMES[k];
    var col = map[name];
    msg += '  - ' + name + ' : ' + (col ? col + '列目' : '❌見つかりません') + '\n';
  }

  Browser.msgBox(msg);
}

function sortByField(sheet) {
  var colMap = { '箱ID': 2, '管理番号': 3, 'ブランド': 4, 'サイズ': 5, '性別': 6, 'カテゴリ': 7 };
  var field = sheet.getRange('B4').getValue();
  var colIdx = colMap[field];
  var lastRow = sheet.getLastRow();
  if (colIdx && lastRow >= 7) {
    sheet.getRange(7, 1, lastRow - 6, sheet.getLastColumn()).sort({ column: colIdx, ascending: true });
  }
}

const BASIC_HEADER_ROW = 3;
const BASIC_START_COL = 1;
const BASIC_END_COL = 13;
const FEE_HEADER_ROW = 3;
const FEE_START_COL = 13;
const FEE_NUM_COLS = 3;

function showManual() { const html = HtmlService.createHtmlOutputFromFile('Manual').setTitle('マニュアル'); SpreadsheetApp.getUi().showSidebar(html); }
function showBasicSettings() { const html = HtmlService.createHtmlOutputFromFile('BasicSettings').setTitle('基本設定'); SpreadsheetApp.getUi().showSidebar(html); }
function showFeeSettings() { const html = HtmlService.createHtmlOutputFromFile('FeeSettings').setTitle('手数料設定'); SpreadsheetApp.getUi().showSidebar(html); }
function getBasicHeaders() { const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME); return sh.getRange(BASIC_HEADER_ROW, BASIC_START_COL, 1, BASIC_END_COL - BASIC_START_COL + 1).getValues()[0]; }
function getColumnData(colIndex1Based) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
  const targetCol = BASIC_START_COL + (colIndex1Based - 1);
  const startRow = BASIC_HEADER_ROW + 1;
  const lastRow = Math.max(sh.getLastRow(), startRow);
  const numRows = lastRow - startRow + 1;
  const raw = sh.getRange(startRow, targetCol, numRows, 1).getValues().map(r => r[0]);
  const values = raw.filter(v => String(v).trim() !== '');
  const header = sh.getRange(BASIC_HEADER_ROW, targetCol).getValue();
  return { index: colIndex1Based, header, values };
}
function saveColumn(colIndex1Based, newHeader, values) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
  const targetCol = BASIC_START_COL + (colIndex1Based - 1);
  sh.getRange(BASIC_HEADER_ROW, targetCol).setValue(newHeader || '');
  const startRow = BASIC_HEADER_ROW + 1;
  const data = (values || []).map(v => [v]);
  const maxNeeded = startRow + data.length - 1;
  if (maxNeeded > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), maxNeeded - sh.getMaxRows());
  const lastRow = Math.max(sh.getLastRow(), startRow);
  const clearRows = Math.max(lastRow - startRow + 1, 1);
  sh.getRange(startRow, targetCol, clearRows, 1).clearContent();
  if (data.length) sh.getRange(startRow, targetCol, data.length, 1).setValues(data);
  return 'OK';
}
function getFeeSettings() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
  const headerRange = sh.getRange(FEE_HEADER_ROW, FEE_START_COL, 1, FEE_NUM_COLS);
  if (!headerRange.getValues()[0][0]) headerRange.setValues([['販売場所名', '手数料率', '有効フラグ']]);
  const last = sh.getLastRow();
  const values = last > FEE_HEADER_ROW ? sh.getRange(FEE_HEADER_ROW + 1, FEE_START_COL, last - FEE_HEADER_ROW, FEE_NUM_COLS).getValues() : [];
  return values.filter(r => String(r[0] || '').trim() !== '');
}
function saveFeeSettings(rows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
  sh.getRange(FEE_HEADER_ROW, FEE_START_COL, 1, FEE_NUM_COLS).setValues([['販売場所名', '手数料率', '有効フラグ']]);
  const startRow = FEE_HEADER_ROW + 1;
  const clearRows = Math.max(sh.getMaxRows() - FEE_HEADER_ROW, 1);
  sh.getRange(startRow, FEE_START_COL, clearRows, FEE_NUM_COLS).clearContent();
  if (rows && rows.length) {
    const norm = rows.map(r => [r[0] || '', parseFloat(r[1]) || 0, r[2] === true || String(r[2]).toUpperCase() === 'TRUE']);
    sh.getRange(startRow, FEE_START_COL, norm.length, FEE_NUM_COLS).setValues(norm);
  }
  return 'OK';
}
function warmUp() { SpreadsheetApp.getActive(); return 'OK'; }

function runHighBrandSort() {
  const TARGET_SHEET_NAME = '回収完了';
  const AI_BRAND_SHEET_NAME = 'ブランドAI判定';
  const TARGET_START_ROW = 7;
  const TARGET_BRAND_COL = 4;
  const AI_DEFAULT_BRAND_COL = 1;
  const AI_DEFAULT_MEDIAN_COL = 5;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const aiSheet = ss.getSheetByName(AI_BRAND_SHEET_NAME);

  if (!targetSheet) throw new Error('回収完了 シートが見つかりません: ' + TARGET_SHEET_NAME);
  if (!aiSheet) throw new Error('ブランドAI判定 シートが見つかりません: ' + AI_BRAND_SHEET_NAME);

  const medianMap = buildMedianMap_(aiSheet, AI_DEFAULT_BRAND_COL, AI_DEFAULT_MEDIAN_COL);

  const lastRow = targetSheet.getLastRow();
  const lastCol = targetSheet.getLastColumn();
  if (lastRow < TARGET_START_ROW) return;

  const numRows = lastRow - TARGET_START_ROW + 1;

  targetSheet.insertColumnAfter(lastCol);
  const helperCol = lastCol + 1;

  const headerRow = TARGET_START_ROW - 1;
  if (headerRow >= 1) {
    targetSheet.getRange(headerRow, helperCol).setValue('AI_中央値キー');
  }

  const brandValues = targetSheet.getRange(TARGET_START_ROW, TARGET_BRAND_COL, numRows, 1).getValues();
  const helperValues = new Array(numRows);

  for (let i = 0; i < numRows; i++) {
    const brand = normalizeBrand_(brandValues[i][0]);
    const key = brand && medianMap.has(brand) ? medianMap.get(brand) : -1;
    helperValues[i] = [key];
  }

  targetSheet.getRange(TARGET_START_ROW, helperCol, numRows, 1).setValues(helperValues);

  const sortRange = targetSheet.getRange(TARGET_START_ROW, 1, numRows, helperCol);
  sortRange.sort([
    { column: helperCol, ascending: false },
    { column: TARGET_BRAND_COL, ascending: true }
  ]);

  targetSheet.deleteColumn(helperCol);
}

function buildMedianMap_(aiSheet, fallbackBrandCol, fallbackMedianCol) {
  const lastRow = aiSheet.getLastRow();
  const lastCol = aiSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return new Map();

  const header = aiSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let brandCol = findColByCandidates_(header, ['ブランド', 'brand', 'Brand', 'BRAND']);
  let medianCol = findColByCandidates_(header, ['中央値', '推定中央値', 'AI_推定中央値', 'AI_推定中央値(円)', 'median', 'Median']);

  if (!brandCol) brandCol = fallbackBrandCol;
  if (!medianCol) medianCol = fallbackMedianCol;

  const values = aiSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const brand = normalizeBrand_(row[brandCol - 1]);
    if (!brand) continue;

    const median = toNumber_(row[medianCol - 1]);
    if (median === null) continue;

    if (!map.has(brand) || median > map.get(brand)) {
      map.set(brand, median);
    }
  }

  return map;
}

// findHeaderCol_, toNumber_ は Utils.gs に統合済み
// normalizeBrand_ は trim のみのため Utils.normalizeText_ でカバー

function normalizeBrand_(v) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim();
  if (!s) return '';
  return s;
}
