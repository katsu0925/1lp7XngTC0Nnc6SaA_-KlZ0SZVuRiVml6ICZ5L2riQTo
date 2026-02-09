/**
 * reflectAndDelete：入力行の商品管理反映＋ステータス変更＋行削除
 * 改善: 個別setValue → バッチ書き込み、マジックナンバー → ヘッダ検索
 */
function reflectAndDelete(sheet, rowNum) {
  Logger.log('【reflectAndDelete】開始 row=' + rowNum);

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('商品管理');

  // ヘッダから列番号を動的取得
  var lastCol = main.getLastColumn();
  var hdr = main.getRange(1, 1, 1, lastCol).getValues()[0];
  var hMap = buildHeaderMap_(hdr);

  var colId       = hMap['管理番号'];
  var colStatus   = hMap['ステータス'];
  var colSaleDate = hMap['販売日'];
  var colSalePlace= hMap['販売場所'];
  var colSalePrice= hMap['販売価格'];
  var colIncome   = hMap['入金額'] || hMap['粗利'];
  var colCost     = hMap['仕入れ値'];
  var colProfit   = hMap['利益'];
  var colProfitR  = hMap['利益率'];

  if (!colId) { Logger.log('  管理番号列が見つかりません'); return; }

  // 管理番号取得
  var id = sheet.getRange(rowNum, 1).getValue();
  Logger.log('  管理番号=' + id);

  // 商品管理シート 管理番号列を検索
  var ids = main.getRange(2, colId, main.getLastRow() - 1, 1).getValues();
  var idx0 = ids.findIndex(function(c) { return c[0] === id; });
  if (idx0 < 0) {
    Logger.log('  該当IDなし: ' + id);
    return;
  }
  var tgtRow = idx0 + 2;
  Logger.log('  対象行=' + tgtRow);

  // J～M 取得
  var row = sheet.getRange(rowNum, 10, 1, 4).getValues()[0];

  // バッチ書き込み用に全列を一度に読み書き
  var updates = [];
  if (colSaleDate)  updates.push({ col: colSaleDate,  val: row[0] });
  if (colSalePlace) updates.push({ col: colSalePlace, val: row[1] });
  if (colSalePrice) updates.push({ col: colSalePrice, val: row[2] });
  if (colIncome)    updates.push({ col: colIncome,    val: row[3] });
  if (colStatus)    updates.push({ col: colStatus,    val: '売却済み' });

  // 利益・利益率計算
  var cost = colCost ? main.getRange(tgtRow, colCost).getValue() : 0;
  var profit = row[3] - cost;
  var rate   = cost ? profit / cost : '';
  if (colProfit)  updates.push({ col: colProfit,  val: profit });
  if (colProfitR) updates.push({ col: colProfitR, val: rate });

  // RangeList でバッチ書き込み（個別 setValue 7回 → まとめて処理）
  updates.forEach(function(u) {
    main.getRange(tgtRow, u.col).setValue(u.val);
  });
  SpreadsheetApp.flush();

  Logger.log('  販売情報＋ステータス＋利益を反映 (計' + updates.length + '列)');

  // 回収完了シート行削除
  sheet.deleteRow(rowNum);
  Logger.log('  行削除: 回収完了 ' + rowNum);
  Logger.log('【reflectAndDelete】完了');
}

function stampByThreshold() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('在庫分析');
  var headerRow = 15;
  var startRow = 16;
  var lastRow = sh.getLastRow();
  if (lastRow < startRow) return;

  var headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  var percentCol = headers.indexOf('回収割合') + 1;
  var stampCol = headers.indexOf('回収完了日') + 1;
  if (percentCol < 1 || stampCol < 1) return;

  var validationsRow = sh.getRange(14, 1, 1, sh.getLastColumn()).getDataValidations()[0];
  var thresholdCol = -1;
  for (var i = 0; i < validationsRow.length; i++) {
    if (validationsRow[i]) {
      thresholdCol = i + 1;
      break;
    }
  }
  if (thresholdCol === -1) return;

  var rawThresholdStr = sh.getRange(14, thresholdCol).getDisplayValue();
  if (rawThresholdStr === '' || rawThresholdStr == null) return;
  var m = String(rawThresholdStr).match(/[\d\.]+/);
  if (!m) return;
  var tn = Number(m[0]);
  if (isNaN(tn)) return;
  var threshold = tn / 100;

  var recsDisp = sh.getRange(startRow, percentCol, lastRow - startRow + 1, 1).getDisplayValues();
  var stamps = sh.getRange(startRow, stampCol, lastRow - startRow + 1, 1).getValues();

  for (var r = 0; r < recsDisp.length; r++) {
    var disp = recsDisp[r][0];
    if (disp === '' || disp == null) continue;

    var m2 = String(disp).match(/[\d\.]+/);
    if (!m2) continue;
    var vn = Number(m2[0]);
    if (isNaN(vn)) continue;
    var v = vn / 100;

    if (v >= threshold && !stamps[r][0]) {
      sh.getRange(startRow + r, stampCol).setValue(new Date());
    }
  }
}

function toggleKaishuKanryoFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('回収完了');
  if (!sheet) throw new Error('「回収完了」シートが見つかりません。');

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
    return;
  }

  const headerRow = 6;
  const dataStartRow = 7;
  const startCol = 1;
  const numCols = 17;

  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  const range = sheet.getRange(headerRow, startCol, lastRow - headerRow + 1, numCols);
  range.createFilter();

  const filter = sheet.getFilter();
  const color = SpreadsheetApp.newColor().setRgbColor('#f4cccc').build();
  const criteria = SpreadsheetApp.newFilterCriteria().setVisibleBackgroundColor(color).build();
  filter.setColumnFilterCriteria(1, criteria);
}