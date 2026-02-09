/**
 * reflectAndDelete：入力行の商品管理反映＋ステータス変更＋行削除
 */
function reflectAndDelete(sheet, rowNum) {
  Logger.log('【reflectAndDelete】開始 row=' + rowNum);
  
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName('商品管理');
  
  // 管理番号取得
  var id = sheet.getRange(rowNum, 1).getValue();
  Logger.log('  管理番号=' + id);
  
  // 商品管理シート F列(管理番号)検索
  var ids = main.getRange(2, 6, main.getLastRow() - 1, 1).getValues();
  var idx0 = ids.findIndex(function(c) { return c[0] === id; });
  if (idx0 < 0) {
    Logger.log('  該当IDなし: ' + id);
    return;
  }
  var tgtRow = idx0 + 2;
  Logger.log('  対象行=' + tgtRow);
  
  // J～M 取得→ 商品管理 AP～AU 反映
  var row = sheet.getRange(rowNum, 10, 1, 4).getValues()[0];
  main.getRange(tgtRow, 42).setValue(row[0]); // AP: 販売日
  main.getRange(tgtRow, 43).setValue(row[1]); // AQ: 販売場所
  main.getRange(tgtRow, 44).setValue(row[2]); // AR: 販売価格
  main.getRange(tgtRow, 47).setValue(row[3]); // AU: 入金額
  Logger.log('  販売情報反映 AP～AU');

  // E列を「売却済み」に設定
  main.getRange(tgtRow, 5).setValue('売却済み');
  Logger.log('  ステータス変更 E' + tgtRow + ' = 売却済み');

  // 利益・利益率計算
  var cost   = main.getRange(tgtRow, 41).getValue(); // AO: 仕入れ値
  var profit = row[3] - cost;
  var rate   = cost ? profit / cost : '';
  main.getRange(tgtRow, 48).setValue(profit); // AV: 利益
  main.getRange(tgtRow, 49).setValue(rate);   // AW: 利益率
  Logger.log('  利益計算: 仕入=' + cost + ' → 利益=' + profit + ' 率=' + rate);

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