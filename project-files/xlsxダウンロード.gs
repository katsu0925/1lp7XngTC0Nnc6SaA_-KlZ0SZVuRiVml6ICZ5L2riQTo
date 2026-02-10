// ── 定数（変更箇所） ──
// REQUEST_PUT_COL は廃止 → ヘッダで動的に探す
// REQUEST_MATCH_COL も廃止 → ヘッダで動的に探す

const SOURCE_SPREADSHEET_ID = '1lp7XngTC0Nnc6SaA_-KlZ0SZVuRiVml6ICZ5L2riQTo';
const SOURCE_SHEET_GID = 1614333946;

const NAME_SHEET_NAME = '配布用リスト';
const NAME_CELL_A1 = 'E1';
const RECEIPT_CELL = 'I1';           // ★追加: 受付番号セル

const EXPORT_FOLDER_ID = '1lq8Xb_dVwz5skrXlGvrS5epTwEc_yEts';

const REQUEST_SPREADSHEET_ID = '1eDkAMm_QUDFHbSzkL4IMaFeB2YV6_Gw5Dgi-HqIB2Sc';
const REQUEST_SHEET_NAME = '依頼管理';

// ★変更: ヘッダ名で列を特定する
const HEADER_RECEIPT = '受付番号';   // A列にあるはずのヘッダ名
const HEADER_NAME = '会社名/氏名';   // 名前列のヘッダ
const HEADER_LINK = '確認リンク';        // URL を入れる列のヘッダ名

// ──────────────────────────────────
// メイン関数（変更箇所のみコメント付き）
// ──────────────────────────────────
function exportDistributionList() {
  const srcSs = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);

  const nameSheet = srcSs.getSheetByName(NAME_SHEET_NAME);
  if (!nameSheet) throw new Error('配布用リスト が見つかりません');

  const rawName = String(nameSheet.getRange(NAME_CELL_A1).getDisplayValue() || '').trim();
  if (!rawName) throw new Error('配布用リスト!E1 が空です');

  // ★追加: I1 から受付番号を取得
  const receiptNo = String(nameSheet.getRange(RECEIPT_CELL).getDisplayValue() || '').trim();
  if (!receiptNo) throw new Error('配布用リスト!I1（受付番号）が空です');

  const baseName = rawName + '様';
  const exportFileName = baseName + '.xlsx';

  const folder = DriveApp.getFolderById(EXPORT_FOLDER_ID);

  if (folderHasFileName_(folder, exportFileName)) {
    return { ok: false, message: '同名ファイルが既に存在するため処理を中止しました', fileName: exportFileName };
  }

  const srcSheet = getSheetById_(srcSs, SOURCE_SHEET_GID);

  const tmpSs = SpreadsheetApp.create('tmp_' + baseName + '_' + Date.now());
  const tmpId = tmpSs.getId();

  const copied = srcSheet.copyTo(tmpSs);
  copied.setName(srcSheet.getName());

  deleteAllExceptSheet_(tmpSs, copied.getSheetId());

  trimColumnBAfterSecondHyphen_(copied);

  trimToDataBoundsStrict_(copied);

  SpreadsheetApp.flush();

  const xlsxBlob = exportSpreadsheetAsXlsxBlob_(tmpId, exportFileName);

  const outFile = folder.createFile(xlsxBlob);
  outFile.setName(exportFileName);
  outFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const url = outFile.getUrl();

  // ★変更: 受付番号も渡す
  updateRequestSheetLink_(rawName, receiptNo, url);

  DriveApp.getFileById(tmpId).setTrashed(true);

  return { ok: true, url: url, fileName: exportFileName };
}

// ──────────────────────────────────
// ★ 書き換えた関数
// ──────────────────────────────────
function updateRequestSheetLink_(name, receiptNo, url) {
  const ss = SpreadsheetApp.openById(REQUEST_SPREADSHEET_ID);
  const sh = ss.getSheetByName(REQUEST_SHEET_NAME);
  if (!sh) throw new Error('依頼管理 シートが見つかりません');

  const lastRow = Math.max(sh.getLastRow(), 1);
  const lastCol = Math.max(sh.getLastColumn(), 1);

  // ── 1行目ヘッダから列番号を特定 ──
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  const receiptCol = findColByName_(headers, HEADER_RECEIPT);  // 受付番号列
  const nameCol    = findColByName_(headers, HEADER_NAME);     // 会社名/氏名列
  const linkCol    = findColByName_(headers, HEADER_LINK);     // 確認リンク列

  if (receiptCol === -1) throw new Error('ヘッダに「' + HEADER_RECEIPT + '」が見つかりません');
  if (nameCol === -1)    throw new Error('ヘッダに「' + HEADER_NAME + '」が見つかりません');
  if (linkCol === -1)    throw new Error('ヘッダに「' + HEADER_LINK + '」が見つかりません');

  // ── データ取得（受付番号列と名前列を読む） ──
  const dataRows = lastRow - 1;
  const receiptVals = sh.getRange(2, receiptCol, dataRows, 1).getDisplayValues();
  const nameVals    = sh.getRange(2, nameCol, dataRows, 1).getDisplayValues();

  const targetReceipt = String(receiptNo || '').trim();
  const targetName    = String(name || '').trim();

  let found = false;
  for (let i = 0; i < dataRows; i++) {
    const r = String(receiptVals[i][0] || '').trim();
    const n = String(nameVals[i][0] || '').trim();
    // ★ 受付番号と名前の両方が一致する行のみ書き込む
    if (r === targetReceipt && n === targetName) {
      sh.getRange(i + 2, linkCol).setValue(url);   // +2 = ヘッダ1行 + 0始まり補正
      found = true;
    }
  }

  if (!found) {
    // 一致する行がなければ最終行の次に追記
    const newRow = lastRow + 1;
    sh.getRange(newRow, receiptCol).setValue(targetReceipt);
    sh.getRange(newRow, nameCol).setValue(targetName);
    sh.getRange(newRow, linkCol).setValue(url);
  }
}

// findColByHeader_ は Utils.gs の findColByName_ に統合済み

// ──────────────────────────────────
// 以下は変更なし（そのまま残す）
// ──────────────────────────────────
function folderHasFileName_(folder, filename) {
  const it = folder.getFilesByName(filename);
  return it.hasNext();
}

function getSheetById_(ss, gid) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  throw new Error('指定gidのシートが見つかりません: ' + gid);
}

function deleteAllExceptSheet_(ss, keepSheetId) {
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i >= 0; i--) {
    const sh = sheets[i];
    if (sh.getSheetId() !== keepSheetId) {
      if (ss.getSheets().length > 1) ss.deleteSheet(sh);
    }
  }
}

function trimColumnBAfterSecondHyphen_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const rng = sheet.getRange(1, 2, lastRow, 1);
  const vals = rng.getDisplayValues();

  for (let i = 0; i < vals.length; i++) {
    const s = String(vals[i][0] || '');
    if (!s) {
      vals[i][0] = '';
      continue;
    }
    const parts = s.split('-');
    if (parts.length >= 2) {
      vals[i][0] = parts[0] + '-' + parts[1];
    } else {
      vals[i][0] = s;
    }
  }

  rng.setValues(vals);
}

function trimToDataBoundsStrict_(sheet) {
  const rowCand = Math.max(sheet.getLastRow(), 1);
  const colCand = Math.max(sheet.getLastColumn(), 1);

  const vals = sheet.getRange(1, 1, rowCand, colCand).getDisplayValues();

  let lastR = 1;
  let lastC = 1;

  for (let r = 0; r < vals.length; r++) {
    const row = vals[r];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c] || '').trim() !== '') {
        if (r + 1 > lastR) lastR = r + 1;
        if (c + 1 > lastC) lastC = c + 1;
      }
    }
  }

  const maxR = sheet.getMaxRows();
  const maxC = sheet.getMaxColumns();

  if (maxR > lastR) sheet.deleteRows(lastR + 1, maxR - lastR);
  if (maxC > lastC) sheet.deleteColumns(lastC + 1, maxC - lastC);
}

function exportSpreadsheetAsXlsxBlob_(spreadsheetId, filename) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('XLSXエクスポートに失敗しました: ' + code + ' / ' + res.getContentText());
  }

  return res.getBlob().setName(filename);
}
