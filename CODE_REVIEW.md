# コードレビュー報告書 - 仕入れ管理Ver.2

## 概要

全20ファイル（.gs 16ファイル、.html 4ファイル）を精査し、重複コード・改善点を以下にまとめます。

---

## 1. 重複コード（Duplications）

### 1-1. `findHeaderCol_` 関数が3ファイルで重複定義

| ファイル | 関数名 | 動作の違い |
|---|---|---|
| `メニュー.gs:664` | `findHeaderCol_` | 候補配列を受け取り、部分一致も許可。見つからなければ `0` を返す |
| `返送済みステータス変更.gs:103` | `findHeaderCol_` | 単一文字列を受け取り、完全一致のみ。見つからなければ **throw Error** |
| `xlsxダウンロード.gs:130` | `findColByHeader_` | 単一文字列、完全一致。見つからなければ `-1` を返す |

**推奨:** 共通ユーティリティとして1つに統合し、オプションで `throwOnMissing` や `partialMatch` を指定できるようにする。

### 1-2. 列番号⇔列文字の変換関数が3箇所で重複

| ファイル | 関数名 | 方向 |
|---|---|---|
| `メニュー.gs:464` | `_colToA1Letter_` | 数値 → 文字（A, B, ...） |
| `経費・仕入れ数報告の通知用アドレス.gs:123` | `columnA1ToNumber` | 文字 → 数値 |
| `報酬更新.gs:14` | `col_` | 文字 → 数値 |

**推奨:** `columnA1ToNumber` と `col_` は同じロジック。1つに統合する。

### 1-3. ヘッダからインデックスマップを構築するパターンが6箇所以上で重複

```javascript
// パターンA: Object map
headers.forEach(function (h, i) { idx[h] = i; });
// → メニュー.gs:135, 280, 366-368, 501-502

// パターンB: indexOf 連続呼び出し
const idxBrand = hdr.indexOf("ブランド");
const idxSize  = hdr.indexOf("メルカリサイズ");
// → Code.gs:23-41, 回収完了.gs:59-60, 在庫日数更新.gs:18-20

// パターンC: 専用関数
buildHeaderIndex_()  // → EC管理自動反映.gs:160
findCol_()           // → キーワードAPI.gs:335
```

**推奨:** `buildHeaderIndex_` を共通関数として全ファイルで共有する。

### 1-4. 文字列正規化関数が4箇所で類似重複

| ファイル | 関数名 | 特徴 |
|---|---|---|
| `メニュー.gs:676` | `normalizeBrand_` | trim のみ |
| `返送済みステータス変更.gs:129` | `normalizeText_` | trim + 全角数字変換 + ゼロ幅文字除去 |
| `EC管理自動反映.gs:176` | `normalizeKeyPart_` | trim + Date 対応 |
| `キーワードAPI.gs:380` | `safe_` | 500文字で切り詰め |

**推奨:** `normalizeText_` をベースに統合。特殊ケースはオプション引数で対応。

### 1-5. 数値変換関数が2箇所で重複

| ファイル | 関数名 |
|---|---|
| `メニュー.gs:683` | `toNumber_` — 円記号・カンマ除去、NaN/null 対応 |
| `報酬更新.gs:15` | `toNum` — 簡易版、記号除去 |

**推奨:** `toNumber_` に統合。

### 1-6. 日付ヘルパー関数が3ファイルで重複

| ファイル | 関数群 |
|---|---|
| `作業分析更新.gs:22-50` | `toMonthKey`, `toDateObj`, `toDayKey` |
| `報酬更新.gs:10-13` | `ymKey`, `pad2`, `parseYM`, `mkIndex` |
| `棚卸し.gs:448-451` | `normalizeDate`, `toYMD`, `parseYMD`, `parseISODate` |

**推奨:** 共通の日付ユーティリティファイルに統合。

### 1-7. トリガー設定・削除パターンが4箇所で重複

```javascript
// 同一パターン: 既存トリガー削除 → 新規作成
const triggers = ScriptApp.getProjectTriggers();
triggers.forEach(t => { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
ScriptApp.newTrigger(fn).timeBased()...create();
```

| ファイル | 関数 |
|---|---|
| `トリガー設定.gs:5` | `FULL_RESTORE_ALL` |
| `EC管理自動反映.gs:25` | `setupBaseOrderSync` |
| `返送済みステータス変更.gs:13` | `setupHourlyTrigger_updateReturnStatus` |
| `報酬更新.gs:229` | `setupDailyTrigger` |

**推奨:** 共通関数 `replaceTrigger_(fnName, config)` を作成。

### 1-8. ロック取得パターンが5箇所で重複

```javascript
const lock = LockService.getScriptLock();
if (!lock.tryLock(25000)) return;
try { ... } finally { lock.releaseLock(); }
```

使用箇所: `キーワードAPI.gs`, `まとめID発行.gs`, `EC管理自動反映.gs`, `経費・仕入れ数報告の通知用アドレス.gs`, `返送済みステータス変更.gs`

**推奨:** `withScriptLock_(fn, timeout)` ラッパー関数を用意する。

### 1-9. `handleChange_Mailer` と `handleEdit_Mailer` が完全同一

`経費・仕入れ数報告の通知用アドレス.gs:23-41` で2つの関数が全く同じ内容。

**推奨:** 片方を削除し、もう片方から呼び出す。

### 1-10. HTML内の `setStatus` 関数が3ファイルで重複

`BasicSettings.html`, `FeeSettings.html`, `Manual.html` すべてに同じ `setStatus` 関数が定義されている。

---

## 2. 改善点（Improvements）

### 2-1. マジックナンバーの多用（高優先度）

| ファイル:行 | 問題 | 推奨 |
|---|---|---|
| `メニュー.gs:376` | `var summaryCol = 67;` | ヘッダ名で列番号を動的取得すべき |
| `メニュー.gs:384` | `main.getRange(2, 6, ...)` | 列6 = 管理番号 を定数化 |
| `回収完了.gs:26-29` | `42, 43, 44, 47, 48, 49` | AP〜AW列をヘッダ名で動的取得 |
| `作業分析更新.gs:13` | `startCol = 33; width = 26;` | ヘッダ検索で特定すべき |
| `報酬更新.gs:77-92` | `col_('AI')`, `col_('AJ')` 等の大量呼び出し | ヘッダ名検索に変更 |
| `棚卸し.gs:286-289` | `42, 51, 60, 61` | ヘッダから取得 |

**影響:** 列の追加・並び替えで即座に壊れる。最も修正優先度が高い。

### 2-2. コーディングスタイルの不統一

- `メニュー.gs`, `回収完了.gs`, `作業分析更新.gs`, `報酬更新.gs` は `var` を使用（ES5スタイル）
- `Code.gs`, `まとめID発行.gs`, `キーワードAPI.gs` は `const`/`let` を使用（ES6スタイル）
- アロー関数の使用も不統一

**推奨:** 全ファイルを `const`/`let` + アロー関数に統一（GAS V8ランタイムでサポート済み）。

### 2-3. パフォーマンス問題

#### a) 個別セル書き込みの連続呼び出し
`回収完了.gs:26-41` の `reflectAndDelete` は6回の個別 `setValue` を実行：
```javascript
main.getRange(tgtRow, 42).setValue(row[0]);
main.getRange(tgtRow, 43).setValue(row[1]);
// ... 6回続く
```
**推奨:** `setValues` でバッチ書き込みに変更。

#### b) deleteRow のループ実行
`メニュー.gs:458-460`:
```javascript
rowsToDelete.forEach(function (r) { sh.deleteRow(r); });
```
**推奨:** GASでは行削除は下から上に1行ずつ行う必要があるが、大量削除時は一時シートにコピーする方法が高速。

#### c) 複数範囲の個別読み取り
`EC管理自動反映.gs:191-195` の `findAppendRowByActualData_` は5列を別々に読み取り：
```javascript
const rngOrderKey = sh.getRange(2, cols.orderKey, scanRows, 1).getDisplayValues();
const rngChannel  = sh.getRange(2, cols.channel, scanRows, 1).getDisplayValues();
// ... 5回
```
**推奨:** まとめて1回で読み取り、列インデックスで参照。

#### d) 報酬更新.gs の大量個別列読み取り
`報酬更新.gs:77-92` は13列を個別に getRange → getValues：
```javascript
var AI = nP? shP.getRange(2,col_('AI'),nP,1).getValues().flat():[];
var AJ = nP? shP.getRange(2,col_('AJ'),nP,1).getValues().flat():[];
// ... 13回
```
**推奨:** 必要な範囲を1回の `getRange` でまとめて取得し、列オフセットで参照。

### 2-4. エラーハンドリングの不整合

- `generateCompletionList` (`メニュー.gs:51`): `main`, `analysis`, `out` シートが null の場合のチェックなし
- `回収完了.gs:4` `reflectAndDelete`: Logger.log のみでユーザーへのフィードバックなし
- 一部の関数は `throw Error`、他は `Browser.msgBox`、他は `ss.toast` — 統一されていない

**推奨:** エラー通知のパターンを統一（例: UI操作は `ss.toast`、バックグラウンド処理は `Logger.log`）。

### 2-5. `reflectAndDelete` 関数の非使用疑い

`回収完了.gs` の `reflectAndDelete` 関数は `processSelectedSales`（メニュー.gs）の機能と大幅に重複。
- `reflectAndDelete`: 1行ずつ処理（個別setValue × 6回）
- `processSelectedSales`: バッチ処理（RangeList + setValues）

**推奨:** `reflectAndDelete` がどこからも呼ばれていない場合は削除を検討。

### 2-6. `BUSY_KEY` による排他制御（棚卸し.gs）

`棚卸し.gs` は `PropertiesService` の `BUSY_KEY` で排他制御しているが、`finally` ブロックで削除が保証されているとはいえ、GASの実行時間制限（6分）に引っかかった場合にキーが残留してデッドロックになるリスクがある。

**推奨:** タイムスタンプ付きの `BUSY_KEY` にし、一定時間経過したらロック解放する仕組みを入れる。

### 2-7. API キーのハードコード

`EC管理自動反映.gs:2-5` と `xlsxダウンロード.gs:5-14` にスプレッドシートIDやフォルダIDがソースコードにハードコードされている。

**推奨:** `PropertiesService.getScriptProperties()` で管理するか、設定シートから読み込む。

### 2-8. 未使用変数

- `在庫日数更新.gs:35`: `var tz = Session.getScriptTimeZone()` — 宣言されているが使用されていない
- `報酬更新.gs:20`: `var ts = new Date()` — ログメッセージのみに使用

---

## 3. 推奨リファクタリング構成

### 共通ユーティリティファイル `Utils.gs` を新規作成

以下の関数を集約：

```
// ヘッダ関連
buildHeaderIndex_(headerRow)
findHeaderCol_(headerRow, name, options)
columnLetterToNumber(a1)
columnNumberToLetter(col)

// 文字列・数値
normalizeText_(v)
toNumber_(v)

// 日付
toMonthKey(d)
toYMD(d)
parseYMD(s)
normalizeDate(d)

// 制御
withLock_(lockType, timeout, fn)
replaceTrigger_(fnName, triggerConfig)
```

---

## 4. 優先度まとめ

| 優先度 | 項目 | 理由 |
|---|---|---|
| **高** | マジックナンバー排除 (2-1) | 列変更で即障害。運用上の最大リスク |
| **高** | `findHeaderCol_` 統合 (1-1) | 名前衝突リスク（同名関数が異なる動作） |
| **中** | パフォーマンス改善 (2-3) | GAS実行時間制限（6分）に達するリスク |
| **中** | 日付/数値ユーティリティ統合 (1-5, 1-6) | メンテナンス性向上 |
| **低** | コーディングスタイル統一 (2-2) | 機能には影響なし |
| **低** | HTML `setStatus` 重複 (1-10) | GASではinclude不可のため許容範囲 |
