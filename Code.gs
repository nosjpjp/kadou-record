/**
 * KADOU RECORD – Google Sheets バックエンド
 *
 * 列の並び（左から固定）:
 *   A: 日付 → 出勤 → 退勤 → 休憩 → 移動 → 各プロジェクト（追加順）
 *   固定列は常に日付の直後にこの順で並びます。プロジェクト列はその後ろに追加されます。
 *
 * 使い方:
 *  1. 記録用スプレッドシートで「拡張機能 → Apps Script」を開く
 *  2. このコードを貼り付けて保存
 *  3. デプロイ → 新しいデプロイ → ウェブアプリ（実行: 自分 / アクセス: 全員）
 *  4. /exec URL をアプリの設定に貼り付け
 *  ※ コードを変更したら「デプロイを管理 → 編集 → 新バージョン → デプロイ」で再デプロイ
 *
 * チーム利用:
 *   全員が同じスプレッドシート＋同じ /exec URL を設定。各自の「メンバー名」ごとに
 *   専用タブが自動作成され、同じ列構成で記録されます（メンバー名が空なら 'Log' タブ）。
 */

var DEFAULT_SHEET = 'Log';

// 集計タブ名
var SUMMARY_SHEET = '集計';

// 端末間同期用の状態タブ（メンバーごとに全状態をJSONで保持）
var STATE_SHEET = '_state';

/**
 * 合言葉（簡易トークン）。
 *   ここを各チーム独自の文字列に変更してください（例: 'nosight-2026-xyz'）。
 *   アプリの「設定 → 合言葉」に同じ文字列を入れた人だけが書き込み・集計できます。
 *   空文字 '' のままにすると認証なし（誰でも書き込み可）になります。
 */
var ACCESS_TOKEN = 'Nos@1202';

// 固定列の並び（日付=A列の次に、この順で配置）
var FIXED_COLS = ['出勤', '退勤', '休憩', '移動'];

// 合言葉の照合（ACCESS_TOKEN が空なら常に許可）
function authOK_(token) {
  if (!ACCESS_TOKEN) return true;
  return String(token || '') === ACCESS_TOKEN;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // 稼働: {date, project, member, hours}
    // 打刻: {kind:'punch', date, member, which:'in'|'out', time:'HH:MM'}
    var data = JSON.parse(e.postData.contents);

    // 合言葉チェック（不一致なら拒否）
    if (!authOK_(data.token)) {
      return json_({ ok: false, error: 'unauthorized' });
    }

    // 集計更新リクエスト（アプリの「集計更新」ボタン）
    if (data.action === 'summary') {
      var n = rebuildSummary();
      return json_({ ok: true, summary: true, rows: n });
    }

    // 端末間同期：状態を保存（より新しいタイムスタンプのみ採用＝last-write-wins）
    if (data.action === 'pushState') {
      var ssh = stateSheet_();
      var srow = findStateRow_(ssh, data.member);
      var inc = Number(data.updatedAt || 0);
      if (srow) {
        var cur = Number(ssh.getRange(srow, 3).getValue() || 0);
        if (inc >= cur) {
          ssh.getRange(srow, 2).setValue(String(data.json || ''));
          ssh.getRange(srow, 3).setValue(inc);
        }
      } else {
        ssh.appendRow([String(data.member || ''), String(data.json || ''), inc]);
      }
      return json_({ ok: true });
    }

    var sh = getSheet_(data.member);
    if (sh.getLastRow() < 1) sh.getRange(1, 1).setValue('日付');

    var row = findOrCreateDateRow_(sh, data.date);
    var col, value;
    if (data.kind === 'punch') {
      col = findOrCreateCol_(sh, data.which === 'in' ? '出勤' : '退勤');
      value = data.time;            // 'HH:MM' 文字列（Sheets 側で時刻として表示）
    } else {
      col = findOrCreateCol_(sh, data.project);
      value = Number(data.hours);   // 稼働時間（絶対値で上書き／再送しても二重加算なし）
    }
    sh.getRange(row, col).setValue(value);

    return json_({ ok: true, row: row, col: col });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ヘッダー名から列を検索（無ければ正しい位置に作成）
function findOrCreateCol_(sh, name) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // 既存の列はそのまま使う
  for (var i = 1; i < headers.length; i++) {
    if (String(headers[i]).trim() === String(name).trim()) return i + 1;
  }

  var fixedIndex = FIXED_COLS.indexOf(name);
  if (fixedIndex >= 0) {
    // 固定列：日付(1列目)の直後、FIXED_COLS の順で入る位置を計算して挿入
    var insertAt = 2; // 日付の次から
    for (var f = 0; f < fixedIndex; f++) {
      if (headers.indexOf(FIXED_COLS[f]) >= 0) insertAt++;
    }
    sh.insertColumnBefore(insertAt);
    sh.getRange(1, insertAt).setValue(name);
    return insertAt;
  } else {
    // プロジェクト列：末尾に追加
    var col = lastCol + 1;
    sh.getRange(1, col).setValue(name);
    return col;
  }
}

// 日付から行を検索（無ければ末尾に追加）
function findOrCreateDateRow_(sh, date) {
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var dates = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var r = 0; r < dates.length; r++) {
      if (fmt_(dates[r][0]) === String(date)) return r + 2;
    }
  }
  var row = Math.max(lastRow + 1, 2);
  sh.getRange(row, 1).setValue(date);
  return row;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'summary') {
    if (!authOK_(p.token)) return json_({ ok: false, error: 'unauthorized' });
    var n = rebuildSummary();
    return json_({ ok: true, summary: true, rows: n });
  }
  // 端末間同期：状態を取得（JSONP で返す＝CORS回避）
  if (p.action === 'pullState') {
    var cb = p.callback || 'callback';
    if (!authOK_(p.token)) return jsonp_(cb, { ok: false, error: 'unauthorized' });
    var ssh = stateSheet_();
    var srow = findStateRow_(ssh, p.member);
    if (!srow) return jsonp_(cb, { ok: true, json: null, updatedAt: 0 });
    var js = ssh.getRange(srow, 2).getValue();
    var up = Number(ssh.getRange(srow, 3).getValue() || 0);
    return jsonp_(cb, { ok: true, json: String(js || ''), updatedAt: up });
  }
  return json_({ ok: true, message: 'KADOU RECORD backend is alive' });
}

function jsonp_(cb, obj) {
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(obj) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function stateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STATE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STATE_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['member', 'json', 'updatedAt']]);
  }
  return sh;
}

function findStateRow_(sh, member) {
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(member)) return i + 2;
  }
  return null;
}

function getSheet_(member) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = sanitizeTab_(member) || DEFAULT_SHEET;
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// タブ名に使えない文字を除去し 90 文字に制限
function sanitizeTab_(s) {
  if (!s) return '';
  return String(s).replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90);
}

function fmt_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).trim();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 既存のバラバラな列を、正しい並びに一括で直すツール（任意・1回だけ）
 *
 * 使い方：関数選択で reorderAllTabs を選んで「実行」。
 *   全タブを「日付 → 出勤 → 退勤 → 休憩 → 移動 → 各プロジェクト（元の順）」に並べ替えます。
 *   ※ データは保持されます。実行前にシートのコピーを取っておくと安心です。
 */
function reorderAllTabs() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var s = 0; s < sheets.length; s++) reorderSheet_(sheets[s]);
}

function reorderSheet_(sh) {
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });

  // ヘッダーに「日付」が無いタブはスキップ（Log/メンバータブのみ対象）
  if (headers.indexOf('日付') === -1) return;

  // 望ましい並び：日付 → 固定列（存在するもの）→ 残り（プロジェクト等を元の順）
  var order = [];
  function pushIfExists(name) {
    var idx = headers.indexOf(name);
    if (idx >= 0 && order.indexOf(idx) === -1) order.push(idx);
  }
  pushIfExists('日付');
  FIXED_COLS.forEach(pushIfExists);
  for (var c = 0; c < headers.length; c++) {
    if (order.indexOf(c) === -1) order.push(c);
  }

  // 並べ替え後の2次元配列を作成して書き戻す
  var out = values.map(function (rowArr) {
    return order.map(function (ci) { return rowArr[ci]; });
  });

  sh.clearContents();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
}

/**
 * 全員集計（行=プロジェクト × 列=各メンバー）を「集計」タブに書き出す。
 *
 * グルーピング規則:
 *   - '9999_' で始まる名称は、名称ごとに別行（統合しない）。
 *   - それ以外で頭4桁が数字のものは、4桁コードが一致すれば同一プロジェクトとして統合。
 *   - 4桁コードが無い名称は、その名称を1グループとして集計。
 *   - 休憩・移動は別行で集計。出勤・退勤は集計対象外。
 *
 * 代表名称（行ラベル）:
 *   管理者が「集計」タブの A 列に事前記入した正式名称を採用。
 *   その行の頭4桁（'9999_'系は名称完全一致）で各メンバーの該当PJを合算する。
 *   集計シートに未記載のPJはタブ下部に「未登録」として自動追記。
 */
function rebuildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summary = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);

  // 1) 集約: agg[key][member] = 合計時間, repName[key] = 代表名称（未登録時の表示用）
  var agg = {}, repName = {}, members = [];
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var name = sh.getName();
    if (name === SUMMARY_SHEET) continue;
    var values = sh.getDataRange().getValues();
    if (values.length < 1) continue;
    var headers = values[0];
    // 「日付」列を持つタブのみメンバータブとみなす
    var hasDate = false;
    for (var h = 0; h < headers.length; h++) { if (String(headers[h]).trim() === '日付') { hasDate = true; break; } }
    if (!hasDate) continue;

    members.push(name);
    for (var c = 0; c < headers.length; c++) {
      var head = String(headers[c]).trim();
      if (head === '' || head === '日付' || head === '出勤' || head === '退勤') continue; // 出勤/退勤は除外
      var sum = 0;
      for (var r = 1; r < values.length; r++) { var v = values[r][c]; if (typeof v === 'number' && !isNaN(v)) sum += v; }
      var key = projKey_(head);
      agg[key] = agg[key] || {};
      agg[key][name] = (agg[key][name] || 0) + sum;
      if (!repName[key]) repName[key] = head;
    }
  }

  // 2) 管理者が記入済みの行ラベル（A列）を取得。'合計'/'未登録' 以降は前回の自動出力なので除外
  var adminLabels = [];
  var lastRow = summary.getLastRow();
  if (lastRow > 1) {
    var colA = summary.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var t = String(colA[i][0]).trim();
      if (t === '合計') break;
      if (t.indexOf('未登録') >= 0) break;
      if (t === '') continue;
      adminLabels.push(String(colA[i][0]).trim());
    }
  }

  // 3) 出力を組み立て
  var out = [];
  var used = {};
  // 3-1) 登録済み（管理者ラベル）行
  for (var a = 0; a < adminLabels.length; a++) {
    var label = adminLabels[a];
    var k = projKey_(label);
    used[k] = true;
    var row = [label], tot = 0;
    for (var m = 0; m < members.length; m++) {
      var val = round2((agg[k] && agg[k][members[m]]) || 0);
      row.push(val); tot += val;
    }
    row.push(round2(tot));
    out.push(row);
  }
  // 3-2) 列合計行
  var totRow = ['合計'], grand = 0;
  for (var ci = 0; ci < members.length; ci++) {
    var cs = 0;
    for (var oi = 0; oi < out.length; oi++) cs += Number(out[oi][ci + 1]) || 0;
    cs = round2(cs); totRow.push(cs); grand += cs;
  }
  totRow.push(round2(grand));
  out.push(totRow);
  // 3-3) 未登録（集計シート未記載）
  var unreg = [];
  for (var key in agg) { if (!used[key]) unreg.push(key); }
  if (unreg.length) {
    var sep = ['――― 未登録（集計シート未記載のPJ）―――'];
    for (var z = 0; z < members.length + 1; z++) sep.push('');
    out.push(sep);
    unreg.sort();
    for (var u = 0; u < unreg.length; u++) {
      var uk = unreg[u];
      var ulabel = repName[uk] || uk.replace(/^[CN]:/, '');
      var urow = [ulabel], ut = 0;
      for (var mm = 0; mm < members.length; mm++) {
        var uv = round2((agg[uk] && agg[uk][members[mm]]) || 0);
        urow.push(uv); ut += uv;
      }
      urow.push(round2(ut));
      out.push(urow);
    }
  }

  // 4) 書き出し（ヘッダー＋本体）。A列の管理者ラベルはメモリに退避済みなので clear して再構築
  var header = ['プロジェクト'].concat(members).concat(['合計']);
  summary.clearContents();
  summary.getRange(1, 1, 1, header.length).setValues([header]);
  if (out.length) summary.getRange(2, 1, out.length, header.length).setValues(out);
  summary.setFrozenRows(1);
  summary.setFrozenColumns(1);

  return out.length;
}

// プロジェクト名 → グルーピングキー
function projKey_(name) {
  name = String(name).trim();
  if (name.indexOf('9999_') === 0) return 'N:' + name;       // 9999_ は名称ごと
  var m = name.match(/^(\d{4})(?:_|$)/);
  if (m) return 'C:' + m[1];                                  // 頭4桁コードで統合
  return 'N:' + name;                                         // コードなしは名称
}

function round2(x) { return Math.round((Number(x) || 0) * 100) / 100; }

/**
 * 集計を1時間ごとに自動更新するトリガーを設定（1回だけ実行すればOK）。
 * 既存の rebuildSummary トリガーがあれば作り直します。
 */
function installSummaryTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) {
    if (ts[i].getHandlerFunction() === 'rebuildSummary') ScriptApp.deleteTrigger(ts[i]);
  }
  ScriptApp.newTrigger('rebuildSummary').timeBased().everyHours(1).create();
  Logger.log('1時間ごとの集計トリガーを設定しました');
}

// 動作確認・権限承認用（必要なら実行）
function testWrite() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(ss ? ('対象シート: ' + ss.getName()) : 'NULL（バインドされていません）');
  var sh = getSheet_('');
  sh.getRange(1, 1).setValue('日付');
  sh.getRange(2, 1).setValue('テスト書き込み ' + new Date());
  Logger.log('書き込み完了');
}
