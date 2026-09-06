/* ★日報のLINE送信を止めておく（2026-09-01）。無料枠を見積の成約に回すため。
   また送りたくなったら true に戻す */
var NIPPO_LINE_ON = false;

/**
 * 業務日報 受信・記録スクリプト
 */

// Drive権限エラーが出たら、エディタでこの関数を選んで「実行」→ アクセスを許可する
function 権限を承認する() {
  DriveApp.getRootFolder().getName();
  SpreadsheetApp.getActiveSpreadsheet().getName();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  ScriptApp.getProjectTriggers();  // トリガー管理スコープ（PDF非同期化に必要）の承認
  CalendarApp.getAllCalendars().length;  // カレンダー読取スコープ（日報の抜けチェックに必要）の承認
  // メール送信スコープの承認。※getRemainingDailyQuota()では許可画面が出ないため実際に1通送る
  MailApp.sendEmail('yawata51@gmail.com', '【権限確認】マックライン通知', 'メール送信の権限が有効になりました。');
  Logger.log('権限OK');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏠 日報メニュー')
    .addItem('📱 業務日報アプリを開く', 'openKoujiApp')
    .addToUi();
}

function openKoujiApp() {
  const url = 'https://yawata51-dotcom.github.io/himawari-report/kouji.html';
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '"); google.script.host.close();</script>'
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, '業務日報アプリを開いています...');
}

// 顧客検索対応
function doGet(e) {
  // 全顧客リスト（フロント側で商品見積と同じ即時絞り込みをするため）
  if (e && e.parameter && e.parameter.action === 'customers') {
    const ss = SpreadsheetApp.openById(CUSTOMER_MASTER_SS_ID);
    const s = ss.getSheetByName('顧客マスタ') || ss.getSheetByName('顧客ﾏｽﾀ');
    if (!s) return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
    const d = s.getDataRange().getValues();
    const out = d.slice(1).filter(function(r) { return r[1]; }).map(function(r) {
      return {
        name: String(r[1] || ''),
        kana: String(r[2] || ''),
        tel:  String(r[3] || ''),
        tels: [r[3], r[4], r[5], r[6], r[7]].map(function(t) { return String(t || '').replace(/[-\s]/g, ''); }).filter(String).join(','),
        address: String(r[15] || '') + String(r[16] || '')
      };
    });
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  // 日報まとめのトリガーを設定し直す
  if (e && e.parameter && e.parameter.action === 'trigset') {
    return ContentService.createTextOutput(日報まとめ20時トリガー設定());
  }
  // トリガーの状態を確認する
  if (e && e.parameter && e.parameter.action === 'trigdiag') {
    var _ts = ScriptApp.getProjectTriggers();
    var _o = ['トリガー数=' + _ts.length];
    _ts.forEach(function(t) {
      _o.push('  ' + t.getHandlerFunction() + '  種類=' + t.getEventType());
    });
    // 最近の実行ログ（エラー行）も見る
    var _sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var _last = _sh.getLastRow();
    var _v = _sh.getRange(Math.max(1, _last - 30), 1, Math.min(30, _last), 4).getValues();
    _o.push('', '■ エラー記録');
    _v.forEach(function(r) {
      if (String(r[0]).indexOf('エラー') >= 0) _o.push('  ' + r.join(' | ').slice(0, 200));
    });
    return ContentService.createTextOutput(_o.join('\n'));
  }
  // AIが呼べるか確認する（キーの設定チェック用）
  if (e && e.parameter && e.parameter.action === 'aitest') {
    var _k = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if (!_k) return ContentService.createTextOutput('❌ ANTHROPIC_API_KEY がGASに設定されていません');
    var _r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': _k, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-opus-4-8', max_tokens: 40,
        messages: [{ role: 'user', content: '「テスト成功」とだけ返してください' }]
      }),
      muteHttpExceptions: true
    });
    var _c = _r.getResponseCode();
    if (_c !== 200) return ContentService.createTextOutput('❌ AI呼び出し失敗 status=' + _c + '\n' + _r.getContentText().slice(0, 300));
    var _j = JSON.parse(_r.getContentText());
    return ContentService.createTextOutput('✅ AIが使えます\nAIの返事: ' + _j.content[0].text +
      '\nキーの先頭: ' + _k.slice(0, 14) + '...');
  }
  // 見込みの集計用（日報のS列）
  // スタッフ目標の達成状況（?action=mokuhyou&ym=2026-08）
  if (e && e.parameter && e.parameter.action === 'mokuhyou') {
    return ContentService.createTextOutput(JSON.stringify(目標の達成状況(e.parameter.ym)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 見込シートの中身を見る（動作確認用）
  if (e && e.parameter && e.parameter.action === 'mikomisheet') {
    var _ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('見込');
    if (!_ms) {
      var _q0 = JSON.parse(PropertiesService.getScriptProperties().getProperty('MEIBAN_QUEUE') || '[]');
      var _tg = ScriptApp.getProjectTriggers().filter(function(t) {
        return t.getHandlerFunction() === 'processMeibanQueue';
      }).length;
      return ContentService.createTextOutput(
        '見込シートはまだありません\n待ち行列=' + _q0.length + '件　処理トリガー=' + _tg + '本');
    }
    var _mv = _ms.getRange(1, 1, _ms.getLastRow(), 16).getValues();
    var _mo = ['行数=' + (_mv.length - 1), ''];
    _mv.forEach(function(r, i) {
      _mo.push((i === 0 ? '■ ' : '  ') + r.map(function(x) {
        return (x instanceof Date) ? Utilities.formatDate(x, 'Asia/Tokyo', 'MM/dd') : String(x || '');
      }).join(' | '));
    });
    return ContentService.createTextOutput(_mo.join('\n'));
  }
  // 見込の顧客引き当てを試す（?action=mikomicust&name=鹿内典子）
  if (e && e.parameter && e.parameter.action === 'mikomicust') {
    var _n = String(e.parameter.name || '');
    var _ci = mikomiCustomer_(_n, hagakiCustomerIndex_());
    return ContentService.createTextOutput(
      '入力=' + _n + '\n引き当て=' + JSON.stringify(_ci));
  }
  // 見込シートの行を消す（?action=mikomidel&row=3）
  if (e && e.parameter && e.parameter.action === 'mikomidel') {
    var _ds = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('見込');
    if (!_ds) return ContentService.createTextOutput('見込シートがありません');
    var _dr = Number(e.parameter.row);
    if (!_dr || _dr < 2) return ContentService.createTextOutput('rowを指定してください（2以上）');
    var _got = _ds.getRange(_dr, 1, 1, 16).getValues()[0];
    _ds.deleteRow(_dr);
    return ContentService.createTextOutput('消しました：' + _got.slice(1, 8).join(' | '));
  }
  // 見込の読み取り待ちを今すぐ処理する
  if (e && e.parameter && e.parameter.action === 'mikomirun') {
    processMeibanQueue();
    return ContentService.createTextOutput('処理しました');
  }
  if (e && e.parameter && e.parameter.action === 'mikomi') {
    var _sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('日報');
    var _v = _sh.getRange(2, 1, _sh.getLastRow() - 1, 21).getValues();
    var _o = [];
    _v.forEach(function(r) {
      var t = String(r[18] || '').trim();
      if (!t) return;
      _o.push({ ymd: nippoYmd_(r[0]), staff: String(r[1] || ''), visit: String(r[3] || ''),
                work: String(r[4] || ''), mikomi: t });
    });
    return ContentService.createTextOutput(JSON.stringify(_o))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 顧客のランク集計用（区・顧客ランク・RFMランク）
  if (e && e.parameter && e.parameter.action === 'custrank') {
    var _ss = SpreadsheetApp.openById(CUSTOMER_MASTER_SS_ID);
    var _s = _ss.getSheetByName('顧客マスタ') || _ss.getSheetByName('顧客ﾏｽﾀ');
    var _d = _s.getDataRange().getValues();
    var _out = _d.slice(1).filter(function(r) { return r[1]; }).map(function(r) {
      return { ku: String(r[8] || ''), rank: String(r[9] || ''), rfm: String(r[10] || '') };
    });
    return ContentService.createTextOutput(JSON.stringify(_out))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 保有家電（そろそろリスト）
  if (e && e.parameter && e.parameter.action === 'kaden') {
    return ContentService.createTextOutput(JSON.stringify(getKadenList_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // お礼ハガキ：未訪問リスト
  if (e && e.parameter && e.parameter.action === 'hagaki') {
    return ContentService.createTextOutput(JSON.stringify(getHagakiPending_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'search') {
    const results = searchCustomer(e.parameter.query);
    return ContentService.createTextOutput(JSON.stringify(results))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 商品マスタ（商品見積SSの「商品マスタ」＋「商品ﾏｽﾀ2」）をカテゴリ→商品名で返す
  if (e && e.parameter && e.parameter.action === 'products') {
    return ContentService.createTextOutput(JSON.stringify(getProductMaster_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 商品見積SSの商品マスタを読み、{order:[カテゴリ...], items:{カテゴリ:[商品名...]}} を返す
function getProductMaster_() {
  var ss = SpreadsheetApp.openById(CUSTOMER_MASTER_SS_ID); // 商品見積SS(1Pqb)
  var order = [];
  var items = {};
  function add(cat, name) {
    cat = String(cat || '').trim(); name = String(name || '').trim();
    if (!cat || !name) return;
    if (!items[cat]) { items[cat] = []; order.push(cat); }
    if (items[cat].indexOf(name) < 0) items[cat].push(name);
  }
  var s1 = ss.getSheetByName('商品マスタ') || ss.getSheetByName('商品ﾏｽﾀ');
  if (s1 && s1.getLastRow() > 6) {
    s1.getRange(7, 1, s1.getLastRow() - 6, 3).getValues().forEach(function(r) { add(r[1], r[2]); });
  }
  var s2 = ss.getSheetByName('商品ﾏｽﾀ2');
  if (s2 && s2.getLastRow() > 1) {
    s2.getRange(2, 1, s2.getLastRow() - 1, 3).getValues().forEach(function(r) { add(r[1], r[2]); });
  }
  return { order: order, items: items };
}

// 一時：完了写真を報告書フォルダから分離
  if (e && e.parameter && e.parameter.action === 'splitkanryo') {
    return ContentService.createTextOutput(完了写真を報告書フォルダから分離());
  }
  // 一時：日報まとめの20時トリガー設定（実行後に削除する）
  if (e && e.parameter && e.parameter.action === 'setupdaily') {
    return ContentService.createTextOutput(日報まとめ20時トリガー設定());
  }
  // 一時：ログ（1枚目シートの末尾）を確認
  if (e && e.parameter && e.parameter.action === 'readlog') {
    var ls = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var lr = ls.getLastRow();
    var start = Math.max(1, lr - 14);
    var rows = ls.getRange(start, 1, lr - start + 1, Math.min(4, ls.getLastColumn())).getValues();
    return ContentService.createTextOutput(JSON.stringify(rows, null, 2));
  }
  // 一時：トリガー経由テスト送信（1分後）
  if (e && e.parameter && e.parameter.action === 'firetest') {
    return ContentService.createTextOutput(fireTestTriggerNow_());
  }
  // 一時：トリガー一覧を確認
  if (e && e.parameter && e.parameter.action === 'ssid') {
    return ContentService.createTextOutput('日報SSID=' + SpreadsheetApp.getActiveSpreadsheet().getId() + ' 名前=' + SpreadsheetApp.getActiveSpreadsheet().getName());
  }
  if (e && e.parameter && e.parameter.action === 'houchi') {
    var _hh = 放置案件をまとめる();
    var _hl = 放置案件をさがす();
    var _ho = [_hh, ''];
    _hl.forEach(function(x){ _ho.push('[' + x.kind + '] ' + x.name + '　' + x.date + '（' + x.days + '日）　' + x.what + '　' + x.note); });
    return ContentService.createTextOutput(_ho.join(String.fromCharCode(10)));
  }
  if (e && e.parameter && e.parameter.action === 'listtriggers') {
    var ts = ScriptApp.getProjectTriggers().map(function(t) {
      return { fn: t.getHandlerFunction(), type: String(t.getEventType()), src: String(t.getTriggerSource()) };
    });
    return ContentService.createTextOutput(JSON.stringify(ts, null, 2));
  }
  // 一時：日報まとめを今すぐテスト送信（実行後に削除する）。date指定でその日を送信
  if (e && e.parameter && e.parameter.action === 'testdaily') {
    return ContentService.createTextOutput(sendDailyNippoSummary(e.parameter.date) || '（該当日の日報なし＝送信なし）');
  }
  // 一時：まとめ本文のプレビュー（LINE送信なし）。date未指定なら直近の日報日
  if (e && e.parameter && e.parameter.action === 'previewdaily') {
    var d = e.parameter.date || 直近の日報日_();
    return ContentService.createTextOutput((d ? '[' + d + ']\n' : '') + (buildNippoSummaryText_(d) || '（該当日の日報なし）'));
  }
  // 一時：写真をDrive内で名前検索（q=杣谷 など）。両フォルダ＋キュー残りも確認
  if (e && e.parameter && e.parameter.action === 'findphoto') {
    var q3 = String(e.parameter.q || '');
    var o3 = [];
    [{ id: '1Po0QN0S1VmpxmyEt1eCmAYUk0ynbQecd', n: '現調写真' },
     { id: '1xqPMo6fnCM1v-W4UgQrJMrcDlVY0-iBR', n: '完了写真' }].forEach(function(f) {
      try {
        var fo = DriveApp.getFolderById(f.id);
        var it = q3 ? fo.searchFiles('title contains "' + q3 + '"') : fo.getFiles();
        var c = 0;
        while (it.hasNext() && c < 30) {
          var fi = it.next(); c++;
          o3.push(f.n + ' | ' + fi.getName() + ' | ' + Utilities.formatDate(fi.getDateCreated(), 'Asia/Tokyo', 'MM/dd HH:mm'));
        }
        if (!c) o3.push(f.n + ' | 該当なし');
      } catch (er) { o3.push(f.n + ' | ERR ' + er); }
    });
    try {
      var qf = PropertiesService.getScriptProperties().getProperty('KOUJI_PDF_QUEUE');
      o3.push('--- KOUJI_PDF_QUEUE = ' + (qf || '(空)'));
      var qfid = PropertiesService.getScriptProperties().getProperty('KOUJI_QUEUE_FOLDER_ID');
      if (qfid) {
        var qit = DriveApp.getFolderById(qfid).getFiles(); var qc = 0;
        while (qit.hasNext() && qc < 20) { var qq = qit.next(); qc++; o3.push('queue残 | ' + qq.getName() + ' | ' + Utilities.formatDate(qq.getDateCreated(), 'Asia/Tokyo', 'MM/dd HH:mm')); }
        if (!qc) o3.push('queue残 | なし');
      }
    } catch (er) { o3.push('queue ERR ' + er); }
    return ContentService.createTextOutput(o3.join('\n'));
  }
  // 一時：直近の日報を担当・作業・写真URL有無つきで確認（n件, staff絞込）
  if (e && e.parameter && e.parameter.action === 'rowdiag') {
    var sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('日報') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var lr2 = sh2.getLastRow();
    var n2 = Math.min(parseInt(e.parameter.n || '20', 10), 100);
    var st2 = String(e.parameter.staff || '');
    var from2 = Math.max(2, lr2 - 99);
    var vs2 = sh2.getRange(from2, 1, lr2 - from2 + 1, 21).getValues();
    var out2 = [];
    vs2.forEach(function(r, idx) {
      var staff = String(r[1] || '').trim();
      if (st2 && staff.indexOf(st2) < 0) return;
      var url = String(r[20] || '').trim();
      out2.push({
        row: from2 + idx, date: nippoYmd_(r[0]), staff: staff, visit: String(r[3] || ''),
        work: String(r[4] || ''), 写真URL数: url ? url.split(',').filter(String).length : 0
      });
    });
    return ContentService.createTextOutput(JSON.stringify(out2.slice(-n2), null, 2));
  }
  // マックライン（FXシグナル通知：メール）
  if (e && e.parameter && e.parameter.action === 'mackline_test') {
    マックライン初期設定();
    return ContentService.createTextOutput(テスト送信());
  }
  if (e && e.parameter && e.parameter.action === 'mackline_check') {
    return ContentService.createTextOutput(checkMackline());
  }
  if (e && e.parameter && e.parameter.action === 'mackline_setup') {
    マックライン初期設定();
    return ContentService.createTextOutput(トリガー設定());
  }
  // 一時：カレンダー診断（実行アカウント・見えるカレンダー・エラー内容）
  if (e && e.parameter && e.parameter.action === 'caldiag') {
    var o = [];
    try { o.push('effectiveUser=' + Session.getEffectiveUser().getEmail()); } catch (er) { o.push('effectiveUser ERR ' + er); }
    try { o.push('defaultCalendar=' + CalendarApp.getDefaultCalendar().getName()); } catch (er) { o.push('defaultCalendar ERR ' + er); }
    try {
      var all = CalendarApp.getAllCalendars();
      o.push('getAllCalendars=' + all.length + '件');
      all.forEach(function(c) { o.push('  - ' + c.getId() + ' / ' + c.getName()); });
    } catch (er) { o.push('getAllCalendars ERR ' + er); }
    CHECK_CALENDARS.forEach(function(c) {
      try {
        var cal = CalendarApp.getCalendarById(c.id);
        o.push(c.id + ' => ' + (cal ? 'OK(' + cal.getName() + ')' : 'null'));
      } catch (er) { o.push(c.id + ' => ERR ' + er); }
    });
    return ContentService.createTextOutput(o.join('\n'));
  }
  // 一時：カレンダー照合テスト（date=2026-07-23 または dates=2026-07-23,2026-07-24,…）
  if (e && e.parameter && e.parameter.action === 'calcheck') {
    var ds = (e.parameter.dates || e.parameter.date || '').split(',').map(function(s) { return s.trim(); }).filter(String);
    if (!ds.length) return ContentService.createTextOutput('dates を指定してください');
    var txt = ds.map(function(d) { return '=== ' + d + ' ===\n' + calMissingText_(d); }).join('\n');
    return ContentService.createTextOutput(txt);
  }
  return ContentService.createTextOutput("OK");
}

// 一時：完了写真(画像)を正式な完了写真フォルダ(1xqPMo6)へ集約。
// 報告書PDFフォルダ(1vwl)は画像だけ／旧完了写真フォルダ(17R6)は全部を対象。PDF報告書は動かさない
function 完了写真を報告書フォルダから分離(maxMove) {
  var LIMIT = maxMove || 150;
  var dst = DriveApp.getFolderById('1xqPMo6fnCM1v-W4UgQrJMrcDlVY0-iBR'); // 完了写真(正式)
  var sources = [
    { id: '1vwliwaWcBaCNddsfyktCXqiMVwZWPA7W', imageOnly: true },  // 報告書PDFフォルダ→画像だけ
    { id: '17R6uVjwN6kei_VdLjNJX4LeA6TEINwc4', imageOnly: false }  // 旧完了写真フォルダ→全部
  ];
  var toMove = [];
  for (var i = 0; i < sources.length && toMove.length < LIMIT; i++) {
    if (sources[i].id === dst.getId()) continue;
    var src;
    try { src = DriveApp.getFolderById(sources[i].id); } catch (e) { continue; }
    var it = src.getFiles();
    while (it.hasNext() && toMove.length < LIMIT) {
      var f = it.next();
      var nm = f.getName();
      if (nm === '_writetest.txt' || nm === '_diagtest.txt' || nm === '_whoami.txt') continue;
      if (sources[i].imageOnly && (f.getMimeType() || '').indexOf('image/') !== 0) continue;
      toMove.push(f);
    }
  }
  var more = toMove.length >= LIMIT;
  var moved = 0, err = '';
  toMove.forEach(function(f) {
    try { f.moveTo(dst); moved++; } catch (e) { err = e.message; }
  });
  return JSON.stringify({ moved: moved, more: more, err: err });
}

function 診断_フォルダ書込() {
  var out = {};
  try { out.effectiveUser = Session.getEffectiveUser().getEmail(); } catch (e) { out.effectiveUser = 'err:' + e.message; }
  try { out.activeUser = Session.getActiveUser().getEmail(); } catch (e) { out.activeUser = 'err:' + e.message; }
  // 実行アカウントの特定：書ける旧完了フォルダにテストファイルを作り、その所有者＝実行アカウント
  try {
    var probe = DriveApp.getFolderById('17R6uVjwN6kei_VdLjNJX4LeA6TEINwc4')
                  .createFile(Utilities.newBlob('probe', 'text/plain', '_whoami.txt'));
    out.実行アカウント = probe.getOwner().getEmail();
    probe.setTrashed(true);
  } catch (e) { out.実行アカウント = 'err:' + e.message; }
  var targets = {
    '現調_1Po0': '1Po0QN0S1VmpxmyEt1eCmAYUk0ynbQecd',
    '完了_1vwl': '1vwliwaWcBaCNddsfyktCXqiMVwZWPA7W',
    '旧完了_17R6': '17R6uVjwN6kei_VdLjNJX4LeA6TEINwc4',
    '旧現調_1gy': '1--gy-_U-U-onPEZd3i8oWo290NDMZC9w'
  };
  out.folders = {};
  Object.keys(targets).forEach(function(k) {
    var r = {};
    try {
      var f = DriveApp.getFolderById(targets[k]);
      r.name = f.getName();
      try { r.owner = f.getOwner() ? f.getOwner().getEmail() : '(共有ドライブ等)'; } catch (e2) { r.owner = 'err:' + e2.message; }
      r.access = String(f.getSharingAccess());
      try {
        var tf = f.createFile(Utilities.newBlob('t', 'text/plain', '_diagtest.txt'));
        tf.setTrashed(true);
        r.write = 'OK';
      } catch (e3) { r.write = 'NG:' + e3.message; }
    } catch (e1) { r.error = 'フォルダ取得失敗:' + e1.message; }
    out.folders[k] = r;
  });
  return JSON.stringify(out, null, 2);
}

function doPost(e) {

  const FOLDER_IDS = {
    // 現調＝商品見積と共通の現調写真(JPEG)フォルダ／完了＝専用の完了写真フォルダ（報告書PDFとは別）
    "現調": "1Po0QN0S1VmpxmyEt1eCmAYUk0ynbQecd",
    "完了": "1xqPMo6fnCM1v-W4UgQrJMrcDlVY0-iBR"
  };

  try {
    const data = JSON.parse(e.postData.contents);

    // お礼ハガキ：訪問済の書き戻し
    if (data.action === 'hagakiDone') {
      const ok = markHagakiDone_(data.key, data.memo);
      return ContentService.createTextOutput(JSON.stringify({ ok: ok }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("日報") || ss.getSheets()[0];

    const workAmount = parseInt(data.workAmount) || 0;
    const peopleStr = data.people || "";
    let perPersonAmount = 0;
    if (workAmount > 0 && peopleStr) {
      const numMatch = peopleStr.match(/\d+/);
      const peopleNum = numMatch ? parseInt(numMatch[0]) : 1;
      perPersonAmount = Math.floor(workAmount / peopleNum);
    }

    // 写真タグの「見込商品」（旧名「銘板」）は廃止。古いアプリからの送信だけ受け付ける
    const isMeiban = p => p.type === '見込商品' || p.type === '銘板';
    const meibanPhotos = (data.photos || []).filter(isMeiban);
    if (meibanPhotos.length) data.photos = (data.photos || []).filter(p => !isMeiban(p));

    // 見込商品：スタッフは「写真」と「気づいたこと」だけ入れる。
    // 商品名・メーカー・型番・製造年は写真からAIが読み、ランクと時期は備考からAIが判断する。
    const prospects = data.prospectItems || [];
    let prospectText = prospects.map(item => {
      const n = (item.photos || []).length;
      return (item.content || '（写真のみ）') + (n ? ` 📷${n}` : '');
    }).join("\n");
    if (meibanPhotos.length) {
      prospectText += (prospectText ? "\n" : "") + `【写真】見込商品 ${meibanPhotos.length}点`;
    }

    // 見込みをAI読み取りキューへ回す。
    // 写真があれば写真＋備考、写真が無ければ備考だけ（会話で聞いた見込み）でも登録する
    // ★見込み1件＝1行。写真を何枚撮っても1行にまとめる
    prospects.forEach(item => {
      const datas = (item.photos || []).map(ph => ph.data).filter(Boolean);
      if (datas.length || (item.content || '').trim()) {
        meibanPhotos.push({ datas: datas, note: (item.content || '').trim() });
      }
    });

    let photoUrls = [];
    if (data.photos && data.photos.length > 0) {
      photoUrls = data.photos.map((photo, index) => {
        try {
          const folder = getPhotoFolder_(photo.type, FOLDER_IDS);
          const parts = photo.data.split(",");
          const contentType = parts[0].split(":")[1].split(";")[0];
          const bytes = Utilities.base64Decode(parts[1]);
          const fileName = `${data.date}_${data.visitName}_${photo.type}_${index + 1}.jpg`;
          const blob = Utilities.newBlob(bytes, contentType, fileName);
          const file = folder.createFile(blob);
          return file.getUrl();
        } catch (err) {
          return "保存エラー:" + err.message;
        }
      });
    }

    const rowData = [
      data.date,
      data.staff,
      data.time,
      data.visitName,
      data.workType,
      data.inspectionAircon,
      data.inspectionOther,
      data.deliveryItems,
      data.constructionTypes,
      "",                    // K列（販売担当者）空欄で維持
      data.salesAmount,
      data.workAmount,
      perPersonAmount,
      data.people,
      data.payment,
      data.paymentAmount,
      data.creditMethod,
      data.creditTiming,
      prospectText,
      data.remarks,
      photoUrls.join(", ")   // 改行にすると行が縦に膨らむためカンマ区切り1行
    ];

    sheet.appendRow(rowData);

    // ★訪問先で次の約束をしたら、受注日報＋カレンダーに入れる
    if (data.nextVisit) {
      try {
        次の訪問を登録する_(data.nextVisit, {
          date: data.date, staff: data.staff, visitName: data.visitName
        });
      } catch (nvErr) {
        try { sheet.appendRow(['次の訪問エラー', new Date(), String(nvErr)]); } catch (e4) {}
      }
    }

    // ★銘板写真は裏でAIに読ませて「保有家電」に貯める（送信を待たせない）
    if (meibanPhotos.length) {
      try { queueMeiban_(data.visitName, data.staff, meibanPhotos); }
      catch (mErr) { sheet.appendRow(['銘板キューエラー', new Date(), String(mErr)]); }
    }

    // 現調・配達・工事・修理＋写真ありなら報告書PDFを自動生成（見積書ありの場合はスキップ）
    // ★PDF生成＋LINE送信は重いので裏（別実行）に回し、記録・写真保存が済んだ時点ですぐ完了を返す
    const PDF_TRIGGERS = ['工事', '配達', '修理', '現調'];
    if (!data.hasEstimate && data.photos && data.photos.length > 0 &&
        PDF_TRIGGERS.some(function(t) { return String(data.workType || '').includes(t); })) {
      try {
        queueKoujiPdf_(data);           // Driveに一時保存＋トリガー予約（即座に返る）
      } catch (qErr) {
        // キューに積めない場合は従来どおりその場で生成（フォールバック）
        try {
          createKoujiPDF(data);
        } catch (pdfErr) {
          sheet.appendRow(['PDF作成エラー', new Date(), pdfErr.toString()]);
        }
      }
    }

    return ContentService.createTextOutput("Success");

  } catch (err) {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    logSheet.appendRow(["ERRORログ", new Date(), err.toString()]);
    return ContentService.createTextOutput("Error: " + err.toString());
  }
}

// ===== 報告書PDF自動生成：写真をA4縦・2列×2行（4枚/ページ）で「完了報告書(PDF)」フォルダへ =====
function createKoujiPDF(data) {
  const photos = (data.photos || []).slice(0, 20);
  if (photos.length === 0) return;

  // 現調のみ→現調報告書、工事/配達/修理（現調と混在含む）→完了報告書
  const wt = String(data.workType || '');
  const isGencho = wt.includes('現調') && !['工事', '配達', '修理'].some(function(t) { return wt.includes(t); });
  const docTitle = isGencho ? '現調報告書' : '完了報告書';

  // 一時スプレッドシートに写真を配置してPDF化。
  // 4枚＝1シート＝1ページ（シートの切れ目は必ず改ページになるため、写真が途中で切れない）
  const tempSS = SpreadsheetApp.create('Temp_' + docTitle + '_' + Date.now());
  const CAP_ROWS = 2, PHOTO_ROWS = 24, SLOT_ROWS = CAP_ROWS + PHOTO_ROWS; // 1枠=26行×20px=520px（実測の印刷可能高≈1120pxに対しヘッダー52+2段1040=1092で収まる）

  for (let start = 0; start < photos.length; start += 4) {
    const chunk = photos.slice(start, start + 4);
    const pageNo = start / 4 + 1;
    const sh = pageNo === 1 ? tempSS.getSheets()[0] : tempSS.insertSheet('P' + pageNo);
    sh.setColumnWidth(1, 380); sh.setColumnWidth(2, 380);

    let headerRows = 0;
    if (pageNo === 1) {
      sh.getRange('A1:B1').merge().setValue(docTitle).setFontSize(14).setFontWeight('bold');
      sh.getRange('A2:B2').merge().setValue(
        '日付：' + (data.date || '') + '　担当：' + (data.staff || '') +
        '　現場：' + (data.visitName || '') + ' 様　作業：' + wt
      ).setFontSize(9);
      sh.setRowHeight(1, 26); sh.setRowHeight(2, 18); sh.setRowHeight(3, 8);
      headerRows = 3; // 計52px
    }

    const totalRows = headerRows + Math.ceil(chunk.length / 2) * SLOT_ROWS;
    for (let r = headerRows + 1; r <= totalRows; r++) sh.setRowHeight(r, 20);

    chunk.forEach(function(p, ci) {
      const col = (ci % 2) + 1;
      const baseRow = headerRows + Math.floor(ci / 2) * SLOT_ROWS + 1;
      const cap = (p.type ? '【' + p.type + '】' : '') + (p.memo ? ' ' + p.memo : '');
      if (cap) {
        sh.getRange(baseRow, col, CAP_ROWS, 1).merge()
          .setValue(cap).setFontSize(9).setBackground('#f5f5f5').setWrap(true);
      }
      const b64 = String(p.data).split(',').pop(); // dataURL・生base64どちらにも対応
      const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'photo' + (start + ci + 1) + '.jpg');
      const img = sh.insertImage(blob, col, baseRow + CAP_ROWS);
      // 枠内(横360×縦390)に縦横比を保って収める
      const nw = img.getWidth(), nh = img.getHeight();
      const scale = Math.min(360 / nw, 470 / nh);
      img.setWidth(Math.round(nw * scale)); img.setHeight(Math.round(nh * scale));
    });
  }
  SpreadsheetApp.flush();

  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() +
    '/export?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false';
  const pdfRes = UrlFetchApp.fetch(exportUrl, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  if (pdfRes.getResponseCode() !== 200) throw new Error('PDFエクスポート失敗: ' + pdfRes.getResponseCode());

  const fileName = data.date + '_' + data.visitName + '_' + docTitle + '.pdf';
  const file = getOrCreateFolder_('完了報告書(PDF)').createFile(pdfRes.getBlob().setName(fileName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  if (data.noLine) return; // 過去分の一括再生成時はLINEを送らない

  // 報告書PDFをオーナーのLINEへ送信（商品見積と同じ方式）
  try {
    const props = PropertiesService.getScriptProperties();
    const lineToken = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    const ownerLineId = props.getProperty('OWNER_LINE_USER_ID');
    if (NIPPO_LINE_ON && lineToken && ownerLineId) {
      const msg = '📋【業務日報】' + docTitle + 'ができました\n\n担当：' + (data.staff || '') +
        '\n現場：' + (data.visitName || '') + ' 様\n作業：' + (data.workType || '') +
        '\n\n' + file.getUrl();
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + lineToken },
        payload: JSON.stringify({ to: ownerLineId, messages: [{ type: 'text', text: msg }] }),
        muteHttpExceptions: true
      });
    }
  } catch (lineErr) {}
}

// ===== 毎日20時：当日分の業務日報をまとめてLINE送信 =====
// 日付セルを yyyy-MM-dd に正規化（文字列 "2026-07-23"／"2026/7/23"／Date のいずれでも対応）
function nippoYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  var s = String(v).trim();
  var m = s.match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}

// 指定日(yyyy-MM-dd)のまとめ本文を組み立てる。該当なしはnull
function buildNippoSummaryText_(ymd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('日報') || ss.getSheets()[0];
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var vals = sheet.getRange(2, 1, last - 1, 21).getValues(); // A〜U

  var lines = [];
  vals.forEach(function(r) {
    if (nippoYmd_(r[0]) !== ymd) return;
    var staff = String(r[1] || '').trim();
    var visit = String(r[3] || '').trim();
    var work = String(r[4] || '').trim();
    var remarks = String(r[19] || '').trim();
    var t = (staff || '（担当未記入）') + '｜' + (visit || '（現場名なし）') + '｜' + (work || '（作業未記入）');
    if (remarks) t += '　※' + remarks.replace(/\n/g, ' ');
    lines.push(t);
  });
  if (lines.length === 0) return null;

  var d = new Date(ymd + 'T00:00:00+09:00');
  var wd = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
  var mdLabel = (d.getMonth() + 1) + '月' + d.getDate() + '日(' + wd + ')';
  return '📋【業務日報まとめ】' + mdLabel + '（' + lines.length + '件）\n\n' +
    lines.map(function(t, i) { return (i + 1) + '. ' + t; }).join('\n');
}

/**
 * 日報で「次の訪問」を約束したとき、受注日報に1行足して担当者のカレンダーにも入れる。
 * 点検の最中に「冷蔵庫が壊れたから見に来て」と言われる場面のための入口。
 * 受注日報と同じ表に書くので、あとの流れ（完了・集計）は今までどおり。
 */
function 次の訪問を登録する_(nv, base) {
  if (!nv || !nv.訪問日時 || !nv.訪問担当メール || !nv.予定時間) return '';
  var eventId = '';

  /* 1. カレンダーに入れる */
  try {
    var cal = CalendarApp.getCalendarById(nv.訪問担当メール);
    if (cal) {
      var st = new Date(String(nv.訪問日時).replace(/-/g, '/'));
      var hm = String(nv.予定時間).match(/(\d+(\.\d+)?)/);
      var hours = hm ? parseFloat(hm[1]) : 1;
      var en = new Date(st.getTime() + hours * 60 * 60 * 1000);
      if (!isNaN(st.getTime())) {
        var title = base.visitName + ' 様' + (nv.内容 ? ' - ' + nv.内容 : '');
        var desc = ['【顧客名】 ' + base.visitName + ' 様',
                    '【受注種類】 訪問先で約束'];
        if (nv.内容) desc.push('【内容】 ' + nv.内容);
        desc.push('---');
        desc.push('業務日報から登録（' + base.staff + ' / ' + base.date + '）');
        var ev = cal.createEvent(title, st, en, { description: desc.join('\n') });
        eventId = ev.getId();
      }
    }
  } catch (calErr) {
    try {
      var dbg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('デバッグ')
             || SpreadsheetApp.getActiveSpreadsheet().insertSheet('デバッグ');
      dbg.appendRow([new Date(), '次の訪問カレンダーエラー', base.visitName, nv.訪問担当メール, calErr.toString()]);
    } catch (e2) {}
  }

  /* 2. 受注日報の表に1行足す（列の並びは受注日報アプリと同じ） */
  try {
    var jss = SpreadsheetApp.openById('1Cg5FACS_HqpAlchsU6O4ansNHc7GO_FOWAYhpX_72_w');
    var jsh = jss.getSheetByName('受注日報シート') || jss.getSheetByName('受注表')
           || jss.getSheetByName('受注日報') || jss.getSheets()[0];
    if (jsh) {
      jsh.appendRow([
        base.date,                                                   // 日付
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm'),     // 時刻
        base.staff,                                                  // 担当（受けた人）
        '訪問',                                                       // 種類
        base.visitName,                                              // 顧客名
        '', '',                                                       // 住所・電話（日報にはない）
        '', '', '', '',                                               // サービス種別〜商品詳細
        nv.訪問日時,                                                  // 訪問日時
        nv.訪問担当,                                                  // 訪問担当
        nv.予定時間,                                                  // 予定時間
        nv.内容 || '',                                                // 内容
        '',                                                           // 見込み商品
        eventId                                                       // カレンダーID
      ]);
    }
  } catch (jErr) {
    try {
      var dbg2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('デバッグ')
              || SpreadsheetApp.getActiveSpreadsheet().insertSheet('デバッグ');
      dbg2.appendRow([new Date(), '次の訪問 受注表エラー', base.visitName, '', jErr.toString()]);
    } catch (e3) {}
  }
  return eventId;
}

// 受注日報スプレッドシートの集計（件数・担当別）。読めなければ ok:false
function getJuchuTally_(ymd, ym) {
  var out = { ok: false, dayCount: 0, monCount: 0, dayStaff: {}, monStaff: {}, staffOrder: [] };
  try {
    var ss = SpreadsheetApp.openById('1Cg5FACS_HqpAlchsU6O4ansNHc7GO_FOWAYhpX_72_w');
    var sh = ss.getSheetByName('受注日報シート') || ss.getSheetByName('受注表') || ss.getSheetByName('受注日報') || ss.getSheets()[0];
    if (!sh || sh.getLastRow() < 2) return out;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues(); // A日付 C担当
    vals.forEach(function(r) {
      var y = nippoYmd_(r[0]); if (!y) return;
      var st = String(r[2] || '').trim() || '(未記入)';
      if (y.slice(0, 7) === ym) { out.monCount++; out.monStaff[st] = (out.monStaff[st] || 0) + 1; if (out.staffOrder.indexOf(st) < 0) out.staffOrder.push(st); }
      if (y === ymd) { out.dayCount++; out.dayStaff[st] = (out.dayStaff[st] || 0) + 1; }
    });
    out.ok = true;
  } catch (e) {}
  return out;
}

// ひまわり日報スプレッドシートの集計（面談・TEL・ポスティング）。読めなければ ok:false
function getHimawariTally_(ymd, ym) {
  var out = { ok: false, day: { v: 0, t: 0, p: 0, c: 0 }, mon: { v: 0, t: 0, p: 0, c: 0 } };
  try {
    var ss = SpreadsheetApp.openById('1WMsviRURr3Lwfi4zy4-dvH7DHO3uq4WXb4NBScwLqi4');
    var sh = ss.getSheetByName('全データ') || ss.getSheets()[0];
    if (!sh || sh.getLastRow() < 2) return out;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues(); // A日にち D面談 E TEL F ポスティング
    function nz(x) { var n = Number(String(x).replace(/[^\d.-]/g, '')); return isNaN(n) ? 0 : n; }
    vals.forEach(function(r) {
      var y = nippoYmd_(r[0]); if (!y) return;
      var v = nz(r[3]), t = nz(r[4]), p = nz(r[5]);
      if (y.slice(0, 7) === ym) { out.mon.v += v; out.mon.t += t; out.mon.p += p; out.mon.c++; }
      if (y === ymd) { out.day.v += v; out.day.t += t; out.day.p += p; out.day.c++; }
    });
    out.ok = true;
  } catch (e) {}
  return out;
}

// ===== カレンダー照合（日報の抜けチェック） =====
var CHECK_CALENDARS = [
  { id: 'yawata51@gmail.com', name: '高山' },
  { id: 'yahata.hattori@gmail.com', name: '服部' },
  { id: 'yahata.itou@gmail.com', name: '伊藤' }
];
// 訪問予定ではない予定（照合対象外）
var CAL_SKIP_WORDS = [
  // 社内の定型枠（毎日入る予定）
  '在庫確認', '当日急ぎ', '片付け', '日報', '伝票', '準備', '予定あり',
  // 訪問ではないもの
  'TEL】', '☎', '電話連絡', '連絡】', '見積り', '見積書', '発注', '入荷', '納期',
  // 実施しないもの
  'ｷｬﾝｾﾙ', 'キャンセル', '仮】', '中止', '延期',
  // その他
  '休', '定休', '会議', '打合', 'ミーティング', '研修', '朝礼', '出張', '誕生', '祝', 'メンテ'
];

// 名前の正規化：記号・敬称・空白を落として比較しやすくする
function calNorm_(s) {
  return String(s || '')
    .replace(/[☎✨★☆●○◆■▲]/g, '')
    .replace(/^[^぀-ヿ一-龯\w]*/, '')   // 先頭の記号列
    .replace(/[（(].*?[）)]/g, '')                        // 括弧内
    .replace(/(サ|工|配|修|現|見)\s*[】\]]/g, '')          // サ】など先頭タグ
    .replace(/[様さん御中]/g, '')
    .replace(/[\s　]/g, '')
    .trim();
}
// 予定タイトルから顧客名らしい部分を取り出す
function calCustomerName_(title) {
  var t = String(title || '');
  var m = t.match(/^(.*?)\s*様/);          // 「〇〇 様 …」形式
  var base = m ? m[1] : t;
  return calNorm_(base);
}

// 指定日について、カレンダーに予定があるのに日報が無いものを返す
function getCalendarMissing_(ymd) {
  var out = { ok: false, err: '', missing: [], noAccess: [], checked: 0 };
  try {
    var start = new Date(ymd + 'T00:00:00+09:00');
    var end = new Date(ymd + 'T23:59:59+09:00');

    // その日の日報（担当者・訪問先名）
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('日報') || ss.getSheets()[0];
    var rows = [];
    var last = sheet.getLastRow();
    if (last >= 2) {
      sheet.getRange(2, 1, last - 1, 5).getValues().forEach(function(r) {
        if (nippoYmd_(r[0]) !== ymd) return;
        rows.push({ staff: String(r[1] || '').trim(), visit: calNorm_(r[3]) });
      });
    }

    CHECK_CALENDARS.forEach(function(c) {
      var cal = null;
      try { cal = CalendarApp.getCalendarById(c.id); } catch (e) {}
      if (!cal) { out.noAccess.push(c.name + '(' + c.id + ')'); return; }

      var evs = [];
      try { evs = cal.getEvents(start, end); } catch (e) { out.noAccess.push(c.name); return; }

      evs.forEach(function(ev) {
        var title = String(ev.getTitle() || '');
        if (ev.isAllDayEvent()) return;                                   // 終日予定は対象外
        if (CAL_SKIP_WORDS.some(function(w) { return title.indexOf(w) >= 0; })) return;
        var cust = calCustomerName_(title);
        if (!cust || cust.length < 2) return;                             // 名前が取れないものは対象外
        out.checked++;

        // 同じ担当の日報に、その名前が含まれるか（前方一致・部分一致の両方向）
        var hit = rows.some(function(r) {
          if (!r.visit) return false;
          if (r.staff !== c.name) return false;
          return r.visit.indexOf(cust) >= 0 || cust.indexOf(r.visit) >= 0;
        });
        if (hit) return;
        // 別担当で提出されていれば抜けとしない
        var other = rows.filter(function(r) {
          return r.visit && (r.visit.indexOf(cust) >= 0 || cust.indexOf(r.visit) >= 0);
        });
        if (other.length) return;

        var hm = Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'HH:mm');
        out.missing.push({ staff: c.name, time: hm, title: title });
      });
    });
    out.ok = true;
  } catch (e) {
    out.err = String(e);
  }
  return out;
}

// 抜けチェックの結果を文字列にする（まとめLINE/テスト用）
function calMissingText_(ymd) {
  var r = getCalendarMissing_(ymd);
  if (!r.ok) return '⚠ カレンダー照合エラー: ' + r.err;
  var s = '';
  if (r.missing.length) {
    s += '⚠ 日報未提出の可能性（' + r.missing.length + '件）\n';
    r.missing.forEach(function(m) { s += '・' + m.staff + ' ' + m.time + ' ' + m.title + '\n'; });
  } else {
    s += '✅ 抜けなし（カレンダー予定' + r.checked + '件と照合）\n';
  }
  if (r.noAccess.length) s += '※未共有で見られないカレンダー: ' + r.noAccess.join('、') + '\n';
  return s;
}

// 指定日の日報を1件=カード形式（全項目）でPDF化し「日報まとめ(PDF)」へ保存。{file,count,label}返す（該当なしnull）
function buildDailyNippoPdf_(ymd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('日報') || ss.getSheets()[0];
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var vals = sheet.getRange(2, 1, last - 1, 21).getValues(); // A〜U

  var reports = vals.filter(function(r) { return nippoYmd_(r[0]) === ymd; });
  if (!reports.length) return null;

  function num(v) { var n = Number(String(v).replace(/[^\d.-]/g, '')); return isNaN(n) ? 0 : n; }
  // 金額は千円単位で入力・表示（例:200=200千円）。0/空欄は非表示（showZero=trueで0も表示）
  function sen(v, showZero) {
    var n = num(v);
    return (n === 0 && !showZero) ? '' : n.toLocaleString() + '千円';
  }
  function s(v) { return String(v == null ? '' : v).trim(); }

  var d = new Date(ymd + 'T00:00:00+09:00');
  var wd = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
  var mdLabel = (d.getMonth() + 1) + '月' + d.getDate() + '日(' + wd + ')';
  var ym = ymd.slice(0, 7);

  // 集計（0も表示）。工事金額=L(11)、売上=K(10)、見込み件数=Sあり
  var TYPES = ['工事', '配達', '修理', '現調', '点検', 'お困りごと対応', '商談', '集金', 'その他'];
  function tally(list) {
    var t = { count: 0, byType: {}, aircon: 0, other: 0, kouji: 0, people: 0, mikomi: 0, byStaff: {}, staffOrder: [] };
    TYPES.forEach(function(x) { t.byType[x] = 0; });
    list.forEach(function(r) {
      t.count++;
      var types = s(r[4]).split(/[、,・]/).map(function(x) { return x.trim(); });
      TYPES.forEach(function(x) { if (types.indexOf(x) >= 0) t.byType[x]++; });
      t.aircon += num(r[5]); t.other += num(r[6]);
      t.kouji += num(r[11]); t.people += num(r[13]);
      if (s(r[18])) t.mikomi++;
      var st = s(r[1]) || '（担当未記入）';
      if (!t.byStaff[st]) { t.byStaff[st] = { count: 0, kouji: 0 }; t.staffOrder.push(st); }
      t.byStaff[st].count++; t.byStaff[st].kouji += num(r[11]);
    });
    return t;
  }
  var dayT = tally(reports);
  var monthList = vals.filter(function(r) { return nippoYmd_(r[0]).slice(0, 7) === ym; });
  var monT = tally(monthList);

  var NAVY = '#1a3a5c', HEADBG = '#dfe7ef', THBG = '#eef2f6';
  var mm2 = (d.getMonth() + 1);

  // 集計メトリクス（項目を横並び・6項目ずつ折り返して大きく表示）。0も表示
  var SHORT = { 'お困りごと対応': '困りごと' };
  var METRICS = [{ h: '件数', d: dayT.count, m: monT.count }];
  TYPES.forEach(function(x) { METRICS.push({ h: (SHORT[x] || x), d: dayT.byType[x], m: monT.byType[x] }); });
  METRICS.push({ h: 'ｴｱｺﾝ点検', d: dayT.aircon, m: monT.aircon });
  METRICS.push({ h: '他点検', d: dayT.other, m: monT.other });
  METRICS.push({ h: '売上(千円)', d: dayT.kouji, m: monT.kouji });
  // 見込みは「見込」シートの件数で数える（メモから拾った分・写真だけの分も入るため）
  var mk = mikomiCount_(ymd);
  METRICS.push({ h: '見込み', d: Math.max(dayT.mikomi, mk.day), m: Math.max(monT.mikomi, mk.month) });

  var staffAll = monT.staffOrder.slice();
  dayT.staffOrder.forEach(function(st) { if (staffAll.indexOf(st) < 0) staffAll.push(st); });

  var temp = SpreadsheetApp.create('Temp_日報まとめ_' + ymd);

  // ===== 集計シート（横向きで見る前提・項目を1行に横並び） =====
  var ncol = METRICS.length + 1; // ラベル列＋14項目
  var sh = temp.getSheets()[0]; sh.setName('集計');
  sh.setColumnWidth(1, 74);
  for (var c = 2; c <= ncol; c++) sh.setColumnWidth(c, 68);
  var row = 1;
  sh.getRange(row, 1, 1, ncol).merge().setValue('業務日報まとめ　' + mdLabel + '　（' + reports.length + '件）')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground(NAVY).setFontColor('#ffffff');
  row++;
  sh.getRange(row, 1, 1, ncol).merge().setValue('📊 集計').setFontWeight('bold').setBackground(HEADBG); row++;
  sh.getRange(row, 1).setValue('').setBackground(THBG);
  METRICS.forEach(function(mt, i) { sh.getRange(row, 2 + i).setValue(mt.h).setFontWeight('bold').setBackground(THBG).setHorizontalAlignment('center'); });
  row++;
  sh.getRange(row, 1).setValue('本日').setFontWeight('bold').setBackground('#f4f6f8');
  METRICS.forEach(function(mt, i) { sh.getRange(row, 2 + i).setValue(mt.d.toLocaleString()).setHorizontalAlignment('center'); });
  row++;
  sh.getRange(row, 1).setValue('当月(' + mm2 + '月)').setFontWeight('bold').setBackground('#f4f6f8');
  METRICS.forEach(function(mt, i) { sh.getRange(row, 2 + i).setValue(mt.m.toLocaleString()).setHorizontalAlignment('center'); });
  row++;

  sh.getRange(row, 1, 1, ncol).merge().setValue('👤 担当別（件数 / 工事配達売上・千円）').setFontWeight('bold').setBackground(HEADBG); row++;
  var LAB = 3, M1 = 1 + LAB;
  var half = Math.floor((ncol - LAB) / 2);
  var M2 = M1 + half;
  var w2 = ncol - M2 + 1;
  sh.getRange(row, 1, 1, LAB).merge().setValue('担当').setFontWeight('bold').setBackground(THBG);
  sh.getRange(row, M1, 1, half).merge().setValue('本日').setFontWeight('bold').setBackground(THBG).setHorizontalAlignment('center');
  sh.getRange(row, M2, 1, w2).merge().setValue('当月(' + mm2 + '月)').setFontWeight('bold').setBackground(THBG).setHorizontalAlignment('center');
  row++;
  staffAll.forEach(function(st) {
    var dd = dayT.byStaff[st] || { count: 0, kouji: 0 };
    var mmv = monT.byStaff[st] || { count: 0, kouji: 0 };
    sh.getRange(row, 1, 1, LAB).merge().setValue(st);
    sh.getRange(row, M1, 1, half).merge().setValue(dd.count + '件 / ' + dd.kouji.toLocaleString() + '千円').setHorizontalAlignment('center');
    sh.getRange(row, M2, 1, w2).merge().setValue(mmv.count + '件 / ' + mmv.kouji.toLocaleString() + '千円').setHorizontalAlignment('center');
    row++;
  });

  // 受注日報 集計（別スプレッドシート）
  var ju = getJuchuTally_(ymd, ym);
  if (ju.ok) {
    sh.getRange(row, 1, 1, ncol).merge().setValue('📞 受注日報 集計（受注件数）').setFontWeight('bold').setBackground(HEADBG); row++;
    sh.getRange(row, 1, 1, LAB).merge().setValue('担当').setFontWeight('bold').setBackground(THBG);
    sh.getRange(row, M1, 1, half).merge().setValue('本日').setFontWeight('bold').setBackground(THBG).setHorizontalAlignment('center');
    sh.getRange(row, M2, 1, w2).merge().setValue('当月(' + mm2 + '月)').setFontWeight('bold').setBackground(THBG).setHorizontalAlignment('center');
    row++;
    sh.getRange(row, 1, 1, LAB).merge().setValue('合計').setFontWeight('bold');
    sh.getRange(row, M1, 1, half).merge().setValue(ju.dayCount + '件').setHorizontalAlignment('center');
    sh.getRange(row, M2, 1, w2).merge().setValue(ju.monCount + '件').setHorizontalAlignment('center');
    row++;
    ju.staffOrder.forEach(function(st) {
      sh.getRange(row, 1, 1, LAB).merge().setValue(st);
      sh.getRange(row, M1, 1, half).merge().setValue((ju.dayStaff[st] || 0) + '件').setHorizontalAlignment('center');
      sh.getRange(row, M2, 1, w2).merge().setValue((ju.monStaff[st] || 0) + '件').setHorizontalAlignment('center');
      row++;
    });
  }

  // ひまわり日報 集計（別スプレッドシート）
  var hm = getHimawariTally_(ymd, ym);
  if (hm.ok) {
    sh.getRange(row, 1, 1, ncol).merge().setValue('🌻 ひまわり日報 集計').setFontWeight('bold').setBackground(HEADBG); row++;
    var HMET = [
      { h: '報告数', d: hm.day.c, m: hm.mon.c },
      { h: '面談件数', d: hm.day.v, m: hm.mon.v },
      { h: 'TEL件数', d: hm.day.t, m: hm.mon.t },
      { h: 'ポスティング', d: hm.day.p, m: hm.mon.p }
    ];
    sh.getRange(row, 1).setValue('項目').setFontWeight('bold').setBackground(THBG);
    HMET.forEach(function(mt, i) { sh.getRange(row, 2 + i).setValue(mt.h).setFontWeight('bold').setBackground(THBG).setHorizontalAlignment('center'); });
    row++;
    sh.getRange(row, 1).setValue('本日').setFontWeight('bold').setBackground('#f4f6f8');
    HMET.forEach(function(mt, i) { sh.getRange(row, 2 + i).setValue(mt.d.toLocaleString()).setHorizontalAlignment('center'); });
    row++;
    sh.getRange(row, 1).setValue('当月').setFontWeight('bold').setBackground('#f4f6f8');
    HMET.forEach(function(mt, i) { sh.getRange(row, 2 + i).setValue(mt.m.toLocaleString()).setHorizontalAlignment('center'); });
    row++;
  }

  sh.getRange(1, 1, row - 1, ncol).setFontSize(11).setWrap(true).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  // ===== 2ページ目：明細 =====
  var sh2 = temp.insertSheet('明細');
  sh2.setColumnWidth(1, 34); sh2.setColumnWidth(2, 58); sh2.setColumnWidth(3, 130);
  sh2.setColumnWidth(4, 96); sh2.setColumnWidth(5, 40); sh2.setColumnWidth(6, 520);
  var r2 = 1;
  sh2.getRange(r2, 1, 1, 6).merge().setValue('明細　' + mdLabel + '　（' + reports.length + '件）')
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center').setBackground(NAVY).setFontColor('#ffffff'); r2++;
  sh2.getRange(r2, 1, 1, 6).setValues([['No', '担当', '現場', '作業', '人数', '内容・備考']])
    .setFontWeight('bold').setBackground(THBG).setHorizontalAlignment('center'); r2++;
  reports.forEach(function(r, idx) {
    var det = [];
    var insp = [];
    if (s(r[5]) && s(r[5]) !== '0') insp.push('エアコン' + s(r[5]));
    if (s(r[6]) && s(r[6]) !== '0') insp.push('その他' + s(r[6]));
    if (insp.length) det.push('点検: ' + insp.join(' / '));
    if (s(r[7])) det.push('商品: ' + s(r[7]));
    if (s(r[8])) det.push('工事種別: ' + s(r[8]));
    if (sen(r[11])) det.push('金額: 工事' + sen(r[11]) + (sen(r[12]) ? '（1人' + sen(r[12]) + '）' : ''));
    var pay = [];
    if (s(r[14])) pay.push(s(r[14]) + (sen(r[15]) ? ' ' + sen(r[15]) : ''));
    if (s(r[16])) pay.push('クレジット' + s(r[16]) + (s(r[17]) ? '（' + s(r[17]) + '）' : ''));
    if (pay.length) det.push('支払: ' + pay.join(' / '));
    if (s(r[18])) det.push('見込み商品: ' + s(r[18]));
    if (s(r[19])) det.push('備考: ' + s(r[19]));

    sh2.getRange(r2, 1).setValue(idx + 1).setHorizontalAlignment('center');
    sh2.getRange(r2, 2).setValue(s(r[1]));
    sh2.getRange(r2, 3).setValue(s(r[3]));
    sh2.getRange(r2, 4).setValue(s(r[4]) + (s(r[2]) ? '\n(' + s(r[2]) + ')' : ''));
    sh2.getRange(r2, 5).setValue(s(r[13])).setHorizontalAlignment('center');
    sh2.getRange(r2, 6).setValue(det.join('\n'));
    r2++;
  });
  sh2.getRange(1, 1, r2 - 1, 6).setFontSize(9).setWrap(true).setVerticalAlignment('top')
    .setBorder(true, true, true, true, true, true);
  SpreadsheetApp.flush();

  // 両シートとも横向きで1つのPDFに（1枚目=集計、2枚目=明細）
  var ssId = temp.getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId +
    '/export?format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3';
  var res = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  DriveApp.getFileById(ssId).setTrashed(true);
  if (res.getResponseCode() !== 200) throw new Error('PDFエクスポート失敗: ' + res.getResponseCode());
  var folder = getWritableFolder_('NIPPO_SUMMARY_FOLDER_ID', null, '日報まとめ(PDF)');
  var file = folder.createFile(res.getBlob().setName('日報まとめ_' + ymd + '.pdf'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { file: file, count: reports.length, label: mdLabel };
}

function sendDailyNippoSummary(ymd) {
  try {
    return sendDailyNippoSummaryCore_(ymd);
  } catch (err) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
        .appendRow(['まとめLINEエラー', new Date(), (ymd || '(today)'), String(err)]);
    } catch (e) {}
    throw err;
  }
}

function sendDailyNippoSummaryCore_(ymd) {
  // ★トリガーはイベントオブジェクトを第1引数に渡す。日付文字列(YYYY-MM-DD)以外は当日扱いにする
  var target = (typeof ymd === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(ymd))
    ? ymd : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var r = buildDailyNippoPdf_(target);

  // カレンダー照合（抜けチェック）。エラーでも本体の送信は止めない
  var calTxt = '';
  try {
    var cm = getCalendarMissing_(target);
    if (cm.ok && cm.missing.length) {
      calTxt = '\n\n⚠ 日報未提出の可能性（' + cm.missing.length + '件）\n' +
        cm.missing.map(function(m) { return '・' + m.staff + ' ' + m.time + ' ' + m.title; }).join('\n');
    }
  } catch (e) {}

  var body;
  if (r) {
    body = '📋【業務日報まとめ】' + r.label + '（' + r.count + '件）\n' +
      '※スマホは横向きで（1枚目=集計 / 2枚目=明細）\n\n' + r.file.getUrl() + calTxt;
  } else if (calTxt) {
    // 日報0件だがカレンダーに訪問予定あり＝丸ごと未提出の可能性
    var d0 = new Date(target + 'T00:00:00+09:00');
    var wd0 = ['日', '月', '火', '水', '木', '金', '土'][d0.getDay()];
    body = '📋【業務日報まとめ】' + (d0.getMonth() + 1) + '月' + d0.getDate() + '日(' + wd0 + ')\n' +
      '日報の提出はありません。' + calTxt;
  } else {
    return; // 日報も予定もなし＝送らない（週末等の無駄打ち防止）
  }
  var props = PropertiesService.getScriptProperties();
  var lineToken = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  var ownerLineId = props.getProperty('OWNER_LINE_USER_ID');
  if (NIPPO_LINE_ON && lineToken && ownerLineId) {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + lineToken },
      payload: JSON.stringify({ to: ownerLineId, messages: [{ type: 'text', text: body }] }),
      muteHttpExceptions: true
    });
  }
  return body;
}

// 直近で日報がある日を返す（プレビュー用）
function 直近の日報日_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('日報') || ss.getSheets()[0];
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var col = sheet.getRange(2, 1, last - 1, 1).getValues();
  var best = '';
  col.forEach(function(r) { var y = nippoYmd_(r[0]); if (y > best) best = y; });
  return best || null;
}

// 朝8時台に「前日分」のまとめを送る（20時以降に書き込まれた日報を拾い直す用）
function sendMorningNippoSummary() {
  var y = new Date();
  y.setDate(y.getDate() - 1);
  var ymd = Utilities.formatDate(y, 'Asia/Tokyo', 'yyyy-MM-dd');
  return sendDailyNippoSummary(ymd);
}

// まとめLINEのトリガーを設定（8時＝前日分／20時＝当日分。重複しないよう既存を消してから作る）
function 日報まとめ20時トリガー設定() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'sendDailyNippoSummary' || fn === 'sendMorningNippoSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyNippoSummary').timeBased().atHour(20).nearMinute(0).everyDays(1).create();
  ScriptApp.newTrigger('sendMorningNippoSummary').timeBased().atHour(8).nearMinute(0).everyDays(1).create();
  return 'OK: 毎日8時(前日分)と20時(当日分)にまとめLINEを送信します';
}

// トリガー経由の動作テスト：1分後に「昨日(7/23)分」をトリガーで送る
function fireTestTriggerNow_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailySummaryTriggerTest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailySummaryTriggerTest').timeBased().after(60 * 1000).create();
  return 'OK: 約1分後にトリガー経由でテスト送信します';
}
function sendDailySummaryTriggerTest() {
  try { sendDailyNippoSummary('2026-07-23'); }
  catch (e) {}
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailySummaryTriggerTest') ScriptApp.deleteTrigger(t);
  });
}

// ===== PDF生成の非同期化：ジョブをDriveに退避し、time-basedトリガーで別実行に処理させる =====
// 写真base64を含むためCache(100KB)/Properties(9KB)には入りきらない → Driveの一時JSONに保存する
function queueKoujiPdf_(data) {
  const folder = getWritableFolder_('KOUJI_QUEUE_FOLDER_ID', null, '日報PDFキュー');
  const jobFile = folder.createFile(Utilities.newBlob(
    JSON.stringify(data), 'application/json',
    'job_' + Date.now() + '_' + Math.random().toString(36).slice(2) + '.json'
  ));
  const props = PropertiesService.getScriptProperties();
  const q = JSON.parse(props.getProperty('KOUJI_PDF_QUEUE') || '[]');
  q.push(jobFile.getId());
  props.setProperty('KOUJI_PDF_QUEUE', JSON.stringify(q));
  // 約1秒後に別実行でキューを処理（doPostはこの後すぐ return され、体感が数秒に短縮される）
  ScriptApp.newTrigger('processKoujiPdfQueue').timeBased().after(1000).create();
}

// トリガーから呼ばれるワーカー：キューに溜まったジョブをまとめてPDF化＋LINE送信
function processKoujiPdfQueue() {
  // 使い捨てトリガーの後始末（同名トリガーをすべて削除。キュー中のジョブは下でまとめて処理する）
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processKoujiPdfQueue') ScriptApp.deleteTrigger(t);
  });
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) {}
  let q;
  try {
    q = JSON.parse(props.getProperty('KOUJI_PDF_QUEUE') || '[]');
    props.setProperty('KOUJI_PDF_QUEUE', '[]'); // 取り出したのでクリア（二重処理防止）
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  (q || []).forEach(function(fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const data = JSON.parse(file.getBlob().getDataAsString());
      createKoujiPDF(data);
      file.setTrashed(true);
    } catch (jobErr) {
      try {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
          .appendRow(['PDF作成エラー(裏)', new Date(), jobErr.toString()]);
      } catch (e) {}
    }
  });
  // 処理中に積まれた新規ジョブが残っていれば、もう一度自分を予約
  const remain = JSON.parse(props.getProperty('KOUJI_PDF_QUEUE') || '[]');
  if (remain.length) ScriptApp.newTrigger('processKoujiPdfQueue').timeBased().after(1000).create();
}

function getOrCreateFolder_(name) {
  return getWritableFolder_('KOUJI_PDF_FOLDER_ID', null, name);
}

// 写真フォルダ：実行アカウントが書き込める場所を自動確保（キャッシュ付き）
function getPhotoFolder_(type, preferredIds) {
  const id = preferredIds[type] || preferredIds['現調'];
  // ★本来のフォルダ（完了=17R6…／現調=1--gy…）に必ず直接保存する。
  // 以前は getWritableFolder_ のキャッシュが複製フォルダを指し、そちらへ保存し続けていた。
  try {
    const f = DriveApp.getFolderById(id);
    f.createFile(Utilities.newBlob('t', 'text/plain', '_writetest.txt')).setTrashed(true);
    return f;
  } catch (e) {
    // 本来フォルダに書けない時だけ従来の自動確保にフォールバック（写真を失わないため）
    const name = type === '完了' ? '完了_写真' : '現調(下見)_写真';
    return getWritableFolder_('FOLDER_' + name, id, name);
  }
}

// ===== 過去の業務日報の写真を新しい保存先へまとめる（分割実行） =====
// 1回で最大maxMove枚だけ移動し、残りがあれば more:true を返す。実行時間上限を避けるため。
function 写真を正規フォルダに移動(maxMove, onlyType) {
  const LIMIT = maxMove || 120;
  const START = Date.now();
  const TIME_BUDGET = 4 * 60 * 1000; // 4分で安全に打ち切り（上限6分の手前）
  // 移動先（新しい保存先）と、集めてくる元フォルダの一覧
  const PLAN_ALL = {
    '現調(下見)_写真': {
      target: '1Po0QN0S1VmpxmyEt1eCmAYUk0ynbQecd',        // 新・現調（商品見積と共通）
      oldIds: ['1--gy-_U-U-onPEZd3i8oWo290NDMZC9w']       // 旧・現調
    },
    '完了_写真': {
      target: '1vwliwaWcBaCNddsfyktCXqiMVwZWPA7W',        // 新・完了（完了報告書PDFと同じ）
      oldIds: ['17R6uVjwN6kei_VdLjNJX4LeA6TEINwc4']       // 旧・完了
    }
  };
  // onlyType='完了' なら完了だけ、'現調' なら現調だけ処理（並行実行できるように）
  const PLAN = {};
  if (onlyType === '完了') { PLAN['完了_写真'] = PLAN_ALL['完了_写真']; }
  else if (onlyType === '現調') { PLAN['現調(下見)_写真'] = PLAN_ALL['現調(下見)_写真']; }
  else { PLAN['現調(下見)_写真'] = PLAN_ALL['現調(下見)_写真']; PLAN['完了_写真'] = PLAN_ALL['完了_写真']; }
  let moved = 0, remainSeen = false;
  Object.keys(PLAN).forEach(function(name) {
    const targetId = PLAN[name].target;
    const target = DriveApp.getFolderById(targetId);

    // 集める元フォルダ＝旧フォルダID＋同名フォルダ（複製含む）。移動先自身は除外。
    const sources = [];
    const seen = {};
    PLAN[name].oldIds.forEach(function(id) {
      if (id === targetId || seen[id]) return;
      try { sources.push(DriveApp.getFolderById(id)); seen[id] = true; } catch (e) {}
    });
    const it = DriveApp.getFoldersByName(name);
    while (it.hasNext()) {
      const f = it.next();
      if (f.getId() === targetId || seen[f.getId()]) continue;
      sources.push(f); seen[f.getId()] = true;
    }

    sources.forEach(function(folder) {
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (file.getName() === '_writetest.txt') continue;
        if (moved >= LIMIT || (Date.now() - START) > TIME_BUDGET) { remainSeen = true; return; }
        try { file.moveTo(target); moved++; }
        catch (e1) { try { target.addFile(file); folder.removeFile(file); moved++; } catch (e2) {} }
      }
    });
  });
  if (!remainSeen) {
    // 全部終わったのでキャッシュ掃除（次回から新フォルダを使う）
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('FOLDER_完了_写真');
    props.deleteProperty('FOLDER_現調(下見)_写真');
  }
  const msg = (remainSeen ? '一部完了' : '全完了') + '：今回 ' + moved + ' 枚を移動' + (remainSeen ? '（残りあり→もう一度実行）' : '');
  Logger.log(msg);
  return JSON.stringify({ moved: moved, more: remainSeen, msg: msg });
}

// 書き込めるフォルダを確保：①キャッシュ→②優先ID→③同名検索→④新規作成
const FOLDER_CACHE_ = {};
function getWritableFolder_(cacheKey, preferredId, name) {
  if (FOLDER_CACHE_[cacheKey]) return FOLDER_CACHE_[cacheKey];
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(cacheKey);
  if (savedId) {
    try {
      const f = DriveApp.getFolderById(savedId);
      f.createFile(Utilities.newBlob('t', 'text/plain', '_writetest.txt')).setTrashed(true);
      FOLDER_CACHE_[cacheKey] = f;
      return f;
    } catch (e) { props.deleteProperty(cacheKey); }
  }
  const candidates = [];
  if (preferredId) { try { candidates.push(DriveApp.getFolderById(preferredId)); } catch (e) {} }
  const it = DriveApp.getFoldersByName(name);
  while (it.hasNext()) candidates.push(it.next());
  for (let i = 0; i < candidates.length; i++) {
    try {
      candidates[i].createFile(Utilities.newBlob('t', 'text/plain', '_writetest.txt')).setTrashed(true);
      props.setProperty(cacheKey, candidates[i].getId());
      FOLDER_CACHE_[cacheKey] = candidates[i];
      return candidates[i];
    } catch (e) {}
  }
  const nf = DriveApp.createFolder(name);
  props.setProperty(cacheKey, nf.getId());
  FOLDER_CACHE_[cacheKey] = nf;
  return nf;
}

// カナ正規化（半角カタカナ・ひらがな・全角カタカナ → すべて全角カタカナに統一）
function normalizeKana(str) {
  var result = str;
  
  // 1. 半角カタカナの濁音・半濁音（2文字→1文字）を先に変換
  result = result.replace(/ｶﾞ/g,'ガ').replace(/ｷﾞ/g,'ギ').replace(/ｸﾞ/g,'グ').replace(/ｹﾞ/g,'ゲ').replace(/ｺﾞ/g,'ゴ');
  result = result.replace(/ｻﾞ/g,'ザ').replace(/ｼﾞ/g,'ジ').replace(/ｽﾞ/g,'ズ').replace(/ｾﾞ/g,'ゼ').replace(/ｿﾞ/g,'ゾ');
  result = result.replace(/ﾀﾞ/g,'ダ').replace(/ﾁﾞ/g,'ヂ').replace(/ﾂﾞ/g,'ヅ').replace(/ﾃﾞ/g,'デ').replace(/ﾄﾞ/g,'ド');
  result = result.replace(/ﾊﾞ/g,'バ').replace(/ﾋﾞ/g,'ビ').replace(/ﾌﾞ/g,'ブ').replace(/ﾍﾞ/g,'ベ').replace(/ﾎﾞ/g,'ボ');
  result = result.replace(/ﾊﾟ/g,'パ').replace(/ﾋﾟ/g,'ピ').replace(/ﾌﾟ/g,'プ').replace(/ﾍﾟ/g,'ペ').replace(/ﾎﾟ/g,'ポ');
  result = result.replace(/ｳﾞ/g,'ヴ');
  
  // 2. 半角カタカナ（単独文字）→ 全角カタカナ
  var hw = 'ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ';
  var fw = 'ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン';
  for (var j = 0; j < hw.length; j++) {
    result = result.split(hw[j]).join(fw[j]);
  }
  
  // 3. 残った濁点・半濁点マークを除去
  result = result.replace(/ﾞ/g,'').replace(/ﾟ/g,'');
  
  // 4. ひらがな → 全角カタカナ
  result = result.replace(/[\u3041-\u3096]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) + 96);
  });
  
  return result;
}

// 顧客検索（全カナ形式対応、電話番号1〜5対応、カナ前方一致）
// 顧客マスタは商品見積SS（本家）を直接参照（2026-07-19一本化。全アプリ共通）
const CUSTOMER_MASTER_SS_ID = '1Pqb_DY3utvxKhTCIb4PJ0yNGra3hi1aSyFOHrydmuY0'; // 商品見積SS（本家・2026-07-19移設）
function searchCustomer(query) {
  const ss = SpreadsheetApp.openById(CUSTOMER_MASTER_SS_ID);
  const s = ss.getSheetByName("顧客マスタ") || ss.getSheetByName("顧客ﾏｽﾀ");
  if (!s || !query) return [];
  
  const d = s.getDataRange().getValues();
  const cleanQuery = String(query).replace(/-|\s|　/g, "");
  const kanaQuery = normalizeKana(cleanQuery);
  const results = [];

  for (let i = 1; i < d.length; i++) {
    const rowName = String(d[i][1]).replace(/\s|　/g, "");
    const rowKana = normalizeKana(String(d[i][2]).replace(/\s|　/g, ""));
    
    // 電話番号1〜5（D列〜H列 = index 3〜7）
    let telMatch = false;
    if (cleanQuery !== "") {
      for (let t = 3; t <= 7; t++) {
        const tel = String(d[i][t]).replace(/-/g, "");
        if (tel && tel.includes(cleanQuery)) {
          telMatch = true;
          break;
        }
      }
    }
    
    if (rowName.includes(cleanQuery) || rowKana.startsWith(kanaQuery) || telMatch) {
      results.push({
        tel: d[i][3],
        name: d[i][1],
        address: [d[i][15], d[i][16]].filter(v => v).join(' ')
      });
    }
  }
  return results;
}

function setup() {
  DriveApp.getFolderById("1--gy-_U-U-onPEZd3i8oWo290NDMZC9w");
  console.log("接続確認完了");
}