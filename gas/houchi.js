/*******************************************************
 * 放置案件の一覧
 *
 * 「行ったのに、その後なにもしていない」ものを日報から拾い出す。
 *   ① 見積を頼まれたのに、見積を出していない
 *   ② 見積を出したのに、工事をしていない
 *   ③ 一度行ったきり、そのままになっている
 *
 * 日報＝このスプレッドシート、見積台帳＝商品見積システムのスプレッドシート。
 * 2つを突き合わせて「その後の動き」があるかを見る。
 *******************************************************/

var HOUCHI_SHEET   = '放置案件';
var MITSUMORI_SSID = '1Pqb_DY3utvxKhTCIb4PJ0yNGra3hi1aSyFOHrydmuY0';   // 商品見積システム

/* 何日たったら「放置」とみなすか */
var HOUCHI_DAYS = {
  mitsumori: 7,    // 見積を頼まれてから7日
  kouji:    14,    // 見積を出してから14日
  houmon:   30     // 訪問してから30日
};

/* 日報の「作業」で、必ずあとが続くはずのもの。
   点検・集金・配達・お困りごと対応は「行って終わり」が普通なので入れない。
   現調は見積を出すために行くので、見積が無ければ放置 */
var HOUCHI_WORK = ['現調'];

/* これより古いものは追いかけない（日数） */
var HOUCHI_LIMIT = 180;

/* 見積を頼まれたと分かる言葉（作業・備考から探す） */
var HOUCHI_MITSUMORI_WORDS = ['見積', 'みつもり', 'ミツモリ', '検討', '金額を出す', '概算'];

/* ---------- 小さい道具 ---------- */

function hcNorm_(s) {
  return String(s || '').replace(/[\s　]/g, '').replace(/様$/, '');
}

function hcYmd_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  var t = String(v).trim().replace(/\//g, '-');
  var m = t.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function hcToday_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function hcDiffDays_(from, to) {
  if (!from || !to) return 0;
  return Math.floor((new Date(to) - new Date(from)) / 86400000);
}

/* ---------- 見積台帳を読む ---------- */

/**
 * 商品見積システムの「入力」シートから、お客様ごとの見積を集める。
 * 返すのは { 正規化した顧客名: [ {no, date, project, yotei, orderStatus, lost} ] }
 */
function hcMitsumori_() {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(MITSUMORI_SSID);
    var sh = ss.getSheetByName('入力');
    if (!sh || sh.getLastRow() < 2) return map;
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 21).getValues();
    v.forEach(function (r) {
      var cust = hcNorm_(r[2]);
      if (!cust) return;
      if (!map[cust]) map[cust] = [];
      map[cust].push({
        no:      String(r[0] || ''),
        date:    hcYmd_(r[1]),
        project: String(r[3] || ''),
        yotei:   hcYmd_(r[11]),                  // L列＝工事予定日
        order:   String(r[12] || ''),            // M列＝発注状態
        docType: String(r[8] || '')
      });
    });
  } catch (e) {}
  return map;
}

/**
 * 見積メモ（あとで見積）も見る。メモが残っていれば「忘れている」わけではない。
 */
function hcMemo_() {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(MITSUMORI_SSID);
    var sh = ss.getSheetByName('メモ');
    if (!sh || sh.getLastRow() < 2) return map;
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
    v.forEach(function (r) {
      if (String(r[2] || '') === '完了') return;
      var cust = hcNorm_(r[3]);
      if (cust) map[cust] = true;
    });
  } catch (e) {}
  return map;
}

/* ---------- 日報を読む ---------- */

/**
 * 日報から、お客様ごとの「最後に行った日」と、その内容を集める。
 */
function hcNippo_() {
  var out = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('日報') || ss.getSheets()[0];
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 21).getValues();
  v.forEach(function (r) {
    var cust = hcNorm_(r[3]);
    var ymd  = hcYmd_(r[0]);
    if (!cust || !ymd) return;
    var work = String(r[4] || '');
    var memo = String(r[19] || '');
    var cur = out[cust] || { name: String(r[3] || ''), last: '', staff: '', work: '', memo: '',
                             askedMitsumori: '', kouji: '' };
    // 工事・配達をした日を覚えておく（あとが続いている証拠になる）
    if (/工事|配達|取付|交換/.test(work)) {
      if (!cur.kouji || ymd > cur.kouji) cur.kouji = ymd;
    }
    // 見積を頼まれた日
    var line = work + ' ' + memo;
    if (HOUCHI_MITSUMORI_WORDS.some(function (w) { return line.indexOf(w) >= 0; })) {
      if (!cur.askedMitsumori || ymd > cur.askedMitsumori) cur.askedMitsumori = ymd;
    }
    if (!cur.last || ymd >= cur.last) {
      cur.last = ymd; cur.staff = String(r[1] || ''); cur.work = work; cur.memo = memo;
    }
    out[cust] = cur;
  });
  return out;
}

/* ---------- 本体 ---------- */

/**
 * 放置案件を拾って配列で返す。
 * kind は '見積まだ' / '工事まだ' / 'そのまま' の3つ。
 */
function 放置案件をさがす() {
  var today = hcNippo_ ? hcToday_() : '';
  var nippo = hcNippo_();
  var mits  = hcMitsumori_();
  var memo  = hcMemo_();
  var out = [];

  Object.keys(nippo).forEach(function (key) {
    var n = nippo[key];
    var list = mits[key] || [];

    /* ① 見積を頼まれたのに、そのあと見積が無い */
    if (n.askedMitsumori) {
      var after = list.filter(function (m) { return m.date && m.date >= n.askedMitsumori; });
      var days = hcDiffDays_(n.askedMitsumori, today);
      if (!after.length && !memo[key] && days >= HOUCHI_DAYS.mitsumori) {
        out.push({
          kind: '見積まだ', name: n.name, staff: n.staff,
          date: n.askedMitsumori, days: days,
          what: n.work, memo: n.memo,
          note: '日報で見積の話が出てから ' + days + '日'
        });
        return;   // ①に当てはまったら②③は見ない
      }
    }

    /* ② 見積は出したのに、工事をしていない */
    var machi = list.filter(function (m) {
      if (m.docType === 'sokubai') return false;          // その場販売は工事いらず
      if (m.order === '在庫') return false;
      if (!m.date) return false;
      if (m.yotei) return false;                          // 工事日が決まっている＝動いている
      return hcDiffDays_(m.date, today) >= HOUCHI_DAYS.kouji;
    });
    if (machi.length) {
      // いちばん古い見積を代表にする
      machi.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
      var m0 = machi[0];
      var d2 = hcDiffDays_(m0.date, today);
      // 見積のあとに工事をしていれば放置ではない
      if (!(n.kouji && n.kouji >= m0.date)) {
        out.push({
          kind: '工事まだ', name: n.name, staff: n.staff,
          date: m0.date, days: d2,
          what: m0.project, memo: 'No.' + m0.no,
          note: '見積を出してから ' + d2 + '日（工事日が未定）'
        });
        return;
      }
    }

    /* ③ 一度行ったきり、そのまま */
    var d3 = hcDiffDays_(n.last, today);
    var atoAri = list.some(function (m) { return m.date && m.date >= n.last; });
    if (!atoAri && !memo[key] && !n.kouji &&
        d3 >= HOUCHI_DAYS.houmon &&
        HOUCHI_WORK.some(function (w) { return n.work.indexOf(w) >= 0; })) {
      out.push({
        kind: 'そのまま', name: n.name, staff: n.staff,
        date: n.last, days: d3,
        what: n.work, memo: n.memo,
        note: '行ってから ' + d3 + '日、そのあと動きなし'
      });
    }
  });

  // 古すぎるものは実務で追えないので外す
  out = out.filter(function (x) { return x.days <= HOUCHI_LIMIT; });

  // 古い（放置日数が長い）ものを上に
  out.sort(function (a, b) { return b.days - a.days; });
  return out;
}

/**
 * 放置案件をシートに書き出す。
 */
function 放置案件をまとめる() {
  var rows = 放置案件をさがす();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOUCHI_SHEET);
  if (!sh) sh = ss.insertSheet(HOUCHI_SHEET);
  sh.clear();

  sh.getRange(1, 1).setValue('📌 放置案件　' + hcToday_() + ' 時点　' + rows.length + '件')
    .setFontSize(14).setFontWeight('bold');
  sh.getRange(2, 1).setValue('行ったのに、そのあと止まっているものです。上ほど日数がたっています。')
    .setFontSize(10).setFontColor('#666666');

  var HEAD = ['種類', 'お客様', '担当', '起点の日', '日数', '内容', 'メモ', '状況'];
  sh.getRange(4, 1, 1, HEAD.length).setValues([HEAD])
    .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
  sh.setFrozenRows(4);

  if (rows.length) {
    var vals = rows.map(function (r) {
      return [r.kind, r.name, r.staff, r.date, r.days, r.what, r.memo, r.note];
    });
    sh.getRange(5, 1, vals.length, HEAD.length).setValues(vals)
      .setVerticalAlignment('top').setWrap(true).setFontSize(11);

    var area = sh.getRange(5, 1, vals.length, HEAD.length);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A5="見積まだ"')
        .setBackground('#fdecea').setRanges([area]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A5="工事まだ"')
        .setBackground('#fff8e1').setRanges([area]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A5="そのまま"')
        .setBackground('#eef4fb').setRanges([area]).build()
    ]);
  } else {
    sh.getRange(5, 1).setValue('放置しているものはありません。').setFontColor('#1b5e20').setFontWeight('bold');
  }

  [80, 130, 70, 95, 55, 200, 200, 240].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return '放置案件 ' + rows.length + '件をまとめました';
}

/* 画面から呼ぶ用 */
function houchiList() { return 放置案件をさがす(); }
