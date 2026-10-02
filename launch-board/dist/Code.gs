/******************************************************************
 * BNI福岡博多 ローンチ管理 自動化スクリプト（Code.gs 1ファイル完結）
 *
 *  できること
 *   1. CCS追っかけ：担当（リストアップしたメンバー）がCCSまで追う。編集で最終接触日を自動入力／3点✓で正式化日を自動入力
 *   2. 正式化した人の「入会後行動リスト」（本人タスク＋DNAタスク）を行動マスタから自動作成
 *   3. 毎日決まった時刻に、遅れ・放置・LINE未参加をメールでアラート（LD・担当DNA・OM・本人）
 *   4. ローンチ進捗ボード（Webアプリ）：毎日の進捗の見える化・入力をサイトで（ファイル後半）
 *
 *  最初にやること：「使い方」シートの『GASの入れ方』を参照
 ******************************************************************/

// ===== ここだけ触ればOK =====
const CONFIG = {
  ALERT_HOURS: [8, 19],      // 毎日アラートを送る時刻（0〜23の数字をいくつでも）
  TEST_MODE: true,           // true の間は誰にも送らず、OM_EMAIL に全部まとめて届く（最初はこのまま確認）
  OM_EMAIL: '',              // 空欄ならこのスクリプトの持ち主のGmailに届く
  SEND_LAUNCH_MAIL: true,    // ローンチごとのまとめをLD（＋CC）へ
  SEND_EMAIL_DNA: true,      // CCS担当・担当DNAへ「自分の担当分だけ」のメール
  SEND_EMAIL_MEMBER: false,  // 本人へのリマインドメール（慣れてから true 推奨）
  SEND_OM_SUMMARY: true,     // OMへ全ローンチのまとめメール
  LINE_TASK_KEYWORD: 'LINEグループ', // この文字を含むタスクを「LINEグループ参加タスク」として扱う（初回定例会前は⚠表示）
  MAIL_SENDER_NAME: 'BNI福岡博多 ローンチ事務局',
  SPREADSHEET_ID: '154-KWaNppYa_8CSu8NmlD3fP1gns5qYjTzbp5E8ScAc',        // 進捗ボード（サイト）用。通常は空欄でOK（メニュー操作で自動保存）。うまく開けない時だけ、管理表URLの /d/ と /edit の間の文字を貼る
  SHEET: {
    CCS: 'CCS追っかけ', TANTOU: '担当別CCS', TASK: '入会後行動リスト', PROGRESS: '入会後進捗', MASTER: '行動マスタ',
    NOTIFY: '通知先', SETTING: '設定', ALOG: 'アラートログ',
  },
};
// ===========================

const FIRST_ROW = 5;  // 各シートのデータ開始行（4行目が見出し）
const CHECK = '✓';
// CCS追っかけの列番号
const C = { LD: 2, TAN: 3, NAME: 4, KANA: 5, CAT: 6, COMPANY: 7, TEL: 8, MAIL: 9, STAGE: 10, CCSDATE: 11, CCSDNA: 12,
  APP: 13, PAY: 14, AGREE: 15, JUDGE: 16, FORMAL: 17, LAST: 18, IDLE: 19, NEXT: 20, DUE: 21, STATUS: 22, ALERT: 23, TASKED: 24, MEMO: 25 };
// CCS追っかけの「現状」の分類（担当別CCSの集計に使用）
const STAGE_GROUP = {
  'お誘い中': ['説明会お誘い中', 'チャプター見学お誘い中', 'チャプター見学済'],
  'CCS調整中': ['CCS依頼中', 'CCS日程確定'],
  'CCS済': ['CCS済'],
};
// 入会後行動リストの列番号
const T = { NAME: 1, LD: 2, ORD: 3, TASK: 4, KIND: 5, OWNER: 6, DUE: 7, DONE: 8, DONEDATE: 9, STATE: 10, MEMO: 11, KEY: 12 };

/* ---------------- メニュー ---------------- */
function onOpen() {
  try { PropertiesService.getScriptProperties().setProperty('SS_ID', SpreadsheetApp.getActive().getId()); } catch (e) {}
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {
    // Apps Scriptの画面で▶実行した時はメニューを作れないため、ここで終了（エラーにはしない）
    console.log('onOpen はスプレッドシートを開いた時に自動で動きます。スプレッドシートを再読み込みしてください。');
    return;
  }
  ui.createMenu('🚀ローンチ管理')
    .addItem('アラートをプレビュー（送信しない）', 'previewAlerts')
    .addItem('アラートを今すぐ送信', 'sendAlertsNow')
    .addSeparator()
    .addItem('入会者の行動リスト・担当別CCSを更新', 'syncAndRefresh')
    .addItem('➕ 新しいローンチ（LD）を追加', 'addLaunch')
    .addSeparator()
    .addItem('🌐 サイトの利用者シートを用意', 'setupSiteUsers')
    .addItem('🌐 個人用リンクを作成・送信', 'makeSiteLinks')
    .addSeparator()
    .addItem('自動実行をON（毎日）', 'installTriggers')
    .addItem('自動実行をOFF', 'removeTriggers')
    .addItem('初期チェック', 'setup')
    .addToUi();
}

/* ---------------- 共通ヘルパー ---------------- */
function ss_() {
  let a = null;
  try { a = SpreadsheetApp.getActive(); } catch (e) {}
  if (a) return a;
  // Webアプリ（進捗ボード）から呼ばれた時は、保存しておいたIDで開く
  const id = CONFIG.SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('管理表の場所がまだ登録されていません。管理表のメニュー「🚀ローンチ管理」→「初期チェック」を一度実行してから、もう一度開いてください。');
  return SpreadsheetApp.openById(id);
}
// 管理表のIDを保存（メニュー操作のたびに実行）
function rememberSS_() {
  try { const a = SpreadsheetApp.getActive(); if (a) PropertiesService.getScriptProperties().setProperty('SS_ID', a.getId()); } catch (e) {}
}
function sh_(key) {
  const s = ss_().getSheetByName(CONFIG.SHEET[key]);
  if (!s) throw new Error('シート「' + CONFIG.SHEET[key] + '」が見つかりません');
  return s;
}
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function ymd_(d) { return Utilities.formatDate(d, tz_(), 'yyyy/MM/dd'); }
function md_(d) { return Utilities.formatDate(d, tz_(), 'M/d'); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v); }
function dayNum_(d) { const p = ymd_(d).split('/'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 864e5; }
function today_() { return Utilities.parseDate(ymd_(new Date()), tz_(), 'yyyy/MM/dd'); }
function addDays_(d, n) {
  const p = ymd_(d).split('/');
  const u = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return Utilities.parseDate(Utilities.formatDate(u, 'UTC', 'yyyy/MM/dd'), tz_(), 'yyyy/MM/dd');
}
function norm_(s) { return String(s || '').replace(/[\s　・,，.。()（）]/g, '').toLowerCase(); }
function str_(v) { return String(v === null || v === undefined ? '' : v).trim(); }
function lastRow_(sh, col) {
  const n = sh.getMaxRows() - FIRST_ROW + 1;
  if (n < 1) return FIRST_ROW - 1;
  const v = sh.getRange(FIRST_ROW, col, n, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (str_(v[i][0]) !== '') return FIRST_ROW + i;
  return FIRST_ROW - 1;
}
function omEmail_() { return CONFIG.OM_EMAIL || Session.getEffectiveUser().getEmail(); }

function settings_() {
  const v = sh_('SETTING').getRange('A4:B40').getValues();
  const m = {};
  v.forEach(r => { if (str_(r[0])) m[str_(r[0])] = r[1]; });
  const num = (k, d) => (typeof m[k] === 'number' ? m[k] : d);
  return {
    idle: num('CCS放置アラート（日）', 3),
    remind: num('本人リマインド再送間隔（日）', 2),
    warn: num('初回定例会前の警戒期間（日）', 7),
  };
}
function notify_() {
  const sh = sh_('NOTIFY');
  const n = Math.max(1, sh.getLastRow() - FIRST_ROW + 1);
  const v = sh.getRange(FIRST_ROW, 1, n, 7).getValues();
  const launch = {}, dna = {};
  v.forEach(r => {
    if (str_(r[0])) launch[str_(r[0])] = { ldMail: str_(r[1]), cc: str_(r[2]).replace(/[、，\s]+/g, ','), invite: str_(r[3]) };
    if (str_(r[5])) dna[str_(r[5])] = { mail: str_(r[6]) };
  });
  return { launch, dna };
}
function master_() {
  const sh = sh_('MASTER');
  const n = Math.max(1, sh.getLastRow() - FIRST_ROW + 1);
  return sh.getRange(FIRST_ROW, 1, n, 7).getValues()
    .filter(r => str_(r[1]) !== '')
    .map((r, i) => ({ ord: r[0] || i + 1, task: str_(r[1]), kind: str_(r[2]) || '本人', role: str_(r[3]) || '本人',
      days: Number(r[4]) || 0, remind: str_(r[5]) === '要' }));
}
function firstMeeting_(ld) {
  const s = ss_().getSheetByName(ld);
  if (!s) return null;
  const v = s.getRange('B8').getValue();
  return isDate_(v) ? v : null;
}

/* ---------------- 編集時の自動入力 ---------------- */
function onEdit(e) {
  try {
    const rg = e.range, sh = rg.getSheet(), name = sh.getName();
    const row = rg.getRow(), col = rg.getColumn();
    if (row < FIRST_ROW || rg.getNumRows() > 1 || rg.getNumColumns() > 1) return;

    if (name === CONFIG.SHEET.CCS) {
      const r = sh.getRange(row, 1, 1, C.MEMO).getValues()[0];
      // 見込み客名を入れた時、または現状・CCS日・✓・次アクション・期日を変えた時 → 最終接触日を今日に
      const touched = [C.STAGE, C.CCSDATE, C.APP, C.PAY, C.AGREE, C.NEXT, C.DUE].includes(col) && str_(e.value) !== '';
      const newName = col === C.NAME && str_(e.value) !== '' && !isDate_(r[C.LAST - 1]);
      if (touched || newName) sh.getRange(row, C.LAST).setValue(today_()).setNumberFormat('yyyy/mm/dd');
      // 3点✓ → 正式化日＋行動リスト作成
      if ([C.APP, C.PAY, C.AGREE].includes(col) &&
          r[C.APP - 1] === CHECK && r[C.PAY - 1] === CHECK && r[C.AGREE - 1] === CHECK) {
        if (!isDate_(r[C.FORMAL - 1])) sh.getRange(row, C.FORMAL).setValue(today_()).setNumberFormat('yyyy/mm/dd');
        const made = createTasksForRow_(sh, row);
        if (made) ss_().toast(r[C.NAME - 1] + 'さんの入会後行動リストを' + made + '件作成しました', '🎉正式プリコア', 5);
      }
    } else if (name === CONFIG.SHEET.TASK && col === T.DONE) {
      sh.getRange(row, T.DONEDATE).setValue(e.value === CHECK ? today_() : '').setNumberFormat('yyyy/mm/dd');
    }
  } catch (err) { console.error(err); }
}

/* ---------------- 入会後行動リストの作成 ---------------- */
function createTasksForRow_(ccsSh, row) {
  const r = ccsSh.getRange(row, 1, 1, C.MEMO).getValues()[0];
  const name = str_(r[C.NAME - 1]), ld = str_(r[C.LD - 1]);
  if (!name || !ld || str_(r[C.TASKED - 1]) === '作成済') return 0;
  const base = isDate_(r[C.FORMAL - 1]) ? r[C.FORMAL - 1] : today_();
  const tan = str_(r[C.TAN - 1]), ccsDna = str_(r[C.CCSDNA - 1]);
  const master = master_();
  if (!master.length) return 0;

  const tsh = sh_('TASK');
  const start = lastRow_(tsh, T.NAME) + 1;
  const rows = master.map(m => {
    const owner = m.role === '担当' ? (tan || ld)
      : (m.role === 'CCS実施' || m.role === 'CCS担当') ? (ccsDna || ld)
      : m.role === 'LD' ? ld : name;
    return [name, ld, m.ord, m.task, m.kind, owner, addDays_(base, m.days), '', '', '', '', ld + '|' + name + '|' + m.ord];
  });
  if (start + rows.length - 1 > tsh.getMaxRows()) tsh.insertRowsAfter(tsh.getMaxRows(), rows.length + 100);
  tsh.getRange(start, 1, rows.length, 12).setValues(rows);
  tsh.getRange(start, T.DUE, rows.length, 1).setNumberFormat('yyyy/mm/dd');
  tsh.getRange(start, T.DONEDATE, rows.length, 1).setNumberFormat('yyyy/mm/dd');
  const f = rows.map((_, i) => {
    const k = start + i;
    return [`=IF(D${k}="","",IF(H${k}="✓","完了",IF(G${k}="","",IF(G${k}<TODAY(),"期限超過",IF(G${k}=TODAY(),"今日まで","予定")))))`];
  });
  tsh.getRange(start, T.STATE, rows.length, 1).setFormulas(f);
  ccsSh.getRange(row, C.TASKED).setValue('作成済');
  return rows.length;
}

// 貼り付け等でonEditが動かなかった正式者も拾う
function syncAll_() {
  const sh = sh_('CCS');
  const last = lastRow_(sh, C.NAME);
  if (last < FIRST_ROW) return 0;
  const v = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, C.MEMO).getValues();
  let made = 0;
  v.forEach((r, i) => {
    const ok = r[C.APP - 1] === CHECK && r[C.PAY - 1] === CHECK && r[C.AGREE - 1] === CHECK;
    if (ok && str_(r[C.NAME - 1]) && str_(r[C.TASKED - 1]) !== '作成済') {
      if (!isDate_(r[C.FORMAL - 1])) sh.getRange(FIRST_ROW + i, C.FORMAL).setValue(today_()).setNumberFormat('yyyy/mm/dd');
      if (createTasksForRow_(sh, FIRST_ROW + i)) made++;
    }
  });
  return made;
}

/* ---------------- 入会後進捗（マトリクス） ---------------- */
function refreshProgress_() {
  const tsh = sh_('TASK');
  const last = lastRow_(tsh, T.NAME);
  const ps = ss_().getSheetByName(CONFIG.SHEET.PROGRESS) || ss_().insertSheet(CONFIG.SHEET.PROGRESS);
  const master = master_();
  const today = dayNum_(today_());

  if (ps.getMaxRows() >= 4) ps.getRange(4, 1, ps.getMaxRows() - 3, ps.getMaxColumns()).clear();
  ps.getRange('A1').setValue('入会後進捗（入会者 × 行動）').setFontSize(14).setFontWeight('bold').setFontColor('#1F3A5F');
  ps.getRange('A2').setValue('最終更新：' + Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm') +
    '　✓＝完了／✗＝期限超過／日付＝期限　【DNA】はDNAのタスク').setFontSize(9).setFontColor('#595959');

  const head = ['ローンチ', '氏名', '完了率'].concat(master.map(m => (m.kind === 'DNA' ? '【DNA】' : '') + m.task));
  const map = {}, order = [];
  if (last >= FIRST_ROW) {
    tsh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, T.KEY).getValues().forEach(r => {
      if (!str_(r[T.NAME - 1])) return;
      const k = r[T.LD - 1] + '|' + r[T.NAME - 1];
      if (!map[k]) { map[k] = { ld: r[T.LD - 1], name: r[T.NAME - 1], t: {} }; order.push(k); }
      map[k].t[str_(r[T.TASK - 1])] = r;
    });
  }
  const body = order.map(k => {
    const m = map[k]; let done = 0, total = 0;
    const cells = master.map(ms => {
      const r = m.t[ms.task];
      if (!r) return '－';
      total++;
      if (r[T.DONE - 1] === CHECK) { done++; return '✓'; }
      if (isDate_(r[T.DUE - 1])) return (dayNum_(r[T.DUE - 1]) < today ? '✗ ' : '') + md_(r[T.DUE - 1]);
      return '';
    });
    return [m.ld, m.name, total ? done / total : ''].concat(cells);
  });

  ps.getRange(4, 1, 1, head.length).setValues([head]).setBackground('#1F3A5F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  ps.setRowHeight(4, 48);
  if (body.length) {
    const rg = ps.getRange(5, 1, body.length, head.length);
    rg.setValues(body).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, '#BFBFBF', null);
    ps.getRange(5, 3, body.length, 1).setNumberFormat('0%');
    rg.setBackgrounds(body.map(r => r.map((c, j) =>
      j < 3 ? '#FFFFFF' : String(c).indexOf('✓') === 0 ? '#C6EFCE' : String(c).indexOf('✗') === 0 ? '#F8CBAD' : '#FFFFFF')));
  }
  ps.setColumnWidth(1, 110); ps.setColumnWidth(2, 110); ps.setColumnWidth(3, 70);
  if (master.length) ps.setColumnWidths(4, master.length, 115);
  ps.setFrozenRows(4); ps.setFrozenColumns(2);
}

/* ---------------- 新しいローンチ（LD）の追加 ---------------- */
// LDシートの複製＋ダッシュボード・設定のLD一覧・通知先への登録をまとめて行う
function addLaunch() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('新しいローンチを追加',
    'シート名にするローンチ名を入力してください。\n例：石川　／　同じLDが2つ持つ場合は「石川（ハルモニア）」のように区別してください。',
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const name = str_(r1.getResponseText()).replace(/[\/\\?*\[\]:']/g, '');
  if (!name) return;
  if (ss_().getSheetByName(name)) { ui.alert('「' + name + '」というシートはすでにあります。別の名前にしてください。'); return; }
  const r2 = ui.prompt('管轄エリア（チーム）', '例：JJ／Believe／チェキン（空欄でもOK）', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const team = str_(r2.getResponseText());

  // ① 複製元のLDシートを探す（設定のLD一覧に載っている既存シート）
  const st = sh_('SETTING');
  const head = st.getRange(3, 1, 1, st.getLastColumn()).getValues()[0];
  const ldCol = head.indexOf('LD一覧') + 1;
  if (!ldCol) { ui.alert('設定シートに「LD一覧」の列が見つかりません'); return; }
  const ldVals = st.getRange(4, ldCol, 20, 1).getValues().map(r => str_(r[0]));
  const tplName = ldVals.find(n => n && ss_().getSheetByName(n));
  if (!tplName) { ui.alert('複製元になるLDシートが見つかりません'); return; }

  // ② 空きがあるか先に確認（ダッシュボード・設定・通知先 各20枠）
  const db = ss_().getSheetByName('📊ダッシュボード');
  const dbVals = db ? db.getRange('A8:A27').getValues().map(r => str_(r[0])) : [];
  const dbIdx = dbVals.indexOf('');
  const ldIdx = ldVals.indexOf('');
  const nt = sh_('NOTIFY');
  const ntVals = nt.getRange('A5:A24').getValues().map(r => str_(r[0]));
  const ntIdx = ntVals.indexOf('');
  if (!db || dbIdx < 0 || ldIdx < 0 || ntIdx < 0) {
    ui.alert('登録できる枠（20ローンチ分）がいっぱいです。\n発足済・解散したローンチをダッシュボードA列・設定のLD一覧・通知先から消してから、もう一度実行してください。');
    return;
  }

  // ③ LDシートを複製して中身をリセット
  const tpl = ss_().getSheetByName(tplName);
  const sh = tpl.copyTo(ss_()).setName(name);
  const lastLd = ldVals.filter(n => n && ss_().getSheetByName(n)).map(n => ss_().getSheetByName(n).getIndex());
  ss_().setActiveSheet(sh);
  ss_().moveActiveSheet(Math.max.apply(null, lastLd) + 1);
  sh.getRange('A1').setValue(name + '｜ローンチチーム形成シート');
  sh.getRange('B4').setValue(team);
  sh.getRange('B5').setValue(name);
  sh.getRange('B6:B8').clearContent();
  sh.getRange('B9:D21').clearContent();
  sh.getRange('E9:E21').setValue('未着手');
  sh.getRange('B24').setValue('準備中');
  sh.getRange('B25:B28').clearContent();

  // ④ 各所に登録
  db.getRange(8 + dbIdx, 1).setValue(name);
  st.getRange(4 + ldIdx, ldCol).setValue(name);
  nt.getRange(5 + ntIdx, 1).setValue(name);

  ui.alert('「' + name + '」を追加しました✅\n\n' +
    '自動で登録した場所：LDシート／ダッシュボード／全体管理・週次推移（ダッシュボードに連動）／設定のLD一覧／通知先\n\n' +
    'あと入力してほしいところ：\n・' + name + 'シートのチャプター名・定例会日時・初回定例会目安・体制\n' +
    '・通知先シートのLDメール（とCC）\n・担当メンバーの名前とメール（通知先シート右側）');
}

/* ---------------- 担当別CCS（ローンチ × 担当メンバー） ---------------- */
function refreshTantou_() {
  const cs = sh_('CCS');
  const last = lastRow_(cs, C.NAME);
  const ts = ss_().getSheetByName(CONFIG.SHEET.TANTOU) || ss_().insertSheet(CONFIG.SHEET.TANTOU);
  const head = ['ローンチ', '担当（メンバー）', 'リストアップ', 'お誘い中', 'CCS調整中', 'CCS済', '入会（正式）', '辞退・保留', 'アラート', '最終更新'];
  const map = {}, order = [];
  if (last >= FIRST_ROW) {
    cs.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, C.MEMO).getValues().forEach(r => {
      if (!str_(r[C.NAME - 1])) return;
      const ld = str_(r[C.LD - 1]) || '（未入力）', tan = str_(r[C.TAN - 1]) || '（担当未入力）', stage = str_(r[C.STAGE - 1]);
      const k = ld + '|' + tan;
      if (!map[k]) { map[k] = [ld, tan, 0, 0, 0, 0, 0, 0, 0]; order.push(k); }
      const m = map[k], status = str_(r[C.STATUS - 1]);
      const formal = r[C.APP - 1] === CHECK && r[C.PAY - 1] === CHECK && r[C.AGREE - 1] === CHECK;
      m[2]++;
      if (status === '辞退' || status === '保留') { m[7]++; return; }
      if (formal) { m[5]++; m[6]++; }
      else if (STAGE_GROUP['CCS済'].indexOf(stage) >= 0) m[5]++;
      else if (STAGE_GROUP['CCS調整中'].indexOf(stage) >= 0) m[4]++;
      else if (STAGE_GROUP['お誘い中'].indexOf(stage) >= 0) m[3]++;
      if (str_(r[C.ALERT - 1])) m[8]++;
    });
  }
  const now = Utilities.formatDate(new Date(), tz_(), 'M/d HH:mm');
  const body = order.map(k => map[k].concat([now]))
    .sort((a, b) => a[0] === b[0] ? (b[5] - a[5]) || (b[2] - a[2]) : String(a[0]).localeCompare(String(b[0])));
  if (ts.getMaxRows() >= 4) ts.getRange(4, 1, ts.getMaxRows() - 3, Math.max(ts.getMaxColumns(), head.length)).clear();
  ts.getRange('A1').setValue('担当別CCS状況（ローンチ × 担当メンバー）').setFontSize(14).setFontWeight('bold').setFontColor('#1F3A5F');
  ts.getRange('A2').setValue('CCS済＝CCSまでつないだ人数（入会済みを含む）。アラート＝CCS追っかけでアラートが出ている人数').setFontSize(9).setFontColor('#595959');
  ts.getRange(4, 1, 1, head.length).setValues([head]).setBackground('#1F3A5F').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  if (body.length) {
    ts.getRange(5, 1, body.length, head.length).setValues(body).setBorder(true, true, true, true, true, true, '#BFBFBF', null)
      .setHorizontalAlignment('center');
    ts.getRange(5, 6, body.length, 1).setBackground('#DDEBF7').setFontWeight('bold');
    ts.getRange(5, 7, body.length, 1).setBackground('#C6EFCE').setFontWeight('bold');
    ts.getRange(5, 9, body.length, 1).setFontColor('#C00000');
  }
  ts.setColumnWidths(1, head.length, 95); ts.setColumnWidth(2, 130);
  ts.setFrozenRows(4);
}

/* ---------------- アラートの組み立て ---------------- */
function buildAlerts_() {
  const st = settings_(), nt = notify_(), today = dayNum_(today_());
  const L = {};        // ローンチ → { pre:[], post:[] }
  const D = {};        // DNA名 → [行]
  const members = [];  // 本人リマインド候補
  const get = ld => (L[ld] = L[ld] || { pre: [], post: [] });
  const addD = (who, line) => { if (who) (D[who] = D[who] || []).push(line); };

  // --- 入会前（CCS追っかけ）---
  const cs = sh_('CCS');
  const cl = lastRow_(cs, C.NAME);
  const info = {};  // ld|氏名 → {mail, tan}
  if (cl >= FIRST_ROW) {
    cs.getRange(FIRST_ROW, 1, cl - FIRST_ROW + 1, C.MEMO).getValues().forEach(r => {
      const name = str_(r[C.NAME - 1]), ld = str_(r[C.LD - 1]) || '（ローンチ未入力）';
      if (!name) return;
      const tan = str_(r[C.TAN - 1]), stage = str_(r[C.STAGE - 1]);
      info[ld + '|' + name] = { mail: str_(r[C.MAIL - 1]), tan };
      const status = str_(r[C.STATUS - 1]);
      const formal = r[C.APP - 1] === CHECK && r[C.PAY - 1] === CHECK && r[C.AGREE - 1] === CHECK;
      if (status === '辞退' || status === '保留' || formal) return;

      const is = [];
      if (isDate_(r[C.DUE - 1]) && dayNum_(r[C.DUE - 1]) < today) is.push('期日超過(' + md_(r[C.DUE - 1]) + ')');
      if (isDate_(r[C.LAST - 1]) && today - dayNum_(r[C.LAST - 1]) >= st.idle) is.push('放置' + (today - dayNum_(r[C.LAST - 1])) + '日');
      if (stage === 'CCS済' && r[C.APP - 1] !== CHECK) is.push('CCS後の申込待ち');
      if (r[C.APP - 1] === CHECK && r[C.PAY - 1] !== CHECK) is.push('振込待ち');
      if (r[C.PAY - 1] === CHECK && r[C.AGREE - 1] !== CHECK) is.push('同意書待ち');
      if (!str_(r[C.NEXT - 1])) is.push('次アクション未設定');
      if (!is.length) return;

      const company = str_(r[C.COMPANY - 1]);
      const line = `${name}${company ? '（' + company + '）' : ''} 担当:${tan || '未定'}${stage ? '｜現状:' + stage : ''}｜${is.join('・')}` +
        (str_(r[C.NEXT - 1]) ? '｜次:' + str_(r[C.NEXT - 1]) : '');
      get(ld).pre.push('・' + line);
      addD(tan, `[${ld}] ${line}`);
    });
  }

  // --- 入会後（行動リスト）---
  const ms = {}; master_().forEach(m => { ms[m.task] = m; });
  const ts = sh_('TASK');
  const tl = lastRow_(ts, T.NAME);
  const meet = {};
  if (tl >= FIRST_ROW) {
    ts.getRange(FIRST_ROW, 1, tl - FIRST_ROW + 1, T.KEY).getValues().forEach(r => {
      if (!str_(r[T.NAME - 1]) || r[T.DONE - 1] === CHECK || !isDate_(r[T.DUE - 1])) return;
      const due = dayNum_(r[T.DUE - 1]);
      if (due > today) return;
      const ld = str_(r[T.LD - 1]), name = str_(r[T.NAME - 1]), task = str_(r[T.TASK - 1]);
      const kind = str_(r[T.KIND - 1]), owner = str_(r[T.OWNER - 1]);
      const isLine = task.indexOf(CONFIG.LINE_TASK_KEYWORD) >= 0;
      if (!(ld in meet)) meet[ld] = firstMeeting_(ld);
      let warn = '';
      if (isLine && meet[ld]) {
        const left = dayNum_(meet[ld]) - today;
        if (left >= 0 && left <= st.warn) warn = `⚠初回定例会まであと${left}日 `;
      }
      const late = due < today ? `期限${md_(r[T.DUE - 1])}超過` : '今日まで';
      const line = `${warn}${name}：${task}（${kind === 'DNA' ? '担当:' + owner : '本人'}）${late}`;
      get(ld).post.push({ line: '・' + line, warn: warn ? 1 : 0 });

      const tan = (info[ld + '|' + name] || {}).tan;
      if (kind === 'DNA') addD(owner, `[${ld}] ${line}`);
      else addD(tan, `[${ld}] ${line}　→ 本人へ声かけをお願いします`);   // 本人タスクは紹介した担当メンバーが声かけ

      const m = ms[task];
      if (kind !== 'DNA' && m && m.remind) {
        members.push({ ld, name, task, due: r[T.DUE - 1], isLine, key: str_(r[T.KEY - 1]) || (ld + '|' + name + '|' + task),
          mail: (info[ld + '|' + name] || {}).mail || '' });
      }
    });
  }
  Object.keys(L).forEach(ld => L[ld].post.sort((a, b) => b.warn - a.warn));
  return { L, D, members, nt, st };
}

/* ---------------- メールの本文 ---------------- */
function sheetUrl_(key) { return ss_().getUrl() + '#gid=' + sh_(key).getSheetId(); }
function esc_(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function hl_(t) {  // 要注意ワードを赤太字に
  return esc_(t).replace(/(⚠[^ ]+|期日超過\([^)]*\)|放置\d+日|期限[\d/]+超過|今日まで|CCS後の申込待ち|振込待ち|同意書待ち|次アクション未設定)/g,
    '<b style="color:#C00000">$1</b>');
}
function listHtml_(title, items, color) {
  if (!items.length) return '';
  return `<div style="margin:14px 0 4px;font-weight:bold;color:${color}">${esc_(title)}（${items.length}件）</div>` +
    '<ul style="margin:0;padding-left:20px;line-height:1.7">' +
    items.map(t => `<li>${hl_(String(t).replace(/^・/, ''))}</li>`).join('') + '</ul>';
}
function button_(label, url) {
  return `<a href="${url}" style="display:inline-block;margin:6px 8px 0 0;background:#C8102E;color:#fff;padding:8px 16px;border-radius:6px;text-decoration:none;font-size:13px">${esc_(label)}</a>`;
}
function wrapHtml_(title, inner, lead, site) {
  return '<div style="font-family:\'Hiragino Sans\',\'Meiryo\',sans-serif;max-width:680px;color:#222;font-size:14px">' +
    `<div style="background:#1F3A5F;color:#fff;padding:12px 16px;border-radius:8px 8px 0 0;font-weight:bold">${esc_(title)}</div>` +
    '<div style="border:1px solid #ddd;border-top:0;padding:12px 16px 16px;border-radius:0 0 8px 8px">' +
    (lead ? `<p style="margin:0 0 6px">${esc_(lead).replace(/\n/g, '<br>')}</p>` : '') + inner +
    '<div style="margin-top:14px">' + (site ? button_('進捗ボードを開く', site) + '<br>' : '') +
    button_('CCS追っかけ（シート）', sheetUrl_('CCS')) + button_('入会後行動リスト（シート）', sheetUrl_('TASK')) + '</div>' +
    '<p style="margin:14px 0 0;font-size:12px;color:#777">対応したらシートの✓・最終接触日・次のアクションを更新してください。このメールは自動送信です。</p>' +
    '</div></div>';
}
function launchTitle_(ld) { return `【${ld}ローンチ】${Utilities.formatDate(new Date(), tz_(), 'M/d HH:mm')} 追っかけアラート`; }
function launchInner_(o) {
  return listHtml_('■ 入会前（CCSまでの追っかけ）', o.pre, '#1F3A5F') + listHtml_('■ 入会後の行動', o.post.map(p => p.line), '#2E75B6');
}
function launchText_(ld, o) {
  let s = launchTitle_(ld) + '\n';
  if (o.pre.length) s += `\n■入会前（CCS追っかけ）${o.pre.length}件\n` + o.pre.join('\n') + '\n';
  if (o.post.length) s += `\n■入会後の行動 ${o.post.length}件\n` + o.post.map(p => p.line).join('\n') + '\n';
  return s + '\nシート：' + ss_().getUrl();
}

// 送信内容のリストを作る（まだ送らない）。並び順＝送信の優先順
function composeMessages_() {
  const a = buildAlerts_();
  const out = [];
  const active = Object.keys(a.L).filter(ld => a.L[ld].pre.length || a.L[ld].post.length);
  const date = Utilities.formatDate(new Date(), tz_(), 'M/d');
  const links = siteLinks_();

  // ① OM：全体まとめ
  if (CONFIG.SEND_OM_SUMMARY && active.length) {
    const total = active.reduce((n, ld) => n + a.L[ld].pre.length + a.L[ld].post.length, 0);
    out.push({ to: omEmail_(), label: 'OM（全体まとめ）', subject: `【ローンチ全体】${date} 追っかけアラート ${total}件`,
      text: active.map(ld => launchText_(ld, a.L[ld])).join('\n\n────────\n\n'),
      html: wrapHtml_(`ローンチ全体 ${date} 追っかけアラート（${total}件）`,
        active.map(ld => `<div style="margin-top:16px;padding-top:8px;border-top:2px solid #DDEBF7;font-weight:bold;font-size:15px">${esc_(ld)}ローンチ</div>` + launchInner_(a.L[ld])).join(''), '', links.__om) });
  }
  // ② LD＋CC：ローンチまとめ
  if (CONFIG.SEND_LAUNCH_MAIL) {
    active.forEach(ld => {
      const n = a.nt.launch[ld] || {};
      if (!n.ldMail) return;
      out.push({ to: n.ldMail, cc: n.cc, label: ld + '（LD' + (n.cc ? '＋CC' : '') + '）', subject: `【ローンチアラート】${ld} ${date}`,
        text: launchText_(ld, a.L[ld]), html: wrapHtml_(launchTitle_(ld), launchInner_(a.L[ld]), '', links[ld]) });
    });
  }
  // ③ 各DNA：自分の担当分だけ
  if (CONFIG.SEND_EMAIL_DNA) {
    Object.keys(a.D).forEach(who => {
      const mail = (a.nt.dna[who] || {}).mail;
      if (!mail) return;
      const lead = `${who}さん\nいつもありがとうございます。今日声かけ・対応をお願いしたい方の一覧です。一緒に入会・定着まで伴走しましょう！`;
      out.push({ to: mail, label: who + '（担当）', subject: `【要対応】${who}さんの追っかけリスト ${a.D[who].length}件（${date}）`,
        text: lead + '\n\n・' + a.D[who].join('\n・') + '\n\n' + (links[who] ? '進捗ボード：' + links[who] : 'シート：' + ss_().getUrl()),
        html: wrapHtml_(`${who}さんの追っかけリスト（${a.D[who].length}件）`, listHtml_('■ 今日の対応', a.D[who], '#1F3A5F'), lead, links[who]) });
    });
  }
  // ④ 本人：リマインド
  if (CONFIG.SEND_EMAIL_MEMBER) {
    const log = alertLogMap_(), today = dayNum_(today_());
    a.members.forEach(m => {
      if (!m.mail) return;
      const key = 'member|' + m.key;
      if (log[key] !== undefined && today - log[key] < a.st.remind) return;
      const inv = (a.nt.launch[m.ld] || {}).invite;
      out.push({ to: m.mail, label: m.name + '様（本人）', key, subject: `【BNI】「${m.task}」のお願い`,
        text: `${m.name}様\n\nいつもありがとうございます。BNI福岡博多 ${m.ld}ローンチ事務局です。\n` +
          `「${m.task}」の期限（${md_(m.due)}）となっております。お手すきの際にご対応をお願いいたします。` +
          (m.isLine && inv ? `\n\nLINEグループへのご参加はこちらからお願いいたします。\n${inv}` : '') +
          '\n\nご不明な点は担当DNAまでお気軽にご連絡ください。\n引き続きよろしくお願いいたします。' });
    });
  }
  return out;
}

/* ---------------- 送信 ---------------- */
function sendAlerts_() {
  const out = composeMessages_();
  if (!out.length) return 0;
  const om = omEmail_();

  if (CONFIG.TEST_MODE) {
    const html = out.map(o => `<div style="margin:24px 0 6px;padding:6px 10px;background:#FFF2CC;font-size:13px">▼ 宛先：${esc_(o.label)}　${esc_(o.to)}` +
      (o.cc ? '　CC：' + esc_(o.cc) : '') + '<br>件名：' + esc_(o.subject) + '</div>' +
      (o.html || '<pre style="white-space:pre-wrap;font-family:inherit">' + esc_(o.text) + '</pre>')).join('');
    MailApp.sendEmail({ to: om, subject: `[TEST] ローンチアラート ${out.length}通分`, body: out.map(o => o.label + '\n' + o.text).join('\n\n=====\n\n'),
      htmlBody: '<p><b>TEST_MODE中です。</b>本番では以下がそれぞれの宛先に届きます。</p>' + html, name: CONFIG.MAIL_SENDER_NAME });
    logAlert_('TEST', om, 'テスト送信', out.length + '通分をOMに送信', '');
    return out.length;
  }

  let quota = MailApp.getRemainingDailyQuota(), sent = 0, skipped = [];
  out.forEach(o => {
    const need = 1 + (o.cc ? o.cc.split(',').filter(String).length : 0);
    if (quota < need) { skipped.push(o.label); return; }
    try {
      const opt = { to: o.to, subject: o.subject, body: o.text, name: CONFIG.MAIL_SENDER_NAME };
      if (o.cc) opt.cc = o.cc;
      if (o.html) opt.htmlBody = o.html;
      MailApp.sendEmail(opt);
      quota -= need; sent++;
      logAlert_('MAIL', o.to + (o.cc ? ' / CC:' + o.cc : ''), o.label, o.subject, o.key || '');
    } catch (err) {
      logAlert_('ERROR', o.to, o.label, String(err).slice(0, 200), '');
    }
  });
  if (skipped.length) logAlert_('SKIP', '', '送信上限のため未送信', skipped.join('、'), '');
  return sent;
}

function alertLogMap_() {
  const sh = sh_('ALOG'), last = sh.getLastRow(), m = {};
  if (last < FIRST_ROW) return m;
  sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 6).getValues().forEach(r => {
    if (str_(r[5]) && isDate_(r[0])) m[str_(r[5])] = Math.max(m[str_(r[5])] || 0, dayNum_(r[0]));
  });
  return m;
}
function logAlert_(ch, to, label, text, key) {
  const sh = sh_('ALOG');
  sh.getRange(Math.max(sh.getLastRow() + 1, FIRST_ROW), 1, 1, 6).setValues([[new Date(), ch, to, label, text, key]]);
}

/* ---------------- メニューから呼ぶ関数 ---------------- */
function previewAlerts() {
  rememberSS_();
  const out = composeMessages_();
  const html = out.length
    ? out.map(o => `<div style="margin:18px 0 6px;padding:6px 10px;background:#FFF2CC;font-size:12px">▼ ${esc_(o.label)}　${esc_(o.to)}` +
        (o.cc ? '　CC：' + esc_(o.cc) : '') + '<br>件名：' + esc_(o.subject) + '</div>' +
        (o.html || '<pre style="white-space:pre-wrap;font-size:12px">' + esc_(o.text) + '</pre>')).join('')
    : '<p>今日アラート対象はありません🎉</p>';
  const quota = MailApp.getRemainingDailyQuota();
  const head = `<p style="font-size:12px">${CONFIG.TEST_MODE ? '※TEST_MODE中：送信するとOMにだけまとめて届きます' : '※本番モード：送信すると各宛先に届きます'}` +
    `　／　メール${out.length}通　／　今日の残り送信可能数：${quota}</p>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(head + html).setWidth(760).setHeight(600), 'アラートのプレビュー');
}
function sendAlertsNow() {
  rememberSS_();
  const ui = SpreadsheetApp.getUi();
  const mode = CONFIG.TEST_MODE ? 'TEST_MODE（OMにだけ届きます）' : '本番（各宛先に届きます）';
  if (ui.alert('アラートを今すぐ送信しますか？\n' + mode, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  syncAll_();
  const n = sendAlerts_();
  ui.alert(n ? n + '通送信しました（詳細はアラートログ）' : '今日アラート対象はありません🎉');
}
function syncAndRefresh() {
  rememberSS_();
  const n = syncAll_();
  refreshProgress_();
  refreshTantou_();
  ss_().toast((n ? n + '人分の行動リストを作成し、' : '') + '入会後進捗・担当別CCSを更新しました', '完了', 5);
}
// 時間トリガーから毎日呼ばれる
function runScheduledAlerts() {
  syncAll_();
  refreshProgress_();
  refreshTantou_();
  sendAlerts_();
}
function installTriggers() {
  rememberSS_();
  removeTriggers_(true);
  CONFIG.ALERT_HOURS.forEach(h => {
    ScriptApp.newTrigger('runScheduledAlerts').timeBased().everyDays(1).atHour(h).nearMinute(0).inTimezone(tz_()).create();
  });
  SpreadsheetApp.getUi().alert('自動実行をONにしました：毎日 ' + CONFIG.ALERT_HOURS.join('時・') + '時ごろ' +
    (CONFIG.TEST_MODE ? '\n※TEST_MODE中なのでOMにだけ届きます' : ''));
}
function removeTriggers() { removeTriggers_(false); }
function removeTriggers_(silent) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'runScheduledAlerts') ScriptApp.deleteTrigger(t); });
  if (!silent) SpreadsheetApp.getUi().alert('自動実行をOFFにしました');
}
function setup() {
  rememberSS_();
  const need = ['CCS', 'TASK', 'MASTER', 'NOTIFY', 'SETTING'];
  const miss = need.filter(k => !ss_().getSheetByName(CONFIG.SHEET[k])).map(k => CONFIG.SHEET[k]);
  if (!ss_().getSheetByName(CONFIG.SHEET.ALOG)) {
    const s = ss_().insertSheet(CONFIG.SHEET.ALOG);
    s.getRange(4, 1, 1, 6).setValues([['日時', '種類', '宛先', '対象', '内容', 'キー']]).setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
  }
  if (!ss_().getSheetByName(CONFIG.SHEET.PROGRESS)) ss_().insertSheet(CONFIG.SHEET.PROGRESS);
  const nt = notify_();
  const noMail = Object.keys(nt.launch).filter(ld => !nt.launch[ld].ldMail);
  SpreadsheetApp.getUi().alert(
    (miss.length ? '⚠ 見つからないシート：' + miss.join('、') + '\n管理表（xlsx）から取り込んだファイルで実行してください。\n\n' : '✅ 必要なシートはそろっています\n\n') +
    'OMの受信先：' + omEmail_() + '\n' +
    'LDメール未入力のローンチ：' + (noMail.length ? noMail.join('、') : 'なし') + '\n' +
    'メール登録済みの担当メンバー・DNA：' + Object.keys(nt.dna).filter(k => nt.dna[k].mail).length + '人\n' +
    '今日の残り送信可能数：' + MailApp.getRemainingDailyQuota() + '通\n' +
    'モード：' + (CONFIG.TEST_MODE ? 'TEST_MODE（OMにだけ届く）' : '本番') + '\n\n' +
    '次は「アラートをプレビュー」で内容を確認してください。');
}

/******************************************************************
 * ローンチ進捗ボード（Webアプリ）
 *  ・データはこのスプレッドシートにそのまま貯まります
 *  ・誰が開いたかは「サイト利用者」シートの個人用リンク（?k=…）で判定します
 *  ・OM＝全ローンチ／LD＝自分のローンチ／担当＝自分の見込み客と、その人の入会後タスク
 ******************************************************************/
const SITE = {
  USERS: 'サイト利用者', DASH: '📊ダッシュボード', LOG: '日次ログ',
  DAILY_GOAL: 10,           // 声かけ目標の既定値（設定シートの「声かけ目標（1日）」が空のローンチに使う）
  TITLE: 'ローンチ進捗ボード',
};
const U = { NAME: 1, ROLE: 2, LAUNCH: 3, MAIL: 4, KEY: 5, URL: 6 };  // サイト利用者の列

// Webアプリの入口
function doGet(e) {
  const key = String((e && e.parameter && e.parameter.k) || '').trim();
  const boot = JSON.stringify({ key: key }).replace(/</g, '\\u003c');
  const html = PAGE_HTML_LINES.join('\n').replace('__BOOT__', boot);
  return HtmlService.createHtmlOutput(html)
    .setTitle(SITE.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 利用者の判定 ---------- */
function usersSheet_() { return ss_().getSheetByName(SITE.USERS); }
function readUsers_() {
  const sh = usersSheet_();
  if (!sh) return [];
  const last = lastRow_(sh, U.NAME);
  if (last < FIRST_ROW) return [];
  return sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 6).getValues().map((r, i) => ({
    row: FIRST_ROW + i, name: str_(r[0]), roleText: str_(r[1]), launches: str_(r[2]).split(/[、,，\n]/).map(str_).filter(String),
    mail: str_(r[3]).toLowerCase(), key: str_(r[4]), url: str_(r[5]),
  })).filter(u => u.name);
}
function roleOf_(t) { return /OM|オペレーション|管理/i.test(t) ? 'om' : /LD|ローンチ|ディレクター/i.test(t) ? 'ld' : 'tan'; }
function auth_(key) {
  const users = readUsers_();
  let u = key ? users.find(x => x.key && x.key === key) : null;
  if (!u) {
    let mail = '';
    try { mail = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
    if (mail) u = users.find(x => x.mail && x.mail === mail);
  }
  if (!u) return null;
  const role = roleOf_(u.roleText);
  return { name: u.name, role: role, launches: role === 'ld' ? (u.launches.length ? u.launches : [u.name]) : u.launches };
}
function denied_() { return { error: 'このリンクでは開けません。管理者から届いた「あなた専用のリンク」から開いてください。' }; }

/* ---------- 読み込み ---------- */
function launchNames_() {
  const db = ss_().getSheetByName(SITE.DASH);
  if (!db) return [];
  return db.getRange('A8:A27').getValues().map(r => str_(r[0])).filter(n => n && ss_().getSheetByName(n));
}
function readLaunch_(name) {
  const v = ss_().getSheetByName(name).getRange('B4:B28').getValues().map(r => r[0]);
  const d = x => (isDate_(x) ? ymd_(x) : '');
  return { ld: name, team: str_(v[0]), ch: str_(v[2]), firstMtg: d(v[4]), phase: str_(v[20]) || '準備中',
    start: d(v[21]), target: typeof v[22] === 'number' ? v[22] : Number(v[22]) || null, issue: str_(v[23]) };
}
function siteSettings_() {
  const v = sh_('SETTING').getRange('A4:B40').getValues();
  const m = {}; v.forEach(r => { if (str_(r[0])) m[str_(r[0])] = r[1]; });
  const n = (k, d) => (typeof m[k] === 'number' ? m[k] : d);
  return { p1: n('フェーズ1期限（日）', 42), p2: n('フェーズ2期限（日）', 28), yline: n('黄色判定ライン（目標ペース比）', 0.7),
    near: n('期限間近アラート（残日数）', 7), nolog: n('ログ未入力アラート（日）', 3), idle: n('CCS放置アラート（日）', 3),
    weekStart: isDate_(m['週次推移の開始日（月曜）']) ? ymd_(m['週次推移の開始日（月曜）']) : '', dailyGoal: SITE.DAILY_GOAL, goals: goalMap_() };
}
// 設定シートの「LD一覧」と同じ行にある「声かけ目標（1日）」を、ローンチ名 → 人数で返す
function goalMap_() {
  const st = sh_('SETTING');
  const head = st.getRange(3, 1, 1, st.getLastColumn()).getValues()[0].map(str_);
  const ldCol = head.indexOf('LD一覧') + 1, gCol = head.indexOf('声かけ目標（1日）') + 1;
  const m = {};
  if (!ldCol || !gCol) return m;
  const v = st.getRange(4, 1, 20, Math.max(ldCol, gCol)).getValues();
  v.forEach(r => { const n = str_(r[ldCol - 1]), g = Number(r[gCol - 1]); if (n && r[gCol - 1] !== '' && g >= 0) m[n] = g; });
  return m;
}
function stageList_() {
  const st = sh_('SETTING');
  const head = st.getRange(3, 1, 1, st.getLastColumn()).getValues()[0].map(str_);
  const col = head.indexOf('現状') + 1;
  const list = col ? st.getRange(4, col, 20, 1).getValues().map(r => str_(r[0])).filter(String) : [];
  return list.length ? list : ['リストアップ済', '説明会お誘い中', 'チャプター見学お誘い中', 'チャプター見学済', 'CCS依頼中', 'CCS日程確定', 'CCS済'];
}
function readPros_() {
  const cs = sh_('CCS');
  const last = lastRow_(cs, C.NAME);
  if (last < FIRST_ROW) return [];
  const d = x => (isDate_(x) ? ymd_(x) : '');
  return cs.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, C.MEMO).getValues().map((r, i) => ({
    row: FIRST_ROW + i, ld: str_(r[C.LD - 1]), tan: str_(r[C.TAN - 1]), name: str_(r[C.NAME - 1]), co: str_(r[C.COMPANY - 1]),
    cat: str_(r[C.CAT - 1]), stage: str_(r[C.STAGE - 1]), app: r[C.APP - 1] === CHECK, pay: r[C.PAY - 1] === CHECK, agr: r[C.AGREE - 1] === CHECK,
    formal: d(r[C.FORMAL - 1]), last: d(r[C.LAST - 1]), next: str_(r[C.NEXT - 1]), due: d(r[C.DUE - 1]), status: str_(r[C.STATUS - 1]) || '進行中',
  })).filter(p => p.name);
}
function readTasks_() {
  const ts = sh_('TASK');
  const last = lastRow_(ts, T.NAME);
  if (last < FIRST_ROW) return [];
  return ts.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, T.KEY).getValues().map((r, i) => ({
    row: FIRST_ROW + i, name: str_(r[T.NAME - 1]), ld: str_(r[T.LD - 1]), ord: Number(r[T.ORD - 1]) || 0, task: str_(r[T.TASK - 1]),
    kind: str_(r[T.KIND - 1]), owner: str_(r[T.OWNER - 1]), due: isDate_(r[T.DUE - 1]) ? ymd_(r[T.DUE - 1]) : '', done: r[T.DONE - 1] === CHECK,
  })).filter(t => t.name && t.task);
}
function readLogs_(fromDay) {
  const sh = ss_().getSheetByName(SITE.LOG);
  if (!sh) return [];
  const last = lastRow_(sh, 1);
  if (last < FIRST_ROW) return [];
  return sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 10).getValues()
    .filter(r => isDate_(r[0]) && str_(r[1]) && dayNum_(r[0]) >= fromDay)
    .map(r => ({ d: ymd_(r[0]), ld: str_(r[1]), by: str_(r[2]), koe: Number(r[3]) || 0, info: Number(r[4]) || 0, men: Number(r[5]) || 0, memo: str_(r[9]) }));
}

/* ---------- 見える範囲の判定 ---------- */
function canSeePro_(me, p) {
  return me.role === 'om' || (me.role === 'ld' && me.launches.indexOf(p.ld) >= 0) || (me.role === 'tan' && p.tan === me.name);
}
function canSeeTask_(me, t, myRecruits) {
  if (me.role === 'om') return true;
  if (me.role === 'ld') return me.launches.indexOf(t.ld) >= 0;
  return myRecruits[t.ld + '|' + t.name] || t.owner === me.name;
}

function webGetData(key) {
  try {
    const me = auth_(key);
    if (!me) return denied_();
    const today = today_(), set = siteSettings_();
    const all = launchNames_();
    const pros = readPros_().filter(p => canSeePro_(me, p));
    const recruits = {}; pros.forEach(p => { recruits[p.ld + '|' + p.name] = true; });
    const tasks = readTasks_().filter(t => canSeeTask_(me, t, recruits));
    let launches = [];
    if (me.role === 'om') launches = all;
    else if (me.role === 'ld') launches = all.filter(n => me.launches.indexOf(n) >= 0);
    let from = dayNum_(addDays_(today, -13));
    if (set.weekStart) from = Math.min(from, dayNum_(Utilities.parseDate(set.weekStart, tz_(), 'yyyy/MM/dd')));
    const logs = me.role === 'tan' ? [] : readLogs_(from).filter(x => launches.indexOf(x.ld) >= 0);
    return { me: me, today: ymd_(today), set: set, stages: stageList_(), allLaunches: all,
      launches: launches.map(readLaunch_), logs: logs, pros: pros, tasks: tasks };
  } catch (err) { return { error: '読み込みに失敗しました：' + err.message }; }
}

/* ---------- 書き込み ---------- */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { error: 'ほかの人が保存中です。少し待ってからもう一度お試しください。' };
  try { return fn(); } catch (err) { return { error: '保存に失敗しました：' + err.message }; } finally { lock.releaseLock(); }
}
function toDate_(s) { return s ? Utilities.parseDate(String(s), tz_(), 'yyyy/MM/dd') : ''; }

function webSaveLog(key, o) {
  return withLock_(() => {
    const me = auth_(key);
    if (!me) return denied_();
    if (me.role === 'tan') return { error: '日次実績の入力はLD・OMが行います。' };
    if (me.role === 'ld' && me.launches.indexOf(o.ld) < 0) return { error: '自分のローンチだけ入力できます。' };
    const sh = ss_().getSheetByName(SITE.LOG);
    const date = toDate_(o.date);
    const last = lastRow_(sh, 1);
    let row = 0;
    if (last >= FIRST_ROW) {
      const v = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 2).getValues();
      for (let i = 0; i < v.length; i++) {
        if (isDate_(v[i][0]) && ymd_(v[i][0]) === ymd_(date) && str_(v[i][1]) === o.ld) { row = FIRST_ROW + i; break; }
      }
    }
    if (!row) row = last + 1;
    sh.getRange(row, 1, 1, 6).setValues([[date, o.ld, me.name, Number(o.koe) || 0, Number(o.info) || 0, Number(o.men) || 0]]);
    sh.getRange(row, 1).setNumberFormat('yyyy/mm/dd');
    sh.getRange(row, 10).setValue(str_(o.memo));
    return { ok: true };
  });
}

function webUpdateProspect(key, row, patch) {
  return withLock_(() => {
    const me = auth_(key);
    if (!me) return denied_();
    const cs = sh_('CCS');
    row = Number(row);
    const r = cs.getRange(row, 1, 1, C.MEMO).getValues()[0];
    const p = { ld: str_(r[C.LD - 1]), tan: str_(r[C.TAN - 1]), name: str_(r[C.NAME - 1]) };
    if (!p.name || !canSeePro_(me, p)) return { error: 'この見込み客は更新できません。' };
    if ('stage' in patch) cs.getRange(row, C.STAGE).setValue(str_(patch.stage));
    ['app', 'pay', 'agr'].forEach(k => {
      if (k in patch) cs.getRange(row, { app: C.APP, pay: C.PAY, agr: C.AGREE }[k]).setValue(patch[k] ? CHECK : '');
    });
    if ('next' in patch) cs.getRange(row, C.NEXT).setValue(str_(patch.next));
    if ('due' in patch) cs.getRange(row, C.DUE).setValue(toDate_(patch.due)).setNumberFormat('yyyy/mm/dd');
    if ('status' in patch) cs.getRange(row, C.STATUS).setValue(str_(patch.status));
    cs.getRange(row, C.LAST).setValue(today_()).setNumberFormat('yyyy/mm/dd');
    const n = cs.getRange(row, 1, 1, C.MEMO).getValues()[0];
    const formal = n[C.APP - 1] === CHECK && n[C.PAY - 1] === CHECK && n[C.AGREE - 1] === CHECK;
    let created = 0;
    if (formal) {
      if (!isDate_(n[C.FORMAL - 1])) cs.getRange(row, C.FORMAL).setValue(today_()).setNumberFormat('yyyy/mm/dd');
      created = createTasksForRow_(cs, row);
    }
    return { ok: true, formal: formal && created > 0 };
  });
}

function webAddProspect(key, o) {
  return withLock_(() => {
    const me = auth_(key);
    if (!me) return denied_();
    const name = str_(o.name);
    if (!name) return { error: 'お名前を入れてください。' };
    const ld = str_(o.ld);
    if (launchNames_().indexOf(ld) < 0) return { error: 'ローンチを選んでください。' };
    if (me.role === 'ld' && me.launches.indexOf(ld) < 0) return { error: '自分のローンチにだけ追加できます。' };
    const tan = me.role === 'tan' ? me.name : (str_(o.tan) || me.name);
    const cs = sh_('CCS');
    const row = lastRow_(cs, C.NAME) + 1;
    if (row > cs.getMaxRows()) cs.insertRowsAfter(cs.getMaxRows(), 100);
    // B〜O列（数式のない範囲）とR・T・U・V列だけに書き込む
    cs.getRange(row, C.LD, 1, C.AGREE - C.LD + 1).setValues([[ld, tan, name, '', str_(o.cat), str_(o.co), str_(o.tel), '', str_(o.stage), '', '', '', '', '']]);
    cs.getRange(row, C.LAST).setValue(today_()).setNumberFormat('yyyy/mm/dd');
    cs.getRange(row, C.NEXT).setValue(str_(o.next));
    if (o.due) cs.getRange(row, C.DUE).setValue(toDate_(o.due)).setNumberFormat('yyyy/mm/dd');
    cs.getRange(row, C.STATUS).setValue('進行中');
    return { ok: true, row: row };
  });
}

function webSetTaskDone(key, row, done) {
  return withLock_(() => {
    const me = auth_(key);
    if (!me) return denied_();
    const ts = sh_('TASK');
    row = Number(row);
    const r = ts.getRange(row, 1, 1, T.KEY).getValues()[0];
    const t = { name: str_(r[T.NAME - 1]), ld: str_(r[T.LD - 1]), owner: str_(r[T.OWNER - 1]) };
    const recruits = {};
    if (me.role === 'tan') readPros_().filter(p => p.tan === me.name).forEach(p => { recruits[p.ld + '|' + p.name] = true; });
    if (!t.name || !canSeeTask_(me, t, recruits)) return { error: 'このタスクは更新できません。' };
    ts.getRange(row, T.DONE).setValue(done ? CHECK : '');
    ts.getRange(row, T.DONEDATE).setValue(done ? today_() : '').setNumberFormat('yyyy/mm/dd');
    return { ok: true };
  });
}

/* ---------- 利用者シートと個人用リンク（メニューから実行） ---------- */
function setupSiteUsers() {
  const ui = SpreadsheetApp.getUi();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss_().getId());
  let sh = usersSheet_();
  if (!sh) {
    sh = ss_().insertSheet(SITE.USERS);
    sh.getRange('A1').setValue('サイト利用者（ローンチ進捗ボードを開ける人）').setFontSize(14).setFontWeight('bold').setFontColor('#1F3A5F');
    sh.getRange('A2').setValue('役割：OM＝全ローンチ／LD＝担当ローンチ／担当＝自分の見込み客だけ。名前はCCS追っかけの「担当」、LDシート名と同じ表記に。キーとURLは自動で入ります（キーを消すと、そのリンクは使えなくなります）')
      .setFontSize(9).setFontColor('#595959');
    sh.getRange(4, 1, 1, 6).setValues([['名前', '役割（OM／LD／担当）', '担当ローンチ（LDのみ・複数は「、」区切り）', 'メール', 'キー（自動）', '個人用リンク（自動）']])
      .setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF').setWrap(true);
    sh.setColumnWidths(1, 6, 150); sh.setColumnWidth(3, 220); sh.setColumnWidth(6, 360); sh.setFrozenRows(4);
    // 通知先シートから名前とメールを取り込み
    const nt = notify_(), rows = [['OM', 'OM', '', omEmail_()]];
    Object.keys(nt.launch).forEach(ld => rows.push([ld, 'LD', ld, nt.launch[ld].ldMail]));
    Object.keys(nt.dna).forEach(n => { if (!nt.launch[n]) rows.push([n, '担当', '', nt.dna[n].mail]); });
    sh.getRange(FIRST_ROW, 1, rows.length, 4).setValues(rows);
    const dv = SpreadsheetApp.newDataValidation().requireValueInList(['OM', 'LD', '担当'], true).build();
    sh.getRange(FIRST_ROW, 2, 200, 1).setDataValidation(dv);
    ui.alert('「' + SITE.USERS + '」シートを作りました。\n通知先シートから ' + rows.length + '人分を取り込みました。\n\nOMの行の名前を自分の名前に直し、役割・担当ローンチを確認してから、メニュー「個人用リンクを作成・送信」を実行してください。');
  } else {
    ui.alert('「' + SITE.USERS + '」シートはすでにあります。人を追加したら「個人用リンクを作成・送信」を実行してください。');
  }
  ss_().setActiveSheet(sh);
}

function makeSiteLinks() {
  const ui = SpreadsheetApp.getUi();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss_().getId());
  const url = ScriptApp.getService().getUrl();
  if (!url) { ui.alert('先にWebアプリとしてデプロイしてください。\nApps Script → デプロイ → 新しいデプロイ → 種類：ウェブアプリ'); return; }
  const sh = usersSheet_();
  if (!sh) { ui.alert('先にメニュー「サイトの利用者シートを用意」を実行してください。'); return; }
  const users = readUsers_();
  let made = 0;
  users.forEach(u => {
    let k = u.key;
    if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20); sh.getRange(u.row, U.KEY).setValue(k); made++; }
    sh.getRange(u.row, U.URL).setValue(url + '?k=' + k);
  });
  const withMail = readUsers_().filter(u => u.mail);
  const ans = ui.alert('個人用リンクを作成しました（新しく ' + made + '人分）。\n\nメールアドレスがある ' + withMail.length + '人に、リンクをメールで送りますか？',
    ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  let sent = 0;
  withMail.forEach(u => {
    const body = u.name + 'さん\n\nいつもありがとうございます。\nローンチの毎日の進捗を確認・入力できる「' + SITE.TITLE + '」のリンクをお送りします。\n' +
      'あなた専用のリンクなので、ほかの人には転送しないでください。スマホのホーム画面に追加しておくと便利です。\n\n' + u.url +
      '\n\n' + (roleOf_(u.roleText) === 'tan' ? '見込み客の現状や次のアクションを、このページから更新できます。' : '毎日の実績入力もこのページからできます。') +
      '\n\nBNI福岡博多 ローンチ事務局';
    try { MailApp.sendEmail({ to: u.mail, subject: '【' + SITE.TITLE + '】あなた専用のリンク', body: body, name: CONFIG.MAIL_SENDER_NAME }); sent++; }
    catch (err) { logAlert_('ERROR', u.mail, 'リンク送信', String(err).slice(0, 200), ''); }
  });
  logAlert_('LINK', '', '個人用リンク送信', sent + '人に送信', '');
  ui.alert(sent + '人にリンクを送りました。');
}

// アラートメールに入れる個人用リンク（名前 → URL）
function siteLinks_() {
  const m = {};
  try { readUsers_().forEach(u => { if (u.url) { m[u.name] = u.url; if (roleOf_(u.roleText) === 'om') m.__om = u.url; } }); } catch (e) {}
  return m;
}

/* ---------- 画面（HTML）。src/page.html から自動生成。ここは直接編集しない ---------- */
const PAGE_HTML_LINES = [
  "<!DOCTYPE html>",
  "<html lang=\"ja\">",
  "<head>",
  "<meta charset=\"utf-8\">",
  "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">",
  "<base target=\"_top\">",
  "<title>ローンチ進捗ボード</title>",
  "<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?family=Zen+Kaku+Gothic+New:wght@400;500;700&family=IBM+Plex+Mono:wght@500&display=swap\">",
  "<style>",
  ":root{",
  "  --bg:#F5F6F8; --panel:#FFFFFF; --ink:#1A2233; --muted:#5B6578; --line:#DDE1E8; --soft:#EEF1F6;",
  "  --navy:#1F3A5F; --navy-soft:#E3EAF3; --red:#C8102E;",
  "  --ok:#1E7A4F; --ok-soft:#E2F3EA; --warn:#9A6200; --warn-soft:#FFF1D6; --bad:#B42318; --bad-soft:#FDE6E3;",
  "  --bar:#9DB3CF; --bar-hi:#1F3A5F;",
  "  --f-body:\"Zen Kaku Gothic New\",\"Hiragino Sans\",\"Meiryo\",sans-serif;",
  "  --f-num:\"IBM Plex Mono\",ui-monospace,Menlo,monospace;",
  "}",
  "@media (prefers-color-scheme: dark){:root{",
  "  --bg:#12161D; --panel:#1A2029; --ink:#E7EBF2; --muted:#9AA4B5; --line:#2C3441; --soft:#222A35;",
  "  --navy:#8FB0D9; --navy-soft:#223149; --red:#F0667A;",
  "  --ok:#5CC894; --ok-soft:#18342A; --warn:#E6B45A; --warn-soft:#3A2E17; --bad:#F2877C; --bad-soft:#3E1F1C;",
  "  --bar:#3E5578; --bar-hi:#8FB0D9; color-scheme:dark}}",
  "*{box-sizing:border-box}",
  "[hidden]{display:none!important}",
  "body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--f-body);font-size:15px;line-height:1.6}",
  ".wrap{max-width:1080px;margin:0 auto;padding-inline:16px;padding-block:0 48px}",
  "header .wrap{padding-block:0}",
  "header{position:sticky;top:0;z-index:5;background:var(--bg);border-bottom:1px solid var(--line)}",
  ".top{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px;padding-block:12px}",
  ".brand{display:flex;align-items:center;gap:10px}",
  ".mark{width:30px;height:30px;border-radius:7px;background:var(--red);color:#fff;display:grid;place-items:center;font-weight:700;font-size:13px}",
  ".brand b{font-size:17px;font-weight:700}",
  ".brand small{display:block;color:var(--muted);font-size:12px}",
  ".who{font-size:13px;color:var(--muted);display:flex;gap:8px;align-items:center}",
  ".who b{color:var(--ink);font-weight:500}",
  "label.filter{display:block;margin:0 0 12px;font-size:13px}",
  "select,input,textarea{font:inherit;color:var(--ink);background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:7px 10px}",
  "select:focus-visible,input:focus-visible,textarea:focus-visible,button:focus-visible{outline:2px solid var(--navy);outline-offset:1px}",
  "nav{display:flex;gap:4px;overflow-x:auto;padding-bottom:8px}",
  "nav button{flex:none;border:0;background:transparent;color:var(--muted);font:inherit;font-size:14px;padding:7px 12px;border-radius:8px;cursor:pointer}",
  "nav button[aria-selected=\"true\"]{background:var(--navy);color:var(--panel);font-weight:500}",
  "h2{font-size:18px;font-weight:700;margin:22px 0 10px}",
  "h3{font-size:15px;font-weight:700;margin:0}",
  ".muted{color:var(--muted)}",
  ".tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}",
  ".tile{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 14px}",
  ".tile .k{font-size:12px;color:var(--muted);letter-spacing:.04em}",
  ".tile .v{font-family:var(--f-num);font-size:26px;font-variant-numeric:tabular-nums;line-height:1.3}",
  ".tile .v small{font-family:var(--f-body);font-size:12px;color:var(--muted)}",
  ".meter{height:6px;background:var(--soft);border-radius:3px;margin-top:6px;overflow:hidden}",
  ".meter i{display:block;height:100%;background:var(--navy);border-radius:3px}",
  ".panel{background:var(--panel);border:1px solid var(--line);border-radius:10px}",
  ".tbl{width:100%;border-collapse:collapse;font-size:14px}",
  ".tbl th{font-size:12px;font-weight:500;color:var(--muted);text-align:left;padding:10px 12px;border-bottom:1px solid var(--line);white-space:nowrap}",
  ".tbl td{padding:10px 12px;border-bottom:1px solid var(--line);vertical-align:middle}",
  ".tbl tr:last-child td{border-bottom:0}",
  ".num{font-family:var(--f-num);font-variant-numeric:tabular-nums;white-space:nowrap}",
  ".tbl td .muted{white-space:nowrap}",
  ".scroll{overflow-x:auto}",
  ".pill{display:inline-block;font-size:12px;font-weight:500;padding:2px 9px;border-radius:999px;white-space:nowrap}",
  ".p-ok{background:var(--ok-soft);color:var(--ok)} .p-warn{background:var(--warn-soft);color:var(--warn)} .p-bad{background:var(--bad-soft);color:var(--bad)} .p-n{background:var(--soft);color:var(--muted)}",
  ".spark{display:flex;align-items:flex-end;gap:3px;height:28px;width:104px}",
  ".spark span{flex:1;background:var(--bar);border-radius:2px 2px 0 0;min-height:2px}",
  ".spark span:last-child{background:var(--bar-hi)}",
  ".two{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(0,1fr);gap:16px}",
  "@media (max-width:820px){.two{grid-template-columns:minmax(0,1fr)}}",
  ".todo{list-style:none;margin:0;padding:0}",
  ".todo li{display:flex;gap:10px;align-items:flex-start;padding:10px 14px;border-bottom:1px solid var(--line)}",
  ".todo li:last-child{border-bottom:0}",
  ".todo .tag{flex:none;font-size:11px;font-weight:700;padding:2px 7px;border-radius:5px;margin-top:2px}",
  ".todo .t{min-width:0}",
  ".todo .t small{display:block;color:var(--muted);font-size:12px}",
  ".empty{padding:18px 14px;color:var(--muted);font-size:14px}",
  ".btn{font:inherit;font-size:14px;border:1px solid var(--line);background:var(--panel);color:var(--ink);border-radius:8px;padding:7px 12px;cursor:pointer}",
  ".btn:hover{border-color:var(--navy)}",
  ".btn-pri{background:var(--navy);border-color:var(--navy);color:var(--panel);font-weight:500}",
  ".btn:disabled{opacity:.6;cursor:wait}",
  ".form{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;padding:16px}",
  ".form label{display:flex;flex-direction:column;gap:4px;font-size:13px;color:var(--muted)}",
  ".form .wide{grid-column:1/-1}",
  ".stepper{display:flex;align-items:center;gap:6px}",
  ".stepper input{width:64px;text-align:center;font-family:var(--f-num)}",
  ".stepper button{width:34px;height:34px;border-radius:8px;border:1px solid var(--line);background:var(--soft);color:var(--ink);font-size:18px;cursor:pointer}",
  ".cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:12px}",
  ".card{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:14px;display:flex;flex-direction:column;gap:10px;min-width:0}",
  ".card.alert{border-color:var(--bad)}",
  ".card .row{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}",
  ".card .meta{font-size:12px;color:var(--muted)}",
  ".stages{display:flex;flex-wrap:wrap;gap:4px}",
  ".stages button{font:inherit;font-size:12px;padding:3px 8px;border-radius:6px;border:1px solid var(--line);background:var(--panel);color:var(--muted);cursor:pointer}",
  ".stages button.on{background:var(--navy-soft);border-color:var(--navy);color:var(--navy);font-weight:700}",
  ".checks{display:flex;gap:6px}",
  ".checks button{flex:1;font:inherit;font-size:12px;padding:5px 0;border-radius:6px;border:1px dashed var(--line);background:transparent;color:var(--muted);cursor:pointer}",
  ".checks button.on{border-style:solid;border-color:var(--ok);background:var(--ok-soft);color:var(--ok);font-weight:700}",
  ".alerttxt{font-size:12px;color:var(--bad);font-weight:500}",
  ".filters{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}",
  ".tasks{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:2px}",
  ".tasks li{display:flex;align-items:center;gap:8px;font-size:13px}",
  ".tasks input{width:16px;height:16px;padding:0;accent-color:var(--ok)}",
  ".tasks .late{color:var(--bad);font-size:11px;margin-left:auto;white-space:nowrap}",
  ".tasks .due{color:var(--muted);font-size:11px;margin-left:auto;white-space:nowrap}",
  ".dna{font-size:10px;color:var(--navy);background:var(--navy-soft);padding:0 5px;border-radius:4px}",
  ".prog{height:6px;background:var(--soft);border-radius:3px;overflow:hidden}",
  ".prog i{display:block;height:100%;background:var(--ok)}",
  ".chart{width:100%;height:auto;display:block}",
  ".toast{position:fixed;left:50%;bottom:20px;transform:translateX(-50%);background:var(--ink);color:var(--panel);padding:10px 16px;border-radius:8px;font-size:14px;z-index:9;max-width:90vw}",
  ".center{padding:60px 16px;text-align:center;color:var(--muted)}",
  "details.add{margin-bottom:14px}",
  "details.add summary{cursor:pointer;font-weight:500;color:var(--navy);padding:6px 0}",
  "</style>",
  "</head>",
  "<body>",
  "<header>",
  "  <div class=\"wrap\">",
  "    <div class=\"top\">",
  "      <div class=\"brand\"><div class=\"mark\">BNI</div><div><b>ローンチ進捗ボード</b><small>BNI福岡博多リージョン｜<span id=\"today\">読み込み中</span></small></div></div>",
  "      <div class=\"who\" id=\"who\"></div>",
  "    </div>",
  "    <nav role=\"tablist\" id=\"tabs\" hidden>",
  "      <button role=\"tab\" data-tab=\"home\" aria-selected=\"true\">今日</button>",
  "      <button role=\"tab\" data-tab=\"input\" aria-selected=\"false\" data-lead>今日の実績を入力</button>",
  "      <button role=\"tab\" data-tab=\"ccs\" aria-selected=\"false\">CCS追っかけ</button>",
  "      <button role=\"tab\" data-tab=\"member\" aria-selected=\"false\">入会者</button>",
  "      <button role=\"tab\" data-tab=\"trend\" aria-selected=\"false\" data-lead>推移</button>",
  "    </nav>",
  "  </div>",
  "</header>",
  "<main class=\"wrap\">",
  "  <div class=\"center\" id=\"boot\">読み込み中です…</div>",
  "  <section id=\"v-home\" hidden></section>",
  "  <section id=\"v-input\" hidden></section>",
  "  <section id=\"v-ccs\" hidden></section>",
  "  <section id=\"v-member\" hidden></section>",
  "  <section id=\"v-trend\" hidden></section>",
  "</main>",
  "<div class=\"toast\" id=\"toast\" hidden></div>",
  "",
  "<script>",
  "const BOOT = __BOOT__;",
  "const DAY=864e5;",
  "let D=null, TODAY=null, TAB='home', CCSF={stage:'all',alert:false,tan:'all'};",
  "const pd=s=>{if(!s)return null;const a=String(s).split('/').map(Number);return new Date(a[0],a[1]-1,a[2]);};",
  "const fmt=d=>d?(d.getMonth()+1)+'/'+d.getDate():'';",
  "const iso=d=>d?d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'):'';",
  "const addD=(d,n)=>new Date(d.getFullYear(),d.getMonth(),d.getDate()+n);",
  "const diff=(a,b)=>Math.round((a-b)/DAY);",
  "const esc=s=>String(s==null?'':s).replace(/[&<>\"]/g,c=>({\"&\":\"&amp;\",\"<\":\"&lt;\",\">\":\"&gt;\",'\"':\"&quot;\"}[c]));",
  "const WD=['日','月','火','水','木','金','土'];",
  "",
  "function api(fn){const args=[].slice.call(arguments,1);return new Promise((res,rej)=>{",
  "  google.script.run.withSuccessHandler(r=>{ if(r&&r.error){rej(new Error(r.error));} else res(r); }).withFailureHandler(rej)[fn].apply(null,[BOOT.key].concat(args));});}",
  "function toast(m){const t=document.getElementById('toast');t.textContent=m;t.hidden=false;clearTimeout(toast.h);toast.h=setTimeout(()=>t.hidden=true,2800);}",
  "",
  "function load(msg){ return api('webGetData').then(r=>{",
  "    D=r; TODAY=pd(r.today);",
  "    D.launches.forEach(L=>{L.start=pd(L.start);L.firstMtg=pd(L.firstMtg);});",
  "    D.logs.forEach(x=>x.date=pd(x.d));",
  "    D.pros.forEach(p=>{p.last=pd(p.last);p.due=pd(p.due);p.formal=pd(p.formal);});",
  "    D.tasks.forEach(t=>t.due=pd(t.due));",
  "    D.set.weekStart=pd(D.set.weekStart)||addD(TODAY,-((TODAY.getDay()+6)%7));",
  "    document.getElementById('today').textContent=TODAY.getFullYear()+'年'+fmt(TODAY)+'（'+WD[TODAY.getDay()]+'）';",
  "    document.getElementById('who').innerHTML='<span>'+({om:'OM',ld:'LD',tan:'担当メンバー'}[D.me.role])+'</span><b>'+esc(D.me.name)+'</b>';",
  "    document.getElementById('boot').hidden=true; document.getElementById('tabs').hidden=false;",
  "    document.querySelectorAll('[data-lead]').forEach(b=>b.hidden=D.me.role==='tan');",
  "    renderAll(); showTab(TAB); if(msg)toast(msg);",
  "  }).catch(e=>{ document.getElementById('boot').hidden=false; document.getElementById('boot').textContent=e.message||String(e); }); }",
  "",
  "const isFormal=p=>p.app&&p.pay&&p.agr;",
  "const isLead=()=>D.me.role!=='tan';",
  "function logOn(ld,d){return D.logs.filter(x=>x.ld===ld&&diff(x.date,d)===0)}",
  "function sumLog(ld,from,to,key){return D.logs.filter(x=>x.ld===ld&&x.date>=from&&x.date<=to).reduce((a,b)=>a+(+b[key]||0),0)}",
  "function formalCount(ld){return D.pros.filter(p=>p.ld===ld&&isFormal(p)).length}",
  "function lastLog(ld){const d=D.logs.filter(x=>x.ld===ld).map(x=>x.date);return d.length?new Date(Math.max.apply(null,d)):null}",
  "function deadline(L){if(!L.start)return null;return L.phase==='フェーズ1'?addD(L.start,D.set.p1):L.phase==='フェーズ2'?addD(L.start,D.set.p2):null}",
  "function goalOf(ld){const g=D.set.goals&&D.set.goals[ld];return typeof g==='number'?g:D.set.dailyGoal}",
  "function active(L){return /^フェーズ/.test(L.phase||'')}",
  "function signal(L){ if(!L.start||!L.target)return{c:'p-n',t:L.phase||'準備中'}; const J=formalCount(L.ld),G=deadline(L); if(J>=L.target)return{c:'p-ok',t:'達成'};",
  "  if(!G)return{c:'p-warn',t:'進行中'}; const left=diff(G,TODAY); if(left<0)return{c:'p-bad',t:'期限超過'};",
  "  if(left<=D.set.near&&J<L.target*.8)return{c:'p-bad',t:'期限間近'}; const span=Math.max(1,diff(G,L.start)); const pace=L.target*Math.min(1,diff(TODAY,L.start)/span);",
  "  if(J>=pace)return{c:'p-ok',t:'順調'}; if(J>=pace*D.set.yline)return{c:'p-warn',t:'要注意'}; return{c:'p-bad',t:'遅れ'};}",
  "function prosAlerts(p){ if(isFormal(p)||(p.status&&p.status!=='進行中'))return[]; const a=[]; const idle=p.last?diff(TODAY,p.last):null;",
  "  if(p.due&&p.due<TODAY)a.push({lv:'bad',t:'期日超過'}); if(idle!=null&&idle>=D.set.idle)a.push({lv:'warn',t:'放置'+idle+'日'});",
  "  if(p.stage==='CCS済'&&!p.app)a.push({lv:'bad',t:'CCS後の申込待ち'}); if(p.app&&!p.pay)a.push({lv:'warn',t:'振込待ち'});",
  "  if(p.pay&&!p.agr)a.push({lv:'warn',t:'同意書待ち'}); if(!p.next)a.push({lv:'warn',t:'次アクション未設定'}); return a;}",
  "function membersOf(){ const m={}; D.tasks.forEach(t=>{const k=t.ld+'|'+t.name;(m[k]=m[k]||{ld:t.ld,name:t.name,tasks:[]}).tasks.push(t);});",
  "  Object.values(m).forEach(x=>{x.tasks.sort((a,b)=>a.ord-b.ord); x.p=D.pros.find(p=>p.ld===x.ld&&p.name===x.name)||null;}); return Object.values(m);}",
  "",
  "function todos(){ const out=[];",
  "  D.pros.forEach(p=>{ const a=prosAlerts(p); if(a.length)out.push({lv:a.some(x=>x.lv==='bad')?'bad':'warn',tag:a[0].t,title:p.name+(p.co?'（'+p.co+'）':''),sub:'担当 '+(p.tan||'未定')+'｜'+p.ld+'｜'+a.map(x=>x.t).join('・')+(p.next?'｜次：'+p.next:'')}) });",
  "  D.tasks.forEach(t=>{ if(t.done||!t.due||t.due>TODAY)return; if(D.me.role==='tan'&&t.kind==='DNA'&&t.owner!==D.me.name)return;",
  "    out.push({lv:t.due<TODAY?'bad':'warn',tag:t.due<TODAY?'入会後・期限超過':'入会後・今日まで',title:t.name+'：'+t.task,sub:(t.kind==='DNA'?'DNAのタスク（'+t.owner+'）':'本人のタスク')+'｜'+t.ld+'｜期限 '+fmt(t.due)}); });",
  "  if(isLead()) D.launches.forEach(L=>{ if(!active(L))return; const ll=lastLog(L.ld); const n=ll?diff(TODAY,ll):99;",
  "    if(n>=1) out.unshift({lv:n>=D.set.nolog?'bad':'warn',tag:n>=99?'未入力':n>=2?n+'日未入力':'今日まだ未入力',title:L.ld+'ローンチの日次実績',sub:'今日の実績を入力しましょう'}); });",
  "  return out.sort((a,b)=>(a.lv==='bad'?0:1)-(b.lv==='bad'?0:1));}",
  "const tagStyle=lv=>lv==='bad'?'background:var(--bad-soft);color:var(--bad)':'background:var(--warn-soft);color:var(--warn)';",
  "function todoHtml(td){return '<div class=\"panel\"><ul class=\"todo\">'+(td.length?td.map(t=>'<li><span class=\"tag\" style=\"'+tagStyle(t.lv)+'\">'+esc(t.tag)+'</span><div class=\"t\">'+esc(t.title)+'<small>'+esc(t.sub)+'</small></div></li>').join(''):'<li class=\"empty\">今日の要対応はありません。</li>')+'</ul></div>';}",
  "",
  "function renderHome(){ const el=document.getElementById('v-home'); const td=todos();",
  "  if(!isLead()){ const mine=D.pros; const st={}; D.stages.forEach(s=>st[s]=0); mine.filter(p=>!isFormal(p)&&(!p.status||p.status==='進行中')).forEach(p=>{st[p.stage||'（未設定）']=(st[p.stage||'（未設定）']||0)+1;});",
  "    const ccs=mine.filter(p=>p.stage==='CCS済'||isFormal(p)).length, joined=mine.filter(isFormal).length;",
  "    el.innerHTML='<h2>'+esc(D.me.name)+'さんの今日</h2><div class=\"tiles\">'+",
  "      '<div class=\"tile\"><div class=\"k\">今日の要対応</div><div class=\"v\" style=\"color:'+(td.length?'var(--bad)':'var(--ok)')+'\">'+td.length+'<small> 件</small></div></div>'+",
  "      '<div class=\"tile\"><div class=\"k\">担当している見込み客</div><div class=\"v\">'+mine.filter(p=>!isFormal(p)&&(!p.status||p.status==='進行中')).length+'<small> 人</small></div></div>'+",
  "      '<div class=\"tile\"><div class=\"k\">CCSにつないだ人</div><div class=\"v\">'+ccs+'<small> 人</small></div></div>'+",
  "      '<div class=\"tile\"><div class=\"k\">入会（正式）</div><div class=\"v\" style=\"color:var(--ok)\">'+joined+'<small> 人</small></div></div></div>'+",
  "      '<div class=\"two\"><div><h2>今日の要対応</h2>'+todoHtml(td)+'</div><div><h2>見込み客の現状</h2><div class=\"panel\"><table class=\"tbl\"><tbody>'+",
  "      Object.entries(st).map(([k,v])=>'<tr><td>'+esc(k)+'</td><td class=\"num\" style=\"text-align:right\">'+v+'</td></tr>').join('')+'</tbody></table></div>'+",
  "      '<p style=\"margin-top:12px\"><button class=\"btn btn-pri\" onclick=\"showTab(\\'ccs\\');document.querySelector(\\'details.add\\').open=true\">見込み客を追加する</button></p></div></div>';",
  "    return; }",
  "  const Ls=D.launches.filter(active);",
  "  const koe=Ls.reduce((a,L)=>a+sumLog(L.ld,TODAY,TODAY,'koe'),0), men=Ls.reduce((a,L)=>a+sumLog(L.ld,TODAY,TODAY,'men'),0);",
  "  const goal=Ls.reduce((a,L)=>a+goalOf(L.ld),0); const entered=Ls.filter(L=>logOn(L.ld,TODAY).length).length;",
  "  const rows=D.launches.map(L=>{ const s=signal(L); const days=[0,1,2,3,4,5,6].map(i=>sumLog(L.ld,addD(TODAY,i-6),addD(TODAY,i-6),'koe')); const mx=Math.max(10,Math.max.apply(null,days));",
  "    const ll=lastLog(L.ld); const n=ll?diff(TODAY,ll):null; const G=deadline(L);",
  "    const inp=!active(L)?'<span class=\"muted\">―</span>':n===0?'<span style=\"color:var(--ok);font-weight:500\">入力済</span>':'<span style=\"color:'+((n==null||n>=D.set.nolog)?'var(--bad)':'var(--warn)')+';font-weight:500\">'+(n==null?'未入力':n+'日未入力')+'</span>';",
  "    return '<tr><td><b>'+esc(L.ld)+'</b><div class=\"muted\" style=\"font-size:12px\">'+esc(L.team||'')+'｜'+esc(L.phase||'準備中')+'</div></td>'+",
  "      '<td><span class=\"pill '+s.c+'\">'+esc(s.t)+'</span></td><td class=\"num\">'+(L.target?formalCount(L.ld)+' / '+L.target:'―')+'</td>'+",
  "      '<td class=\"num\">'+(G?diff(G,TODAY)+'日':'―')+'</td><td class=\"num\">'+(active(L)?sumLog(L.ld,TODAY,TODAY,'koe')+' ／ '+sumLog(L.ld,TODAY,TODAY,'men'):'―')+'</td>'+",
  "      '<td>'+(active(L)?'<div class=\"spark\" title=\"直近7日の声かけ\">'+days.map(v=>'<span style=\"height:'+Math.round(v/mx*100)+'%\"></span>').join('')+'</div>':'<span class=\"muted\">―</span>')+'</td><td>'+inp+'</td></tr>'}).join('');",
  "  el.innerHTML='<h2>'+(D.me.role==='om'?'全ローンチの今日':esc(D.me.name)+'さんのローンチの今日')+'</h2><div class=\"tiles\">'+",
  "    '<div class=\"tile\"><div class=\"k\">今日の声かけ</div><div class=\"v\">'+koe+'<small> / 目標 '+goal+'</small></div><div class=\"meter\"><i style=\"width:'+(goal?Math.min(100,Math.round(koe/goal*100)):0)+'%\"></i></div></div>'+",
  "    '<div class=\"tile\"><div class=\"k\">今日の面談・CCS</div><div class=\"v\">'+men+'</div></div>'+",
  "    '<div class=\"tile\"><div class=\"k\">正式プリコア（累計）</div><div class=\"v\">'+Ls.reduce((a,L)=>a+formalCount(L.ld),0)+'<small> / '+Ls.reduce((a,L)=>a+(+L.target||0),0)+'</small></div></div>'+",
  "    '<div class=\"tile\"><div class=\"k\">今日の入力</div><div class=\"v\">'+entered+'<small> / '+Ls.length+' ローンチ</small></div></div>'+",
  "    '<div class=\"tile\"><div class=\"k\">要対応</div><div class=\"v\" style=\"color:'+(td.length?'var(--bad)':'var(--ok)')+'\">'+td.length+'<small> 件</small></div></div></div>'+",
  "    '<div class=\"two\"><div><h2>ローンチ一覧</h2><div class=\"panel scroll\"><table class=\"tbl\"><thead><tr><th>ローンチ</th><th>判定</th><th>正式/目標</th><th>残り</th><th>今日 声かけ／面談</th><th>直近7日</th><th>今日の入力</th></tr></thead><tbody>'+(rows||'<tr><td class=\"empty\" colspan=\"7\">表示できるローンチがありません</td></tr>')+'</tbody></table></div>'+",
  "    '<p class=\"muted\" style=\"font-size:12px\">判定：経過日数に対して目標ペース以上＝順調、'+Math.round(D.set.yline*100)+'%以上＝要注意、それ未満＝遅れ</p></div>'+",
  "    '<div><h2>今日の要対応</h2>'+todoHtml(td)+'</div></div>';}",
  "",
  "function stepper(id,label){return '<label for=\"'+id+'\">'+label+'<div class=\"stepper\"><button type=\"button\" data-step=\"-1\" data-for=\"'+id+'\" aria-label=\"'+label+'を1減らす\">−</button><input id=\"'+id+'\" type=\"number\" min=\"0\" value=\"0\" inputmode=\"numeric\"><button type=\"button\" data-step=\"1\" data-for=\"'+id+'\" aria-label=\"'+label+'を1増やす\">＋</button></div></label>'}",
  "function renderInput(){ if(!isLead())return; const Ls=D.launches.filter(L=>L.ld);",
  "  const recent=D.logs.slice().sort((a,b)=>b.date-a.date).slice(0,10);",
  "  document.getElementById('v-input').innerHTML='<h2>今日の実績を入力</h2><p class=\"muted\" style=\"margin:0 0 10px\">その日に増えた件数を入れて送信します。活動がなかった日も0のまま送信してください。同じ日にもう一度送ると上書きされます。</p>'+",
  "    '<form class=\"panel form\" id=\"logform\"><label for=\"f-ld\">ローンチ<select id=\"f-ld\">'+Ls.map(L=>'<option>'+esc(L.ld)+'</option>').join('')+'</select></label>'+",
  "    '<label for=\"f-date\">日付<input id=\"f-date\" type=\"date\" value=\"'+iso(TODAY)+'\"></label>'+stepper('f-koe','声かけ')+stepper('f-info','インフォ参加')+stepper('f-men','面談・CCS')+",
  "    '<label class=\"wide\" for=\"f-memo\">今日のひとこと・課題<textarea id=\"f-memo\" rows=\"2\" placeholder=\"例：CCS後の申込が2名止まっている。明日電話する\"></textarea></label>'+",
  "    '<div class=\"wide\"><button class=\"btn btn-pri\" type=\"submit\" id=\"f-send\">送信する</button></div><div class=\"wide alerttxt\" id=\"f-err\" hidden></div></form>'+",
  "    '<h2>最近の入力</h2><div class=\"panel scroll\"><table class=\"tbl\"><thead><tr><th>日付</th><th>ローンチ</th><th>声かけ</th><th>インフォ</th><th>面談・CCS</th><th>入力者</th><th>ひとこと</th></tr></thead><tbody>'+",
  "    (recent.length?recent.map(x=>'<tr><td class=\"num\">'+fmt(x.date)+'</td><td>'+esc(x.ld)+'</td><td class=\"num\">'+x.koe+'</td><td class=\"num\">'+x.info+'</td><td class=\"num\">'+x.men+'</td><td>'+esc(x.by)+'</td><td>'+esc(x.memo)+'</td></tr>').join(''):'<tr><td class=\"empty\" colspan=\"7\">まだ入力がありません</td></tr>')+'</tbody></table></div>';",
  "  const f=document.getElementById('logform');",
  "  f.addEventListener('click',e=>{const b=e.target.closest('[data-step]');if(!b)return;const i=document.getElementById(b.dataset.for);i.value=Math.max(0,(+i.value||0)+(+b.dataset.step));});",
  "  f.addEventListener('submit',e=>{e.preventDefault(); const err=document.getElementById('f-err'); const dv=document.getElementById('f-date').value;",
  "    if(!dv){err.textContent='日付を入れてください';err.hidden=false;return;} err.hidden=true; const btn=document.getElementById('f-send'); btn.disabled=true; btn.textContent='送信中…';",
  "    const o={ld:document.getElementById('f-ld').value,date:dv.replace(/-/g,'/'),koe:+document.getElementById('f-koe').value||0,info:+document.getElementById('f-info').value||0,men:+document.getElementById('f-men').value||0,memo:document.getElementById('f-memo').value.trim()};",
  "    api('webSaveLog',o).then(()=>load(o.ld+'ローンチの'+fmt(pd(o.date))+'の実績を記録しました')).catch(ex=>{err.textContent=ex.message;err.hidden=false;btn.disabled=false;btn.textContent='送信する';}); });}",
  "",
  "function renderCCS(){ const base=D.pros.filter(p=>!isFormal(p)&&(!p.status||p.status==='進行中'));",
  "  const tans=[...new Set(base.map(p=>p.tan).filter(Boolean))];",
  "  const list=base.filter(p=>CCSF.stage==='all'||p.stage===CCSF.stage).filter(p=>CCSF.tan==='all'||p.tan===CCSF.tan).filter(p=>!CCSF.alert||prosAlerts(p).length).sort((a,b)=>prosAlerts(b).length-prosAlerts(a).length);",
  "  const byTan={}; D.pros.forEach(p=>{const k=p.ld+'｜'+(p.tan||'未定');byTan[k]=byTan[k]||{list:0,ccs:0,formal:0}; byTan[k].list++; if(p.stage==='CCS済'||isFormal(p))byTan[k].ccs++; if(isFormal(p))byTan[k].formal++;});",
  "  const ldOpts=D.allLaunches.map(n=>'<option>'+esc(n)+'</option>').join('');",
  "  document.getElementById('v-ccs').innerHTML='<h2>CCS追っかけ</h2>'+",
  "    '<details class=\"add\" '+(CCSF.addOpen?'open':'')+'><summary>＋ 見込み客を追加する</summary><form class=\"panel form\" id=\"addform\">'+",
  "      '<label for=\"a-name\">お名前（必須）<input id=\"a-name\" required placeholder=\"例：山田 太郎\"></label><label for=\"a-co\">会社名<input id=\"a-co\"></label><label for=\"a-cat\">カテゴリー<input id=\"a-cat\" placeholder=\"例：税理士\"></label>'+",
  "      '<label for=\"a-tel\">電話番号<input id=\"a-tel\" inputmode=\"tel\"></label><label for=\"a-ld\">ローンチ<select id=\"a-ld\">'+ldOpts+'</select></label>'+",
  "      (isLead()?'<label for=\"a-tan\">担当（リストアップした人）<input id=\"a-tan\" placeholder=\"例：井口\"></label>':'')+",
  "      '<label for=\"a-stage\">現状<select id=\"a-stage\">'+D.stages.map(s=>'<option>'+esc(s)+'</option>').join('')+'</select></label>'+",
  "      '<label for=\"a-next\">次のアクション<input id=\"a-next\" placeholder=\"例：説明会に誘う\"></label><label for=\"a-due\">期日<input id=\"a-due\" type=\"date\"></label>'+",
  "      '<div class=\"wide\"><button class=\"btn btn-pri\" type=\"submit\" id=\"a-send\">追加する</button> <span class=\"alerttxt\" id=\"a-err\" hidden></span></div></form></details>'+",
  "    '<div class=\"filters\"><label for=\"c-stage\" class=\"muted\" style=\"font-size:13px\">現状</label><select id=\"c-stage\"><option value=\"all\">すべて（'+base.length+'）</option>'+D.stages.map(s=>'<option '+(CCSF.stage===s?'selected':'')+'>'+esc(s)+'</option>').join('')+'</select>'+",
  "      (isLead()&&tans.length?'<label for=\"c-tan\" class=\"muted\" style=\"font-size:13px\">担当</label><select id=\"c-tan\"><option value=\"all\">全員</option>'+tans.map(t=>'<option '+(CCSF.tan===t?'selected':'')+'>'+esc(t)+'</option>').join('')+'</select>':'')+",
  "      '<label style=\"display:flex;gap:6px;align-items:center;font-size:13px\"><input type=\"checkbox\" id=\"c-alert\" '+(CCSF.alert?'checked':'')+' style=\"width:16px;height:16px;padding:0\">要対応だけ</label></div>'+",
  "    '<div class=\"cards\">'+(list.length?list.map(p=>{ const a=prosAlerts(p);",
  "      return '<div class=\"card '+(a.some(x=>x.lv==='bad')?'alert':'')+'\"><div class=\"row\"><div><h3>'+esc(p.name)+'</h3><div class=\"meta\">'+esc(p.co||'')+'｜担当 '+esc(p.tan||'未定')+'｜'+esc(p.ld)+'</div></div><span class=\"meta num\">'+(p.last?'最終 '+fmt(p.last):'')+'</span></div>'+",
  "        '<div class=\"stages\" role=\"group\" aria-label=\"現状\">'+D.stages.map(s=>'<button data-act=\"stage\" data-row=\"'+p.row+'\" data-v=\"'+esc(s)+'\" class=\"'+(p.stage===s?'on':'')+'\">'+esc(s)+'</button>').join('')+'</div>'+",
  "        '<div class=\"checks\">'+[['app','申込書'],['pay','振込控え'],['agr','同意書']].map(k=>'<button data-act=\"chk\" data-row=\"'+p.row+'\" data-v=\"'+k[0]+'\" class=\"'+(p[k[0]]?'on':'')+'\">'+(p[k[0]]?'✓ ':'')+k[1]+'</button>').join('')+'</div>'+",
  "        '<div style=\"display:flex;gap:6px;flex-wrap:wrap\"><input id=\"n-'+p.row+'\" value=\"'+esc(p.next)+'\" placeholder=\"次のアクション\" style=\"flex:1;min-width:140px;font-size:13px\"><input id=\"d-'+p.row+'\" type=\"date\" value=\"'+iso(p.due)+'\" style=\"font-size:13px\"><button class=\"btn\" data-act=\"save\" data-row=\"'+p.row+'\">保存</button></div>'+",
  "        (a.length?'<div class=\"alerttxt\">'+a.map(x=>x.t).join('・')+'</div>':'')+'</div>'}).join(''):'<div class=\"empty\">条件に合う見込み客はいません。</div>')+'</div>'+",
  "    '<h2>担当別CCS</h2><div class=\"panel scroll\"><table class=\"tbl\"><thead><tr><th>ローンチ｜担当</th><th>リストアップ</th><th>CCS済</th><th>入会（正式）</th></tr></thead><tbody>'+",
  "    (Object.keys(byTan).length?Object.entries(byTan).sort((a,b)=>b[1].ccs-a[1].ccs).map(e=>'<tr><td>'+esc(e[0])+'</td><td class=\"num\">'+e[1].list+'</td><td class=\"num\">'+e[1].ccs+'</td><td class=\"num\" style=\"color:var(--ok);font-weight:500\">'+e[1].formal+'</td></tr>').join(''):'<tr><td class=\"empty\" colspan=\"4\">まだ見込み客がいません</td></tr>')+'</tbody></table></div>';",
  "  document.getElementById('c-stage').onchange=e=>{CCSF.stage=e.target.value;renderCCS();};",
  "  const ct=document.getElementById('c-tan'); if(ct)ct.onchange=e=>{CCSF.tan=e.target.value;renderCCS();};",
  "  document.getElementById('c-alert').onchange=e=>{CCSF.alert=e.target.checked;renderCCS();};",
  "  document.querySelector('details.add').ontoggle=e=>{CCSF.addOpen=e.target.open;};",
  "  document.getElementById('addform').onsubmit=e=>{e.preventDefault(); const err=document.getElementById('a-err'); const name=document.getElementById('a-name').value.trim();",
  "    if(!name){err.textContent='お名前を入れてください';err.hidden=false;return;} const btn=document.getElementById('a-send'); btn.disabled=true; btn.textContent='追加中…';",
  "    const g=id=>{const x=document.getElementById(id);return x?x.value.trim():'';};",
  "    api('webAddProspect',{name,co:g('a-co'),cat:g('a-cat'),tel:g('a-tel'),ld:g('a-ld'),tan:g('a-tan'),stage:g('a-stage'),next:g('a-next'),due:g('a-due').replace(/-/g,'/')})",
  "      .then(()=>{CCSF.addOpen=false;return load(name+'さんを追加しました');}).catch(ex=>{err.textContent=ex.message;err.hidden=false;btn.disabled=false;btn.textContent='追加する';}); };}",
  "document.getElementById('v-ccs').addEventListener('click',e=>{ const b=e.target.closest('[data-act]'); if(!b)return; const row=+b.dataset.row; const p=D.pros.find(x=>x.row===row); if(!p)return;",
  "  let patch={};",
  "  if(b.dataset.act==='stage')patch.stage=b.dataset.v;",
  "  if(b.dataset.act==='chk')patch[b.dataset.v]=!p[b.dataset.v];",
  "  if(b.dataset.act==='save'){patch.next=document.getElementById('n-'+row).value.trim();patch.due=document.getElementById('d-'+row).value.replace(/-/g,'/');}",
  "  b.disabled=true; toast('保存しています…');",
  "  api('webUpdateProspect',row,patch).then(r=>load(r&&r.formal?p.name+'さんが正式プリコアになりました。入会後の行動リストを作成しました':p.name+'さんを更新しました')).catch(ex=>{toast(ex.message);b.disabled=false;}); });",
  "",
  "let MEMBER_LD='';",
  "function renderMember(){ const all=membersOf(); const lds=Array.from(new Set(D.launches.map(L=>L.ld).concat(all.map(m=>m.ld))));",
  "  if(lds.indexOf(MEMBER_LD)<0)MEMBER_LD=''; const ms=MEMBER_LD?all.filter(m=>m.ld===MEMBER_LD):all;",
  "  const sel=lds.length>1?'<label class=\"filter\" for=\"m-ld\">ローンチで絞り込む <select id=\"m-ld\"><option value=\"\">すべてのローンチ（'+all.length+'人）</option>'+lds.map(n=>'<option value=\"'+esc(n)+'\"'+(n===MEMBER_LD?' selected':'')+'>'+esc(n)+'（'+all.filter(m=>m.ld===n).length+'人）</option>').join('')+'</select></label>':'';",
  "  document.getElementById('v-member').innerHTML='<h2>入会者の行動リスト</h2><p class=\"muted\" style=\"margin:0 0 12px\">正式プリコアになると、入会後の流れが自動で並びます。終わったらチェックを入れてください。<span class=\"dna\">DNA</span> はDNA・LDのタスクです。</p>'+sel+",
  "    '<div class=\"cards\">'+(ms.length?ms.map(m=>{ const ts=m.tasks; const done=ts.filter(t=>t.done).length; const late=ts.filter(t=>!t.done&&t.due&&t.due<TODAY).length;",
  "      return '<div class=\"card '+(late?'alert':'')+'\"><div class=\"row\"><div><h3>'+esc(m.name)+'</h3><div class=\"meta\">'+esc(m.p&&m.p.co||'')+(m.p&&m.p.tan?'｜担当 '+esc(m.p.tan):'')+'｜'+esc(m.ld)+'</div></div><span class=\"meta num\">'+(m.p&&m.p.formal?'入会 '+fmt(m.p.formal):'')+'</span></div>'+",
  "        '<div><div class=\"row meta\"><span>'+done+' / '+ts.length+' 完了</span>'+(late?'<span class=\"alerttxt\">'+late+'件 期限超過</span>':'')+'</div><div class=\"prog\"><i style=\"width:'+Math.round(done/Math.max(1,ts.length)*100)+'%\"></i></div></div>'+",
  "        '<ul class=\"tasks\">'+ts.map(t=>'<li><input type=\"checkbox\" id=\"t-'+t.row+'\" data-row=\"'+t.row+'\" '+(t.done?'checked':'')+'><label for=\"t-'+t.row+'\" style=\"'+(t.done?'color:var(--muted);text-decoration:line-through':'')+'\">'+esc(t.task)+'</label>'+(t.kind==='DNA'?'<span class=\"dna\">DNA</span>':'')+(t.done?'':(t.due&&t.due<TODAY)?'<span class=\"late\">'+fmt(t.due)+' 超過</span>':'<span class=\"due\">'+(t.due?fmt(t.due)+'まで':'')+'</span>')+'</li>').join('')+'</ul></div>'}).join(''):'<div class=\"empty\">まだ入会者はいません。CCS追っかけで3点に✓が付くと、ここに表示されます。</div>')+'</div>';}",
  "document.getElementById('v-member').addEventListener('change',e=>{ const c=e.target; if(c.id==='m-ld'){MEMBER_LD=c.value;renderMember();return;} if(!c.dataset.row)return; c.disabled=true;",
  "  api('webSetTaskDone',+c.dataset.row,c.checked).then(()=>load(c.checked?'完了にしました':'未完了に戻しました')).catch(ex=>{toast(ex.message);c.checked=!c.checked;c.disabled=false;}); });",
  "",
  "function renderTrend(){ if(!isLead())return; const Ls=D.launches.filter(active); const N=14; const days=[]; for(let i=0;i<N;i++)days.push(addD(TODAY,i-N+1));",
  "  const tot=days.map(d=>Ls.reduce((a,L)=>a+sumLog(L.ld,d,d,'koe'),0)); const ceil=Math.ceil(Math.max(10,Math.max.apply(null,tot))/10)*10;",
  "  const W=640,H=220,pl=34,pb=26,pt=12,cw=(W-pl-8)/N; const y=v=>pt+(H-pt-pb)*(1-v/ceil);",
  "  const svg='<svg class=\"chart\" viewBox=\"0 0 '+W+' '+H+'\" role=\"img\" aria-label=\"直近14日の声かけ数\"><title>直近14日の声かけ数</title>'+",
  "    [0,ceil/2,ceil].map(t=>'<line x1=\"'+pl+'\" x2=\"'+(W-8)+'\" y1=\"'+y(t)+'\" y2=\"'+y(t)+'\" stroke=\"var(--line)\"/><text x=\"'+(pl-6)+'\" y=\"'+(y(t)+4)+'\" text-anchor=\"end\" font-size=\"11\" fill=\"var(--muted)\">'+t+'</text>').join('')+",
  "    tot.map((v,i)=>'<rect x=\"'+(pl+i*cw+3)+'\" y=\"'+y(v)+'\" width=\"'+(cw-6)+'\" height=\"'+Math.max(0,y(0)-y(v))+'\" rx=\"2\" fill=\"'+(i===N-1?'var(--bar-hi)':'var(--bar)')+'\"/>'+((i%2===1||i===N-1)?'<text x=\"'+(pl+i*cw+cw/2)+'\" y=\"'+(H-8)+'\" text-anchor=\"middle\" font-size=\"11\" fill=\"var(--muted)\">'+fmt(days[i])+'</text>':'')).join('')+",
  "    '<text x=\"'+(pl+(N-1)*cw+cw/2)+'\" y=\"'+(y(tot[N-1])-5)+'\" text-anchor=\"middle\" font-size=\"11\" fill=\"var(--ink)\">'+tot[N-1]+'</text></svg>';",
  "  const wk=[]; for(let k=0;k<13;k++)wk.push(addD(D.set.weekStart,7*k));",
  "  const wkCell=(L,w,key)=>w>TODAY?'<td class=\"num muted\">―</td>':'<td class=\"num\">'+sumLog(L.ld,w,addD(w,6),key)+'</td>';",
  "  document.getElementById('v-trend').innerHTML='<h2>直近14日の声かけ（'+Ls.length+'ローンチ合計）</h2><div class=\"panel\" style=\"padding:12px\">'+svg+'</div>'+",
  "    '<h2>ローンチ別　直近7日の声かけ</h2><div class=\"panel scroll\"><table class=\"tbl\"><thead><tr><th>ローンチ</th>'+days.slice(-7).map(d=>'<th class=\"num\">'+fmt(d)+'</th>').join('')+'<th>7日計</th></tr></thead><tbody>'+",
  "    Ls.map(L=>{const v=days.slice(-7).map(d=>logOn(L.ld,d).length?sumLog(L.ld,d,d,'koe'):null);return '<tr><td>'+esc(L.ld)+'</td>'+v.map(x=>x==null?'<td class=\"num\" style=\"color:var(--bad);background:var(--bad-soft)\">未</td>':'<td class=\"num\">'+x+'</td>').join('')+'<td class=\"num\"><b>'+v.reduce((a,b)=>a+(b||0),0)+'</b></td></tr>'}).join('')+'</tbody></table></div>'+",
  "    '<h2>週次推移（'+fmt(D.set.weekStart)+'スタート）</h2><div class=\"panel scroll\"><table class=\"tbl\"><thead><tr><th>ローンチ</th><th>項目</th>'+wk.map((d,k)=>'<th class=\"num\">第'+(k+1)+'週<br><span style=\"font-weight:400\">'+fmt(d)+'〜</span></th>').join('')+'</tr></thead><tbody>'+",
  "    Ls.map(L=>'<tr><td rowspan=\"2\">'+esc(L.ld)+'</td><td class=\"muted\">声かけ</td>'+wk.map(w=>wkCell(L,w,'koe')).join('')+'</tr><tr><td class=\"muted\">面談</td>'+wk.map(w=>wkCell(L,w,'men')).join('')+'</tr>').join('')+'</tbody></table></div>';}",
  "",
  "function renderAll(){renderHome();renderInput();renderCCS();renderMember();renderTrend();}",
  "function showTab(t){ TAB=t; document.querySelectorAll('#tabs button').forEach(x=>x.setAttribute('aria-selected',x.dataset.tab===t?'true':'false'));",
  "  document.querySelectorAll('main>section').forEach(s=>s.hidden=s.id!=='v-'+t); window.scrollTo(0,0); }",
  "document.getElementById('tabs').addEventListener('click',e=>{const b=e.target.closest('[data-tab]');if(b)showTab(b.dataset.tab);});",
  "load();",
  "</script>",
  "</body>",
  "</html>"
];
