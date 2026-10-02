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
