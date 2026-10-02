/******************************************************************
 * ローンチ進捗ボード（Webアプリ）
 *  ・データはこのスプレッドシートにそのまま貯まります
 *  ・誰が開いたかは「サイト利用者」シートの個人用リンク（?k=…）で判定します
 *  ・OM＝全ローンチ／LD＝自分のローンチ／担当＝自分の見込み客と、その人の入会後タスク
 ******************************************************************/
const SITE = {
  USERS: 'サイト利用者', DASH: '📊ダッシュボード', LOG: '日次ログ',
  DAILY_GOAL: 10,           // 1ローンチあたりの1日の声かけ目標（ホームのメーターに使用）
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
    weekStart: isDate_(m['週次推移の開始日（月曜）']) ? ymd_(m['週次推移の開始日（月曜）']) : '', dailyGoal: SITE.DAILY_GOAL };
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
