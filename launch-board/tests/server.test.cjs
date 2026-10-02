// GASのサービス（SpreadsheetApp など）を模擬して、dist/Code.gs のサーバー関数を動かすテスト
//   node tests/server.test.cjs   （先に node build.mjs で dist/Code.gs を作っておく）
const fs=require('fs'), path=require('path'), assert=require('assert');
eval(fs.readFileSync(path.join(__dirname,'../dist/Code.gs'),'utf8').replace(/^const /gm,'var '));
function mkSheet(name,rows){ const data=rows;
 const sheet={getName:()=>name,getMaxRows:()=>Math.max(data.length,600),getMaxColumns:()=>30,getLastRow:()=>data.length,getLastColumn:()=>15,getSheetId:()=>1,data,insertRowsAfter(){},
  getRange:(r,c,nr=1,nc=1)=>{ if(typeof r==='string'){const m=r.match(/([A-Z])(\d+)(?::([A-Z])(\d+))?/);const cc=x=>x.charCodeAt(0)-64;c=cc(m[1]);const r0=+m[2];nr=m[4]?+m[4]-r0+1:1;nc=m[3]?cc(m[3])-c+1:1;r=r0;}
   const self={getValues:()=>{const o=[];for(let i=0;i<nr;i++){const row=[];for(let j=0;j<nc;j++)row.push(((data[r-1+i]||[])[c-1+j])??'');o.push(row);}return o;},
    getValue:()=>((data[r-1]||[])[c-1])??'',
    setValues:v=>{v.forEach((row,i)=>{data[r-1+i]=data[r-1+i]||[];row.forEach((x,j)=>data[r-1+i][c-1+j]=x);});return P;},
    setValue:x=>{data[r-1]=data[r-1]||[];data[r-1][c-1]=x;return P;},setFormulas:v=>self.setValues(v)};
   const P=new Proxy(self,{get:(t,k)=>k in t?t[k]:()=>P}); return P;}};
 return new Proxy(sheet,{get:(t,k)=>k in t?t[k]:()=>{}});}
const d=(s)=>new Date(s+'T00:00:00+09:00'); const T0='2026-09-30';
const E=()=>[[],[],[],[]];
function ccs(ld,tan,name,stage,app,pay,agr,formal,last,next,due){const r=[];r[1]=ld;r[2]=tan;r[3]=name;r[6]=name+'商事';r[9]=stage;r[12]=app;r[13]=pay;r[14]=agr;r[16]=formal;r[17]=last;r[19]=next;r[20]=due;r[21]='進行中';return r;}
function ldSheet(team,phase,start,target,first){const v=[];for(let i=0;i<28;i++)v.push([]);v[3][1]=team;v[5][1]='テストCH';v[7][1]=first;v[23][1]=phase;v[24][1]=start;v[25][1]=target;v[26][1]='課題メモ';return v;}
const set=[[],[],[,,,'フェーズ','チーム','声かけ状況','候補者の状態','要否','LD一覧','声かけ目標（1日）','現状'],
 ['フェーズ1期限（日）',42,,,,,,,'石川',15,'リストアップ済'],['フェーズ2期限（日）',28,,,,,,,'竹尾','','説明会お誘い中'],['黄色判定ライン（目標ペース比）',0.7,,,,,,,,,'CCS依頼中'],
 ['期限間近アラート（残日数）',7,,,,,,,,,'CCS済'],['ログ未入力アラート（日）',3],['CCS放置アラート（日）',3],['本人リマインド再送間隔（日）',2],['初回定例会前の警戒期間（日）',7],['週次推移の開始日（月曜）',d('2026-10-05')]];
const logs=E(); for(let i=10;i>=0;i--){logs.push([d('2026-09-'+String(30-i).padStart(2,'0')),'石川','石川',5+i%3,1,1,'','','','']); if(i>3)logs.push([d('2026-09-'+String(30-i).padStart(2,'0')),'竹尾','竹尾',4,0,0,'','','','']);}
const sheets={
 'CCS追っかけ':mkSheet('CCS追っかけ',E().concat([
   ccs('石川','村上','山本健一','CCS済','✓','✓','✓',d('2026-09-25'),d('2026-09-25'),'',''),
   ccs('石川','村上','森田翔','CCS依頼中','','','','',d('2026-09-24'),'',''),
   ccs('石川','岡','中島さやか','CCS済','✓','✓','','',d(T0),'同意書回収',''),
   ccs('竹尾','高木','原田大輔','説明会お誘い中','','','','',d(T0),'電話する',d('2026-09-27'))])),
 '入会後行動リスト':mkSheet('入会後行動リスト',E()),
 '行動マスタ':mkSheet('行動マスタ',E().concat([[1,"ウェルカムコール","DNA","CCS実施",1,"不要"],[2,"FB登録","本人","本人",2,"要"],[3,"FB・LINEグループ招待","DNA","LD",3,"不要"]])),
 '通知先':mkSheet('通知先',E().concat([["石川","ishikawa@x.jp","","","","村上","murakami@x.jp"],["竹尾","takeo@x.jp","","","","岡","oka@x.jp"]])),
 '設定':mkSheet('設定',set),'アラートログ':mkSheet('アラートログ',E()),
 '📊ダッシュボード':mkSheet('db',[[],[],[],[],[],[],[],['石川'],['竹尾'],['世良']]),
 '石川':mkSheet('石川',ldSheet('JJ','フェーズ1',d('2026-09-14'),12,d('2026-10-26'))),
 '竹尾':mkSheet('竹尾',ldSheet('JJ','フェーズ1',d('2026-09-07'),12,d('2026-10-19'))),
 '世良':mkSheet('世良',ldSheet('Believe','準備中','','','')),
 '日次ログ':mkSheet('日次ログ',logs),
 'サイト利用者':mkSheet('u',E().concat([['川嶋','OM','','om@x.jp','KOM'],['石川','LD','石川','','KLD'],['村上','担当','','','KTAN']])),
};
global.SpreadsheetApp={getActive:()=>({getSheetByName:n=>sheets[n]||null,getSpreadsheetTimeZone:()=>'Asia/Tokyo',toast:()=>{},getUrl:()=>'https://docs.google.com/x',getId:()=>'id'})};
const pad=n=>String(n).padStart(2,'0');
global.Utilities={formatDate:(dt,tz,f)=>{ if(tz==='UTC')return `${dt.getUTCFullYear()}/${pad(dt.getUTCMonth()+1)}/${pad(dt.getUTCDate())}`;
  const j=new Date(dt.getTime()+9*3600e3);return f.replace('yyyy',j.getUTCFullYear()).replace('MM',pad(j.getUTCMonth()+1)).replace('dd',pad(j.getUTCDate())).replace('HH',pad(j.getUTCHours())).replace('mm',pad(j.getUTCMinutes())).replace(/\bM\b/,j.getUTCMonth()+1).replace(/\bd\b/,j.getUTCDate()).replace(/\bH\b/,j.getUTCHours());},
 parseDate:(s)=>d(s.replace(/\//g,'-')),getUuid:()=>'abc-def'};
global.Session={getEffectiveUser:()=>({getEmail:()=>'om@x.jp'}),getActiveUser:()=>({getEmail:()=>''})};
global.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})};
global.PropertiesService={getScriptProperties:()=>({getProperty:()=>'',setProperty(){}})};
const RD=Date; global.Date=class extends RD{constructor(...a){if(a.length===0)super('2026-09-30T12:00:00+09:00');else super(...a);} static UTC(...a){return RD.UTC(...a);}};
CONFIG.SPREADSHEET_ID='';
let n=0; const ok=(name,cond)=>{assert.ok(cond,name);n++;console.log('  ok  '+name);};
syncAll_();
ok('正式プリコアの行動リストが作られる', sheets['入会後行動リスト'].data.length-4===3);
const om=webGetData('KOM'), ld=webGetData('KLD'), tan=webGetData('KTAN');
ok('OMは全ローンチが見える', om.launches.length===3 && om.pros.length===4);
ok('LDは自分のローンチだけ', ld.launches.length===1 && ld.pros.every(p=>p.ld==='石川'));
ok('担当は自分の見込み客だけ・日次ログは見えない', tan.pros.every(p=>p.tan==='村上') && tan.logs.length===0);
ok('声かけ目標は設定シートから読む（空欄は既定値）', om.set.goals['石川']===15 && om.set.goals['竹尾']===undefined && om.set.dailyGoal===10);
ok('不正なキーは拒否', !!webGetData('nope').error);
ok('担当は日次実績を入力できない', !!webSaveLog('KTAN',{ld:'石川',date:'2026/09/30'}).error);
ok('LDは他ローンチに入力できない', !!webSaveLog('KLD',{ld:'竹尾',date:'2026/09/30'}).error);
const before=sheets['日次ログ'].data.filter(r=>r[1]==='石川').length;
ok('同じ日の実績は上書き', webSaveLog('KLD',{ld:'石川',date:'2026/09/30',koe:9,info:2,men:1,memo:'テスト'}).ok && sheets['日次ログ'].data.filter(r=>r[1]==='石川').length===before);
ok('担当は他人の見込み客を更新できない', !!webUpdateProspect('KTAN',7,{stage:'CCS済'}).error);
ok('3点✓で正式化し行動リスト作成', webUpdateProspect('KOM',7,{agr:true}).formal===true && sheets['入会後行動リスト'].data.length-4===6);
const add=webAddProspect('KTAN',{name:'新人 花子',ld:'石川',stage:'リストアップ済',next:'電話',due:'2026/10/02'});
ok('担当が見込み客を追加（担当は自分になる）', add.ok && sheets['CCS追っかけ'].data[add.row-1][2]==='村上');
ok('タスク完了を記録', webSetTaskDone('KTAN',5,true).ok && sheets['入会後行動リスト'].data[4][7]==='✓');
let sent=[]; global.MailApp={getRemainingDailyQuota:()=>100,sendEmail:o=>sent.push(o)}; CONFIG.TEST_MODE=false;
sendAlerts_();
ok('アラートメールがOM・LD・担当に届く', sent.some(o=>o.to==='om@x.jp') && sent.some(o=>o.to==='ishikawa@x.jp') && sent.some(o=>o.to==='murakami@x.jp'));
if(process.argv.includes('--dump')){fs.mkdirSync(path.join(__dirname,'out'),{recursive:true});
  fs.writeFileSync(path.join(__dirname,'out/payload_om.json'),JSON.stringify(webGetData('KOM')));
  fs.writeFileSync(path.join(__dirname,'out/payload_tan.json'),JSON.stringify(webGetData('KTAN')));console.log('tests/out に画面確認用データを書き出しました');}
console.log('\n'+n+'件すべて成功');
