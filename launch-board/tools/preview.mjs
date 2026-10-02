// 画面をブラウザで確認するためのプレビューを作る（Googleにデプロイしなくても見た目を確認できる）
//   node tools/preview.mjs   → tests/out/preview_om.html と preview_tan.html をブラウザで開く
import { readFileSync, writeFileSync } from 'node:fs';
import { execSync } from 'node:child_process';
const root = new URL('../', import.meta.url);
execSync('node build.mjs && node tests/server.test.cjs --dump', { cwd: root, stdio: 'inherit' });
const page = readFileSync(new URL('src/page.html', root), 'utf8').replace('__BOOT__', '{"key":"PREVIEW"}');
for (const who of ['om', 'tan']) {
  const data = readFileSync(new URL(`tests/out/payload_${who}.json`, root), 'utf8');
  const stub = `<script>window.google={script:{run:(function(){const h={};const o={withSuccessHandler(f){h.s=f;return o},withFailureHandler(f){h.f=f;return o}};
    o.webGetData=()=>setTimeout(()=>h.s(${data}),50);
    ['webSaveLog','webUpdateProspect','webAddProspect','webSetTaskDone'].forEach(n=>o[n]=()=>setTimeout(()=>h.f(new Error('プレビューでは保存できません')),50));
    return o})()}};</script>`;
  writeFileSync(new URL(`tests/out/preview_${who}.html`, root), page.replace('<head>', '<head>' + stub));
}
console.log('tests/out/preview_om.html（OM画面）と preview_tan.html（担当画面）を作成しました');
