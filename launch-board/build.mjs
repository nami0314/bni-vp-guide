// src/ の3ファイルを、貼り付け用の1ファイル dist/Code.gs にまとめる
//   node build.mjs
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
const read = p => readFileSync(new URL(p, import.meta.url), 'utf8');
const main = read('./src/main.gs').trimEnd();
const site = read('./src/site.gs').trimEnd();
const page = read('./src/page.html').replace(/\n$/, '').split('\n');
if (!page.join('\n').includes('__BOOT__')) throw new Error('page.html に __BOOT__ がありません（doGetが起動情報を差し込む場所）');
const html = '/* ---------- 画面（HTML）。src/page.html から自動生成。ここは直接編集しない ---------- */\n' +
  'const PAGE_HTML_LINES = [\n' + page.map(l => '  ' + JSON.stringify(l)).join(',\n') + '\n];\n';
mkdirSync(new URL('./dist/', import.meta.url), { recursive: true });
writeFileSync(new URL('./dist/Code.gs', import.meta.url), main + '\n\n' + site + '\n\n' + html);
console.log('dist/Code.gs を作成しました（' + (main + site + html).split('\n').length + '行）');
