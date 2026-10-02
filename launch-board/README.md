# ローンチ進捗ボード　Claude Code での進め方

このフォルダをClaude Codeで開くと、`CLAUDE.md` が自動で読み込まれ、プロジェクトの決まりごとを理解した状態で始められます。

## 1. 最初の準備

1. このフォルダを、パソコンの好きな場所に置きます（例：書類 → launch-board）。
2. Claude Codeで、このフォルダを開きます。
3. Node.js が入っていない場合は、Claude Codeに「Node.jsを入れて」と頼んでください（ビルドとテストに使います）。

## 2. 頼み方の例

- 「入会者の画面に、ローンチで絞り込むプルダウンを追加して」
- 「ホームの声かけ目標を、ローンチごとに設定シートから読むようにして」
- 「日次実績の入力に『1to1数』を追加して。管理表の日次ログの列も合わせて」
- 「テストを実行して、プレビューを作って」

Claude Codeは `src/` を直し、`dist/Code.gs` を作り直して、テストまで実行します。

## 3. 管理表（Apps Script）への反映

### 方法A：貼り付ける（今までと同じ）

`dist/Code.gs` の中身をすべてコピーし、管理表の 拡張機能 → Apps Script に貼り替えて保存します。

### 方法B：clasp で自動反映（慣れたらおすすめ）

clasp はGoogle公式の、Apps Scriptをパソコンから更新する道具です。最初の設定はClaude Codeに「claspを設定して」と頼めば、一緒に進められます。

1. https://script.google.com/home/usersettings を開き、「Google Apps Script API」をオンにします。
2. `npx clasp login` を実行し、ブラウザでGoogleアカウントを選んで許可します。
3. `.clasp.json.example` を `.clasp.json` という名前でコピーし、スクリプトIDを入れます。
   - スクリプトIDは、Apps Script画面の左の歯車「プロジェクトの設定」にあります。
4. 以降は `npx clasp push` で、`dist/` の内容がApps Scriptに反映されます。

⚠ `clasp push` は、Apps Script側のファイルを `dist/` の内容で置き換えます。Apps Scriptの画面で直接書き換えた内容がある場合は、先に `src/` に反映してから push してください。

### 反映したあと（方法A・B共通）

サイト（進捗ボード）に反映するには、Apps Scriptで **デプロイ → デプロイを管理 → 鉛筆マーク → バージョン：新バージョン → デプロイ** を押します。
「新しいデプロイ」を押すとURLが変わり、配った個人用リンクが使えなくなるので注意してください。

## 4. 画面の確認（デプロイ前）

`node tools/preview.mjs` を実行すると、`tests/out/preview_om.html`（OMの画面）と `preview_tan.html`（担当メンバーの画面）ができます。ブラウザで開くと、架空のデータで見た目を確認できます（保存はできません）。

## フォルダの中身

| 場所 | 内容 |
|---|---|
| `CLAUDE.md` | Claude Code向けのプロジェクト説明と決まりごと |
| `src/` | 編集するファイル（main.gs／site.gs／page.html） |
| `dist/Code.gs` | 貼り付け用の完成ファイル（自動生成） |
| `tests/` | テスト |
| `tools/` | プレビュー作成、管理表（xlsx）作成スクリプト |
| `docs/` | 管理表のひな形（v5） |
