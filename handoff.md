# 引き継ぎドキュメント（handoff）

このリポジトリを別のアカウント／セッションで引き継ぐための要点をまとめます。
コード上の事実（構造・過去の修正）は git とソースが正なので、ここには **背景・設計判断・流儀・落とし穴** を中心に記します。

## 1. これは何か

**Markdown を、そのまま編集できる Office ファイルに変換するフロントエンド完結の Web アプリ。**
バックエンド無し・無料・内容は外部送信しない（すべてブラウザ内で処理）。起動するとランチャー（ホーム）が開き、2 つのツールを選ぶ。

- **Deckdown** … Markdown（Marp 記法）→ **PowerPoint（.pptx、編集可能）/ PDF**。スライドを「テキストボックス＋画像＋表の集まり（deck）」として直接ビジュアル編集して書き出す。
- **Docdown** … Markdown → **Word（.docx、編集可能）**。文書としてライブプレビューして書き出す（`docx` ライブラリで生成）。

設計思想（ヘルプの「意義」セクションにも記載）：AI に直接スライドを作らせるより **Markdown を書かせる方がトークンが大幅に安い**／**微調整は AI に投げ直さず自分で**／**本物の編集できる成果物**／**どの AI でも・手書きでも**／**ブラウザ完結・無料**。

- 技術: TypeScript + React 19 + Vite。状態は各コンポーネントの useState と localStorage。ルーターライブラリ無し（自前のハッシュルーティング）。
- 公開: **https://mocaluna0117.github.io/md_to_pptx/**
- リポジトリ: `github.com/mocaluna0117/md_to_pptx`（名前は md_to_pptx のままだが実質 2 アプリ）

## 2. デプロイ（GitHub Pages）

- `.github/workflows/deploy.yml` が **main への push で自動ビルド＆デプロイ**（Pages の Source = "GitHub Actions"）。
- プロジェクトサイトのサブパス配信のため、`vite.config.ts` は `base: process.env.BASE_PATH ?? '/'`。workflow が `BASE_PATH=/md_to_pptx/` を渡す。**ローカルの dev/build は `/` のまま**。
- ハッシュルーティング（`#/`・`#/slides`・`#/docx`）なので、サブパス配信でもリロードで壊れない。
- デプロイ確認は `gh run watch <id>` と本番 URL への `curl`。

## 3. 全体アーキテクチャ（3 ページ）

`src/main.tsx` → `src/Root.tsx`（ハッシュルーター）→ 以下を切替表示：
- `home` … `src/Home.tsx`（ランチャー、2 カード）
- `slides` … `src/App.tsx`（**Deckdown 本体**。default export が `App`）
- `docx` … `src/Docdown.tsx`

`navigate(route)`（Root.tsx が export）は `window.location.hash` を書き換えるだけ。各アプリ左上の **⌂** ボタンが `navigate('home')`。

## 4. 主要ファイルと役割

- `src/Root.tsx` — ハッシュルーティング、`navigate()`。
- `src/Home.tsx` / `Home.css` — ランチャー。
- `src/App.tsx` — **Deckdown**。ツールバー／Markdown ドロワー（開閉＆幅ドラッグ）／`VisualEditor`／書き出しメニュー／使い方ダイアログ／undo-redo 履歴。**一番大きいファイル**。
- `src/components/VisualEditor.tsx` — スライド編集キャンバス。サムネイル、テキストボックス／画像／表の移動・リサイズ・編集、太字/斜体/色/フォント/整列、列幅・行高のドラッグ。
- `src/Docdown.tsx` / `Docdown.css` — **Docdown**。Markdown ドロワー＋文書プレビュー＋.docx 書き出し。App.css を流用。
- `src/App.css` — Deckdown＋共通 UI の巨大スタイルシート（ツールバー、ドロワー、`.attached`、ヘルプ等）。Docdown も import する。
- `src/lib/deck.ts` — deck のデータモデル（`Slide`/`Box`/`ImageEl`/`TableEl`/`TextRun`）とヘルパー（`newBox`/`newTable`/`tableColFractions`/`tableRowFractions`/`toHex` など）。
- `src/lib/deckFromRender.ts` — **Marp の実描画結果 → deck**。`deckFromRenderedMarkdown()`。ブロックごとに位置・計算済みスタイルを読んで Box 化。`<table>` → TableEl、`<pre>` → `pre` フラグ付き Box、`<img>` → ImageEl。
- `src/lib/marp.ts` — Marp Core の描画（`renderSlides` など）。
- `src/lib/richText.ts` — Box の runs ⇄ HTML（contentEditable 用、`data-fs` で per-run フォントサイズ）。
- `src/lib/exportDeck.ts` — deck → 編集可能 .pptx（pptxgenjs。addText/addImage/addTable）。
- `src/lib/exportPdf.ts` + `rasterize.ts` — deck / Marp を画像化して .pdf（jspdf。JPEG でサイズ削減）。`rasterizeDeck` は tables/pre/fontFamily も描画。
- `src/lib/exportDocx.ts` — **Markdown → .docx**（`docx` ライブラリ）。markdown-it トークンを Paragraph/Table/ImageRun 等にマッピング。画像は data URI を PNG 化して埋め込み。
- `src/lib/imageAttach.ts` — **相対パス画像の解決**。`resolveImagePaths(md, images)` が `![](fig.png)`/`<img src>` をファイル名一致で添付データ URI に置換（元 md は非破壊）。`readImageFiles`。
- `src/lib/math.ts` — **数式**。`mathToImages(md)` が `$…$`/`$$…$$` を **MathJax の SVG（グリフをパス内包）→ PNG** 化して `<img>`（高さ em 指定）に差し替え。MathJax は動的 import、式ごとにキャッシュ。

## 5. Deckdown（App.tsx）の仕組み

- 起動時／インポート時に Markdown から deck を生成（`buildDeck`）。deck ができるまで「スライドを生成中…」。
- **前処理チェーン**：`buildDeck(src)` は `deckFromRenderedMarkdown(await mathToImages(resolveImagePaths(src, images)))`。つまり **相対パス画像解決 → 数式画像化 → Marp 描画 → deck 化** の順。
- Markdown は左の開閉ドロワー（幅ドラッグ可、localStorage 保存）。**「プレビューに反映」** で `rebuildFromMarkdown`（visual 編集があると上書き確認）。
- deck 編集は VisualEditor 内で完結。**Markdown には戻らない**（座標や per-char 装飾は md で表せないため、md は取り込みの起点）。
- **undo/redo は App.tsx が保持**（`undoRef`/`redoRef`、coalesceKey で 1 ドラッグ＝1 ステップ）。VisualEditor は `onChange(deck, coalesceKey?)` で通知。
- 書き出し：`書き出す ▾` → PPTX（編集可能・`exportDeckToPptx`）/ PDF（`exportDeckToPdf`）。
- 永続化：`localStorage['md-to-pptx:v1']`（markdown/fileName/deck/deckDirty/mdOpen/drawerWidth）。deck が大きい（画像）と quota 超過→ deck を落として md だけ保存。

## 6. Docdown（Docdown.tsx）の仕組み

- 文書は流し込みなので **ビジュアル編集は無し**。左ドロワーの Markdown を **ライブプレビュー**（markdown-it、`html:true`）。
- プレビューは非同期（数式ラスタライズのため）：effect で `mdRender.render(stripFrontmatter(await mathToImages(resolveImagePaths(markdown, images))))` を 200ms デバウンスで計算。
- 書き出し：`exportMarkdownToDocx(await mathToImages(resolveImagePaths(markdown, images)))`。
- 永続化：`localStorage['docdown:v1']`（markdown/fileName/mdOpen/images）。画像で quota 超過→テキストのみ保存。
- markdown-it は `html:true`（`<img>` 等の生 HTML を描画するため）＝貼り付け内容次第で XSS 面はあるが自分の内容前提。

## 7. 共通の仕組み

- **相対パス画像（imageAttach.ts）**：`![](fig.png)` はブラウザからファイル参照できない。ドロワーの「🖼 画像」or 画像 D&D で読み込むと `Record<basename, dataURI>` に入り、描画/書き出し時に `resolveImagePaths` で解決。Deckdown は未編集なら添付時に自動再ビルド、Docdown はライブ反映。`data:`/絶対 URL はそのまま。
- **数式（math.ts）**：KaTeX+html-to-image は **フォント埋め込み失敗で空画像**になったため、**MathJax SVG → canvas → PNG** に変更（フォント非依存）。両アプリとも `<img>` PNG になるので既存の画像パイプラインを通る。通貨の `$` を式と誤認しないよう配慮（必要なら `\$`）。
- **画像埋め込み**：docx（exportDocx.ts）は各画像を Image に読み込み→canvas→PNG 化して ImageRun。CORS 不許可の外部画像は安全にスキップ。

## 8. 開発・ビルド・テスト

```bash
npm install
npm run dev       # http://localhost:5173/（base は '/'）
npm run build     # tsc -b && vite build（型チェック込み）
npm run preview   # 本番相当のローカル確認
npm run lint      # oxlint
```

**E2E 検証パターン（このプロジェクトで多用）**：
- `npm install -D puppeteer-core@latest` →（`npm run preview -- --port 4188` or `npm run dev -- --port 4199`）で起動 → `/Applications/Google Chrome.app/.../Google Chrome`（headless:'shell', --no-sandbox）で駆動 → 終わったら `npm remove puppeteer-core`。
- テスト `.mjs` は **プロジェクト直下**に置く（node_modules 解決のため）。CDP `Browser.setDownloadBehavior` でダウンロード取得。
- 落とし穴：React の controlled input は `el.value=` を無視 → **native setter + `input` イベント**で入力を模倣。pptxgenjs は常に空の `ppt/media/` を作るので「画像あり」は *ファイル* を数える。モジュール単体検証は `npm run dev` + ページ内 `await import('/src/lib/…')`。画像/数式の“見えている”検証は naturalWidth だけでなく **canvas でピクセル（暗い画素数）を確認**（空画像を見抜くため）。

## 9. 開発の流儀（重要）

- **コミットはローカルで作り、push はユーザーが明示指示してから**（毎回聞く）。
- コミットは per-commit の identity を付与：
  `git -c user.name="mocaluna0117" -c user.email="daibon20020117@gmail.com" commit …`
  末尾トレーラ：`Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`
- 変更したら基本 `npm run build`（型）＋ 必要なら E2E で確認してからコミット。
- ユーザーは日本語。UI 文言・説明も日本語。

## 10. 既知の制約・注意点

- Deckdown の deck は **positioned box** モデル。2 カラム等の複雑 CSS レイアウトは完全再現しない。
- Deckdown の **インライン数式は行内埋め込みではなく位置指定の画像**として配置される（文書のような流し込みではない）。表示は読める。
- Docdown の **リンクは青字＋下線の見た目のみ**（クリック可能なハイパーリンクは未実装）。
- 数式 PNG はダーク（#111）固定。暗い背景のスライドでは見えづらい可能性。
- MathJax は動的 import（数式使用時のみロード）。初回は少し重い。

## 11. 未使用・レガシーファイル（削除候補）

過去の構成の名残で、現在の UI からは使われていない：
- `src/lib/exportPptx.ts`（画像方式 PPTX）
- `src/lib/exportPptxNative.ts` ＋ `src/lib/markdownModel.ts`（Markdown 直→ネイティブ pptx）
- `src/lib/frontmatter.ts`（deck.ts の `deckFromMarkdown` フォールバックが一部参照）
- `src/assets/*`（Vite 雛形の残り）

ビルドは通る（無害）。整理するなら参照有無を確認して削除可。

## 12. 今後の候補

- Docdown 用の「使い方」ダイアログ、リンクの実ハイパーリンク化、docx の数式を OMML 化（画像でなくネイティブ数式）
- Deckdown の数式インライン化、テンプレート/テーマ
- 未使用レガシーの整理、バンドル分割
