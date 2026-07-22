# Markdown → Office（Deckdown / Docdown）

Markdown を、そのまま **編集できる PowerPoint / Word** に変換する Web アプリ。
**バックエンド不要 / フロントエンドだけで完結**します。起動するとランチャー（ホーム）が開き、
2 つのツールを選べます。

- **Deckdown** — Markdown → **PowerPoint（.pptx）** / PDF。スライドを直接編集して書き出し。
- **Docdown** — Markdown → **Word（.docx）**。文書としてプレビューして書き出し（`docx` ライブラリで生成）。

3 ページ構成（ホーム／Deckdown／Docdown）で、ハッシュルーティング（`#/`・`#/slides`・`#/docx`）。
各ツール左上の **⌂** でホームに戻れます。以降は主に **Deckdown** の説明です。

## 使い方

```bash
npm install
npm run dev      # 開発サーバー（http://localhost:5173）
npm run build    # 本番ビルド → dist/
npm run preview  # ビルド結果をローカルで確認
```

## 画面の流れ

- **中央：スライド編集画面**（プレビュー兼エディタ）。読み込んだ Markdown がスライドとして表示され、
  PowerPoint のように直接編集できます。
- **左：「Markdown」ドロワー**（開閉式）。元の Markdown を書いたりファイルをインポートしたりして、
  **「プレビューに反映」** で中央のスライドを作り直します。境界の縦タブは、クリックで開閉・
  ドラッグで幅を調整できます（幅は保存されます）。
- **右上：書き出し**。**「書き出す ▾」** から **PPTX（編集可能）** か **PDF** を選んでダウンロード。
- **🔄 初期化**：すべてデフォルトのサンプルに戻します（確認あり・自動保存もクリア）。
- **？ 使い方**：ツールバー左のボタンで操作ガイドを表示。**AI（ChatGPT 等）にスライド用 Markdown を
  作らせるプロンプト**（コピー可）も載っています（もちろん自分で 1 から書くことも可能）。

Markdown ファイルは、ドロワー上部の **「📂 インポート」**（現在の内容を置き換え・対応: `.md` /
`.markdown` / `.txt`）または **ドラッグ&ドロップ**で読み込めます（ファイル名は出力名にも反映）。
2 つめ以降を **「＋ 結合」** で読み込むと、既存の Markdown に `---` 区切りで**連結**されます
（読み込むファイルのフロントマターは除去されます）。読み込むと即プレビューに反映されます。

## スライド編集でできること

- **左に全スライドのサムネイル一覧**（PowerPoint 風）。クリックで切替、× で削除、＋で追加
- **テキストボックスをドラッグで移動**、角をドラッグで**リサイズ**
- **ダブルクリックで文字を直接編集**、編集中に**文字を選択して太字・斜体・色を変更**
- **フォント**・文字サイズ・揃え、テキストボックスの追加／削除、スライドの追加／削除
- **画像も反映**（データ URI／`<img>`／`![](…)`）。移動・リサイズ・削除でき、書き出しにも含まれます
- **表も反映**（Markdown の表 → 編集可能な表）。移動・リサイズ・**セルをダブルクリックで編集**・**列境界をドラッグで列幅調整**でき、書き出しは PowerPoint の**ネイティブ表**になります。ツールバーの **＋表** で新規追加も可
- **コードブロック**（``` フェンス）は**等幅・空白/改行を保持**して表示・書き出し（ASCII 図やフロー図も崩れません）
- **Undo / Redo**（ツールバーの曲がり矢印、または Ctrl/⌘+Z・Ctrl/⌘+Shift+Z）。ドラッグ 1 回＝Undo 1 回
- ↑↓←→ でスライド切替、選択中に Backspace / Delete で削除

Markdown からの取り込みは **Marp の実描画結果ベース**で、テーマ／カスタム CSS の
**背景色・文字色・フォントサイズ・太字・各ブロックの位置**を反映します（インライン HTML も解決）。
※ positioned box のため、完全一致や 2 カラム等の複雑な CSS レイアウトは再現しません。

> 座標や文字ごとの装飾は Markdown では表現できないため、スライド編集の内容は Markdown には戻りません
> （Markdown は読み込みの起点）。編集内容と Markdown は **localStorage に自動保存**され、リロードしても
> 続きから再開できます。

## 書き出し

| 形式 | 中身 | PowerPoint での編集 |
| --- | --- | --- |
| **PPTX（編集可能）** | テキストボックス → ネイティブのテキストボックス、画像 → `addImage`、表 → `addTable`（ネイティブ表） | 可能 |
| **PDF** | 各スライドを画像化し、1 ページずつ収録 | － |

## 仕組み

Markdown を Marp Core で描画し、その結果からスライドを「テキストボックス＋画像の集まり（deck）」に
変換します。deck を pptxgenjs のネイティブ要素として書き出すため、PowerPoint で編集できます。

| 役割 | ライブラリ |
| --- | --- |
| Markdown → スライド描画 | `@marp-team/marp-core` |
| deck → 編集可能 pptx | `pptxgenjs` |
| deck → PDF（画像化） | `html-to-image` + `jspdf` |

> Marp 公式 CLI の `--pptx` はヘッドレス Chrome（Node）が必要でブラウザ単体では使えないため、
> 本アプリは Marp Core（描画）＋ pptxgenjs（生成）を組み合わせています。

### 主なファイル

- [src/lib/marp.ts](src/lib/marp.ts) — Marp の描画
- [src/lib/deckFromRender.ts](src/lib/deckFromRender.ts) — Marp の実描画結果 → deck（叩き台）
- [src/lib/deck.ts](src/lib/deck.ts) — deck のドキュメントモデル
- [src/lib/richText.ts](src/lib/richText.ts) — ボックスの runs ⇄ HTML（contentEditable 用）
- [src/lib/exportDeck.ts](src/lib/exportDeck.ts) — deck → 編集可能 pptx
- [src/lib/exportPdf.ts](src/lib/exportPdf.ts) / [src/lib/rasterize.ts](src/lib/rasterize.ts) — deck → PDF
- [src/components/VisualEditor.tsx](src/components/VisualEditor.tsx) — スライド編集のキャンバス UI
- [src/App.tsx](src/App.tsx) — 全体レイアウト / Markdown ドロワー / 書き出し UI

## 実装メモ

- スライドサイズは Marp 既定の 1280×720（16:9）。書き出す PPTX / PDF も同じアスペクト比です。
- deck の取り込みは `inlineSVG: false` で描画した素の `<section>` を基準に、各ブロックの位置・
  計算済みスタイルを読み取ります。

## 今後の改善候補

- レイアウト精度向上（本文の高さ推定、テキスト＋画像の混在配置、2 カラム対応）
- テンプレート / テーマの追加
- バンドル分割（KaTeX/MathJax の遅延読み込み）でサイズ削減
