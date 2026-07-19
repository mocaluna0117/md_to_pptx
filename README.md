# Marp → PPTX

Markdown（Marp 記法）をブラウザ上で PowerPoint（.pptx）に変換する Web アプリ。
**バックエンド不要 / フロントエンドだけで完結**します。

## 使い方

```bash
npm install
npm run dev      # 開発サーバー（http://localhost:5173）
npm run build    # 本番ビルド → dist/
npm run preview  # ビルド結果をローカルで確認
```

左のエディタに Markdown を書き、`---` でスライドを区切って、右上の **PPTX を書き出す** を押すと `.pptx` がダウンロードされます。

## 表示モード（ツールバー左の切替）

- **Markdown**：エディタ＋ライブプレビュー（従来の画面）。方式A／Bで書き出し。
- **ビジュアル編集**：スライドを「テキストボックスの集まり」として直接編集する画面。
  Markdown を叩き台に読み込み、以後はボックスを操作します。

### ビジュアル編集でできること

- **左に全スライドのサムネイル一覧**（PowerPoint 風）。クリックで切替、× で削除、＋で追加
- **ボックスをドラッグで移動**、角をドラッグで**リサイズ**
- **ダブルクリックで文字を直接編集**
- 編集中に**文字を選択して色を変更**（スウォッチ or カスタム色）
- 文字サイズ・揃え、ボックスの追加／削除、スライドの追加／削除、背景色
- 「Markdown から作り直す」で叩き台を再生成
- 書き出しは**編集可能な PPTX**（各ボックス → ネイティブのテキストボックス）

> 座標や文字ごとの色は Markdown では表現できないため、ビジュアル編集の内容は
> Markdown には戻りません（Markdown は読み込みの起点）。

## 変換方式（ツールバーで切替）

| 方式 | 見た目 | PowerPoint での編集 | 用途 |
| --- | --- | --- | --- |
| **画像方式** | Marp の見た目そのまま（高再現） | 不可（スライド＝画像） | 見た目重視・そのまま配布 |
| **編集可能方式** | 簡易（自前レイアウト） | 可能（本物のテキスト/表） | あとから PowerPoint で編集 |

## 仕組み（方式A：画像方式）

各スライドを Marp で描画 → PNG に変換 → PowerPoint に全面貼り付け、という流れです。
Marp の見た目をそのまま高再現できる代わりに、PowerPoint 上ではテキスト編集はできません（スライド＝画像）。

| 役割 | ライブラリ | 場所 |
| --- | --- | --- |
| Markdown → スライド描画 | `@marp-team/marp-core` | ブラウザ |
| スライドを PNG 化 | `html-to-image` | ブラウザ |
| .pptx 生成・ダウンロード | `pptxgenjs` | ブラウザ |

> Marp 公式 CLI の `--pptx` はヘッドレス Chrome（Node）が必要なため、ブラウザ単体では使えません。
> そのため本アプリは Marp Core（描画）＋ pptxgenjs（生成）を組み合わせています。

## 仕組み（方式B：編集可能方式）

Markdown を解析し、pptxgenjs のネイティブ要素に変換します。テキストが本物のオブジェクトに
なるため PowerPoint で編集できます（ファイルサイズも画像方式より大幅に小さい）。

- 見出し → タイトル / 見出しテキスト
- 段落・箇条書き（ネスト対応）・番号付きリスト → 本文テキストボックス（自動で折返し・縦積み）
- **太字** / *斜体* / `コード` / ~~取消~~ / リンク → テキストランごとに反映
- コードブロック・引用 → 等幅 / イタリックのテキスト
- 表 → `addTable`、画像 → `addImage`（データ URL / CORS 許可の画像）

### 主なファイル

- [src/lib/marp.ts](src/lib/marp.ts) — Marp の描画（プレビュー用と書き出し用の 2 インスタンス）
- [src/lib/exportPptx.ts](src/lib/exportPptx.ts) — 方式A：オフスクリーン描画 → ラスタライズ → pptxgenjs
- [src/lib/markdownModel.ts](src/lib/markdownModel.ts) — 方式B：Markdown → スライド構造モデル（markdown-it）
- [src/lib/exportPptxNative.ts](src/lib/exportPptxNative.ts) — 方式B：モデル → ネイティブ pptx 生成
- [src/lib/frontmatter.ts](src/lib/frontmatter.ts) — フロントマター（背景色・文字色ディレクティブ）の読み取り
- [src/lib/deck.ts](src/lib/deck.ts) — ビジュアル編集のドキュメントモデル＋Markdownからの叩き台生成
- [src/lib/richText.ts](src/lib/richText.ts) — ボックスの runs ⇄ HTML（contentEditable 用）
- [src/lib/exportDeck.ts](src/lib/exportDeck.ts) — ビジュアル deck → 編集可能 pptx
- [src/components/VisualEditor.tsx](src/components/VisualEditor.tsx) — ビジュアル編集のキャンバス UI
- [src/App.tsx](src/App.tsx) — 表示モード切替 / エディタ / プレビュー / スタイル調整 / 書き出し UI

## 実装メモ

- 書き出しは `inlineSVG: false` で描画し、素の `<section>`（ネストした foreignObject なし）を
  キャプチャするため、ラスタライズが安定します。プレビューは既定の inline-SVG で
  ペイン幅にレスポンシブに追従します。
- スライドサイズは Marp 既定の 1280×720（16:9）。書き出す PPTX も同じアスペクト比です。

## 今後の改善候補

- 方式B のレイアウト精度向上（本文の高さ推定、テキスト＋画像の混在配置）
- テーマ切り替え UI（gaia / uncover）、`.md` ファイルのドラッグ＆ドロップ読み込み
- バンドル分割（KaTeX/MathJax の遅延読み込み）でサイズ削減
