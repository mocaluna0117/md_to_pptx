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

### 主なファイル

- [src/lib/marp.ts](src/lib/marp.ts) — Marp の描画（プレビュー用と書き出し用の 2 インスタンス）
- [src/lib/exportPptx.ts](src/lib/exportPptx.ts) — オフスクリーン描画 → ラスタライズ → pptxgenjs
- [src/App.tsx](src/App.tsx) — エディタ / ライブプレビュー / 書き出し UI

## 実装メモ

- 書き出しは `inlineSVG: false` で描画し、素の `<section>`（ネストした foreignObject なし）を
  キャプチャするため、ラスタライズが安定します。プレビューは既定の inline-SVG で
  ペイン幅にレスポンシブに追従します。
- スライドサイズは Marp 既定の 1280×720（16:9）。書き出す PPTX も同じアスペクト比です。

## 今後（方式B：編集可能ネイティブ）

Markdown を解析して pptxgenjs のテキストボックス等に変換すれば、PowerPoint で編集可能な
オブジェクトとして出力できます（テーマ再現は自作が必要）。現状は方式A のみ実装。
