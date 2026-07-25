import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import MarkdownIt from 'markdown-it'
import markdownItCjkFriendly from 'markdown-it-cjk-friendly'
import { navigate } from './Root'
import type { DocSettings } from './lib/exportDocx'
import { resolveImagePaths, readImageFiles, IMAGE_EXT, type AttachedImages } from './lib/imageAttach'
import { mathToImages } from './lib/math'
import { mermaidToImages } from './lib/mermaid'
import { copyText } from './lib/clipboard'
import type { DocBox } from './lib/docBox'
import DocEditor from './components/DocEditor'
import './App.css'
import './Docdown.css'

const STORAGE_KEY = 'docdown:v1'
const mdRender = new MarkdownIt({ html: true, linkify: true, breaks: false }).use(markdownItCjkFriendly)
const MIN_DRAWER = 240
const MAX_DRAWER = 760

const SAMPLE = `# ドキュメントのタイトル

Markdown で書いた文章を、そのまま **編集できる Word（.docx）** に書き出せます。
**太字**・*斜体*・\`コード\`・[リンク](https://example.com) が使えます。

## 使い方

1. 左の「Markdown」に文章を書く（またはインポート）
2. 中央のプレビューで見た目を確認
3. 右上の「Word で書き出す」でダウンロード

## 対応している記法

- 見出し（\`#\`〜\`######\`）
- 箇条書き / 番号付きリスト（ネスト可）
- 表

| 項目 | 説明 |
| --- | --- |
| 見出し | Word の見出しスタイルに変換 |
| 表 | ネイティブの Word の表に変換 |

> 引用も使えます。

## 数式（LaTeX）

インラインは $E = mc^2$、ディスプレイは次のように書けます。

$$\\int_0^1 x^2\\,dx = \\frac{1}{3}$$

\`\`\`
コードブロックは等幅で出力されます
\`\`\`
`

interface Persisted {
  markdown?: string
  fileName?: string
  mdOpen?: boolean
  images?: AttachedImages
  docHtml?: string
  docDirty?: boolean
  boxes?: DocBox[]
  drawerWidth?: number
  mdFileName?: string
  docSettings?: DocSettings
}

/** Preview font stacks for the exportable font names. */
const DOC_FONT_CSS: Record<string, string> = {
  'Yu Mincho': '"Yu Mincho", "Hiragino Mincho ProN", serif',
  'MS Mincho': '"MS Mincho", "Hiragino Mincho ProN", serif',
  'Yu Gothic': '"Yu Gothic", "Hiragino Kaku Gothic ProN", sans-serif',
  Meiryo: 'Meiryo, "Hiragino Kaku Gothic ProN", sans-serif',
}

function loadPersisted(): Persisted {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') as Persisted
  } catch {
    return {}
  }
}
const persisted = loadPersisted()

/** Remove a leading YAML front-matter block (--- … ---). */
function stripFrontmatter(m: string): string {
  return m.replace(/^﻿?---[^\n]*\n[\s\S]*?\n---[^\n]*\n?/, '')
}
/** Append an imported Markdown to the current one (own front-matter dropped). */
function mergeMarkdown(base: string, add: string): string {
  if (!base.trim()) return add
  const body = stripFrontmatter(add).trim()
  if (!body) return base
  return `${base.trimEnd()}\n\n${body}\n`
}

type Status = 'idle' | 'exporting' | { error: string }

/** A prompt users can hand to an AI (ChatGPT etc.) to generate document-ready Markdown. */
const AI_PROMPT = `次の内容を文書（レポート/ドキュメント）にまとめて、Markdown（.md）で出力してください。

# ルール
- 見出し（#〜###）で章立てする
- 箇条書き・番号付きリスト・表を活用する
- 強調は **太字**・*斜体* を使う
- 数式が必要なら $…$ / $$…$$（LaTeX）で書く
- 装飾は Markdown のみ（HTML タグは使わない）

# まとめたい内容
（ここに伝えたい内容や、添付画像の説明を書いてください）`

export default function Docdown() {
  const [markdown, setMarkdown] = useState(persisted.markdown ?? SAMPLE)
  const [fileName, setFileName] = useState(persisted.fileName ?? 'document')
  // Name of the imported Markdown file currently loaded (empty when using the built-in sample).
  const [mdFileName, setMdFileName] = useState<string>(persisted.mdFileName ?? '')
  const [mdOpen, setMdOpen] = useState<boolean>(persisted.mdOpen ?? true)
  const [drawerWidth, setDrawerWidth] = useState<number>(
    Math.min(MAX_DRAWER, Math.max(MIN_DRAWER, persisted.drawerWidth ?? 360)),
  )
  const [resizing, setResizing] = useState(false)
  const [images, setImages] = useState<AttachedImages>(persisted.images ?? {})
  const [status, setStatus] = useState<Status>('idle')
  const [exportMenuOpen, setExportMenuOpen] = useState(false)
  const [docSettings, setDocSettings] = useState<DocSettings>(persisted.docSettings ?? {})
  const [settingsOpen, setSettingsOpen] = useState(false)
  const settingsWrapRef = useRef<HTMLDivElement>(null)
  const patchSettings = (patch: Partial<DocSettings>) => setDocSettings((s) => ({ ...s, ...patch }))

  // Close the 文書設定 popover on an outside click.
  useEffect(() => {
    if (!settingsOpen) return
    const onDown = (e: PointerEvent) => {
      if (!settingsWrapRef.current?.contains(e.target as Node)) setSettingsOpen(false)
    }
    window.addEventListener('pointerdown', onDown)
    return () => window.removeEventListener('pointerdown', onDown)
  }, [settingsOpen])
  const [helpOpen, setHelpOpen] = useState(false)
  const [promptCopied, setPromptCopied] = useState(false)

  async function copyPrompt() {
    if (await copyText(AI_PROMPT)) {
      setPromptCopied(true)
      setTimeout(() => setPromptCopied(false), 1500)
    }
  }

  // Close the "使い方" dialog on Escape.
  useEffect(() => {
    if (!helpOpen) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setHelpOpen(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [helpOpen])
  const fileInputRef = useRef<HTMLInputElement>(null)
  const imageInputRef = useRef<HTMLInputElement>(null)
  const importModeRef = useRef<'replace' | 'append'>('replace')
  const exportWrapRef = useRef<HTMLDivElement>(null)
  const workspaceRef = useRef<HTMLDivElement>(null)

  // Size the filename input to the actual rendered text width (measured via a
  // hidden mirror), so the box hugs short names instead of over-reserving space.
  const fnInputRef = useRef<HTMLInputElement>(null)
  const fnSizerRef = useRef<HTMLSpanElement>(null)
  useLayoutEffect(() => {
    const input = fnInputRef.current
    const sizer = fnSizerRef.current
    if (!input || !sizer) return
    const textW = sizer.getBoundingClientRect().width
    input.style.width = `${Math.min(textW + 6, 320)}px` // +6px caret room, 320px cap
  }, [fileName])

  // Drag the drawer handle to resize; a click without movement toggles it (mirrors Deckdown).
  function onHandlePointerDown(e: React.PointerEvent) {
    if (e.button !== 0) return
    const ws = workspaceRef.current
    if (!ws) return
    const startX = e.clientX
    const wsLeft = ws.getBoundingClientRect().left
    let moved = false
    const onMove = (ev: PointerEvent) => {
      if (!moved && Math.abs(ev.clientX - startX) > 4) {
        moved = true
        setResizing(true)
        setMdOpen(true)
      }
      if (moved) {
        setDrawerWidth(Math.max(MIN_DRAWER, Math.min(MAX_DRAWER, ev.clientX - wsLeft)))
      }
    }
    const onUp = () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
      if (moved) setResizing(false)
      else setMdOpen((o) => !o)
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
  }

  function onHandleKeyDown(e: React.KeyboardEvent) {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault()
      setMdOpen((o) => !o)
    }
  }

  // The edited document (HTML) is the source of truth once the user edits visually;
  // Markdown is the import starting point. `rebuildToken` remounts the editor to reseed.
  const [docHtml, setDocHtml] = useState<string>(persisted.docHtml ?? '')
  const [docDirty, setDocDirty] = useState<boolean>(persisted.docDirty ?? false)
  const [boxes, setBoxes] = useState<DocBox[]>(persisted.boxes ?? [])
  const [rebuildToken, setRebuildToken] = useState(0)
  const docHtmlRef = useRef(docHtml)
  docHtmlRef.current = docHtml
  const boxesRef = useRef(boxes)
  boxesRef.current = boxes
  const imagesRef = useRef(images)
  imagesRef.current = images
  const imageNames = Object.keys(images)

  const exporting = status === 'exporting'
  const error = typeof status === 'object' ? status.error : null

  /** Render Markdown → document HTML (images + math baked in) and reseed the editor. */
  const buildDoc = useCallback(async (src: string) => {
    const withDiagrams = await mermaidToImages(resolveImagePaths(src, imagesRef.current))
    const prepared = await mathToImages(withDiagrams)
    const rendered = mdRender.render(stripFrontmatter(prepared))
    setDocHtml(rendered)
    docHtmlRef.current = rendered
    setDocDirty(false)
    setRebuildToken((t) => t + 1)
  }, [])

  // First boot: build the document from Markdown unless a saved document exists.
  const bootedRef = useRef(false)
  useEffect(() => {
    if (bootedRef.current) return
    bootedRef.current = true
    if (!docHtmlRef.current) void buildDoc(markdown)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  function handleDocChange(next: string) {
    setDocHtml(next)
    docHtmlRef.current = next
    setDocDirty(true)
  }

  /** Rebuild the document from the current Markdown (warns if there are visual edits). */
  async function rebuildFromMarkdown(): Promise<boolean> {
    if (docHtmlRef.current && docDirty && !window.confirm('現在の Markdown から文書を作り直します。編集した内容は上書きされます。よろしいですか？')) {
      return false
    }
    await buildDoc(markdown)
    return true
  }

  useEffect(() => {
    const id = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, mdFileName, mdOpen, drawerWidth, images, docHtml, docDirty, boxes, docSettings }))
      } catch {
        // Document HTML / images / boxes may exceed the storage quota: keep at least the text.
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, mdFileName, mdOpen, drawerWidth, docSettings }))
        } catch {
          /* storage unavailable */
        }
      }
    }, 300)
    return () => clearTimeout(id)
  }, [markdown, fileName, mdFileName, mdOpen, drawerWidth, images, docHtml, docDirty, boxes, docSettings])

  async function addImageFiles(files: File[]) {
    const imgs = files.filter((f) => IMAGE_EXT.test(f.name) || f.type.startsWith('image/'))
    if (imgs.length === 0) return
    const loaded = await readImageFiles(imgs)
    const merged = { ...imagesRef.current, ...loaded }
    imagesRef.current = merged
    setImages(merged)
    // With no visual edits yet, rebuild so relative-path images appear inline.
    if (!docDirty) void buildDoc(markdown)
  }

  async function importFile(file: File, mode: 'replace' | 'append') {
    try {
      const text = await file.text()
      if (mode === 'append') {
        setMarkdown((cur) => mergeMarkdown(cur, text))
      } else {
        if (docHtmlRef.current && docDirty && !window.confirm('読み込んだ Markdown で文書を作り直します。編集した内容は破棄されます。よろしいですか？')) {
          return
        }
        setMarkdown(text)
        setMdFileName(file.name)
        const base = file.name.replace(/\.[^.]+$/, '')
        if (base) setFileName(base)
        await buildDoc(text)
      }
    } catch {
      setStatus({ error: 'ファイルの読み込みに失敗しました。' })
    }
  }

  async function handleExportDocx() {
    setExportMenuOpen(false)
    setStatus('exporting')
    try {
      // The docx lib loads on demand, keeping it out of the app's initial chunk.
      const { exportHtmlToDocx } = await import('./lib/exportDocx')
      await exportHtmlToDocx(docHtmlRef.current, boxesRef.current, { fileName, settings: docSettings })
      setStatus('idle')
    } catch (err) {
      const message =
        err instanceof Error
          ? err.message
          : typeof err === 'string'
            ? err
            : '書き出しに失敗しました。読み込めない画像が含まれていないかご確認ください。'
      setStatus({ error: message })
    }
  }

  /**
   * PDF export via the browser's print-to-PDF: text stays vector (crisp at any zoom,
   * selectable, correct Japanese), which a rasterized PDF can't match. The @media print
   * stylesheet isolates the document page (+ boxes). The title becomes the default file name.
   */
  function handlePrintPdf() {
    setExportMenuOpen(false)
    const prev = document.title
    document.title = fileName?.trim() || 'document'
    const restore = () => {
      document.title = prev
      window.removeEventListener('afterprint', restore)
    }
    window.addEventListener('afterprint', restore)
    window.print()
  }

  // Close the export dropdown on an outside click.
  useEffect(() => {
    if (!exportMenuOpen) return
    const onDown = (e: PointerEvent) => {
      if (!exportWrapRef.current?.contains(e.target as Node)) setExportMenuOpen(false)
    }
    window.addEventListener('pointerdown', onDown)
    return () => window.removeEventListener('pointerdown', onDown)
  }, [exportMenuOpen])

  function resetToDefault() {
    if (!window.confirm('内容を初期状態に戻します。よろしいですか？')) return
    setMarkdown(SAMPLE)
    setFileName('document')
    setMdFileName('')
    imagesRef.current = {}
    setImages({})
    boxesRef.current = []
    setBoxes([])
    setStatus('idle')
    void buildDoc(SAMPLE)
    try {
      localStorage.removeItem(STORAGE_KEY)
    } catch {
      /* ignore */
    }
  }

  return (
    <div className="app">
      <header className="toolbar">
        <div className="tb-left">
          <button className="home-btn" onClick={() => navigate('home')} title="ホームに戻る" aria-label="ホームに戻る">
            ⌂
          </button>
          <button className="help-btn" onClick={() => setHelpOpen(true)} aria-haspopup="dialog">
            ？ 使い方
          </button>
        </div>
        <div className="brand">
          <h1>Docdown</h1>
          <span className="tagline">Markdown → Word</span>
        </div>
        <div className="actions">
          <div className="export-wrap" ref={settingsWrapRef}>
            <button
              className="reset"
              onClick={() => setSettingsOpen((o) => !o)}
              aria-haspopup="dialog"
              aria-expanded={settingsOpen}
              title="用紙・フォント・ページ番号などの文書設定"
            >
              ⚙ 文書設定
            </button>
            {settingsOpen && (
              <div className="export-menu doc-settings" role="dialog" aria-label="文書設定">
                <label className="ds-row">
                  <span>本文フォント</span>
                  <select
                    value={docSettings.font ?? ''}
                    onChange={(e) => patchSettings({ font: e.target.value || undefined })}
                  >
                    <option value="">既定（Calibri）</option>
                    <option value="Yu Mincho">游明朝</option>
                    <option value="MS Mincho">ＭＳ 明朝</option>
                    <option value="Yu Gothic">游ゴシック</option>
                    <option value="Meiryo">メイリオ</option>
                  </select>
                </label>
                <label className="ds-row">
                  <span>余白（A4）</span>
                  <select
                    value={docSettings.margin ?? 'normal'}
                    onChange={(e) => patchSettings({ margin: e.target.value as DocSettings['margin'] })}
                  >
                    <option value="narrow">狭い（12.7mm）</option>
                    <option value="normal">標準（25.4mm）</option>
                    <option value="wide">広い（38.1mm）</option>
                  </select>
                </label>
                <label className="ds-row">
                  <span>ヘッダー（右上）</span>
                  <input
                    type="text"
                    value={docSettings.headerText ?? ''}
                    placeholder="例: 学籍番号・氏名"
                    onChange={(e) => patchSettings({ headerText: e.target.value || undefined })}
                  />
                </label>
                <label className="ds-check">
                  <input
                    type="checkbox"
                    checked={!!docSettings.pageNumbers}
                    onChange={(e) => patchSettings({ pageNumbers: e.target.checked })}
                  />
                  ページ番号を付ける（下部中央）
                </label>
                <label className="ds-check">
                  <input
                    type="checkbox"
                    checked={!!docSettings.toc}
                    onChange={(e) => patchSettings({ toc: e.target.checked })}
                  />
                  先頭に目次を挿入
                </label>
                <p className="ds-note">
                  ヘッダーと本文フォントはプレビュー・PDF にも反映されます。用紙/余白/ページ番号/目次は
                  Word 書き出し用です（目次は Word で開いて F9 で更新）。
                </p>
              </div>
            )}
          </div>
          <button className="reset" onClick={resetToDefault} title="内容を初期状態に戻す">
            🔄 初期化
          </button>
          <label className="filename">
            <span className="fn-label">ファイル名</span>
            <input
              ref={fnInputRef}
              type="text"
              value={fileName}
              onChange={(e) => setFileName(e.target.value)}
              placeholder="document"
              spellCheck={false}
              aria-label="ファイル名"
            />
            <span ref={fnSizerRef} className="fn-sizer" aria-hidden>
              {fileName || 'document'}
            </span>
          </label>
          <div className="export-wrap" ref={exportWrapRef}>
            <button
              className="export"
              onClick={() => setExportMenuOpen((o) => !o)}
              disabled={exporting}
              aria-haspopup="menu"
              aria-expanded={exportMenuOpen}
            >
              {exporting ? '書き出し中…' : '書き出す ▾'}
            </button>
            {exportMenuOpen && !exporting && (
              <div className="export-menu" role="menu">
                <button role="menuitem" onClick={handleExportDocx}>
                  <span className="mi-title">Word（.docx）</span>
                  <span className="mi-desc">編集できる Word 文書</span>
                </button>
                <button role="menuitem" onClick={handlePrintPdf}>
                  <span className="mi-title">PDF（印刷から保存）</span>
                  <span className="mi-desc">高画質・文字を選択できる PDF</span>
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      <div className="banner info">
        左の「Markdown」で下書きし <b>「反映」</b> で文書化、中央のプレビューを <b>直接編集</b>（太字・見出し・表）。<b>＋テキストボックス</b> で自由配置のボックスも追加でき（ドラッグ移動・ダブルクリックで編集）、右上から <b>Word（.docx）/ PDF</b> に書き出せます。
      </div>

      {error && (
        <div className="banner error" role="alert">
          ⚠️ {error}
        </div>
      )}

      <div className={`workspace${resizing ? ' resizing' : ''}`} ref={workspaceRef}>
        <aside className={`md-drawer${mdOpen ? ' open' : ''}`} style={{ width: mdOpen ? drawerWidth : 0 }}>
          <div
            className="md-inner"
            style={{ width: drawerWidth }}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault()
              const files = Array.from(e.dataTransfer.files ?? [])
              const mdFile = files.find((f) => /\.(md|markdown|mdown|txt)$/i.test(f.name))
              if (mdFile) importFile(mdFile, 'replace')
              addImageFiles(files)
            }}
          >
            <div className="pane-head">
              <span>Markdown</span>
              <span className="loadmd-group">
                <button
                  className="loadmd"
                  onClick={() => {
                    importModeRef.current = 'replace'
                    fileInputRef.current?.click()
                  }}
                  data-tip="Markdown を読み込み（現在の内容を置き換え・対応: .md / .markdown / .txt）"
                  aria-label="Markdown ファイルをインポート（現在の内容を置き換え）"
                >
                  📂 インポート
                </button>
                <button
                  className="loadmd"
                  onClick={() => {
                    importModeRef.current = 'append'
                    fileInputRef.current?.click()
                  }}
                  data-tip="別の Markdown を現在の内容に結合して読み込み"
                  aria-label="別の Markdown を現在の内容に結合して読み込み"
                >
                  ＋ 結合
                </button>
                <button
                  className="loadmd"
                  onClick={() => imageInputRef.current?.click()}
                  data-tip="画像ファイルを読み込み、Markdown 内の相対パス（例: ![](fig1.png)）に紐づけます"
                  aria-label="画像ファイルを読み込む"
                >
                  🖼 画像
                </button>
              </span>
              <input
                ref={fileInputRef}
                type="file"
                accept=".md,.markdown,.mdown,.txt,text/markdown,text/plain"
                hidden
                onChange={(e) => {
                  const f = e.target.files?.[0]
                  if (f) importFile(f, importModeRef.current)
                  e.target.value = ''
                }}
              />
              <input
                ref={imageInputRef}
                type="file"
                accept="image/*"
                multiple
                hidden
                onChange={(e) => {
                  addImageFiles(Array.from(e.target.files ?? []))
                  e.target.value = ''
                }}
              />
            </div>
            {mdFileName && (
              <div className="md-source" title={`読み込み中: ${mdFileName}`}>
                <span className="md-source-icon" aria-hidden>
                  📄
                </span>
                <span className="md-source-name">{mdFileName}</span>
                <button className="md-source-clear" onClick={() => setMdFileName('')} title="ファイル名表示を消す">
                  ✕
                </button>
              </div>
            )}
            {imageNames.length > 0 && (
              <div className="attached">
                <span className="attached-label">画像 {imageNames.length} 枚:</span>
                <span className="attached-names" title={imageNames.join(', ')}>
                  {imageNames.join(', ')}
                </span>
                <button className="attached-clear" onClick={() => setImages({})} title="添付画像をすべて外す">
                  クリア
                </button>
              </div>
            )}
            <textarea
              className="editor"
              value={markdown}
              onChange={(e) => setMarkdown(e.target.value)}
              spellCheck={false}
            />
            <div className="md-foot">
              <button
                className="apply"
                onClick={() => void rebuildFromMarkdown()}
                title="現在の Markdown から文書を作り直す（編集内容は上書き）"
              >
                プレビューに反映{docDirty ? ' ●' : ''}
              </button>
            </div>
          </div>
        </aside>

        <button
          className="md-handle"
          onPointerDown={onHandlePointerDown}
          onKeyDown={onHandleKeyDown}
          aria-expanded={mdOpen}
          title={mdOpen ? 'ドラッグで幅を調整 / クリックで閉じる' : 'クリックで開く（ドラッグで幅調整）'}
        >
          <span className="md-handle-text">Markdown</span>
          <span className="md-handle-arrow" aria-hidden>
            {mdOpen ? '◀' : '▶'}
          </span>
        </button>

        <main
          className="doc-main"
          style={docSettings.font ? ({ '--doc-font': DOC_FONT_CSS[docSettings.font] } as React.CSSProperties) : undefined}
        >
          <DocEditor
            key={rebuildToken}
            html={docHtml}
            images={images}
            onChange={handleDocChange}
            boxes={boxes}
            onBoxesChange={setBoxes}
            onRegenerate={() => void rebuildFromMarkdown()}
            headerText={docSettings.headerText}
          />
        </main>
      </div>

      {helpOpen && (
        <div className="help-overlay" onClick={() => setHelpOpen(false)}>
          <div className="help-modal" role="dialog" aria-modal="true" aria-label="使い方" onClick={(e) => e.stopPropagation()}>
            <div className="help-head">
              <h2>使い方</h2>
              <button className="help-close" onClick={() => setHelpOpen(false)} aria-label="閉じる">
                ×
              </button>
            </div>
            <div className="help-body">
              <p className="help-lead">
                Docdown は、Markdown で書いた文章を <b>編集できる Word（.docx）/ PDF</b> に変換するツールです。
              </p>

              <section>
                <h3>1. Markdown を用意する</h3>
                <p>
                  左の「Markdown」パネルに直接書くか、<b>📂 インポート</b> で .md ファイルを読み込みます。
                  AI に書いてもらう場合は、次のプロンプトが便利です。
                </p>
                <div className="help-prompt">
                  <div className="help-prompt-head">
                    <span>AI 用プロンプト</span>
                    <button className="help-copy" onClick={() => void copyPrompt()}>
                      {promptCopied ? '✓ コピーしました' : 'コピー'}
                    </button>
                  </div>
                  <pre>{AI_PROMPT}</pre>
                </div>
                <p className="help-sub">
                  相対パスの画像（例: <code>![](fig1.png)</code>）は <b>🖼 画像</b> からファイルを読み込むと表示されます。
                  数式は <code>$…$</code> / <code>$$…$$</code>（LaTeX）に対応しています。
                </p>
              </section>

              <section>
                <h3>2. 文書に反映する</h3>
                <p>
                  「<b>プレビューに反映</b>」を押すと、Markdown から文書を作り直します。
                  文書を直接編集した後に押すと上書き確認が出ます（<b>●</b> は未反映の編集がある印です）。
                </p>
              </section>

              <section>
                <h3>3. 文書を直接編集する</h3>
                <ul>
                  <li>本文はプレビューをクリックしてそのまま編集（太字・下線・色・見出し・リスト・表など）</li>
                  <li><b>＋テキストボックス</b> で自由配置のボックスを追加（ドラッグ移動・ダブルクリックで編集・行間調整）</li>
                  <li>表の中にカーソルを置くと <b>＋行/−行/＋列/−列</b> が使えます</li>
                  <li>Ctrl/⌘+Z で元に戻す（テキストボックスの操作にも効きます）</li>
                </ul>
              </section>

              <section>
                <h3>4. 書き出す</h3>
                <ul>
                  <li><b>Word（.docx）</b> … 見出し・表・画像・テキストボックスまで、Word でそのまま編集できる形式</li>
                  <li><b>PDF（印刷から保存）</b> … 印刷ダイアログで「PDF に保存」を選ぶと、文字を選択できる高画質 PDF</li>
                </ul>
              </section>

              <section className="help-why">
                <h3>💡 なぜ Markdown から文書を作るの?</h3>
                <ul>
                  <li>
                    <b>AI と相性が良い</b> ― AI に Word ファイルを直接作らせるより、Markdown を書かせる方が速くて安定します。
                  </li>
                  <li>
                    <b>微調整は自分の手で</b> ― 文言の直しやレイアウト調整を AI に投げ直す必要がありません。
                  </li>
                  <li>
                    <b>本物の編集できる成果物</b> ― 画像貼り付けではなく、Word で開いて編集できる文書が手に入ります。
                  </li>
                  <li>
                    <b>ブラウザだけで完結・無料</b> ― サーバーへ送信せず手元で処理するので、内容が外部に上がりません。
                  </li>
                </ul>
              </section>

              <p className="help-note">
                編集内容は自動保存され、リロードしても続きから再開できます。最初の状態に戻すには <b>「🔄 初期化」</b>。
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
