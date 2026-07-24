import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import MarkdownIt from 'markdown-it'
import markdownItCjkFriendly from 'markdown-it-cjk-friendly'
import { navigate } from './Root'
import { exportHtmlToDocx } from './lib/exportDocx'
import { resolveImagePaths, readImageFiles, IMAGE_EXT, type AttachedImages } from './lib/imageAttach'
import { mathToImages } from './lib/math'
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
    const prepared = await mathToImages(resolveImagePaths(src, imagesRef.current))
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
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, mdFileName, mdOpen, drawerWidth, images, docHtml, docDirty, boxes }))
      } catch {
        // Document HTML / images / boxes may exceed the storage quota: keep at least the text.
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, mdFileName, mdOpen, drawerWidth }))
        } catch {
          /* storage unavailable */
        }
      }
    }, 300)
    return () => clearTimeout(id)
  }, [markdown, fileName, mdFileName, mdOpen, drawerWidth, images, docHtml, docDirty, boxes])

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
      await exportHtmlToDocx(docHtmlRef.current, boxesRef.current, { fileName })
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
        <button className="home-btn" onClick={() => navigate('home')} title="ホームに戻る" aria-label="ホームに戻る">
          ⌂
        </button>
        <div className="brand">
          <h1>Docdown</h1>
          <span className="tagline">Markdown → Word</span>
        </div>
        <div className="actions">
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
                  title="Markdown を読み込み（現在の内容を置き換え）"
                >
                  📂 インポート
                </button>
                <button
                  className="loadmd"
                  onClick={() => {
                    importModeRef.current = 'append'
                    fileInputRef.current?.click()
                  }}
                  title="別の Markdown を現在の内容に結合"
                >
                  ＋ 結合
                </button>
                <button
                  className="loadmd"
                  onClick={() => imageInputRef.current?.click()}
                  title="画像ファイルを読み込み、Markdown 内の相対パス（例: ![](fig1.png)）に紐づけます"
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

        <main className="doc-main">
          <DocEditor
            key={rebuildToken}
            html={docHtml}
            images={images}
            onChange={handleDocChange}
            boxes={boxes}
            onBoxesChange={setBoxes}
          />
        </main>
      </div>
    </div>
  )
}
