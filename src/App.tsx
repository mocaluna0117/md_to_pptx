import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { exportDeckToPptx } from './lib/exportDeck'
import { exportDeckToPdf } from './lib/exportPdf'
import { type Deck } from './lib/deck'
import { deckFromRenderedMarkdown } from './lib/deckFromRender'
import VisualEditor from './components/VisualEditor'
import './App.css'

type ExportTarget = 'pptx' | 'pdf'

const STORAGE_KEY = 'md-to-pptx:v1'

interface Persisted {
  markdown?: string
  fileName?: string
  deck?: Deck | null
  deckDirty?: boolean
  mdOpen?: boolean
  drawerWidth?: number
}

const MIN_DRAWER = 240
const MAX_DRAWER = 760

function loadPersisted(): Persisted {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') as Persisted
  } catch {
    return {}
  }
}

const persisted = loadPersisted()

const SAMPLE = `---
marp: true
theme: default
paginate: true
---

# Markdown → PowerPoint

Marp で書いた Markdown を
そのまま .pptx に変換

---

## 使い方

1. 左の「Markdown」から Markdown を書く / インポート
2. \`---\` でスライドを区切る
3. プレビュー上で直接レイアウトを調整
4. 右上の **書き出す** から PPTX / PDF

---

## こんなことができます

- **太字** や *斜体*、\`コード\`
- 箇条書き / 番号付きリスト
- テーマやページ番号の指定

---

# ありがとう 🎉
`

type Status =
  | { kind: 'idle' }
  | { kind: 'exporting'; done: number; total: number }
  | { kind: 'error'; message: string }

function App() {
  const [markdown, setMarkdown] = useState(persisted.markdown ?? SAMPLE)
  const [fileName, setFileName] = useState(persisted.fileName ?? 'slides')
  const [deck, setDeck] = useState<Deck | null>(persisted.deck ?? null)
  // True when the deck has edits not derived from the current Markdown.
  const [deckDirty, setDeckDirty] = useState<boolean>(persisted.deckDirty ?? false)
  const [mdOpen, setMdOpen] = useState<boolean>(persisted.mdOpen ?? true)
  const [drawerWidth, setDrawerWidth] = useState<number>(
    Math.min(MAX_DRAWER, Math.max(MIN_DRAWER, persisted.drawerWidth ?? 380)),
  )
  const [resizing, setResizing] = useState(false)
  const [status, setStatus] = useState<Status>({ kind: 'idle' })
  const [exportMenuOpen, setExportMenuOpen] = useState(false)
  const exportWrapRef = useRef<HTMLDivElement>(null)
  const workspaceRef = useRef<HTMLDivElement>(null)

  // Close the export menu when clicking outside it.
  useEffect(() => {
    if (!exportMenuOpen) return
    const onDown = (e: PointerEvent) => {
      if (!exportWrapRef.current?.contains(e.target as Node)) setExportMenuOpen(false)
    }
    document.addEventListener('pointerdown', onDown)
    return () => document.removeEventListener('pointerdown', onDown)
  }, [exportMenuOpen])

  const exporting = status.kind === 'exporting'

  // Persist Markdown + deck so a reload/close doesn't lose work.
  useEffect(() => {
    const id = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, deck, deckDirty, mdOpen, drawerWidth }))
      } catch {
        // Deck too big for storage (e.g. embedded images): keep at least the Markdown.
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, deckDirty, mdOpen, drawerWidth }))
        } catch {
          /* storage unavailable */
        }
      }
    }, 300)
    return () => clearTimeout(id)
  }, [markdown, fileName, deck, deckDirty, mdOpen, drawerWidth])

  /** Drag the Markdown drawer's edge to resize; releasing without a drag toggles it. */
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
      // A drag resized the drawer; a press without movement is a toggle.
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

  // ---- Undo / redo history for the deck ----
  const deckRef = useRef(deck)
  deckRef.current = deck
  const undoRef = useRef<Deck[]>([])
  const redoRef = useRef<Deck[]>([])
  // Consecutive changes sharing the same non-null key (e.g. one drag) are one undo step.
  const lastKeyRef = useRef<number | null>(null)
  const [, setHistoryTick] = useState(0)
  const bumpHistory = () => setHistoryTick((v) => v + 1)

  function clearHistory() {
    undoRef.current = []
    redoRef.current = []
    lastKeyRef.current = null
    bumpHistory()
  }

  const handleDeckChange = useCallback((next: Deck, coalesceKey?: number) => {
    const prev = deckRef.current
    if (prev) {
      const sameBurst = coalesceKey != null && coalesceKey === lastKeyRef.current
      if (!sameBurst) {
        undoRef.current.push(prev)
        if (undoRef.current.length > 100) undoRef.current.shift()
        redoRef.current = []
      }
      lastKeyRef.current = coalesceKey ?? null
    }
    deckRef.current = next
    setDeck(next)
    setDeckDirty(true)
    bumpHistory()
  }, [])

  const undo = useCallback(() => {
    if (undoRef.current.length === 0) return
    lastKeyRef.current = null
    const prev = undoRef.current.pop() as Deck
    if (deckRef.current) redoRef.current.push(deckRef.current)
    deckRef.current = prev
    setDeck(prev)
    setDeckDirty(true)
    bumpHistory()
  }, [])

  const redo = useCallback(() => {
    if (redoRef.current.length === 0) return
    lastKeyRef.current = null
    const next = redoRef.current.pop() as Deck
    if (deckRef.current) undoRef.current.push(deckRef.current)
    deckRef.current = next
    setDeck(next)
    setDeckDirty(true)
    bumpHistory()
  }, [])

  // The Markdown the current deck was built from.
  const deckSourceRef = useRef<string | null>(null)

  /** Build the deck from Markdown (rendered via Marp) and reset history. */
  const buildDeck = useCallback(async (src: string) => {
    const d = await deckFromRenderedMarkdown(src)
    deckRef.current = d
    setDeck(d)
    deckSourceRef.current = src
    setDeckDirty(false)
    clearHistory()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // On first load, build a deck from the Markdown if we don't have one yet.
  const bootedRef = useRef(false)
  useEffect(() => {
    if (bootedRef.current) return
    bootedRef.current = true
    if (!deckRef.current) void buildDeck(markdown)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  /** Rebuild the deck from the current Markdown, confirming if edits would be lost. */
  async function rebuildFromMarkdown(): Promise<boolean> {
    if (deckRef.current && deckDirty && !window.confirm('現在の Markdown からスライドを作り直します。編集した内容は上書きされます。よろしいですか？')) {
      return false
    }
    await buildDeck(markdown)
    return true
  }

  const fileInputRef = useRef<HTMLInputElement>(null)

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

  async function importMarkdownFile(file: File) {
    if (deckDirty && !window.confirm('ファイルを読み込むと、現在の編集内容は破棄されます。よろしいですか？')) {
      return
    }
    try {
      const text = await file.text()
      setMarkdown(text)
      const base = file.name.replace(/\.[^.]+$/, '')
      if (base) setFileName(base)
      // Reflect the imported Markdown in the preview immediately.
      await buildDeck(text)
    } catch {
      setStatus({ kind: 'error', message: 'ファイルの読み込みに失敗しました。' })
    }
  }

  /** Reset everything to the default sample (the state before importing a file). */
  async function resetToDefault() {
    if (!window.confirm('すべて初期状態に戻します。現在の編集内容は破棄されます。よろしいですか？')) {
      return
    }
    setMarkdown(SAMPLE)
    setFileName('slides')
    setStatus({ kind: 'idle' })
    try {
      localStorage.removeItem(STORAGE_KEY)
    } catch {
      /* ignore */
    }
    await buildDeck(SAMPLE)
  }

  async function runExport(target: ExportTarget) {
    setExportMenuOpen(false)
    if (!deckRef.current) {
      setStatus({ kind: 'error', message: 'スライドがありません。' })
      return
    }
    setStatus({ kind: 'exporting', done: 0, total: 0 })
    const onProgress = (done: number, total: number) => setStatus({ kind: 'exporting', done, total })
    try {
      if (target === 'pptx') {
        await exportDeckToPptx(deckRef.current, { fileName, onProgress })
      } else {
        await exportDeckToPdf(deckRef.current, { fileName, onProgress })
      }
      setStatus({ kind: 'idle' })
    } catch (err) {
      setStatus({ kind: 'error', message: err instanceof Error ? err.message : String(err) })
    }
  }

  return (
    <div className="app">
      <header className="toolbar">
        <div className="brand">
          <h1>Deckdown</h1>
          <span className="tagline">Markdown → PowerPoint</span>
        </div>
        <div className="actions">
          <button
            className="reset"
            onClick={resetToDefault}
            title="すべて初期状態（デフォルトのサンプル）に戻す"
          >
            🔄 初期化
          </button>
          <label className="filename">
            <span className="fn-label">ファイル名</span>
            <input
              ref={fnInputRef}
              type="text"
              value={fileName}
              onChange={(e) => setFileName(e.target.value)}
              placeholder="slides"
              spellCheck={false}
              aria-label="ファイル名"
            />
            <span ref={fnSizerRef} className="fn-sizer" aria-hidden>
              {fileName || 'slides'}
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
              {exporting
                ? status.total > 0
                  ? `書き出し中 ${status.done}/${status.total}`
                  : '準備中…'
                : '書き出す ▾'}
            </button>
            {exportMenuOpen && !exporting && (
              <div className="export-menu" role="menu">
                <button role="menuitem" onClick={() => runExport('pptx')}>
                  <span className="mi-title">PPTX（編集可能）</span>
                  <span className="mi-desc">
                    各テキストボックスを、PowerPoint で編集できる状態のまま書き出します。
                  </span>
                </button>
                <button role="menuitem" onClick={() => runExport('pdf')}>
                  <span className="mi-title">PDF</span>
                  <span className="mi-desc">
                    各スライドを画像にして、1 ページずつ収めた PDF を書き出します。
                  </span>
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      <div className="banner info">
        スライドを PowerPoint のように直接編集できます（ドラッグで移動・角でサイズ変更・ダブルクリックで文字編集）。左の「Markdown」タブで元の Markdown を編集・インポート。書き出しは編集可能な PPTX と PDF。
      </div>

      {status.kind === 'error' && (
        <div className="banner error" role="alert">
          ⚠️ {status.message}
        </div>
      )}

      <div className={`workspace${resizing ? ' resizing' : ''}`} ref={workspaceRef}>
        <aside
          className={`md-drawer${mdOpen ? ' open' : ''}`}
          style={{ width: mdOpen ? drawerWidth : 0 }}
        >
          <div
            className="md-inner"
            style={{ width: drawerWidth }}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault()
              const f = e.dataTransfer.files?.[0]
              if (f && /\.(md|markdown|mdown|txt)$/i.test(f.name)) importMarkdownFile(f)
            }}
          >
            <div className="pane-head">
              <span>Markdown</span>
              <button
                className="loadmd"
                onClick={() => fileInputRef.current?.click()}
                data-tip="対応: .md / .markdown / .txt"
                aria-label="Markdown ファイルをインポート（対応: .md / .markdown / .txt）"
              >
                📂 ファイルをインポート
              </button>
              <input
                ref={fileInputRef}
                type="file"
                accept=".md,.markdown,.mdown,.txt,text/markdown,text/plain"
                hidden
                onChange={(e) => {
                  const f = e.target.files?.[0]
                  if (f) importMarkdownFile(f)
                  e.target.value = ''
                }}
              />
            </div>
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
                title="現在の Markdown からスライドを作り直します（編集した内容がある場合は上書き確認）"
              >
                プレビューに反映{deckDirty ? ' ●' : ''}
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

        <main className="stage-area">
          {deck ? (
            <VisualEditor
              deck={deck}
              onChange={handleDeckChange}
              onRegenerate={rebuildFromMarkdown}
              onUndo={undo}
              onRedo={redo}
              canUndo={undoRef.current.length > 0}
              canRedo={redoRef.current.length > 0}
            />
          ) : (
            <div className="stage-loading">スライドを生成中…</div>
          )}
        </main>
      </div>
    </div>
  )
}

export default App
