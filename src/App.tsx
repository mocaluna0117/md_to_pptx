import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { renderPreview } from './lib/marp'
import { exportPptx } from './lib/exportPptx'
import { exportPptxNative } from './lib/exportPptxNative'
import { exportDeckToPptx } from './lib/exportDeck'
import { exportMarkdownToPdf, exportDeckToPdf } from './lib/exportPdf'
import { type Deck } from './lib/deck'
import { deckFromRenderedMarkdown } from './lib/deckFromRender'
import VisualEditor from './components/VisualEditor'
import './App.css'

type Mode = 'image' | 'native'
type View = 'markdown' | 'visual'
type Format = 'pptx' | 'pdf'

const STORAGE_KEY = 'md-to-pptx:v1'

interface Persisted {
  markdown?: string
  fileName?: string
  deck?: Deck | null
  deckDirty?: boolean
}

function loadPersisted(): Persisted {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') as Persisted
  } catch {
    return {}
  }
}

const persisted = loadPersisted()

const MODE_LABEL: Record<Mode, string> = {
  image: '画像（見た目そのまま）',
  native: '編集可能（テキスト）',
}

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

1. 左のエディタに Markdown を書く
2. \`---\` でスライドを区切る
3. 右上の **PPTX を書き出す** を押す

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
  const [mode, setMode] = useState<Mode>('image')
  const [view, setView] = useState<View>('markdown')
  const [deck, setDeck] = useState<Deck | null>(persisted.deck ?? null)
  // True when the visual deck has edits not derived from the current Markdown.
  const [deckDirty, setDeckDirty] = useState<boolean>(persisted.deckDirty ?? false)
  const [status, setStatus] = useState<Status>({ kind: 'idle' })
  const [exportMenuOpen, setExportMenuOpen] = useState(false)
  const exportWrapRef = useRef<HTMLDivElement>(null)

  // Close the export menu when clicking outside it.
  useEffect(() => {
    if (!exportMenuOpen) return
    const onDown = (e: PointerEvent) => {
      if (!exportWrapRef.current?.contains(e.target as Node)) setExportMenuOpen(false)
    }
    document.addEventListener('pointerdown', onDown)
    return () => document.removeEventListener('pointerdown', onDown)
  }, [exportMenuOpen])

  const preview = useMemo(() => renderPreview(markdown), [markdown])
  const exporting = status.kind === 'exporting'

  // Persist Markdown + visual deck so a reload/close doesn't lose work.
  useEffect(() => {
    const id = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, deck, deckDirty }))
      } catch {
        // Deck too big for storage (e.g. embedded images): keep at least the Markdown.
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, deckDirty }))
        } catch {
          /* storage unavailable */
        }
      }
    }, 300)
    return () => clearTimeout(id)
  }, [markdown, fileName, deck, deckDirty])

  // ---- Undo / redo history for the visual deck ----
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

  // The Markdown the current deck was built from (to auto-refresh when it changes).
  const deckSourceRef = useRef<string | null>(null)

  async function buildDeck() {
    const d = await deckFromRenderedMarkdown(markdown)
    deckRef.current = d
    setDeck(d)
    deckSourceRef.current = markdown
    setDeckDirty(false)
    clearHistory()
  }

  async function enterVisual() {
    // Rebuild from the current Markdown unless there are unsaved visual edits.
    if (!deck || (markdown !== deckSourceRef.current && !deckDirty)) {
      await buildDeck()
    }
    setView('visual')
  }

  /** Rebuild the deck from the current Markdown, confirming if edits would be lost. */
  async function rebuildFromMarkdown(): Promise<boolean> {
    if (deck && deckDirty && !window.confirm('現在の Markdown からスライドを作り直します。ビジュアル編集の変更は上書きされます。よろしいですか？')) {
      return false
    }
    await buildDeck()
    return true
  }

  async function applyMarkdownToVisual() {
    if (await rebuildFromMarkdown()) setView('visual')
  }

  const fileInputRef = useRef<HTMLInputElement>(null)

  async function importMarkdownFile(file: File) {
    if (deckDirty && !window.confirm('ファイルを読み込むと、現在のビジュアル編集の変更は破棄されます。よろしいですか？')) {
      return
    }
    try {
      const text = await file.text()
      setMarkdown(text)
      const base = file.name.replace(/\.[^.]+$/, '')
      if (base) setFileName(base)
      // New content: drop the old deck so entering the visual tab rebuilds fresh.
      deckRef.current = null
      setDeck(null)
      setDeckDirty(false)
      deckSourceRef.current = null
      clearHistory()
      setView('markdown')
    } catch {
      setStatus({ kind: 'error', message: 'ファイルの読み込みに失敗しました。' })
    }
  }

  /** Reset everything to the default sample (the state before importing a file). */
  function resetToDefault() {
    if (!window.confirm('すべて初期状態に戻します。現在の Markdown とビジュアル編集は破棄されます。よろしいですか？')) {
      return
    }
    setMarkdown(SAMPLE)
    setFileName('slides')
    setMode('image')
    deckRef.current = null
    setDeck(null)
    setDeckDirty(false)
    deckSourceRef.current = null
    clearHistory()
    setStatus({ kind: 'idle' })
    setView('markdown')
    try {
      localStorage.removeItem(STORAGE_KEY)
    } catch {
      /* ignore */
    }
  }

  async function handleExport(format: Format) {
    setExportMenuOpen(false)
    setStatus({ kind: 'exporting', done: 0, total: 0 })
    const onProgress = (done: number, total: number) => setStatus({ kind: 'exporting', done, total })
    try {
      if (format === 'pdf') {
        if (view === 'visual') {
          if (!deck) throw new Error('スライドがありません。')
          await exportDeckToPdf(deck, { fileName, onProgress })
        } else {
          await exportMarkdownToPdf(markdown, { fileName, onProgress })
        }
      } else if (view === 'visual') {
        if (!deck) throw new Error('スライドがありません。')
        await exportDeckToPptx(deck, { fileName, onProgress })
      } else {
        const run = mode === 'image' ? exportPptx : exportPptxNative
        await run(markdown, { fileName, onProgress })
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
          <span className="logo" aria-hidden>
            🖥️
          </span>
          <h1>Marp → PPTX</h1>
          <div className="view" role="group" aria-label="表示">
            <button
              type="button"
              className={view === 'markdown' ? 'active' : ''}
              onClick={() => setView('markdown')}
            >
              Markdown
            </button>
            <button type="button" className={view === 'visual' ? 'active' : ''} onClick={enterVisual}>
              ビジュアル編集
            </button>
          </div>
        </div>
        <div className="actions">
          {view === 'markdown' && (
            <div className="mode" role="group" aria-label="変換方式">
              {(['image', 'native'] as Mode[]).map((m) => (
                <button
                  key={m}
                  type="button"
                  className={mode === m ? 'active' : ''}
                  onClick={() => setMode(m)}
                  disabled={exporting}
                  title={MODE_LABEL[m]}
                >
                  {MODE_LABEL[m]}
                </button>
              ))}
            </div>
          )}
          {view === 'markdown' && (
            <button
              className="apply"
              onClick={applyMarkdownToVisual}
              title="現在の Markdown からビジュアル編集を作成／更新します（ビジュアルの編集がある場合は上書き確認）"
            >
              ビジュアルに反映{deckDirty ? ' ●' : ''}
            </button>
          )}
          <label className="filename">
            <input
              type="text"
              value={fileName}
              onChange={(e) => setFileName(e.target.value)}
              spellCheck={false}
              aria-label="ファイル名"
            />
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
                <button role="menuitem" onClick={() => handleExport('pptx')}>
                  PPTX で書き出す
                </button>
                <button role="menuitem" onClick={() => handleExport('pdf')}>
                  PDF で書き出す
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      <div className="banner info">
        {view === 'visual'
          ? 'ビジュアル編集：ボックスをドラッグで移動、ダブルクリックで文字編集、選択して色変更。書き出しは編集可能な PPTX です。'
          : mode === 'image'
            ? '画像方式：Marp の見た目をそのまま再現。PowerPoint 上では画像になり、テキスト編集はできません。'
            : '編集可能方式：見出し・本文・箇条書き・コード・引用・表を編集可能なテキストに変換。テーマの再現は簡易です。'}
      </div>

      {view === 'markdown' && deckDirty && (
        <div className="banner warn">
          ● ビジュアル編集に変更があります。「ビジュアルに反映」を押すと Markdown の内容で上書きされます（変更は自動保存され、リロードしても保持されます）。
        </div>
      )}

      {status.kind === 'error' && (
        <div className="banner error" role="alert">
          ⚠️ {status.message}
        </div>
      )}

      {view === 'visual' && deck ? (
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
        <main className="panes">
          <section
            className="pane editor-pane"
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault()
              const f = e.dataTransfer.files?.[0]
              if (f && /\.(md|markdown|mdown|txt)$/i.test(f.name)) importMarkdownFile(f)
            }}
          >
            <div className="pane-head">
              <span>Markdown</span>
              <span className="loadmd-group">
                <span className="ext-hint">対応: .md / .markdown / .txt</span>
                <button
                  className="loadmd"
                  onClick={resetToDefault}
                  title="すべて初期状態（デフォルトのサンプル）に戻す"
                >
                  🔄 初期化
                </button>
                <button
                  className="loadmd"
                  onClick={() => fileInputRef.current?.click()}
                  title="Markdown ファイルをインポート（.md / .markdown / .mdown / .txt）"
                >
                  📂 ファイルをインポート
                </button>
              </span>
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
          </section>

        <section className="pane preview-pane">
          <div className="pane-head">プレビュー</div>
          <div className="preview-scroll">
            <style>{preview.css}</style>
            <div className="preview" dangerouslySetInnerHTML={{ __html: preview.html }} />
          </div>
        </section>
        </main>
      )}
    </div>
  )
}

export default App
