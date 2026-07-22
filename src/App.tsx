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

/** Remove a leading YAML front-matter block (--- … ---) from Markdown, if present. */
function stripFrontmatter(md: string): string {
  return md.replace(/^﻿?---[^\n]*\n[\s\S]*?\n---[^\n]*\n?/, '')
}

/** Merge an imported Markdown into the current one as extra slides (own front-matter dropped). */
function mergeMarkdown(base: string, add: string): string {
  if (!base.trim()) return add
  const body = stripFrontmatter(add).trim()
  if (!body) return base
  return `${base.trimEnd()}\n\n---\n\n${body}\n`
}

/** Copy text to the clipboard, falling back to execCommand when the async API is unavailable. */
async function copyText(text: string): Promise<boolean> {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text)
      return true
    }
  } catch {
    /* fall through to the legacy path */
  }
  try {
    const ta = document.createElement('textarea')
    ta.value = text
    ta.style.position = 'fixed'
    ta.style.opacity = '0'
    document.body.appendChild(ta)
    ta.focus()
    ta.select()
    const ok = document.execCommand('copy')
    document.body.removeChild(ta)
    return ok
  } catch {
    return false
  }
}

/** A prompt users can hand to an AI (ChatGPT etc.) to generate slide-ready Markdown. */
const AI_PROMPT = `次の内容をプレゼン用スライドにまとめて、Marp 記法の Markdown（.md）で出力してください。

# ルール
- 先頭にフロントマターを付ける：
---
marp: true
paginate: true
---
- スライドは「---」で区切る
- 各スライドは「#」（表紙）または「##」（見出し）で始める
- 箇条書き・表・コードブロックを活用する
- 装飾は Markdown のみ（HTML タグは使わない）

# まとめたい内容
（ここに伝えたい内容や、添付画像の説明を書いてください）`

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
  const [helpOpen, setHelpOpen] = useState(false)
  const [promptCopied, setPromptCopied] = useState(false)

  async function copyPrompt() {
    if (await copyText(AI_PROMPT)) {
      setPromptCopied(true)
      setTimeout(() => setPromptCopied(false), 1500)
    }
  }
  const exportWrapRef = useRef<HTMLDivElement>(null)
  const workspaceRef = useRef<HTMLDivElement>(null)

  // Close the "使い方" dialog on Escape.
  useEffect(() => {
    if (!helpOpen) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setHelpOpen(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [helpOpen])

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
  // Whether the next file pick replaces the Markdown or is merged into it.
  const importModeRef = useRef<'replace' | 'append'>('replace')

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

  async function importMarkdownFile(file: File, mode: 'replace' | 'append' = 'replace') {
    const confirmMsg =
      mode === 'append'
        ? '読み込んだ Markdown を現在の内容に結合します。ビジュアル編集の変更は破棄され、作り直されます。よろしいですか？'
        : 'ファイルを読み込むと、現在の編集内容は破棄されます。よろしいですか？'
    if (deckDirty && !window.confirm(confirmMsg)) {
      return
    }
    try {
      const text = await file.text()
      if (mode === 'append') {
        const merged = mergeMarkdown(markdown, text)
        setMarkdown(merged)
        await buildDeck(merged)
      } else {
        setMarkdown(text)
        const base = file.name.replace(/\.[^.]+$/, '')
        if (base) setFileName(base)
        await buildDeck(text)
      }
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
        <button className="help-btn" onClick={() => setHelpOpen(true)} aria-haspopup="dialog">
          ？ 使い方
        </button>
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
              </span>
              <input
                ref={fileInputRef}
                type="file"
                accept=".md,.markdown,.mdown,.txt,text/markdown,text/plain"
                hidden
                onChange={(e) => {
                  const f = e.target.files?.[0]
                  if (f) importMarkdownFile(f, importModeRef.current)
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

      {helpOpen && (
        <div className="help-overlay" onClick={() => setHelpOpen(false)}>
          <div
            className="help-modal"
            role="dialog"
            aria-modal="true"
            aria-label="使い方"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="help-head">
              <h2>使い方</h2>
              <button className="help-close" onClick={() => setHelpOpen(false)} aria-label="閉じる">
                ×
              </button>
            </div>
            <div className="help-body">
              <p className="help-lead">
                Markdown を貼り付けて、PowerPoint のように直接編集し、編集できる PPTX / PDF に書き出せます。
              </p>

              <section>
                <h3>1. Markdown を用意する</h3>
                <p className="help-sub">
                  基本は AI に作ってもらうのが簡単です。下のプロンプトを ChatGPT などに渡してください。
                </p>
                <div className="help-prompt">
                  <div className="help-prompt-head">
                    <span>AI へのプロンプト例</span>
                    <button className="help-copy" onClick={copyPrompt}>
                      {promptCopied ? '✓ コピーしました' : 'コピー'}
                    </button>
                  </div>
                  <pre>{AI_PROMPT}</pre>
                </div>
                <p className="help-sub">
                  かんたんに <b>「添付画像をスライドにまとめたいから、Marp 形式の .md を作って」</b> のように頼んでもOK。
                </p>
                <ul>
                  <li>できた Markdown を <b>「📂 インポート」</b>／ドロワーに貼り付け／ドラッグ＆ドロップ。</li>
                  <li><b>「＋ 結合」</b>で 2 つめ以降の Markdown を現在の内容に連結。</li>
                  <li>
                    もちろん <b>自分で 1 から書く</b>こともできます（左端の縦タブ「Markdown」を開いて直接入力・
                    境界のドラッグで幅調整）。
                  </li>
                </ul>
              </section>

              <section>
                <h3>2. スライドに反映する</h3>
                <ul>
                  <li>ドロワー下の <b>「プレビューに反映」</b>で、Markdown から中央のスライドを作成／作り直し。</li>
                  <li>見出しや <code>---</code>、箇条書き・表・コードブロックがそのままスライドになります。</li>
                </ul>
              </section>

              <section>
                <h3>3. スライドを直接編集する</h3>
                <ul>
                  <li><b>テキストボックス</b>：ドラッグで移動／角でリサイズ／<b>ダブルクリックで文字編集</b>／文字を選択して<b>太字・斜体</b>・色・サイズ変更／<b>フォント</b>変更。</li>
                  <li><b>表</b>：セルを<b>ダブルクリックで編集</b>、<b>列・行の境界をドラッグ</b>で幅・高さを調整。</li>
                  <li>ツールバーの <b>＋テキストボックス / ＋表 / ＋スライド</b> で追加。</li>
                  <li><b>↑↓←→</b> でスライド切替、<b>Backspace / Delete</b> で選択中の要素を削除、<b>Ctrl/⌘+Z</b> で元に戻す。</li>
                </ul>
              </section>

              <section>
                <h3>4. 書き出す</h3>
                <ul>
                  <li>右上の <b>「書き出す ▾」</b> から <b>PPTX（編集可能）</b> か <b>PDF</b> を選んでダウンロード。</li>
                  <li>ファイル名は右上の「ファイル名」欄で変更できます。</li>
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

export default App
