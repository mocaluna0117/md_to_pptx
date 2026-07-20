import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { renderPreview } from './lib/marp'
import { exportPptx } from './lib/exportPptx'
import { exportPptxNative } from './lib/exportPptxNative'
import { exportDeckToPptx } from './lib/exportDeck'
import { type Deck } from './lib/deck'
import { deckFromRenderedMarkdown } from './lib/deckFromRender'
import VisualEditor from './components/VisualEditor'
import './App.css'

type Mode = 'image' | 'native'
type View = 'markdown' | 'visual'

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

  const preview = useMemo(() => renderPreview(markdown), [markdown])
  const exporting = status.kind === 'exporting'

  // Persist Markdown + visual deck so a reload/close doesn't lose work.
  useEffect(() => {
    const id = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, deck, deckDirty }))
      } catch {
        /* storage full or unavailable */
      }
    }, 300)
    return () => clearTimeout(id)
  }, [markdown, fileName, deck, deckDirty])

  const handleDeckChange = useCallback((d: Deck) => {
    setDeck(d)
    setDeckDirty(true)
  }, [])

  async function enterVisual() {
    if (!deck) setDeck(await deckFromRenderedMarkdown(markdown))
    setView('visual')
  }

  /** Rebuild the deck from the current Markdown, confirming if edits would be lost. */
  async function rebuildFromMarkdown(): Promise<boolean> {
    if (deck && deckDirty && !window.confirm('現在の Markdown からスライドを作り直します。ビジュアル編集の変更は上書きされます。よろしいですか？')) {
      return false
    }
    setDeck(await deckFromRenderedMarkdown(markdown))
    setDeckDirty(false)
    return true
  }

  async function applyMarkdownToVisual() {
    if (await rebuildFromMarkdown()) setView('visual')
  }

  const fileInputRef = useRef<HTMLInputElement>(null)

  async function importMarkdownFile(file: File) {
    try {
      const text = await file.text()
      setMarkdown(text)
      const base = file.name.replace(/\.[^.]+$/, '')
      if (base) setFileName(base)
      setView('markdown')
    } catch {
      setStatus({ kind: 'error', message: 'ファイルの読み込みに失敗しました。' })
    }
  }

  async function handleExport() {
    setStatus({ kind: 'exporting', done: 0, total: 0 })
    const onProgress = (done: number, total: number) => setStatus({ kind: 'exporting', done, total })
    try {
      if (view === 'visual') {
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
            <span className="ext">.pptx</span>
          </label>
          <button className="export" onClick={handleExport} disabled={exporting}>
            {exporting
              ? status.total > 0
                ? `書き出し中 ${status.done}/${status.total}`
                : '準備中…'
              : 'PPTX を書き出す'}
          </button>
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
        <VisualEditor deck={deck} onChange={handleDeckChange} onRegenerate={rebuildFromMarkdown} />
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
