import { useMemo, useState } from 'react'
import { renderPreview } from './lib/marp'
import { exportPptx } from './lib/exportPptx'
import { exportPptxNative } from './lib/exportPptxNative'
import './App.css'

type Mode = 'image' | 'native'

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
  const [markdown, setMarkdown] = useState(SAMPLE)
  const [fileName, setFileName] = useState('slides')
  const [mode, setMode] = useState<Mode>('image')
  const [status, setStatus] = useState<Status>({ kind: 'idle' })

  const preview = useMemo(() => renderPreview(markdown), [markdown])
  const exporting = status.kind === 'exporting'

  async function handleExport() {
    setStatus({ kind: 'exporting', done: 0, total: 0 })
    const run = mode === 'image' ? exportPptx : exportPptxNative
    try {
      await run(markdown, {
        fileName,
        onProgress: (done, total) => setStatus({ kind: 'exporting', done, total }),
      })
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
        </div>
        <div className="actions">
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
        {mode === 'image'
          ? '画像方式：Marp の見た目をそのまま再現。PowerPoint 上では画像になり、テキスト編集はできません。'
          : '編集可能方式：見出し・本文・箇条書き・コード・引用・表を編集可能なテキストに変換。テーマの再現は簡易です。'}
      </div>

      {status.kind === 'error' && (
        <div className="banner error" role="alert">
          ⚠️ {status.message}
        </div>
      )}

      <main className="panes">
        <section className="pane editor-pane">
          <div className="pane-head">Markdown</div>
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
    </div>
  )
}

export default App
