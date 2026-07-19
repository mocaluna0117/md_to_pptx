import { useMemo, useState } from 'react'
import { renderPreview } from './lib/marp'
import { exportPptx } from './lib/exportPptx'
import './App.css'

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
  const [status, setStatus] = useState<Status>({ kind: 'idle' })

  const preview = useMemo(() => renderPreview(markdown), [markdown])
  const exporting = status.kind === 'exporting'

  async function handleExport() {
    setStatus({ kind: 'exporting', done: 0, total: 0 })
    try {
      await exportPptx(markdown, {
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
