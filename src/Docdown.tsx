import { useEffect, useRef, useState } from 'react'
import MarkdownIt from 'markdown-it'
import { navigate } from './Root'
import { exportMarkdownToDocx } from './lib/exportDocx'
import { resolveImagePaths, readImageFiles, IMAGE_EXT, type AttachedImages } from './lib/imageAttach'
import { mathToImages } from './lib/math'
import './App.css'
import './Docdown.css'

const STORAGE_KEY = 'docdown:v1'
const mdRender = new MarkdownIt({ html: true, linkify: true, breaks: false })

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
  const [mdOpen, setMdOpen] = useState<boolean>(persisted.mdOpen ?? true)
  const [images, setImages] = useState<AttachedImages>(persisted.images ?? {})
  const [status, setStatus] = useState<Status>('idle')
  const fileInputRef = useRef<HTMLInputElement>(null)
  const imageInputRef = useRef<HTMLInputElement>(null)
  const importModeRef = useRef<'replace' | 'append'>('replace')

  const [html, setHtml] = useState('')
  const imageNames = Object.keys(images)

  // Preview renders asynchronously (math is rasterized); debounced + math-cached.
  useEffect(() => {
    let cancelled = false
    const id = setTimeout(async () => {
      const prepared = await mathToImages(resolveImagePaths(markdown, images))
      if (!cancelled) setHtml(mdRender.render(stripFrontmatter(prepared)))
    }, 200)
    return () => {
      cancelled = true
      clearTimeout(id)
    }
  }, [markdown, images])
  const exporting = status === 'exporting'
  const error = typeof status === 'object' ? status.error : null

  useEffect(() => {
    const id = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, mdOpen, images }))
      } catch {
        // Images may exceed the storage quota: keep at least the text.
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify({ markdown, fileName, mdOpen }))
        } catch {
          /* storage unavailable */
        }
      }
    }, 300)
    return () => clearTimeout(id)
  }, [markdown, fileName, mdOpen, images])

  async function addImageFiles(files: File[]) {
    const imgs = files.filter((f) => IMAGE_EXT.test(f.name) || f.type.startsWith('image/'))
    if (imgs.length === 0) return
    const loaded = await readImageFiles(imgs)
    setImages((cur) => ({ ...cur, ...loaded }))
  }

  async function importFile(file: File, mode: 'replace' | 'append') {
    try {
      const text = await file.text()
      if (mode === 'append') {
        setMarkdown((cur) => mergeMarkdown(cur, text))
      } else {
        setMarkdown(text)
        const base = file.name.replace(/\.[^.]+$/, '')
        if (base) setFileName(base)
      }
    } catch {
      setStatus({ error: 'ファイルの読み込みに失敗しました。' })
    }
  }

  async function handleExport() {
    setStatus('exporting')
    try {
      await exportMarkdownToDocx(await mathToImages(resolveImagePaths(markdown, images)), { fileName })
      setStatus('idle')
    } catch (err) {
      setStatus({ error: err instanceof Error ? err.message : String(err) })
    }
  }

  function resetToDefault() {
    if (!window.confirm('内容を初期状態に戻します。よろしいですか？')) return
    setMarkdown(SAMPLE)
    setFileName('document')
    setImages({})
    setStatus('idle')
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
              type="text"
              value={fileName}
              onChange={(e) => setFileName(e.target.value)}
              placeholder="document"
              spellCheck={false}
              aria-label="ファイル名"
            />
          </label>
          <button className="export" onClick={handleExport} disabled={exporting}>
            {exporting ? '書き出し中…' : 'Word で書き出す'}
          </button>
        </div>
      </header>

      <div className="banner info">
        左の「Markdown」タブで文章を書く／インポートし、中央のプレビューで確認して、右上から <b>Word（.docx）</b> に書き出せます。
      </div>

      {error && (
        <div className="banner error" role="alert">
          ⚠️ {error}
        </div>
      )}

      <div className="workspace">
        <aside className={`md-drawer${mdOpen ? ' open' : ''}`} style={{ width: mdOpen ? 360 : 0 }}>
          <div
            className="md-inner"
            style={{ width: 360 }}
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
          </div>
        </aside>

        <button
          className="md-handle toggle"
          onClick={() => setMdOpen((o) => !o)}
          aria-expanded={mdOpen}
          title={mdOpen ? 'Markdown を閉じる' : 'Markdown を開く'}
        >
          <span className="md-handle-text">Markdown</span>
          <span className="md-handle-arrow" aria-hidden>
            {mdOpen ? '◀' : '▶'}
          </span>
        </button>

        <main className="doc-main">
          <div className="doc-scroll">
            <div className="doc-page" dangerouslySetInnerHTML={{ __html: html }} />
          </div>
        </main>
      </div>
    </div>
  )
}
