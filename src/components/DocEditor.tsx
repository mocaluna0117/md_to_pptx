import { useEffect, useRef, useState } from 'react'
import type { AttachedImages } from '../lib/imageAttach'

interface Props {
  /** Initial document HTML. Seeded into the editable area once on mount. */
  html: string
  /** Attached images (basename → data URI) offered in the insert-image menu. */
  images: AttachedImages
  /** Called (debounced) with the edited HTML whenever the document changes. */
  onChange: (html: string) => void
}

const DEFAULT_PT = 11
const MIN_PT = 6
const MAX_PT = 96

const BLOCK_OPTIONS: { value: string; label: string }[] = [
  { value: 'p', label: '本文' },
  { value: 'h1', label: '見出し 1' },
  { value: 'h2', label: '見出し 2' },
  { value: 'h3', label: '見出し 3' },
  { value: 'h4', label: '見出し 4' },
  { value: 'blockquote', label: '引用' },
  { value: 'pre', label: 'コード' },
]

const FONT_OPTIONS: { value: string; label: string }[] = [
  { value: '', label: 'フォント' },
  { value: 'Arial', label: 'Arial（ゴシック）' },
  { value: 'Times New Roman', label: 'Times（明朝）' },
  { value: 'Meiryo', label: 'メイリオ' },
  { value: 'Yu Gothic', label: '游ゴシック' },
  { value: 'Yu Mincho', label: '游明朝' },
  { value: 'Courier New', label: 'Courier（等幅）' },
]

const COLORS = ['111111', 'E11D48', '2563EB', '059669', 'D97706', '7C3AED', '6B7280', 'FFFFFF']

/**
 * WYSIWYG document editor: a single contentEditable area plus a formatting toolbar.
 * The edited HTML is the source of truth (Docdown exports it directly to .docx).
 * Content is seeded once on mount — the parent remounts (via `key`) to reseed on rebuild —
 * so React never overwrites the live DOM and the caret is preserved while typing.
 */
export default function DocEditor({ html, images, onChange }: Props) {
  const editorRef = useRef<HTMLDivElement>(null)
  const savedRange = useRef<Range | null>(null)
  const emitTimer = useRef<number | null>(null)
  const [imgMenuOpen, setImgMenuOpen] = useState(false)
  const [active, setActive] = useState<{ bold: boolean; italic: boolean; strike: boolean; block: string; inTable: boolean }>({
    bold: false,
    italic: false,
    strike: false,
    block: 'p',
    inTable: false,
  })
  const imageNames = Object.keys(images)

  // Seed the editable content once on mount.
  useEffect(() => {
    const el = editorRef.current
    if (!el) return
    el.innerHTML = html
    try {
      document.execCommand('defaultParagraphSeparator', false, 'p')
      document.execCommand('styleWithCSS', false, 'false')
    } catch {
      /* older engines */
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Track the live selection so toolbar buttons can operate on it after focus moves to them.
  useEffect(() => {
    const onSelChange = () => {
      const el = editorRef.current
      const sel = window.getSelection()
      if (!el || !sel || sel.rangeCount === 0) return
      const range = sel.getRangeAt(0)
      if (el.contains(range.commonAncestorContainer)) {
        savedRange.current = range.cloneRange()
        refreshActive()
      }
    }
    document.addEventListener('selectionchange', onSelChange)
    return () => document.removeEventListener('selectionchange', onSelChange)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  function refreshActive() {
    const el = editorRef.current
    const sel = window.getSelection()
    if (!el || !sel || sel.rangeCount === 0) return
    let node: Node | null = sel.getRangeAt(0).startContainer
    let block = 'p'
    let inTable = false
    while (node && node !== el) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const tag = (node as HTMLElement).tagName
        if ((tag === 'TD' || tag === 'TH')) inTable = true
        if (block === 'p' && /^(H[1-6]|BLOCKQUOTE|PRE|P)$/.test(tag)) block = tag.toLowerCase()
      }
      node = node.parentNode
    }
    let bold = false
    let italic = false
    let strike = false
    try {
      bold = document.queryCommandState('bold')
      italic = document.queryCommandState('italic')
      strike = document.queryCommandState('strikeThrough')
    } catch {
      /* ignore */
    }
    setActive({ bold, italic, strike, block, inTable })
  }

  function emitSoon() {
    if (emitTimer.current) window.clearTimeout(emitTimer.current)
    emitTimer.current = window.setTimeout(() => {
      if (editorRef.current) onChange(editorRef.current.innerHTML)
    }, 250)
  }
  function emitNow() {
    if (emitTimer.current) window.clearTimeout(emitTimer.current)
    if (editorRef.current) onChange(editorRef.current.innerHTML)
  }

  /** Refocus the editor and restore the last selection before running a command. */
  function restore() {
    const el = editorRef.current
    if (!el) return
    el.focus()
    const sel = window.getSelection()
    if (sel && savedRange.current) {
      sel.removeAllRanges()
      sel.addRange(savedRange.current)
    }
  }

  function exec(command: string, value?: string, css = false) {
    restore()
    try {
      document.execCommand('styleWithCSS', false, css ? 'true' : 'false')
      document.execCommand(command, false, value)
    } catch {
      /* ignore */
    }
    refreshActive()
    emitNow()
  }

  function setBlock(tag: string) {
    // formatBlock wants an angle-bracketed tag name on most engines.
    exec('formatBlock', `<${tag}>`)
  }

  function insertHtml(htmlStr: string) {
    exec('insertHTML', htmlStr)
  }

  function addLink() {
    restore()
    const url = window.prompt('リンク先の URL を入力してください', 'https://')
    if (!url) return
    exec('createLink', url)
  }

  function insertImage(src: string) {
    setImgMenuOpen(false)
    exec('insertImage', src)
  }

  function insertTable() {
    const rows = 2
    const cols = 2
    const head = `<tr>${Array.from({ length: cols }, (_, c) => `<th>見出し${c + 1}</th>`).join('')}</tr>`
    const body = Array.from({ length: rows }, () => `<tr>${Array.from({ length: cols }, () => '<td>&nbsp;</td>').join('')}</tr>`).join('')
    insertHtml(`<table><thead>${head}</thead><tbody>${body}</tbody></table><p><br></p>`)
  }

  // ---- Table structural edits (operate on the cell containing the caret) ----

  function currentCell(): HTMLTableCellElement | null {
    let node: Node | null = savedRange.current?.startContainer ?? null
    while (node && node !== editorRef.current) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const tag = (node as HTMLElement).tagName
        if (tag === 'TD' || tag === 'TH') return node as HTMLTableCellElement
      }
      node = node.parentNode
    }
    return null
  }

  function withCell(fn: (cell: HTMLTableCellElement, row: HTMLTableRowElement, table: HTMLTableElement) => void) {
    const cell = currentCell()
    const row = cell?.parentElement as HTMLTableRowElement | undefined
    const table = cell?.closest('table') as HTMLTableElement | null
    if (!cell || !row || !table) return
    fn(cell, row, table)
    refreshActive()
    emitNow()
  }

  function addRow() {
    withCell((_cell, row) => {
      const cols = row.children.length
      const tr = document.createElement('tr')
      for (let c = 0; c < cols; c++) {
        const td = document.createElement('td')
        td.innerHTML = '&nbsp;'
        tr.appendChild(td)
      }
      row.after(tr)
    })
  }
  function delRow() {
    withCell((_cell, row, table) => {
      if (table.querySelectorAll('tr').length > 1) row.remove()
    })
  }
  function addCol() {
    withCell((cell, _row, table) => {
      const index = Array.from((cell.parentElement as HTMLTableRowElement).children).indexOf(cell)
      for (const tr of Array.from(table.querySelectorAll('tr'))) {
        const ref = tr.children[index]
        const isHead = ref && ref.tagName === 'TH'
        const nc = document.createElement(isHead ? 'th' : 'td')
        nc.innerHTML = '&nbsp;'
        if (ref) ref.after(nc)
        else tr.appendChild(nc)
      }
    })
  }
  function delCol() {
    withCell((cell, _row, table) => {
      const index = Array.from((cell.parentElement as HTMLTableRowElement).children).indexOf(cell)
      const firstRowCells = table.querySelector('tr')?.children.length ?? 0
      if (firstRowCells <= 1) return
      for (const tr of Array.from(table.querySelectorAll('tr'))) {
        tr.children[index]?.remove()
      }
    })
  }

  // ---- Font size (point-based, carried on data-fs so the .docx export can read it) ----

  function effectiveFs(node: Node | null): number {
    let n: Node | null = node
    while (n && n !== editorRef.current) {
      if (n.nodeType === Node.ELEMENT_NODE) {
        const fs = (n as HTMLElement).dataset?.fs
        if (fs) return Number(fs)
      }
      n = n.parentNode
    }
    return DEFAULT_PT
  }

  function changeFontSize(delta: number) {
    restore()
    const sel = window.getSelection()
    if (!sel || sel.rangeCount === 0) return
    const range = sel.getRangeAt(0)
    const cur = effectiveFs(range.startContainer)
    const nextPt = Math.max(MIN_PT, Math.min(MAX_PT, cur + delta))
    if (range.collapsed) return
    const frag = range.extractContents()
    // Preserve relative sizing on already-sized descendants.
    frag.querySelectorAll?.('[data-fs]').forEach((elm) => {
      const el = elm as HTMLElement
      const bumped = Math.max(MIN_PT, Math.min(MAX_PT, Number(el.dataset.fs) + delta))
      el.dataset.fs = String(bumped)
      el.style.fontSize = `${bumped}pt`
    })
    const span = document.createElement('span')
    span.dataset.fs = String(nextPt)
    span.style.fontSize = `${nextPt}pt`
    span.appendChild(frag)
    range.insertNode(span)
    // Reselect the wrapped content.
    const newRange = document.createRange()
    newRange.selectNodeContents(span)
    sel.removeAllRanges()
    sel.addRange(newRange)
    savedRange.current = newRange.cloneRange()
    emitNow()
  }

  const btn = (label: string, title: string, onClick: () => void, isActive = false, extraClass = '') => (
    <button
      type="button"
      className={`det-btn${isActive ? ' active' : ''}${extraClass ? ' ' + extraClass : ''}`}
      title={title}
      aria-label={title}
      aria-pressed={isActive}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {label}
    </button>
  )

  return (
    <div className="doc-editor">
      <div className="doc-editor-toolbar" role="toolbar" aria-label="書式">
        <div className="det-group">
          {btn('↶', '元に戻す (Ctrl+Z)', () => exec('undo'))}
          {btn('↷', 'やり直し (Ctrl+Y)', () => exec('redo'))}
        </div>

        <div className="det-group">
          <select
            className="det-select"
            aria-label="段落スタイル"
            value={active.block}
            onMouseDown={() => restore()}
            onChange={(e) => setBlock(e.target.value)}
          >
            {BLOCK_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </div>

        <div className="det-group">
          {btn('B', '太字', () => exec('bold'), active.bold, 'det-b')}
          {btn('I', '斜体', () => exec('italic'), active.italic, 'det-i')}
          {btn('S', '取り消し線', () => exec('strikeThrough'), active.strike, 'det-s')}
          {btn('U', '下線', () => exec('underline'), false, 'det-u')}
        </div>

        <div className="det-group">
          {btn('A−', '文字を小さく', () => changeFontSize(-1))}
          {btn('A＋', '文字を大きく', () => changeFontSize(1))}
          <select
            className="det-select"
            aria-label="フォント"
            defaultValue=""
            onMouseDown={() => restore()}
            onChange={(e) => {
              if (e.target.value) exec('fontName', e.target.value, true)
              e.target.value = ''
            }}
          >
            {FONT_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </div>

        <div className="det-group det-colors">
          {COLORS.map((c) => (
            <button
              key={c}
              type="button"
              className="det-swatch"
              title={`文字色 #${c}`}
              aria-label={`文字色 #${c}`}
              style={{ background: `#${c}`, ...(c === 'FFFFFF' ? { boxShadow: 'inset 0 0 0 1px #ccc' } : {}) }}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => exec('foreColor', `#${c}`, true)}
            />
          ))}
        </div>

        <div className="det-group">
          {btn('⯀', '左揃え', () => exec('justifyLeft'))}
          {btn('☰', '中央揃え', () => exec('justifyCenter'))}
          {btn('▟', '右揃え', () => exec('justifyRight'))}
        </div>

        <div className="det-group">
          {btn('•', '箇条書き', () => exec('insertUnorderedList'))}
          {btn('1.', '番号付きリスト', () => exec('insertOrderedList'))}
          {btn('🔗', 'リンク', addLink)}
          {btn('―', '水平線', () => exec('insertHorizontalRule'))}
        </div>

        <div className="det-group det-img-wrap">
          {btn('🖼', '画像を挿入', () => setImgMenuOpen((o) => !o))}
          {imgMenuOpen && (
            <div className="det-img-menu" role="menu">
              {imageNames.length === 0 ? (
                <div className="det-img-empty">左の「🖼 画像」から画像を読み込むと挿入できます</div>
              ) : (
                imageNames.map((name) => (
                  <button key={name} type="button" role="menuitem" className="det-img-item" onMouseDown={(e) => e.preventDefault()} onClick={() => insertImage(images[name])}>
                    <img src={images[name]} alt="" />
                    <span>{name}</span>
                  </button>
                ))
              )}
            </div>
          )}
        </div>

        <div className="det-group">
          {btn('▦', '表を挿入', insertTable)}
          {active.inTable && (
            <>
              {btn('＋行', '行を追加', addRow)}
              {btn('−行', '行を削除', delRow)}
              {btn('＋列', '列を追加', addCol)}
              {btn('−列', '列を削除', delCol)}
            </>
          )}
        </div>
      </div>

      <div className="doc-scroll">
        <div
          ref={editorRef}
          className="doc-page doc-editable"
          contentEditable
          suppressContentEditableWarning
          spellCheck={false}
          onInput={emitSoon}
          onBlur={emitNow}
        />
      </div>
    </div>
  )
}
