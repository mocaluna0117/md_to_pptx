import { useEffect, useRef, useState } from 'react'
import type { AttachedImages } from '../lib/imageAttach'
import { newDocBox, type DocBox } from '../lib/docBox'

interface Props {
  /** Initial document HTML. Seeded into the editable area once on mount. */
  html: string
  /** Attached images (basename → data URI) offered in the insert-image menu. */
  images: AttachedImages
  /** Called (debounced) with the edited HTML whenever the flowing document changes. */
  onChange: (html: string) => void
  /** Free-floating text boxes layered over the document. */
  boxes: DocBox[]
  onBoxesChange: (boxes: DocBox[]) => void
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

// Alignment icon shared with Deckdown's visual editor (rows of bars).
type Align = 'left' | 'center' | 'right'
const ALIGN_BARS: Record<Align, [number, number][]> = {
  left: [[1, 14], [1, 8], [1, 14], [1, 8]],
  center: [[1, 14], [4, 8], [1, 14], [4, 8]],
  right: [[1, 14], [7, 8], [1, 14], [7, 8]],
}
function AlignIcon({ dir }: { dir: Align }) {
  return (
    <svg viewBox="0 0 16 16" width="15" height="15" fill="currentColor" aria-hidden focusable="false">
      {ALIGN_BARS[dir].map(([x, w], i) => (
        <rect key={i} x={x} y={2.5 + i * 3.7} width={w} height="1.6" rx="0.8" />
      ))}
    </svg>
  )
}

const BLOCK_TAGS = /^(P|DIV|H[1-6]|LI|BLOCKQUOTE|PRE|TD|TH|UL|OL|TABLE|FIGURE|SECTION)$/

/** Nearest block-level ancestor of `node` within `root` (or null). */
function blockAncestor(node: Node | null, root: HTMLElement): HTMLElement | null {
  let n: Node | null = node
  while (n && n !== root) {
    if (n.nodeType === Node.ELEMENT_NODE && BLOCK_TAGS.test((n as HTMLElement).tagName)) return n as HTMLElement
    n = n.parentNode
  }
  return null
}

// Curved-arrow undo/redo glyphs shared with Deckdown's visual editor.
function UndoIcon() {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden focusable="false">
      <path d="M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z" />
    </svg>
  )
}
function RedoIcon() {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden focusable="false">
      <path d="M18.4 10.6C16.55 8.99 14.15 8 11.5 8c-4.65 0-8.58 3.03-9.96 7.22L3.9 16c1.05-3.19 4.05-5.5 7.6-5.5 1.95 0 3.73.72 5.12 1.88L13 16h9V7l-3.6 3.6z" />
    </svg>
  )
}

/**
 * WYSIWYG document editor: a flowing contentEditable page plus a layer of free-floating
 * text boxes, driven by one shared formatting toolbar. The edited HTML (flow) and the box
 * list are the source of truth (Docdown exports them directly to .docx).
 *
 * Flow content is seeded once on mount — the parent remounts (via `key`) to reseed on
 * rebuild — so React never overwrites the live DOM and the caret is preserved while typing.
 * The toolbar operates on whichever editable surface (the page or a box) last held the
 * selection, tracked via `activeEditableRef`.
 */
export default function DocEditor({ html, images, onChange, boxes, onBoxesChange }: Props) {
  const editorRef = useRef<HTMLDivElement>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const savedRange = useRef<Range | null>(null)
  const activeEditableRef = useRef<HTMLElement | null>(null)
  const emitTimer = useRef<number | null>(null)
  const boxesRef = useRef(boxes)
  boxesRef.current = boxes
  const boxBodyRefs = useRef<Map<string, HTMLDivElement>>(new Map())
  const dragRef = useRef<{ mode: 'move' | 'resize'; id: string; sx: number; sy: number; ox: number; oy: number; ow: number; oh: number } | null>(null)

  const [imgMenuOpen, setImgMenuOpen] = useState(false)
  const [selectedBoxId, setSelectedBoxId] = useState<string | null>(null)
  const [editingBoxId, setEditingBoxId] = useState<string | null>(null)
  const [active, setActive] = useState({ bold: false, italic: false, strike: false, block: 'p', align: 'left' })
  const imageNames = Object.keys(images)

  // Seed the flowing content once on mount.
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

  // Track the live selection (page or box) so toolbar buttons operate on the right surface.
  useEffect(() => {
    const onSelChange = () => {
      const sel = window.getSelection()
      if (!sel || sel.rangeCount === 0) return
      const range = sel.getRangeAt(0)
      const surface = editableAncestor(range.commonAncestorContainer)
      if (!surface) return
      savedRange.current = range.cloneRange()
      activeEditableRef.current = surface
      refreshActive(surface)
    }
    document.addEventListener('selectionchange', onSelChange)
    return () => document.removeEventListener('selectionchange', onSelChange)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  /** Nearest ancestor that is an editable surface (the page or a box body). */
  function editableAncestor(node: Node | null): HTMLElement | null {
    let n: Node | null = node
    while (n) {
      if (n.nodeType === Node.ELEMENT_NODE) {
        const el = n as HTMLElement
        if (el.classList?.contains('doc-editable') || el.classList?.contains('doc-box-body')) return el
      }
      n = n.parentNode
    }
    return null
  }

  function refreshActive(surface: HTMLElement) {
    const sel = window.getSelection()
    if (!sel || sel.rangeCount === 0) return
    let node: Node | null = sel.getRangeAt(0).startContainer
    let block = 'p'
    while (node && node !== surface) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const tag = (node as HTMLElement).tagName
        if (block === 'p' && /^(H[1-6]|BLOCKQUOTE|PRE|P)$/.test(tag)) block = tag.toLowerCase()
      }
      node = node.parentNode
    }
    let align = 'left'
    const blockEl = blockAncestor(sel.getRangeAt(0).startContainer, surface)
    if (blockEl) {
      const ta = getComputedStyle(blockEl).textAlign
      align = ta === 'center' ? 'center' : ta === 'right' || ta === 'end' ? 'right' : ta === 'justify' ? 'justify' : 'left'
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
    setActive({ bold, italic, strike, block, align })
  }

  /**
   * Set paragraph alignment deterministically on the block(s) the selection touches,
   * instead of document.execCommand('justify*') — which produces broken nested <p> when
   * lines are joined by <br>, so some lines wouldn't align.
   */
  function applyAlign(dir: 'left' | 'center' | 'right') {
    restore()
    const root = activeEditableRef.current
    const sel = window.getSelection()
    if (!root || !sel || sel.rangeCount === 0) return
    const range = sel.getRangeAt(0)
    const blocks: HTMLElement[] = []
    for (const child of Array.from(root.children)) {
      if (range.intersectsNode(child)) blocks.push(child as HTMLElement)
    }
    if (blocks.length === 0) {
      const b = blockAncestor(range.startContainer, root)
      if (b) blocks.push(b)
    }
    for (const el of blocks) el.style.textAlign = dir
    refreshActive(root)
    syncActive()
  }

  /** Push the current content of whichever surface was last active back to the parent. */
  function syncActive() {
    const surface = activeEditableRef.current
    if (!surface) return
    if (surface.classList.contains('doc-editable')) {
      onChange(surface.innerHTML)
    } else {
      const id = surface.dataset.boxId
      if (id) patchBox(id, { html: surface.innerHTML })
    }
  }

  function emitFlowSoon() {
    if (emitTimer.current) window.clearTimeout(emitTimer.current)
    emitTimer.current = window.setTimeout(() => {
      if (editorRef.current) onChange(editorRef.current.innerHTML)
    }, 250)
  }

  /** Refocus the last active surface and restore its selection before running a command. */
  function restore() {
    const el = activeEditableRef.current ?? editorRef.current
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
    if (activeEditableRef.current) refreshActive(activeEditableRef.current)
    syncActive()
  }

  function setBlock(tag: string) {
    exec('formatBlock', `<${tag}>`)
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
    const cols = 2
    const head = `<tr>${Array.from({ length: cols }, (_, c) => `<th>見出し${c + 1}</th>`).join('')}</tr>`
    const body = `<tr>${Array.from({ length: cols }, () => '<td>&nbsp;</td>').join('')}</tr>`
    exec('insertHTML', `<table><thead>${head}</thead><tbody>${body}</tbody></table><p><br></p>`)
  }

  // ---- Font size (points, carried on data-fs so the .docx export can read it) ----

  function effectiveFs(node: Node | null, root: HTMLElement | null): number {
    let n: Node | null = node
    while (n && n !== root) {
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
    if (range.collapsed) return
    const cur = effectiveFs(range.startContainer, activeEditableRef.current)
    const nextPt = Math.max(MIN_PT, Math.min(MAX_PT, cur + delta))
    const frag = range.extractContents()
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
    const newRange = document.createRange()
    newRange.selectNodeContents(span)
    sel.removeAllRanges()
    sel.addRange(newRange)
    savedRange.current = newRange.cloneRange()
    syncActive()
  }

  // ---- Text boxes ----

  function patchBox(id: string, patch: Partial<DocBox>) {
    onBoxesChange(boxesRef.current.map((b) => (b.id === id ? { ...b, ...patch } : b)))
  }

  function addBox() {
    const box = newDocBox()
    onBoxesChange([...boxesRef.current, box])
    setSelectedBoxId(box.id)
    setEditingBoxId(null)
  }

  function deleteBox(id: string) {
    onBoxesChange(boxesRef.current.filter((b) => b.id !== id))
    if (selectedBoxId === id) setSelectedBoxId(null)
    if (editingBoxId === id) setEditingBoxId(null)
  }

  function startBoxGesture(e: React.PointerEvent, box: DocBox, mode: 'move' | 'resize') {
    e.preventDefault()
    e.stopPropagation()
    setSelectedBoxId(box.id)
    dragRef.current = { mode, id: box.id, sx: e.clientX, sy: e.clientY, ox: box.x, oy: box.y, ow: box.w, oh: box.h }
    window.addEventListener('pointermove', onGestureMove)
    window.addEventListener('pointerup', onGestureUp)
  }
  function onGestureMove(e: PointerEvent) {
    const d = dragRef.current
    if (!d) return
    const dx = e.clientX - d.sx
    const dy = e.clientY - d.sy
    if (d.mode === 'move') {
      patchBox(d.id, { x: Math.max(0, Math.round(d.ox + dx)), y: Math.max(0, Math.round(d.oy + dy)) })
    } else {
      patchBox(d.id, { w: Math.max(60, Math.round(d.ow + dx)), h: Math.max(36, Math.round(d.oh + dy)) })
    }
  }
  function onGestureUp() {
    dragRef.current = null
    window.removeEventListener('pointermove', onGestureMove)
    window.removeEventListener('pointerup', onGestureUp)
  }

  const btn = (label: React.ReactNode, title: string, onClick: () => void, isActive = false, extraClass = '') => (
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
          {btn(<UndoIcon />, '元に戻す (Ctrl+Z)', () => exec('undo'))}
          {btn(<RedoIcon />, 'やり直し (Ctrl+Y)', () => exec('redo'))}
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
          {btn(<AlignIcon dir="left" />, '左揃え', () => applyAlign('left'), active.align === 'left')}
          {btn(<AlignIcon dir="center" />, '中央揃え', () => applyAlign('center'), active.align === 'center')}
          {btn(<AlignIcon dir="right" />, '右揃え', () => applyAlign('right'), active.align === 'right')}
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
          {btn('＋テキストボックス', 'テキストボックスを追加', addBox, false, 'det-box-add')}
        </div>
      </div>

      <div className="doc-scroll">
        <div className="doc-page-wrap" ref={wrapRef}>
          <div
            ref={editorRef}
            className="doc-page doc-editable"
            contentEditable
            suppressContentEditableWarning
            spellCheck={false}
            onInput={emitFlowSoon}
            onBlur={() => {
              if (editorRef.current) onChange(editorRef.current.innerHTML)
            }}
            onMouseDown={() => setSelectedBoxId(null)}
          />
          <div className="doc-box-layer">
            {boxes.map((box) => (
              <DocBoxView
                key={box.id}
                box={box}
                selected={selectedBoxId === box.id}
                editing={editingBoxId === box.id}
                onSelect={() => setSelectedBoxId(box.id)}
                onStartMove={(e) => startBoxGesture(e, box, 'move')}
                onStartResize={(e) => startBoxGesture(e, box, 'resize')}
                onEdit={() => {
                  setSelectedBoxId(box.id)
                  setEditingBoxId(box.id)
                }}
                onStopEdit={(nextHtml) => {
                  patchBox(box.id, { html: nextHtml })
                  setEditingBoxId(null)
                }}
                onChangeHtml={(nextHtml) => patchBox(box.id, { html: nextHtml })}
                onDelete={() => deleteBox(box.id)}
                registerBody={(el) => {
                  if (el) boxBodyRefs.current.set(box.id, el)
                  else boxBodyRefs.current.delete(box.id)
                }}
              />
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}

interface BoxProps {
  box: DocBox
  selected: boolean
  editing: boolean
  onSelect: () => void
  onStartMove: (e: React.PointerEvent) => void
  onStartResize: (e: React.PointerEvent) => void
  onEdit: () => void
  onStopEdit: (html: string) => void
  onChangeHtml: (html: string) => void
  onDelete: () => void
  registerBody: (el: HTMLDivElement | null) => void
}

/** One floating text box. Static (draggable) until double-clicked into edit mode. */
function DocBoxView({ box, selected, editing, onSelect, onStartMove, onStartResize, onEdit, onStopEdit, onChangeHtml, onDelete, registerBody }: BoxProps) {
  const bodyRef = useRef<HTMLDivElement>(null)
  const emit = useRef<number | null>(null)

  // Seed the editable body once when entering edit mode; place the caret at the end.
  useEffect(() => {
    if (!editing) return
    const el = bodyRef.current
    if (!el) return
    el.innerHTML = box.html
    registerBody(el)
    el.focus()
    const range = document.createRange()
    range.selectNodeContents(el)
    range.collapse(false)
    const sel = window.getSelection()
    sel?.removeAllRanges()
    sel?.addRange(range)
    return () => registerBody(null)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editing])

  const style: React.CSSProperties = { left: box.x, top: box.y, width: box.w, height: box.h }

  return (
    <div
      className={`doc-box${selected ? ' selected' : ''}`}
      style={style}
      onPointerDown={(e) => {
        if (editing) return
        // Left-button drag moves the box; also selects it.
        if (e.button === 0) onStartMove(e)
      }}
      onClick={(e) => {
        e.stopPropagation()
        onSelect()
      }}
      onDoubleClick={(e) => {
        e.stopPropagation()
        onEdit()
      }}
    >
      {editing ? (
        <div
          ref={bodyRef}
          className="doc-box-body"
          data-box-id={box.id}
          contentEditable
          suppressContentEditableWarning
          spellCheck={false}
          onInput={() => {
            if (emit.current) window.clearTimeout(emit.current)
            emit.current = window.setTimeout(() => {
              if (bodyRef.current) onChangeHtml(bodyRef.current.innerHTML)
            }, 200)
          }}
          onBlur={() => {
            if (bodyRef.current) onStopEdit(bodyRef.current.innerHTML)
          }}
          onKeyDown={(e) => {
            if (e.key === 'Escape') {
              e.preventDefault()
              ;(e.target as HTMLElement).blur()
            }
          }}
        />
      ) : (
        <div className="doc-box-body" dangerouslySetInnerHTML={{ __html: box.html || '<span style="opacity:.5">ダブルクリックで編集</span>' }} />
      )}
      {selected && (
        <>
          <button className="doc-box-del" title="ボックスを削除" aria-label="ボックスを削除" onMouseDown={(e) => e.preventDefault()} onClick={(e) => { e.stopPropagation(); onDelete() }}>
            ×
          </button>
          <span className="doc-box-resize" onPointerDown={onStartResize} title="サイズ変更" />
        </>
      )}
    </div>
  )
}
