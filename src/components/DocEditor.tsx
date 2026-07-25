import { useEffect, useRef, useState } from 'react'
import type { AttachedImages } from '../lib/imageAttach'
import { newDocBox, DEFAULT_BOX_LINE_HEIGHT, type DocBox } from '../lib/docBox'

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
  /** Rebuild the document from the current Markdown (confirms when dirty). */
  onRegenerate: () => void
  /** Header text (文書設定) shown at the sheet's top-right; repeats per page in print. */
  headerText?: string
  /** Show a live table-of-contents sheet above the document (mirrors the Word TOC). */
  tocEnabled?: boolean
  /** Show the page-number indicator at the sheet's bottom center. */
  pageNumbers?: boolean
}

/** Distance (px) within which a dragged box edge/center snaps to a guide line. */
const SNAP_PX = 6

/**
 * PowerPoint-style smart guides (same idea as Deckdown's snapMove): snap the dragged
 * box's edges/center to the page's edges & horizontal center, the page top, and other
 * boxes' edges & centers. Returns the snapped position plus guide lines to draw (px).
 */
function snapBoxMove(
  x: number,
  y: number,
  w: number,
  h: number,
  pageW: number,
  others: DocBox[],
): { x: number; y: number; v: number | null; h: number | null } {
  const xLines = [pageW / 2, 0, pageW]
  const yLines = [0]
  for (const o of others) {
    xLines.push(o.x, o.x + o.w / 2, o.x + o.w)
    yLines.push(o.y, o.y + o.h / 2, o.y + o.h)
  }
  const snap1D = (points: number[], lines: number[]) => {
    let best = SNAP_PX
    let line: number | null = null
    let point = 0
    for (const p of points) {
      for (const l of lines) {
        const d = Math.abs(p - l)
        if (d < best) {
          best = d
          line = l
          point = p
        }
      }
    }
    return line === null ? { delta: 0, guide: null as number | null } : { delta: line - point, guide: line }
  }
  const sx = snap1D([x, x + w / 2, x + w], xLines)
  const sy = snap1D([y, y + h / 2, y + h], yLines)
  return { x: x + sx.delta, y: y + sy.delta, v: sx.guide, h: sy.guide }
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

const LINE_HEIGHT_OPTIONS = [1.0, 1.15, 1.3, 1.45, 1.6, 1.8, 2.0]

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
export default function DocEditor({ html, images, onChange, boxes, onBoxesChange, onRegenerate, headerText, tocEnabled, pageNumbers }: Props) {
  const editorRef = useRef<HTMLDivElement>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const savedRange = useRef<Range | null>(null)
  const activeEditableRef = useRef<HTMLElement | null>(null)
  const emitTimer = useRef<number | null>(null)
  const boxesRef = useRef(boxes)
  boxesRef.current = boxes
  const boxBodyRefs = useRef<Map<string, HTMLDivElement>>(new Map())
  const dragRef = useRef<{ mode: 'move' | 'resize'; id: string; sx: number; sy: number; ox: number; oy: number; ow: number; oh: number; key: number } | null>(null)

  const [imgMenuOpen, setImgMenuOpen] = useState(false)
  const [selectedBoxId, setSelectedBoxId] = useState<string | null>(null)
  const [editingBoxId, setEditingBoxId] = useState<string | null>(null)
  // Smart-guide lines shown while dragging a box (positions in px, or null).
  const [guides, setGuides] = useState<{ v: number | null; h: number | null }>({ v: null, h: null })
  const [active, setActive] = useState({ bold: false, italic: false, strike: false, underline: false, block: 'p', align: 'left', inTable: false, font: '' })
  const imageNames = Object.keys(images)
  const selectedBoxIdRef = useRef(selectedBoxId)
  selectedBoxIdRef.current = selectedBoxId
  const editingBoxIdRef = useRef(editingBoxId)
  editingBoxIdRef.current = editingBoxId

  // Keyboard shortcuts, guarded so they never fire while typing in a form field
  // or contentEditable (there the browser's native undo/editing applies):
  // - Ctrl/⌘+Z / Ctrl+Y / ⌘+Shift+Z → undo/redo BOX operations
  // - Backspace/Delete → remove the selected box; Escape → deselect (mirrors Deckdown)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement | null
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return

      if ((e.metaKey || e.ctrlKey) && (e.key === 'z' || e.key === 'Z')) {
        // Only claim the shortcut when a box step exists; otherwise let the
        // browser's document-level undo handle the flow text.
        if (e.shiftKey ? boxRedo() : boxUndo()) e.preventDefault()
        return
      }
      if ((e.metaKey || e.ctrlKey) && (e.key === 'y' || e.key === 'Y')) {
        if (boxRedo()) e.preventDefault()
        return
      }

      const id = selectedBoxIdRef.current
      if (!id || editingBoxIdRef.current) return
      if (e.key === 'Backspace' || e.key === 'Delete') {
        e.preventDefault()
        deleteBox(id)
      } else if (e.key === 'Escape') {
        setSelectedBoxId(null)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

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

  // Live TOC entries (heading levels 1–3), kept in sync with the editable DOM.
  const [tocItems, setTocItems] = useState<{ level: number; text: string }[]>([])
  useEffect(() => {
    const el = editorRef.current
    if (!el || !tocEnabled) return
    let timer: number | null = null
    const refresh = () => {
      const items = Array.from(el.querySelectorAll('h1, h2, h3')).map((h) => ({
        level: Number(h.tagName[1]),
        text: h.textContent?.trim() || '（無題の見出し）',
      }))
      setTocItems(items)
    }
    refresh()
    // A MutationObserver catches every edit path (typing, execCommand, table ops).
    const mo = new MutationObserver(() => {
      if (timer) window.clearTimeout(timer)
      timer = window.setTimeout(refresh, 300)
    })
    mo.observe(el, { childList: true, characterData: true, subtree: true })
    return () => {
      mo.disconnect()
      if (timer) window.clearTimeout(timer)
    }
  }, [tocEnabled])

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
    let inTable = false
    while (node && node !== surface) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const tag = (node as HTMLElement).tagName
        if (tag === 'TD' || tag === 'TH') inTable = true
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
    let underline = false
    let font = ''
    try {
      bold = document.queryCommandState('bold')
      italic = document.queryCommandState('italic')
      strike = document.queryCommandState('strikeThrough')
      underline = document.queryCommandState('underline')
      // Reflect the selection's font in the toolbar select (first family, unquoted).
      const raw = String(document.queryCommandValue('fontName') || '')
      const first = raw.split(',')[0].replace(/['"]/g, '').trim()
      font = FONT_OPTIONS.find((o) => o.value.toLowerCase() === first.toLowerCase())?.value ?? ''
    } catch {
      /* ignore */
    }
    setActive({ bold, italic, strike, underline, block, align, inTable, font })
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

  // ---- Table structural edits (operate on the cell containing the caret) ----

  function currentCell(): HTMLTableCellElement | null {
    let node: Node | null = savedRange.current?.startContainer ?? null
    while (node) {
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
    if (activeEditableRef.current) refreshActive(activeEditableRef.current)
    syncActive()
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

  // Snapshot history for box operations (add/move/resize/edit/delete), mirroring
  // Deckdown's undoRef/redoRef + coalesceKey model: every mutation goes through
  // commitBoxes(), and mutations sharing a key (one drag, one edit session)
  // collapse into a single undo step. Flow-text undo stays native (execCommand).
  const boxUndoRef = useRef<DocBox[][]>([])
  const boxRedoRef = useRef<DocBox[][]>([])
  const lastKeyRef = useRef<number | null>(null)
  const keyCounterRef = useRef(0)
  const editSessionKeyRef = useRef<number | null>(null)

  function commitBoxes(next: DocBox[], coalesceKey?: number) {
    const sameBurst = coalesceKey != null && coalesceKey === lastKeyRef.current
    if (!sameBurst) {
      boxUndoRef.current.push(boxesRef.current)
      if (boxUndoRef.current.length > 100) boxUndoRef.current.shift()
      boxRedoRef.current = []
    }
    lastKeyRef.current = coalesceKey ?? null
    onBoxesChange(next)
  }

  /** Undo the last box operation. Returns false when there is nothing to undo. */
  function boxUndo(): boolean {
    const prev = boxUndoRef.current.pop()
    if (!prev) return false
    boxRedoRef.current.push(boxesRef.current)
    lastKeyRef.current = null
    onBoxesChange(prev)
    if (selectedBoxIdRef.current && !prev.some((b) => b.id === selectedBoxIdRef.current)) setSelectedBoxId(null)
    if (editingBoxIdRef.current && !prev.some((b) => b.id === editingBoxIdRef.current)) setEditingBoxId(null)
    return true
  }

  function boxRedo(): boolean {
    const next = boxRedoRef.current.pop()
    if (!next) return false
    boxUndoRef.current.push(boxesRef.current)
    lastKeyRef.current = null
    onBoxesChange(next)
    if (selectedBoxIdRef.current && !next.some((b) => b.id === selectedBoxIdRef.current)) setSelectedBoxId(null)
    if (editingBoxIdRef.current && !next.some((b) => b.id === editingBoxIdRef.current)) setEditingBoxId(null)
    return true
  }

  function patchBox(id: string, patch: Partial<DocBox>, coalesceKey?: number) {
    commitBoxes(boxesRef.current.map((b) => (b.id === id ? { ...b, ...patch } : b)), coalesceKey)
  }

  function addBox() {
    const box = newDocBox()
    commitBoxes([...boxesRef.current, box])
    setSelectedBoxId(box.id)
    setEditingBoxId(null)
  }

  function deleteBox(id: string) {
    commitBoxes(boxesRef.current.filter((b) => b.id !== id))
    if (selectedBoxId === id) setSelectedBoxId(null)
    if (editingBoxId === id) setEditingBoxId(null)
  }

  // NB: stopPropagation only — preventDefault on pointerdown would suppress the
  // compat mouse events, breaking double-click-to-edit and the ✕ button's click
  // (Deckdown's startDrag works the same way). Text selection during a drag is
  // prevented via CSS user-select on the static box body.
  function startBoxGesture(e: React.PointerEvent, box: DocBox, mode: 'move' | 'resize') {
    e.stopPropagation()
    // Drop focus from the page/box editable so a later Backspace deletes this box
    // instead of editing the previously focused text.
    ;(document.activeElement as HTMLElement | null)?.blur?.()
    // Suppress native selection for the whole gesture: a selection-drag anchored in
    // the flow editable would FOCUS it on pointerup, misrouting the next Ctrl+Z.
    document.body.classList.add('doc-box-gesture')
    setSelectedBoxId(box.id)
    dragRef.current = {
      mode,
      id: box.id,
      sx: e.clientX,
      sy: e.clientY,
      ox: box.x,
      oy: box.y,
      ow: box.w,
      oh: box.h,
      key: (keyCounterRef.current += 1), // one drag = one undo step
    }
    window.addEventListener('pointermove', onGestureMove)
    window.addEventListener('pointerup', onGestureUp)
  }
  function onGestureMove(e: PointerEvent) {
    const d = dragRef.current
    if (!d) return
    const dx = e.clientX - d.sx
    const dy = e.clientY - d.sy
    if (d.mode === 'move') {
      const px = Math.max(0, Math.round(d.ox + dx))
      const py = Math.max(0, Math.round(d.oy + dy))
      const box = boxesRef.current.find((b) => b.id === d.id)
      const pageW = wrapRef.current?.clientWidth ?? 760
      const others = boxesRef.current.filter((b) => b.id !== d.id)
      const snapped = snapBoxMove(px, py, box?.w ?? d.ow, box?.h ?? d.oh, pageW, others)
      patchBox(d.id, { x: Math.max(0, Math.round(snapped.x)), y: Math.max(0, Math.round(snapped.y)) }, d.key)
      setGuides({ v: snapped.v, h: snapped.h })
    } else {
      patchBox(d.id, { w: Math.max(60, Math.round(d.ow + dx)), h: Math.max(36, Math.round(d.oh + dy)) }, d.key)
    }
  }
  function onGestureUp() {
    if (dragRef.current) setGuides({ v: null, h: null })
    dragRef.current = null
    document.body.classList.remove('doc-box-gesture')
    // Belt and braces: drop any stray selection so focus stays off the editables.
    const sel = window.getSelection()
    if (sel && sel.rangeCount > 0 && (document.activeElement as HTMLElement | null)?.isContentEditable) {
      sel.removeAllRanges()
      ;(document.activeElement as HTMLElement | null)?.blur?.()
    }
    window.removeEventListener('pointermove', onGestureMove)
    window.removeEventListener('pointerup', onGestureUp)
  }

  /**
   * Toolbar undo/redo, routed like the keyboard shortcut: text typed in an
   * editable uses the browser's native undo; otherwise box operations undo first,
   * falling back to native undo when the box history is empty.
   */
  function doUndo() {
    const ae = document.activeElement as HTMLElement | null
    if (ae && (ae.isContentEditable || ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA')) {
      exec('undo')
      return
    }
    if (!boxUndo()) exec('undo')
  }
  function doRedo() {
    const ae = document.activeElement as HTMLElement | null
    if (ae && (ae.isContentEditable || ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA')) {
      exec('redo')
      return
    }
    if (!boxRedo()) exec('redo')
  }

  // data-tip (Deckdown's CSS tooltip, 0.5s delay) instead of the native title
  // attribute, whose browser-controlled delay is noticeably slower.
  const btn = (label: React.ReactNode, title: string, onClick: () => void, isActive = false, extraClass = '') => (
    <button
      type="button"
      className={`det-btn${isActive ? ' active' : ''}${extraClass ? ' ' + extraClass : ''}`}
      data-tip={title}
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
          {btn(<UndoIcon />, '元に戻す (Ctrl+Z)', doUndo)}
          {btn(<RedoIcon />, 'やり直し (Ctrl+Y)', doRedo)}
        </div>

        <div className="det-group">
          {/* selects can't render ::after tooltips — wrap them (like Deckdown's .vtipwrap) */}
          <span className="vtipwrap" data-tip="段落スタイル（見出し・引用など）">
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
          </span>
        </div>

        <div className="det-group">
          {btn('B', '太字', () => exec('bold'), active.bold, 'det-b')}
          {btn('I', '斜体', () => exec('italic'), active.italic, 'det-i')}
          {btn('S', '取り消し線', () => exec('strikeThrough'), active.strike, 'det-s')}
          {btn('U', '下線', () => exec('underline'), active.underline, 'det-u')}
        </div>

        <div className="det-group">
          {btn('A−', '文字を小さく', () => changeFontSize(-1))}
          {btn('A＋', '文字を大きく', () => changeFontSize(1))}
          <span className="vtipwrap" data-tip="フォント（選択範囲に適用）">
            <select
              className="det-select"
              aria-label="フォント"
              value={active.font}
              onMouseDown={() => restore()}
              onChange={(e) => {
                if (e.target.value) exec('fontName', e.target.value, true)
              }}
            >
              {FONT_OPTIONS.map((o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ))}
            </select>
          </span>
        </div>

        <div className="det-group det-colors">
          {COLORS.map((c) => (
            <button
              key={c}
              type="button"
              className="det-swatch"
              data-tip={`文字色 #${c}`}
              aria-label={`文字色 #${c}`}
              style={{ background: `#${c}`, ...(c === 'FFFFFF' ? { boxShadow: 'inset 0 0 0 1px #ccc' } : {}) }}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => exec('foreColor', `#${c}`, true)}
            />
          ))}
          <span className="vtipwrap" data-tip="カスタム色">
            <input
              type="color"
              className="det-color-input"
              aria-label="カスタム色"
              onMouseDown={() => restore()}
              onChange={(e) => exec('foreColor', e.target.value, true)}
            />
          </span>
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
          {active.inTable && (
            <>
              {btn('＋行', '行を追加（カーソル行の下）', addRow)}
              {btn('−行', 'カーソル行を削除', delRow)}
              {btn('＋列', '列を追加（カーソル列の右）', addCol)}
              {btn('−列', 'カーソル列を削除', delCol)}
            </>
          )}
          {btn('＋テキストボックス', 'テキストボックスを追加', addBox, false, 'det-box-add')}
          {selectedBoxId && (
            <span className="vtipwrap" data-tip="選択しているテキストボックスの行間">
              <select
                className="det-select"
                aria-label="行間"
                value={String(boxes.find((b) => b.id === selectedBoxId)?.lineHeight ?? DEFAULT_BOX_LINE_HEIGHT)}
                onChange={(e) => patchBox(selectedBoxId, { lineHeight: Number(e.target.value) })}
              >
                {LINE_HEIGHT_OPTIONS.map((v) => (
                  <option key={v} value={String(v)}>
                    行間 {v.toFixed(2).replace(/0$/, '').replace(/\.$/, '')}
                  </option>
                ))}
              </select>
            </span>
          )}
          {selectedBoxId &&
            btn(
              '🗑 選択しているテキストボックスを削除',
              '選択しているテキストボックスを削除（Backspace / Delete でも削除できます）',
              () => deleteBox(selectedBoxId),
              false,
              'det-box-add',
            )}
        </div>

        <div className="det-group det-grow">
          {btn('Markdownから作り直す', '現在のMarkdownから文書を作り直す（編集内容は破棄）', onRegenerate, false, 'det-box-add')}
        </div>
      </div>

      <div className="doc-scroll">
        {tocEnabled && (
          <div className="doc-toc-sheet" aria-label="目次プレビュー">
            <div className="doc-toc-title">目次</div>
            {tocItems.length === 0 ? (
              <p className="doc-toc-empty">見出し（h1〜h3）を追加すると項目が表示されます</p>
            ) : (
              tocItems.map((it, i) => (
                <div key={i} className={`doc-toc-item lv${it.level}`}>
                  {it.text}
                </div>
              ))
            )}
            <p className="doc-toc-note">Word ではページ番号付きの目次になります（開いて F9 で更新）</p>
          </div>
        )}
        {tocEnabled && <div className="doc-print-break" aria-hidden />}
        <div className="doc-page-wrap" ref={wrapRef}>
          {headerText && (
            <div className="doc-header-preview" aria-hidden>
              {headerText}
            </div>
          )}
          {pageNumbers && (
            <div className="doc-footer-preview" aria-hidden>
              1
            </div>
          )}
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
            {guides.v != null && <div className="doc-guide doc-guide-v" style={{ left: guides.v }} aria-hidden />}
            {guides.h != null && <div className="doc-guide doc-guide-h" style={{ top: guides.h }} aria-hidden />}
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
                  // One editing session = one undo step (typing inside still has native undo).
                  editSessionKeyRef.current = keyCounterRef.current += 1
                  setSelectedBoxId(box.id)
                  setEditingBoxId(box.id)
                }}
                onStopEdit={(nextHtml) => {
                  patchBox(box.id, { html: nextHtml }, editSessionKeyRef.current ?? undefined)
                  setEditingBoxId(null)
                  editSessionKeyRef.current = null
                }}
                onChangeHtml={(nextHtml) => patchBox(box.id, { html: nextHtml }, editSessionKeyRef.current ?? undefined)}
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
          style={{ lineHeight: box.lineHeight ?? DEFAULT_BOX_LINE_HEIGHT }}
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
        <div
          className="doc-box-body"
          style={{ lineHeight: box.lineHeight ?? DEFAULT_BOX_LINE_HEIGHT }}
          dangerouslySetInnerHTML={{ __html: box.html || '<span style="opacity:.5">ダブルクリックで編集</span>' }}
        />
      )}
      {selected && (
        <>
          {/* pointerdown must not reach the box, or a move-drag starts and swallows the click. */}
          <button
            className="doc-box-del"
            title="ボックスを削除（Backspace / Delete でも削除できます）"
            aria-label="ボックスを削除"
            onPointerDown={(e) => {
              e.stopPropagation()
              e.preventDefault()
            }}
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
              e.stopPropagation()
              onDelete()
            }}
          >
            ×
          </button>
          <span className="doc-box-resize" onPointerDown={onStartResize} title="サイズ変更" />
        </>
      )}
    </div>
  )
}
