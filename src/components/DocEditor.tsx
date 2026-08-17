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
  /** Word-like page view: split the sheet visually at computed page boundaries. */
  pageView: boolean
  onPageViewChange: (v: boolean) => void
  /** Changes when something outside the DOM affects layout (e.g. the body font). */
  reflowKey?: string
}

/**
 * Flow height of one printed A4 page, in sheet pixels: 281mm printable
 * (297 − 2×8mm @page margins) at the print scale of 0.9647 (194mm / 760px),
 * minus a small safety margin — a chunk that exactly fills the paper height
 * can spill a line over sub-pixel rounding, forcing blank extra pages.
 */
const PAGE_CONTENT_H = Math.floor((281 * 96) / 25.4 / 0.9647) - 24

/** Distance (px) within which a dragged box edge/center snaps to a guide line. */
/** Caret position as a path of child indices from an editable surface's root. */
interface SelPath {
  /** null = the flowing page; otherwise the id of the text box being edited. */
  boxId: string | null
  path: number[]
  offset: number
}

/** One undo step: the whole document plus where the caret was. */
interface Snapshot {
  flow: string
  boxes: DocBox[]
  sel: SelPath | null
}

/** A CSS length is only ours to interpret when it is a percentage. */
function stylePct(v: string): number {
  return v.trim().endsWith('%') ? parseFloat(v) : NaN
}

/** A grabbable table border: an interior column boundary, or the table's own left/right edge. */
interface EdgeHit {
  table: HTMLTableElement
  kind: 'left' | 'inner' | 'right'
  /** For 'inner', the column whose right border is grabbed; else the end column. */
  index: number
}

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
export default function DocEditor({ html, images, onChange, boxes, onBoxesChange, onRegenerate, headerText, tocEnabled, pageNumbers, pageView, onPageViewChange, reflowKey }: Props) {
  const editorRef = useRef<HTMLDivElement>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const savedRange = useRef<Range | null>(null)
  const activeEditableRef = useRef<HTMLElement | null>(null)
  const emitTimer = useRef<number | null>(null)
  const boxesRef = useRef(boxes)
  boxesRef.current = boxes
  const boxBodyRefs = useRef<Map<string, HTMLDivElement>>(new Map())
  const dragRef = useRef<{ mode: 'move' | 'resize'; id: string; sx: number; sy: number; ox: number; oy: number; ow: number; oh: number; tx: Snapshot; moved?: boolean } | null>(null)
  // Column/table-edge drag state captured when the drag starts.
  const colDragRef = useRef<
    | (EdgeHit & {
        /** Document state at pointerdown, so the whole drag commits as one step. */
        tx: Snapshot
        startX: number
        /** Column widths (% of the table) at drag start. */
        widths: number[]
        /** Rendered table width (px) at drag start. */
        tableW: number
        /** Containing-block width (px) that the table's % width/margin resolve against. */
        hostW: number
        /** The table's own span at drag start, as % of the containing block. */
        widthPct: number
        indentPct: number
        /** Set once the pointer has actually moved, so a plain click changes nothing. */
        moved: boolean
      })
    | null
  >(null)

  const [imgMenuOpen, setImgMenuOpen] = useState(false)
  const [selectedBoxId, setSelectedBoxId] = useState<string | null>(null)
  const [editingBoxId, setEditingBoxId] = useState<string | null>(null)
  // Smart-guide lines shown while dragging a box (positions in px, or null).
  const [guides, setGuides] = useState<{ v: number | null; h: number | null }>({ v: null, h: null })
  const [active, setActive] = useState({ bold: false, italic: false, strike: false, underline: false, block: 'p', align: 'left', inTable: false, font: '' })
  const imageNames = Object.keys(images)
  const selectedBoxIdRef = useRef(selectedBoxId)
  selectedBoxIdRef.current = selectedBoxId
  // When each surface last changed, so Ctrl+Z can pick the most recent one.
  const fieldInputAtRef = useRef(0)
  const historyAtRef = useRef(0)
  const editingBoxIdRef = useRef(editingBoxId)
  editingBoxIdRef.current = editingBoxId

  // Undo/redo shortcuts, captured before the browser so its own (unusable) undo
  // never runs. Most-recent-action wins: a text field (the Markdown pane, the file
  // name…) keeps its native undo only while ITS typing is the latest thing that
  // happened. Otherwise Ctrl+Z steps back the document, even though the caret
  // happens to sit in a field — clicking into the Markdown pane after a table edit
  // must not strand that edit.
  useEffect(() => {
    const onFieldInput = (e: Event) => {
      const t = e.target as HTMLElement | null
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA')) fieldInputAtRef.current = Date.now()
    }
    document.addEventListener('input', onFieldInput, true)
    const onUndoKey = (e: KeyboardEvent) => {
      if (!(e.metaKey || e.ctrlKey)) return
      const t = e.target as HTMLElement | null
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA') && fieldInputAtRef.current > historyAtRef.current) {
        return // that field's own typing is the most recent action: let it undo itself
      }
      const k = e.key.toLowerCase()
      if (k === 'z') {
        e.preventDefault()
        if (e.shiftKey) redo()
        else undo()
      } else if (k === 'y') {
        e.preventDefault()
        redo()
      }
    }
    window.addEventListener('keydown', onUndoKey, true)
    return () => {
      window.removeEventListener('keydown', onUndoKey, true)
      document.removeEventListener('input', onFieldInput, true)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Box shortcuts, guarded so they never fire while typing in a field or contentEditable:
  // Backspace/Delete → remove the selected box; Escape → deselect (mirrors Deckdown)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement | null
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return

      const id = selectedBoxIdRef.current
      if (!id || editingBoxIdRef.current) return
      if (e.key === 'Backspace' || e.key === 'Delete') {
        e.preventDefault()
        deleteBox(id)
      } else if (e.key === 'Enter' || e.key === 'F2') {
        e.preventDefault()
        startEditBox(id)
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

  // ---- Word-like page view -------------------------------------------------
  // The sheet stays ONE contentEditable; non-editable spacer elements are inserted
  // at computed page boundaries so blocks visually start on the next page. Spacers
  // are stripped from every emitted/exported HTML, and in print they collapse to a
  // forced page break — so the PDF paginates exactly like the preview.

  /** The editable's HTML without the pagination spacers (what we persist/export). */
  function cleanFlowHtml(el: HTMLElement): string {
    const clone = el.cloneNode(true) as HTMLElement
    clone.querySelectorAll('.page-spacer').forEach((n) => n.remove())
    return clone.innerHTML
  }

  // Total pages in page view (drives the last page's number in the footer).
  const [pageCount, setPageCount] = useState(1)

  const paginate = () => {
    const el = editorRef.current
    if (!el) return
    el.querySelectorAll(':scope > .page-spacer').forEach((n) => n.remove())
    if (!pageView) {
      el.style.minHeight = ''
      setPageCount(1)
      return
    }
    let pageNo = 1
    let limit = PAGE_CONTENT_H
    const kids = el.children // live collection
    let i = 0
    while (i < kids.length) {
      const child = kids[i] as HTMLElement
      const top = child.offsetTop
      const h = child.offsetHeight
      const bottom = top + h
      if (bottom <= limit) {
        i += 1
        continue
      }
      if (h <= PAGE_CONTENT_H) {
        // First block of the next page — either straddling the boundary, or already
        // past it because of inter-block margins. Insert the page gap before it.
        pageNo += 1
        const spacer = document.createElement('div')
        spacer.className = 'page-spacer'
        spacer.setAttribute('contenteditable', 'false')
        // The fill closes the ENDING page — its page number (pageNo − 1) sits there.
        const num = pageNumbers ? `<div class="psp-num">${pageNo - 1}</div>` : ''
        spacer.innerHTML =
          `<div class="psp-fill" style="height:${Math.max(0, limit - top)}px">${num}</div>` +
          `<div class="psp-gap"><span>${pageNo} ページ</span></div>`
        el.insertBefore(spacer, child)
        // Fresh measure: the new page starts where the pushed block now sits.
        limit = (kids[i + 1] as HTMLElement).offsetTop + PAGE_CONTENT_H
        i += 2 // past spacer + pushed block
        continue
      }
      // Taller than a full page: leave it (it prints sliced, like Word slices
      // very long content) and advance whole pages until the boundary clears it.
      while (limit < bottom) {
        limit += PAGE_CONTENT_H
        pageNo += 1
      }
      i += 1
    }
    // Fill the last page so the sheet ends on a full-page edge.
    el.style.minHeight = `${limit}px`
    setPageCount(pageNo)
  }
  const paginateRef = useRef(paginate)
  paginateRef.current = paginate

  // Re-paginate on edits (MutationObserver), image loads, and layout-affecting props.
  useEffect(() => {
    const el = editorRef.current
    if (!el) return
    let timer: number | null = null
    const schedule = () => {
      if (timer) window.clearTimeout(timer)
      timer = window.setTimeout(() => paginateRef.current(), 250)
    }
    const isSpacer = (n: Node) => n instanceof HTMLElement && n.classList.contains('page-spacer')
    const mo = new MutationObserver((muts) => {
      // Ignore our own spacer inserts/removals or the loop never settles.
      const external = muts.some(
        (m) =>
          !(
            [...m.addedNodes, ...m.removedNodes].every(isSpacer) &&
            m.addedNodes.length + m.removedNodes.length > 0
          ) && !(m.target instanceof HTMLElement && m.target.closest('.page-spacer')),
      )
      if (external) schedule()
    })
    mo.observe(el, { childList: true, characterData: true, subtree: true })
    const onImgLoad = (e: Event) => {
      if ((e.target as HTMLElement)?.tagName === 'IMG') schedule()
    }
    el.addEventListener('load', onImgLoad, true)
    paginateRef.current()
    return () => {
      mo.disconnect()
      el.removeEventListener('load', onImgLoad, true)
      if (timer) window.clearTimeout(timer)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pageView, reflowKey, pageNumbers])

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
    withHistory(() => {
      for (const el of blocks) el.style.textAlign = dir
    })
    refreshActive(root)
    syncActive({ silent: true })
  }

  /**
   * Push the current content of whichever surface was last active back to the parent.
   * `silent` skips the history push — the caller already opened its own step.
   */
  function syncActive(opts: { silent?: boolean } = {}) {
    const surface = activeEditableRef.current
    if (!surface) return
    if (surface.classList.contains('doc-editable')) {
      onChange(cleanFlowHtml(surface))
    } else {
      const id = surface.dataset.boxId
      if (id) patchBox(id, { html: surface.innerHTML }, { silent: opts.silent })
    }
  }

  function emitFlowSoon() {
    if (emitTimer.current) window.clearTimeout(emitTimer.current)
    emitTimer.current = window.setTimeout(() => {
      if (editorRef.current) onChange(cleanFlowHtml(editorRef.current))
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
    withHistory(() => {
      try {
        document.execCommand('styleWithCSS', false, css ? 'true' : 'false')
        document.execCommand(command, false, value)
      } catch {
        /* ignore */
      }
    })
    if (activeEditableRef.current) refreshActive(activeEditableRef.current)
    syncActive({ silent: true })
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

  // ---- Column widths (drag a column border; widths live in the table's <colgroup>) ----

  /** Number of columns = widest row (tables here have no colspan). */
  function colCount(table: HTMLTableElement): number {
    return Math.max(1, ...Array.from(table.querySelectorAll('tr')).map((tr) => tr.children.length))
  }

  /**
   * Current column widths as percentages. Read from <colgroup> when present,
   * otherwise measured from the rendered first row so a drag starts from what
   * the user actually sees.
   */
  function readColWidths(table: HTMLTableElement): number[] {
    const n = colCount(table)
    const cols = Array.from(table.querySelectorAll('col'))
    if (cols.length === n) {
      const pct = cols.map((c) => parseFloat((c as HTMLElement).style.width))
      const sum = pct.reduce((a, b) => a + (Number.isFinite(b) ? b : 0), 0)
      if (pct.every((p) => Number.isFinite(p) && p > 0) && sum > 0) return pct.map((p) => (p / sum) * 100)
    }
    // Measure the widest row — the one that spans every column. A first row shortened
    // by a colspan would otherwise fall through to a fabricated equal split.
    const row = Array.from(table.querySelectorAll('tr')).reduce<HTMLTableRowElement | null>(
      (a, b) => (!a || b.children.length > a.children.length ? (b as HTMLTableRowElement) : a),
      null,
    )
    const total = table.getBoundingClientRect().width || 1
    const cells = row ? Array.from(row.children) : []
    if (cells.length === n) {
      return cells.map((c) => (c.getBoundingClientRect().width / total) * 100)
    }
    return Array(n).fill(100 / n)
  }

  /** Write widths back as a <colgroup> so they survive export and re-render. */
  function writeColWidths(table: HTMLTableElement, pct: number[]) {
    let group = table.querySelector('colgroup')
    if (!group) {
      group = document.createElement('colgroup')
      table.insertBefore(group, table.firstChild)
    }
    while (group.children.length > pct.length) group.lastElementChild?.remove()
    while (group.children.length < pct.length) group.appendChild(document.createElement('col'))
    pct.forEach((p, i) => {
      ;(group!.children[i] as HTMLElement).style.width = `${p.toFixed(3)}%`
    })
    // table-layout:fixed makes the browser honour the column widths exactly.
    table.style.tableLayout = 'fixed'
    // Only seed the default span; never clobber a width set by an outer-edge drag.
    if (!table.style.width) table.style.width = '100%'
  }

  /**
   * Keep <colgroup> in step with a structural column insert/remove. Without this the
   * grid, the cells and the table width describe three different tables — visible as a
   * hairline new column, and as a mismatched w:tblGrid in the exported .docx.
   */
  function syncColGroup(table: HTMLTableElement, index: number, mode: 'add' | 'del') {
    if (!table.querySelector('colgroup')) return // no explicit widths: nothing to sync
    const pct = Array.from(table.querySelectorAll('col')).map((c) => parseFloat((c as HTMLElement).style.width) || 0)
    if (mode === 'add') {
      // The new column splits the one it was inserted after: the table's span is unchanged.
      const half = (pct[index] ?? 100 / (pct.length + 1)) / 2
      if (pct[index] !== undefined) pct[index] = half
      pct.splice(index + 1, 0, half)
    } else {
      pct.splice(index, 1)
    }
    const sum = pct.reduce((a, b) => a + b, 0) || 1
    writeColWidths(table, pct.map((p) => (p / sum) * 100))
  }

  /** Width of the block that a table's percentage width/margin resolve against. */
  function hostWidth(table: HTMLTableElement): number {
    const host = table.parentElement as HTMLElement | null
    if (!host) return table.getBoundingClientRect().width || 1
    const cs = getComputedStyle(host)
    const w = host.clientWidth - parseFloat(cs.paddingLeft || '0') - parseFloat(cs.paddingRight || '0')
    return w > 0 ? w : table.getBoundingClientRect().width || 1
  }

  /** The table's own span: width and left offset as % of its containing block. */
  function readTableGeom(table: HTMLTableElement, hostW: number): { widthPct: number; indentPct: number } {
    // Only a percentage may be taken at face value. A length in px/pt — raw HTML in
    // the Markdown, or a table pasted from Word/Excel — would otherwise be read as
    // that many PERCENT and fling the table off the sheet, so measure it instead.
    const w = stylePct(table.style.width)
    const m = stylePct(table.style.marginLeft)
    const r = table.getBoundingClientRect()
    const host = table.parentElement
    // clientLeft is the host's left border: getBoundingClientRect() measures the border
    // box, so omitting it would report a 1px indent for a table that has none.
    const hostLeft = host
      ? host.getBoundingClientRect().left + host.clientLeft + parseFloat(getComputedStyle(host).paddingLeft || '0')
      : r.left
    return {
      widthPct: Number.isFinite(w) && w > 0 ? w : Math.min(100, (r.width / hostW) * 100),
      indentPct: Number.isFinite(m) && m > 0 ? m : Math.max(0, ((r.left - hostLeft) / hostW) * 100),
    }
  }
  function writeTableGeom(table: HTMLTableElement, widthPct: number, indentPct: number) {
    table.style.width = `${widthPct.toFixed(3)}%`
    if (indentPct > 0.001) table.style.marginLeft = `${indentPct.toFixed(3)}%`
    else table.style.removeProperty('margin-left')
  }

  /**
   * Width a column would need to show its longest cell on one line (Excel's
   * "best fit"). Measured off-screen with wrapping disabled, inside the table's own
   * parent so the probe inherits the sheet's font; explicit <br> breaks still count,
   * so the result is the widest rendered line.
   */
  function naturalColWidthPx(table: HTMLTableElement, index: number): number {
    const host = table.parentElement ?? document.body
    const probe = document.createElement('div')
    probe.style.cssText = 'position:absolute;left:-99999px;top:0;width:max-content;visibility:hidden;'
    host.appendChild(probe)
    let content = 0
    let chrome = 0
    try {
      for (const tr of Array.from(table.querySelectorAll('tr'))) {
        const cell = tr.children[index] as HTMLElement | undefined
        if (!cell) continue
        const cs = getComputedStyle(cell)
        chrome = Math.max(
          chrome,
          parseFloat(cs.paddingLeft || '0') +
            parseFloat(cs.paddingRight || '0') +
            parseFloat(cs.borderLeftWidth || '0') +
            parseFloat(cs.borderRightWidth || '0'),
        )
        const line = document.createElement('div')
        line.style.cssText = `display:inline-block;white-space:nowrap;font-family:${cs.fontFamily};font-size:${cs.fontSize};font-weight:${cs.fontWeight};font-style:${cs.fontStyle};letter-spacing:${cs.letterSpacing};`
        line.innerHTML = cell.innerHTML
        probe.appendChild(line)
        content = Math.max(content, line.getBoundingClientRect().width)
        line.remove()
      }
    } finally {
      probe.remove()
    }
    return Math.ceil(content + chrome) + 1 // +1 so the last glyph never clips
  }

  /**
   * Excel-style best fit: size the column at this border to its longest content.
   * An interior border trades with the next column (the table keeps its span); an
   * outer edge resizes the table itself, exactly as dragging that border would.
   */
  function autoFitColumn(table: HTMLTableElement, kind: EdgeHit['kind'], index: number) {
    withHistory(() => {
      const widths = readColWidths(table)
      writeColWidths(table, widths) // fixed layout, starting from what's on screen
      const tableW = table.getBoundingClientRect().width || 1
      const hostW = hostWidth(table)
      const geom = readTableGeom(table, hostW)
      const px = widths.map((p) => (p / 100) * tableW)
      const minPx = (MIN_COL_PCT / 100) * tableW
      const target = Math.max(minPx, naturalColWidthPx(table, index))
      const next = px.slice()

      if (kind === 'inner') {
        const pair = px[index] + px[index + 1]
        next[index] = Math.max(minPx, Math.min(pair - minPx, target))
        next[index + 1] = pair - next[index]
        const sum = next.reduce((a, b) => a + b, 0) || 1
        writeColWidths(table, next.map((w) => (w / sum) * 100))
        return
      }

      // Outer edge: the end column takes its ideal width and the table follows.
      const others = px.reduce((a, b) => a + b, 0) - px[index]
      const minTableW = (MIN_TABLE_PCT / 100) * hostW
      let indentPx = (geom.indentPct / 100) * hostW
      let newTableW = others + target
      if (kind === 'left') {
        const right = indentPx + (geom.widthPct / 100) * hostW
        indentPx = Math.max(0, right - newTableW)
        newTableW = right - indentPx
      }
      newTableW = Math.min(newTableW, hostW - indentPx)
      if (newTableW < minTableW) newTableW = minTableW
      indentPx = Math.min(indentPx, Math.max(0, hostW - newTableW))
      next[index] = Math.max(minPx, newTableW - others)
      let sum = next.reduce((a, b) => a + b, 0) || 1
      if (sum > newTableW) {
        const k = Math.max(0, (newTableW - next[index]) / (others || 1))
        next.forEach((w, i) => {
          if (i !== index) next[i] = w * k
        })
        sum = next.reduce((a, b) => a + b, 0) || 1
      }
      writeColWidths(table, next.map((w) => (w / sum) * 100))
      writeTableGeom(table, (newTableW / hostW) * 100, (indentPx / hostW) * 100)
    })
    if (editorRef.current) onChange(cleanFlowHtml(editorRef.current))
    paginateRef.current?.()
  }

  /** Start a resize drag: an interior column border, or the table's own left/right edge. */
  function startColResize(e: React.PointerEvent, hit: EdgeHit) {
    const { table, kind, index } = hit
    if (e.button !== 0) return // right/middle click keeps its native behaviour
    e.preventDefault()
    e.stopPropagation()
    const widths = readColWidths(table)
    if (kind === 'inner' && index >= widths.length - 1) return
    const tableW = table.getBoundingClientRect().width || 1
    const hostW = hostWidth(table)
    const geom = readTableGeom(table, hostW)
    const tx = beginTx()
    colDragRef.current = {
      tx,
      table,
      kind,
      index,
      startX: e.clientX,
      widths,
      tableW,
      hostW,
      widthPct: geom.widthPct,
      indentPct: geom.indentPct,
      moved: false,
    }
    document.body.classList.add('doc-col-resizing')
    window.addEventListener('pointermove', onColResizeMove)
    window.addEventListener('pointerup', onColResizeUp)
    window.addEventListener('pointercancel', onColResizeUp)
  }

  const MIN_COL_PCT = 5 // of the table
  const MIN_TABLE_PCT = 15 // of the containing block

  function onColResizeMove(e: PointerEvent) {
    const d = colDragRef.current
    if (!d) return
    const dx = e.clientX - d.startX
    if (!d.moved) {
      if (Math.abs(dx) < 4) return // a click with no drag must not touch the document
      d.moved = true
      // Materialize the colgroup from what was on screen at pointerdown, so
      // switching to fixed layout doesn't visibly reflow the table mid-drag.
      writeColWidths(d.table, d.widths)
    }

    if (d.kind === 'inner') {
      // Interior border: the two adjacent columns trade width, the table doesn't move.
      const deltaPct = (dx / d.tableW) * 100
      const pair = d.widths[d.index] + d.widths[d.index + 1]
      const left = Math.max(MIN_COL_PCT, Math.min(pair - MIN_COL_PCT, d.widths[d.index] + deltaPct))
      const next = d.widths.slice()
      next[d.index] = left
      next[d.index + 1] = pair - left
      writeColWidths(d.table, next)
      return
    }

    // Outer edge (Word semantics): the opposite edge stays pinned and the end
    // column nearest the grabbed edge absorbs the whole delta, so the other
    // columns keep their absolute widths.
    const px = d.widths.map((p) => (p / 100) * d.tableW)
    const end = d.kind === 'left' ? 0 : px.length - 1
    const minPx = (MIN_COL_PCT / 100) * d.tableW
    const minTableW = (MIN_TABLE_PCT / 100) * d.hostW
    const others = px.reduce((a, b) => a + b, 0) - px[end]

    // A left drag shrinks the first column as it moves right; a right drag grows the last.
    let endPx = d.kind === 'left' ? px[end] - dx : px[end] + dx
    endPx = Math.max(minPx, endPx)
    if (others + endPx < minTableW) endPx = minTableW - others

    let tableW = others + endPx
    let indentPx = (d.indentPct / 100) * d.hostW
    if (d.kind === 'left') {
      // The right edge is invariant: the indent takes exactly what the width loses.
      const right = indentPx + (d.widthPct / 100) * d.hostW
      indentPx = Math.max(0, right - tableW)
      tableW = right - indentPx // clamped by indent >= 0
    } else if (indentPx + tableW > d.hostW) {
      tableW = d.hostW - indentPx // never spill past the right margin
    }
    // Belt and braces for both kinds: the span can never exceed the sheet.
    tableW = Math.min(tableW, d.hostW - indentPx)
    if (tableW < minTableW) tableW = minTableW
    indentPx = Math.min(indentPx, Math.max(0, d.hostW - tableW))

    // Re-derive the end column from the (possibly clamped) table width. When the other
    // columns alone already fill it, scale them down rather than letting the end
    // column's minimum re-inflate the table past the clamps above.
    const nextPx = px.slice()
    nextPx[end] = Math.max(minPx, tableW - others)
    let sum = nextPx.reduce((a, b) => a + b, 0) || 1
    if (sum > tableW) {
      const k = Math.max(0, (tableW - nextPx[end]) / (others || 1))
      nextPx.forEach((w, i) => {
        if (i !== end) nextPx[i] = w * k
      })
      sum = nextPx.reduce((a, b) => a + b, 0) || 1
    }
    writeColWidths(d.table, nextPx.map((w) => (w / sum) * 100))
    // The geometry must come from the clamped tableW, never from the column sum.
    writeTableGeom(d.table, (tableW / d.hostW) * 100, (indentPx / d.hostW) * 100)
  }

  function onColResizeUp() {
    const d = colDragRef.current
    if (!d) return
    colDragRef.current = null
    document.body.classList.remove('doc-col-resizing')
    window.removeEventListener('pointermove', onColResizeMove)
    window.removeEventListener('pointerup', onColResizeUp)
    window.removeEventListener('pointercancel', onColResizeUp)
    if (!d.moved) return // a click beside a border must not rewrite the document
    // Commit the flow HTML directly: a column drag never places a caret, so
    // syncActive()'s "last active surface" may still be unset.
    if (editorRef.current) onChange(cleanFlowHtml(editorRef.current))
    commitTx(d.tx) // the whole drag is one undo step
    // Style-only mutations aren't observed, and a resized table changes row
    // heights — repaginate so page breaks stay correct.
    paginateRef.current?.()
  }

  /**
   * Cursor affordance + drag start for table borders. Bound on the editable surface
   * (tables come from user HTML, so per-cell React handlers aren't an option).
   *
   * Two tiers: inside a cell we can identify the exact column border; outside any
   * cell (the sheet padding beside a table) only the table's own rect is available —
   * which is the only way the outer edges become grabbable at all.
   */
  const COL_GRAB_IN = 5
  const COL_GRAB_OUT = 8
  function edgeHit(e: React.PointerEvent | React.MouseEvent): EdgeHit | null {
    const target = e.target as HTMLElement | null
    const cell = target?.closest('td, th') as HTMLTableCellElement | null
    if (cell) {
      const table = cell.closest('table') as HTMLTableElement | null
      if (!table) return null
      // The table's own edges win: with a colspan or a ragged row a cell's index in
      // its row is not its column index, so only the table rect identifies an edge.
      const tr = table.getBoundingClientRect()
      if (Math.abs(e.clientX - tr.right) <= COL_GRAB_IN) return { table, kind: 'right', index: colCount(table) - 1 }
      if (Math.abs(e.clientX - tr.left) <= COL_GRAB_IN) return { table, kind: 'left', index: 0 }
      const r = cell.getBoundingClientRect()
      const index = Array.from((cell.parentElement as HTMLTableRowElement).children).indexOf(cell)
      if (e.clientX >= r.right - COL_GRAB_IN && index < colCount(table) - 1) return { table, kind: 'inner', index }
      return null
    }
    // Not over a cell: hit-test the outer edges against each table's own rect.
    const editor = editorRef.current
    if (!editor) return null
    for (const t of Array.from(editor.querySelectorAll('table'))) {
      const table = t as HTMLTableElement
      const r = table.getBoundingClientRect()
      if (e.clientY < r.top || e.clientY > r.bottom) continue
      if (Math.abs(e.clientX - r.left) <= COL_GRAB_OUT) return { table, kind: 'left', index: 0 }
      if (Math.abs(e.clientX - r.right) <= COL_GRAB_OUT) return { table, kind: 'right', index: colCount(table) - 1 }
    }
    return null
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
    withHistory(() => fn(cell, row, table))
    if (activeEditableRef.current) refreshActive(activeEditableRef.current)
    syncActive({ silent: true })
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
      syncColGroup(table, index, 'add')
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
      syncColGroup(table, index, 'del')
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
    withHistory(() => {
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
    })
    syncActive({ silent: true })
  }

  // ---- Text boxes ----

  // ---- Undo / redo -------------------------------------------------------
  // ONE history for the whole document: the flowing HTML, the text boxes and the
  // caret are snapshotted together, so every action is exactly one step in
  // chronological order. The browser's own undo is intercepted and never used —
  // its stack is document-wide (an undo inside a box would rewind the page), it
  // only records execCommand edits (alignment, table edits and column drags mutate
  // the DOM directly and were silently un-undoable), and its granularity is out of
  // our hands. Typing is coalesced into bursts so it still feels native.

  const undoStackRef = useRef<Snapshot[]>([])
  const redoStackRef = useRef<Snapshot[]>([])
  /** Non-null while a typing burst is open; the key identifies the surface. */
  const burstRef = useRef<string | null>(null)
  const burstTimer = useRef<number | null>(null)
  /** True while applySnapshot rewrites the DOM, so its own events are ignored. */
  const restoringRef = useRef(false)
  const composingRef = useRef(false)
  // Re-render the toolbar when the stacks change so the buttons reflect them.
  const [histTick, setHistTick] = useState(0)

  /** Path from a surface root to the caret, skipping pagination spacers. */
  function captureSel(): SelPath | null {
    const sel = window.getSelection()
    if (!sel || sel.rangeCount === 0) return null
    const range = sel.getRangeAt(0)
    const surface = editableAncestor(range.startContainer)
    if (!surface) return null
    const path: number[] = []
    let node: Node = range.startContainer
    while (node !== surface) {
      const parent: Node | null = node.parentNode
      if (!parent) return null
      let i = 0
      for (const sib of Array.from(parent.childNodes)) {
        if (sib === node) break
        if (!(sib instanceof HTMLElement && sib.classList.contains('page-spacer'))) i++
      }
      path.unshift(i)
      node = parent
    }
    return { boxId: surface.dataset.boxId ?? null, path, offset: range.startOffset }
  }

  /** Put the caret back where captureSel() found it (best effort, never throws). */
  function applySel(sel: SelPath | null) {
    if (!sel) return
    const root = sel.boxId ? boxBodyRefs.current.get(sel.boxId) : editorRef.current
    if (!root) return
    let node: Node = root
    for (const index of sel.path) {
      const kids = Array.from(node.childNodes).filter(
        (n) => !(n instanceof HTMLElement && n.classList.contains('page-spacer')),
      )
      const next = kids[Math.min(index, kids.length - 1)]
      if (!next) break
      node = next
    }
    const max = node.nodeType === Node.TEXT_NODE ? (node.textContent ?? '').length : node.childNodes.length
    const range = document.createRange()
    try {
      range.setStart(node, Math.min(sel.offset, max))
    } catch {
      range.selectNodeContents(root)
      range.collapse(false)
    }
    range.collapse(true)
    root.focus()
    const s = window.getSelection()
    s?.removeAllRanges()
    s?.addRange(range)
    // restore()/currentCell() would otherwise hold a Range into the replaced DOM.
    savedRange.current = range.cloneRange()
    activeEditableRef.current = root
  }

  /** The document as it stands right now. */
  function readLive(): Snapshot {
    const editingId = editingBoxIdRef.current
    const liveBody = editingId ? boxBodyRefs.current.get(editingId) : undefined
    // A box being edited has its latest text only in the DOM (the commit is debounced).
    const boxes = liveBody
      ? boxesRef.current.map((b) => (b.id === editingId ? { ...b, html: liveBody.innerHTML } : b))
      : boxesRef.current
    return { flow: editorRef.current ? cleanFlowHtml(editorRef.current) : '', boxes, sel: captureSel() }
  }

  function sameDoc(a: Snapshot, b: Snapshot): boolean {
    return a.flow === b.flow && JSON.stringify(a.boxes) === JSON.stringify(b.boxes)
  }

  function pushSnapshot(snap: Snapshot) {
    historyAtRef.current = Date.now()
    undoStackRef.current.push(snap)
    if (undoStackRef.current.length > 200) undoStackRef.current.shift()
    redoStackRef.current = []
    setHistTick((t) => t + 1)
  }

  /** Run a programmatic mutation as exactly one undo step (no step if nothing changed). */
  function withHistory<T>(fn: () => T): T {
    endBurst()
    const before = readLive()
    const result = fn()
    if (!sameDoc(before, readLive())) pushSnapshot(before)
    return result
  }

  /** For gestures: capture at pointerdown, commit at pointerup only if changed. */
  function beginTx(): Snapshot {
    endBurst()
    return readLive()
  }
  function commitTx(before: Snapshot) {
    if (!sameDoc(before, readLive())) pushSnapshot(before)
  }

  // A run of typing in one surface collapses into a single step, ended by an idle
  // pause, a caret move, a click, or any other kind of action.
  const BURST_IDLE_MS = 700
  function beginBurst(key: string) {
    if (burstRef.current === key) return
    endBurst()
    pushSnapshot(readLive())
    burstRef.current = key
  }
  function touchBurst() {
    if (burstTimer.current) window.clearTimeout(burstTimer.current)
    burstTimer.current = window.setTimeout(() => {
      burstRef.current = null
    }, BURST_IDLE_MS)
  }
  function endBurst() {
    if (burstTimer.current) window.clearTimeout(burstTimer.current)
    burstTimer.current = null
    burstRef.current = null
  }

  function applySnapshot(s: Snapshot) {
    restoringRef.current = true
    try {
      const el = editorRef.current
      if (el && cleanFlowHtml(el) !== s.flow) {
        if (emitTimer.current) window.clearTimeout(emitTimer.current)
        el.innerHTML = s.flow
        onChange(s.flow)
      }
      if (JSON.stringify(boxesRef.current) !== JSON.stringify(s.boxes)) {
        boxesRef.current = s.boxes
        onBoxesChange(s.boxes)
      }
      const ids = new Set(s.boxes.map((b) => b.id))
      if (selectedBoxIdRef.current && !ids.has(selectedBoxIdRef.current)) setSelectedBoxId(null)
      if (editingBoxIdRef.current && !ids.has(editingBoxIdRef.current)) setEditingBoxId(null)
    } finally {
      // Let React commit the box list before the caret is placed / pages recomputed.
      requestAnimationFrame(() => {
        const editingId = editingBoxIdRef.current
        if (editingId) {
          const body = boxBodyRefs.current.get(editingId)
          const box = s.boxes.find((b) => b.id === editingId)
          if (body && box && body.innerHTML !== box.html) body.innerHTML = box.html
        }
        applySel(s.sel)
        paginateRef.current?.()
        restoringRef.current = false
      })
    }
  }

  function undo(): boolean {
    if (composingRef.current) return false
    endBurst()
    const prev = undoStackRef.current.pop()
    if (!prev) return false
    redoStackRef.current.push(readLive())
    historyAtRef.current = Date.now()
    applySnapshot(prev)
    setHistTick((t) => t + 1)
    return true
  }
  function redo(): boolean {
    if (composingRef.current) return false
    endBurst()
    const next = redoStackRef.current.pop()
    if (!next) return false
    undoStackRef.current.push(readLive())
    historyAtRef.current = Date.now()
    applySnapshot(next)
    setHistTick((t) => t + 1)
    return true
  }

  // Typing/IME/paste in ANY editable surface (the page or a box body) feeds the
  // same history. Native listeners: React's onBeforeInput is a synthetic polyfill
  // whose inputType is unreliable.
  useEffect(() => {
    const inSurface = (t: EventTarget | null) => (t instanceof Node ? editableAncestor(t) : null)

    const onBeforeInput = (e: Event) => {
      if (restoringRef.current || composingRef.current) return
      const ev = e as InputEvent
      const surface = inSurface(ev.target)
      if (!surface) return
      const key = `type:${surface.dataset.boxId ?? 'flow'}`
      const t = ev.inputType || ''
      if (/^(insertText|insertCompositionText|deleteContent|deleteWord|deleteSoftLineBackward)/.test(t)) {
        beginBurst(key)
        touchBurst()
      } else {
        // Paste, Enter, drag-and-drop of text, formatting… each is its own step.
        endBurst()
        pushSnapshot(readLive())
      }
    }
    const onCompStart = (e: Event) => {
      if (restoringRef.current) return
      const surface = inSurface(e.target)
      if (!surface) return
      endBurst()
      pushSnapshot(readLive())
      burstRef.current = `type:${surface.dataset.boxId ?? 'flow'}`
      composingRef.current = true
    }
    const onCompEnd = () => {
      composingRef.current = false
      touchBurst()
    }
    // A caret move or a click ends the burst, so the next keystroke starts a step.
    const onKeyDownEndBurst = (e: KeyboardEvent) => {
      if (['Enter', 'Tab', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown'].includes(e.key)) {
        endBurst()
      }
    }
    const onPointerDownEndBurst = () => endBurst()

    document.addEventListener('beforeinput', onBeforeInput, true)
    document.addEventListener('compositionstart', onCompStart, true)
    document.addEventListener('compositionend', onCompEnd, true)
    document.addEventListener('keydown', onKeyDownEndBurst, true)
    document.addEventListener('pointerdown', onPointerDownEndBurst, true)
    return () => {
      document.removeEventListener('beforeinput', onBeforeInput, true)
      document.removeEventListener('compositionstart', onCompStart, true)
      document.removeEventListener('compositionend', onCompEnd, true)
      document.removeEventListener('keydown', onKeyDownEndBurst, true)
      document.removeEventListener('pointerdown', onPointerDownEndBurst, true)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  function patchBox(id: string, patch: Partial<DocBox>, opts: { silent?: boolean } = {}) {
    const before = opts.silent ? null : readLive()
    const next = boxesRef.current.map((b) => (b.id === id ? { ...b, ...patch } : b))
    boxesRef.current = next
    if (before) pushSnapshot(before)
    onBoxesChange(next)
  }

  function addBox() {
    withHistory(() => {
      const box = newDocBox()
      boxesRef.current = [...boxesRef.current, box]
      onBoxesChange(boxesRef.current)
      setSelectedBoxId(box.id)
      setEditingBoxId(null)
    })
  }

  function deleteBox(id: string) {
    withHistory(() => {
      boxesRef.current = boxesRef.current.filter((b) => b.id !== id)
      onBoxesChange(boxesRef.current)
      if (selectedBoxId === id) setSelectedBoxId(null)
      if (editingBoxId === id) setEditingBoxId(null)
    })
  }

  /** Enter a box's text-editing mode (a mode change is not a document change). */
  function startEditBox(id: string) {
    endBurst()
    setSelectedBoxId(id)
    setEditingBoxId(id)
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
      tx: beginTx(), // the whole drag commits as one undo step
    }
    window.addEventListener('pointermove', onGestureMove)
    window.addEventListener('pointerup', onGestureUp)
  }
  function onGestureMove(e: PointerEvent) {
    const d = dragRef.current
    if (!d) return
    const dx = e.clientX - d.sx
    const dy = e.clientY - d.sy
    // Dead zone: jitter within a click must not move (or snap!) the box —
    // otherwise a slightly-wobbly double-click breaks and editing never opens.
    if (!d.moved && Math.abs(dx) < 4 && Math.abs(dy) < 4) return
    d.moved = true
    if (d.mode === 'move') {
      const px = Math.max(0, Math.round(d.ox + dx))
      const py = Math.max(0, Math.round(d.oy + dy))
      const box = boxesRef.current.find((b) => b.id === d.id)
      const pageW = wrapRef.current?.clientWidth ?? 760
      const others = boxesRef.current.filter((b) => b.id !== d.id)
      const snapped = snapBoxMove(px, py, box?.w ?? d.ow, box?.h ?? d.oh, pageW, others)
      patchBox(d.id, { x: Math.max(0, Math.round(snapped.x)), y: Math.max(0, Math.round(snapped.y)) }, { silent: true })
      setGuides({ v: snapped.v, h: snapped.h })
    } else {
      patchBox(d.id, { w: Math.max(60, Math.round(d.ow + dx)), h: Math.max(36, Math.round(d.oh + dy)) }, { silent: true })
    }
  }
  function onGestureUp() {
    const d = dragRef.current
    if (d) setGuides({ v: null, h: null })
    dragRef.current = null
    if (d?.moved) commitTx(d.tx)
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
   * Toolbar undo/redo — the same single history the keyboard shortcut drives,
   * so the button and Ctrl+Z can never disagree about what "one step" is.
   */
  function doUndo() {
    undo()
  }
  function doRedo() {
    redo()
  }

  void histTick // the stacks live in refs; this state only forces the re-render
  const canUndo = undoStackRef.current.length > 0
  const canRedo = redoStackRef.current.length > 0

  // data-tip (Deckdown's CSS tooltip, 0.5s delay) instead of the native title
  // attribute, whose browser-controlled delay is noticeably slower.
  const btn = (label: React.ReactNode, title: string, onClick: () => void, isActive = false, extraClass = '', disabled = false) => (
    <button
      type="button"
      className={`det-btn${isActive ? ' active' : ''}${extraClass ? ' ' + extraClass : ''}`}
      data-tip={title}
      aria-label={title}
      aria-pressed={isActive}
      disabled={disabled}
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
          {btn(<UndoIcon />, '元に戻す (Ctrl+Z)', doUndo, false, '', !canUndo)}
          {btn(<RedoIcon />, 'やり直し (Ctrl+Y)', doRedo, false, '', !canRedo)}
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
          {btn(
            'ページ区切り',
            'Word のようにページごとに区切って表示（印刷 PDF も同じ位置で改ページ）',
            () => onPageViewChange(!pageView),
            pageView,
            'det-box-add',
          )}
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
              {pageView ? pageCount : 1}
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
              if (editorRef.current) onChange(cleanFlowHtml(editorRef.current))
            }}
            onMouseDown={() => setSelectedBoxId(null)}
            onPointerDown={(e) => {
              const hit = edgeHit(e)
              if (hit) startColResize(e, hit)
            }}
            onDoubleClick={(e) => {
              // Excel's best fit. The two preceding pointerdowns each started a
              // resize that committed nothing (the pointer never moved), so acting
              // here is safe. Double-clicking inside a cell is untouched: edgeHit
              // only answers near a border.
              const hit = edgeHit(e)
              if (!hit) return
              e.preventDefault()
              e.stopPropagation()
              autoFitColumn(hit.table, hit.kind, hit.index)
            }}
            onMouseMove={(e) => {
              // Show the resize cursor only while hovering a draggable table border.
              const el = editorRef.current
              if (!el || colDragRef.current) return
              el.classList.toggle('col-grab', !!edgeHit(e))
            }}
            onMouseLeave={() => editorRef.current?.classList.remove('col-grab')}
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
                onEdit={() => startEditBox(box.id)}
                onStopEdit={(nextHtml) => {
                  patchBox(box.id, { html: nextHtml }, { silent: true })
                  setEditingBoxId(null)
                }}
                onChangeHtml={(nextHtml) => patchBox(box.id, { html: nextHtml }, { silent: true })}
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
          {!editing && (
            <button
              className="doc-box-edit"
              title="文字を編集（ダブルクリック / Enter でも編集できます）"
              aria-label="文字を編集"
              onPointerDown={(e) => {
                e.stopPropagation()
                e.preventDefault()
              }}
              onMouseDown={(e) => e.preventDefault()}
              onClick={(e) => {
                e.stopPropagation()
                onEdit()
              }}
            >
              ✎
            </button>
          )}
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
