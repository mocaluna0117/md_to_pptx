import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type MutableRefObject,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import { SLIDE_W, SLIDE_H, boxLineHeight, genId, newBox, newSlide, newTable, tableColFractions, tableRowFractions, type Box, type Deck, type ImageEl, type Slide, type TableEl } from '../lib/deck'
import { runsToHtml, htmlToRuns } from '../lib/richText'
import type { AttachedImages } from '../lib/imageAttach'

/** Anything positioned on a slide (a text box, an image, or a table). */
interface Rect {
  id: string
  x: number
  y: number
  w: number
  h: number
}

interface Props {
  deck: Deck
  /** coalesceKey groups consecutive changes (e.g. one drag) into a single undo step. */
  onChange: (deck: Deck, coalesceKey?: number) => void
  onRegenerate: () => void
  onUndo: () => void
  onRedo: () => void
  canUndo: boolean
  canRedo: boolean
  /** Attached images (basename → data URI) offered by the insert-image menu. */
  images: AttachedImages
}

const LINE_HEIGHT_OPTIONS = [1.0, 1.15, 1.3, 1.45, 1.6, 1.8, 2.0]

const SWATCHES = ['000000', 'E03131', '1971C2', '2F9E44', 'F08C00', '7048E8', '868E96', 'FFFFFF']

// Fonts offered for a text box (value = both the CSS family and the pptx fontFace).
const FONTS: { label: string; value: string }[] = [
  { label: '標準（Arial）', value: '' },
  { label: 'Helvetica', value: 'Helvetica' },
  { label: 'Times New Roman', value: 'Times New Roman' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Courier New', value: 'Courier New' },
  { label: 'メイリオ', value: 'Meiryo' },
  { label: '游ゴシック', value: 'Yu Gothic' },
  { label: '游明朝', value: 'Yu Mincho' },
  { label: 'ＭＳ ゴシック', value: 'MS Gothic' },
  { label: 'ＭＳ 明朝', value: 'MS Mincho' },
]

// Word/PowerPoint-style alignment glyphs: stacked bars flushed to the edge.
type Align = 'left' | 'center' | 'right'
const ALIGN_BARS: Record<Align, [number, number][]> = {
  left: [[1, 14], [1, 8], [1, 14], [1, 8]],
  center: [[1, 14], [4, 8], [1, 14], [4, 8]],
  right: [[1, 14], [7, 8], [1, 14], [7, 8]],
}
const ALIGN_LABEL: Record<Align, string> = { left: '左揃え', center: '中央揃え', right: '右揃え' }

function AlignIcon({ dir }: { dir: Align }) {
  return (
    <svg viewBox="0 0 16 16" width="15" height="15" fill="currentColor" aria-hidden focusable="false">
      {ALIGN_BARS[dir].map(([x, w], i) => (
        <rect key={i} x={x} y={2.5 + i * 3.7} width={w} height="1.6" rx="0.8" />
      ))}
    </svg>
  )
}

// Curved-arrow undo/redo glyphs matching Word/Excel/PowerPoint.
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

/** Distance (px) within which a dragged edge/center snaps to a guide line. */
const SNAP_PX = 6

type SnapRect = { x: number; y: number; w: number; h: number }

/**
 * PowerPoint-style smart guides: snap the dragged element's left/center/right (and
 * top/middle/bottom) to the slide's edges & center and to other elements' edges & centers.
 * Returns the (possibly snapped) position plus the guide lines to draw (in inches, or null).
 */
function snapMove(x: number, y: number, w: number, h: number, others: SnapRect[], threshold: number) {
  const xLines = [SLIDE_W / 2, 0, SLIDE_W]
  const yLines = [SLIDE_H / 2, 0, SLIDE_H]
  for (const o of others) {
    xLines.push(o.x, o.x + o.w / 2, o.x + o.w)
    yLines.push(o.y, o.y + o.h / 2, o.y + o.h)
  }
  const snap1D = (points: number[], lines: number[]) => {
    let best = threshold
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

export default function VisualEditor({ deck, onChange, onRegenerate, onUndo, onRedo, canUndo, canRedo, images }: Props) {
  const [si, setSi] = useState(0)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [editingId, setEditingId] = useState<string | null>(null)
  // A table cell currently being edited (plain-text contentEditable).
  const [editingCell, setEditingCell] = useState<{ id: string; r: number; c: number } | null>(null)
  const [ppi, setPpi] = useState(88) // pixels per inch of the stage
  // Smart-guide lines shown while dragging (positions in inches, or null).
  const [guides, setGuides] = useState<{ v: number | null; h: number | null }>({ v: null, h: null })
  const [imgMenuOpen, setImgMenuOpen] = useState(false)
  const imgWrapRef = useRef<HTMLDivElement>(null)

  // Close the insert-image menu on an outside click.
  useEffect(() => {
    if (!imgMenuOpen) return
    const onDown = (e: PointerEvent) => {
      if (!imgWrapRef.current?.contains(e.target as Node)) setImgMenuOpen(false)
    }
    window.addEventListener('pointerdown', onDown)
    return () => window.removeEventListener('pointerdown', onDown)
  }, [imgMenuOpen])

  const stageRef = useRef<HTMLDivElement>(null)
  const railRef = useRef<HTMLDivElement>(null)
  const editRef = useRef<HTMLDivElement | null>(null)
  const savedRange = useRef<Range | null>(null)

  // Mirror latest values into refs for the persistent pointer listeners.
  const deckRef = useRef(deck)
  deckRef.current = deck
  const ppiRef = useRef(ppi)
  ppiRef.current = ppi
  const dragRef = useRef<
    null | {
      id: string
      mode: 'move' | 'resize'
      sx: number
      sy: number
      orig: { x: number; y: number; w: number; h: number }
      key: number
      moved?: boolean
    }
  >(null)
  const keyCounterRef = useRef(0)
  // Column-width / row-height drag (table id, axis, which internal boundary).
  const bandDragRef = useRef<
    null | { id: string; axis: 'col' | 'row'; index: number; s: number; base: number[]; key: number }
  >(null)

  const slideIndex = Math.min(si, deck.slides.length - 1)
  const slide = deck.slides[slideIndex]
  const siRef = useRef(slideIndex)
  siRef.current = slideIndex
  const selectedIdRef = useRef(selectedId)
  selectedIdRef.current = selectedId
  const editingIdRef = useRef(editingId)
  editingIdRef.current = editingId
  const editingCellRef = useRef(editingCell)
  editingCellRef.current = editingCell

  // Measure the stage to convert inches <-> pixels.
  useLayoutEffect(() => {
    const stage = stageRef.current
    if (!stage) return
    const measure = () => setPpi(stage.getBoundingClientRect().width / SLIDE_W)
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(stage)
    return () => ro.disconnect()
  }, [])

  // Persistent drag/resize listeners.
  useEffect(() => {
    const onMove = (e: PointerEvent) => {
      const bd = bandDragRef.current
      if (bd) {
        const tb = (deckRef.current.slides[siRef.current].tables ?? []).find((t) => t.id === bd.id)
        if (!tb) return
        const span = bd.axis === 'col' ? tb.w : tb.h
        const cur = bd.axis === 'col' ? e.clientX : e.clientY
        const dfr = (cur - bd.s) / (span * ppiRef.current)
        const i = bd.index
        const pair = bd.base[i] + bd.base[i + 1]
        const min = 0.05
        const fi = clamp(bd.base[i] + dfr, min, pair - min)
        const next = bd.base.slice()
        next[i] = fi
        next[i + 1] = pair - fi
        patchTable(bd.id, bd.axis === 'col' ? { colFr: next } : { rowFr: next }, bd.key)
        return
      }
      const d = dragRef.current
      if (!d) return
      // Dead zone: jitter within a click must not move (or snap!) the element —
      // otherwise a slightly-wobbly double-click breaks and text editing never opens.
      if (!d.moved && Math.abs(e.clientX - d.sx) < 4 && Math.abs(e.clientY - d.sy) < 4) return
      d.moved = true
      const dx = (e.clientX - d.sx) / ppiRef.current
      const dy = (e.clientY - d.sy) / ppiRef.current
      if (d.mode === 'move') {
        const px = clamp(d.orig.x + dx, 0, SLIDE_W - 0.2)
        const py = clamp(d.orig.y + dy, 0, SLIDE_H - 0.2)
        const s = deckRef.current.slides[siRef.current]
        const others: SnapRect[] = [...s.boxes, ...(s.images ?? []), ...(s.tables ?? [])]
          .filter((o) => o.id !== d.id)
          .map((o) => ({ x: o.x, y: o.y, w: o.w, h: o.h }))
        const snapped = snapMove(px, py, d.orig.w, d.orig.h, others, SNAP_PX / ppiRef.current)
        patchElement(d.id, { x: snapped.x, y: snapped.y }, d.key)
        setGuides({ v: snapped.v, h: snapped.h })
      } else {
        patchElement(d.id, { w: clamp(d.orig.w + dx, 0.3, SLIDE_W), h: clamp(d.orig.h + dy, 0.3, SLIDE_H) }, d.key)
      }
    }
    const onUp = () => {
      if (dragRef.current || bandDragRef.current) setGuides({ v: null, h: null })
      dragRef.current = null
      bandDragRef.current = null
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
    return () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Keyboard: arrow keys switch slides, Backspace/Delete removes the selected
  // box — but never while editing text or when a form field is focused.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (editingIdRef.current) return
      const t = e.target as HTMLElement | null
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return

      // Undo / redo (native undo handles text while editing, guarded above).
      if ((e.metaKey || e.ctrlKey) && (e.key === 'z' || e.key === 'Z')) {
        e.preventDefault()
        if (e.shiftKey) onRedo()
        else onUndo()
        return
      }
      if ((e.metaKey || e.ctrlKey) && (e.key === 'y' || e.key === 'Y')) {
        e.preventDefault()
        onRedo()
        return
      }

      if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
        const len = deckRef.current.slides.length
        const cur = siRef.current
        const next = e.key === 'ArrowUp' || e.key === 'ArrowLeft' ? cur - 1 : cur + 1
        const clamped = Math.max(0, Math.min(next, len - 1))
        if (clamped !== cur) {
          e.preventDefault()
          setSelectedId(null)
          setSi(clamped)
        }
        return
      }

      if (e.key === 'Backspace' || e.key === 'Delete') {
        const id = selectedIdRef.current
        if (!id) return
        e.preventDefault()
        const d = deckRef.current
        commit(
          d.slides.map((s, i) =>
            i !== siRef.current
              ? s
              : {
                  ...s,
                  boxes: s.boxes.filter((b) => b.id !== id),
                  images: (s.images ?? []).filter((im) => im.id !== id),
                  tables: (s.tables ?? []).filter((t) => t.id !== id),
                },
          ),
        )
        setSelectedId(null)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Keep the active thumbnail visible when navigating.
  useEffect(() => {
    railRef.current?.querySelector('.vthumb.active')?.scrollIntoView({ block: 'nearest' })
  }, [slideIndex])

  // Track the current text selection while editing (for the color controls).
  useEffect(() => {
    if (!editingId) return
    const onSel = () => {
      const el = editRef.current
      if (!el) return
      const sel = window.getSelection()
      if (!sel || sel.rangeCount === 0) return
      const range = sel.getRangeAt(0)
      if (el.contains(range.commonAncestorContainer)) savedRange.current = range.cloneRange()
    }
    document.addEventListener('selectionchange', onSel)
    return () => document.removeEventListener('selectionchange', onSel)
  }, [editingId])

  // ---- deck mutations ----
  function commit(slides: Deck['slides'], coalesceKey?: number) {
    onChange({ slides }, coalesceKey)
  }
  function patchBox(id: string, patch: Partial<Box>) {
    const d = deckRef.current
    commit(
      d.slides.map((s, i) =>
        i !== siRef.current ? s : { ...s, boxes: s.boxes.map((b) => (b.id === id ? { ...b, ...patch } : b)) },
      ),
    )
  }
  /** Move/resize a box, image, or table (whichever matches the id). */
  function patchElement(id: string, patch: Partial<Rect>, coalesceKey?: number) {
    const d = deckRef.current
    commit(
      d.slides.map((s, i) =>
        i !== siRef.current
          ? s
          : {
              ...s,
              boxes: s.boxes.map((b) => (b.id === id ? { ...b, ...patch } : b)),
              images: (s.images ?? []).map((im) => (im.id === id ? { ...im, ...patch } : im)),
              tables: (s.tables ?? []).map((t) => (t.id === id ? { ...t, ...patch } : t)),
            },
      ),
      coalesceKey,
    )
  }
  function patchTable(id: string, patch: Partial<TableEl>, coalesceKey?: number) {
    const d = deckRef.current
    commit(
      d.slides.map((s, i) =>
        i !== siRef.current ? s : { ...s, tables: (s.tables ?? []).map((t) => (t.id === id ? { ...t, ...patch } : t)) },
      ),
      coalesceKey,
    )
  }
  /** Write a single edited table cell back into the model (no-op if unchanged). */
  function commitCell(id: string, r: number, c: number, text: string) {
    const s = deckRef.current.slides[siRef.current]
    const tb = (s.tables ?? []).find((t) => t.id === id)
    if (!tb || tb.rows[r]?.[c] === text) return
    const rows = tb.rows.map((row, ri) => (ri !== r ? row : row.map((cell, ci) => (ci !== c ? cell : text))))
    patchTable(id, { rows })
  }
  function patchSlide(patch: Partial<Deck['slides'][number]>) {
    commit(deck.slides.map((s, i) => (i !== slideIndex ? s : { ...s, ...patch })))
  }

  function syncEditing() {
    const el = editRef.current
    if (!el || !editingId) return
    patchBox(editingId, { runs: htmlToRuns(el) })
  }
  function stopEditing() {
    syncEditing()
    setEditingId(null)
  }

  function startDrag(el: Rect, mode: 'move' | 'resize', e: ReactPointerEvent) {
    if (editingId === el.id || editingCellRef.current) return
    e.stopPropagation()
    if (editingId) stopEditing()
    setSelectedId(el.id)
    dragRef.current = {
      id: el.id,
      mode,
      sx: e.clientX,
      sy: e.clientY,
      orig: { x: el.x, y: el.y, w: el.w, h: el.h },
      key: (keyCounterRef.current += 1),
    }
  }

  function addBox() {
    const box = newBox({ x: 2, y: 2, w: 4, h: 1.2, runs: [{ text: 'テキスト' }] })
    commit(deck.slides.map((s, i) => (i !== slideIndex ? s : { ...s, boxes: [...s.boxes, box] })))
    setSelectedId(box.id)
  }
  function addTable() {
    const table = newTable()
    commit(deck.slides.map((s, i) => (i !== slideIndex ? s : { ...s, tables: [...(s.tables ?? []), table] })))
    setSelectedId(table.id)
  }

  /** Insert an attached image onto the current slide, preserving its aspect ratio. */
  function insertImage(src: string) {
    setImgMenuOpen(false)
    const probe = new Image()
    probe.onload = () => addImageEl(src, probe.naturalWidth || 4, probe.naturalHeight || 3)
    probe.onerror = () => addImageEl(src, 4, 3)
    probe.src = src
  }
  function addImageEl(src: string, nw: number, nh: number) {
    const ratio = nh / nw
    let w = 4
    let h = 4 * ratio
    if (h > SLIDE_H - 1) {
      h = SLIDE_H - 1
      w = h / ratio
    }
    if (w > SLIDE_W - 1) {
      w = SLIDE_W - 1
      h = w * ratio
    }
    const im: ImageEl = { id: genId(), x: (SLIDE_W - w) / 2, y: (SLIDE_H - h) / 2, w, h, src }
    // The async onload must not commit against a stale render-scope deck.
    const d = deckRef.current
    const si = siRef.current
    commit(d.slides.map((s, i) => (i !== si ? s : { ...s, images: [...(s.images ?? []), im] })))
    setSelectedId(im.id)
  }

  // ---- Table rows/columns (append/remove at the end; header row 0 stays put) ----

  function tableCols(tb: TableEl): number {
    return Math.max(1, ...tb.rows.map((r) => r.length))
  }
  function addTableRow(tb: TableEl) {
    setEditingCell(null)
    const cols = tableCols(tb)
    const fr = tableRowFractions(tb)
    // Appending 1/n to the normalized fractions keeps existing proportions
    // (normFractions re-normalizes on read).
    patchTable(tb.id, { rows: [...tb.rows, Array(cols).fill('')], rowFr: [...fr, 1 / fr.length] })
  }
  function deleteTableRow(tb: TableEl) {
    if (tb.rows.length <= 1) return
    setEditingCell(null)
    const rowFr = tb.rowFr && tb.rowFr.length === tb.rows.length ? tb.rowFr.slice(0, -1) : undefined
    patchTable(tb.id, { rows: tb.rows.slice(0, -1), rowFr })
  }
  function addTableCol(tb: TableEl) {
    setEditingCell(null)
    const cols = tableCols(tb)
    const rows = tb.rows.map((r) => [...r, ...Array(cols - r.length).fill(''), ''])
    const fr = tableColFractions(tb)
    patchTable(tb.id, { rows, colFr: [...fr, 1 / fr.length] })
  }
  function deleteTableCol(tb: TableEl) {
    const cols = tableCols(tb)
    if (cols <= 1) return
    setEditingCell(null)
    const rows = tb.rows.map((r) => r.slice(0, cols - 1))
    const colFr = tb.colFr && tb.colFr.length === cols ? tb.colFr.slice(0, -1) : undefined
    patchTable(tb.id, { rows, colFr })
  }
  function startCellEdit(id: string, r: number, c: number) {
    if (editingId) stopEditing()
    setSelectedId(id)
    setEditingCell({ id, r, c })
  }
  function startBandDrag(tb: TableEl, axis: 'col' | 'row', index: number, e: ReactPointerEvent) {
    e.stopPropagation()
    setSelectedId(tb.id)
    bandDragRef.current = {
      id: tb.id,
      axis,
      index,
      s: axis === 'col' ? e.clientX : e.clientY,
      base: axis === 'col' ? tableColFractions(tb) : tableRowFractions(tb),
      key: (keyCounterRef.current += 1),
    }
  }
  function deleteSelected() {
    if (!selectedId) return
    const id = selectedId
    setEditingCell(null)
    patchSlide({
      boxes: slide.boxes.filter((b) => b.id !== id),
      images: (slide.images ?? []).filter((im) => im.id !== id),
      tables: (slide.tables ?? []).filter((t) => t.id !== id),
    })
    setSelectedId(null)
  }
  function selectSlide(index: number) {
    if (editingId) stopEditing()
    setEditingCell(null)
    setSelectedId(null)
    setSi(index)
  }
  function addSlide() {
    const s = newSlide(slide?.background ?? 'FFFFFF')
    commit([...deck.slides.slice(0, slideIndex + 1), s, ...deck.slides.slice(slideIndex + 1)])
    setSelectedId(null)
    setSi(slideIndex + 1)
  }
  function deleteSlideAt(index: number) {
    if (deck.slides.length <= 1) return
    const next = index < slideIndex ? slideIndex - 1 : slideIndex
    commit(deck.slides.filter((_, i) => i !== index))
    setSelectedId(null)
    setSi(Math.max(0, Math.min(next, deck.slides.length - 2)))
  }

  /** Deep-copy a slide (fresh ids throughout) and insert it right after the original. */
  function duplicateSlideAt(index: number) {
    const src = deck.slides[index]
    if (!src) return
    const clone: Slide = JSON.parse(JSON.stringify(src))
    clone.id = genId()
    clone.boxes.forEach((b) => (b.id = genId()))
    clone.images?.forEach((im) => (im.id = genId()))
    clone.tables?.forEach((t) => (t.id = genId()))
    commit([...deck.slides.slice(0, index + 1), clone, ...deck.slides.slice(index + 1)])
    setSelectedId(null)
    setSi(index + 1)
  }

  // Drag & drop reordering of the slide rail.
  const dragSlideRef = useRef<number | null>(null)
  function dropSlideOn(target: number) {
    const from = dragSlideRef.current
    dragSlideRef.current = null
    if (from == null || from === target) return
    const slides = [...deck.slides]
    const [moved] = slides.splice(from, 1)
    slides.splice(target, 0, moved)
    commit(slides)
    setSelectedId(null)
    setSi(target)
  }

  function applyColor(hex: string) {
    const el = editRef.current
    if (!el) return
    el.focus()
    const sel = window.getSelection()
    if (savedRange.current && sel) {
      sel.removeAllRanges()
      sel.addRange(savedRange.current)
    }
    if (!sel || sel.isCollapsed) return
    document.execCommand('styleWithCSS', false, 'true')
    document.execCommand('foreColor', false, `#${hex}`)
    syncEditing()
  }

  /** Toggle an inline style on the selection using tags (<b>/<i>/<u>/<strike>), read back by htmlToRuns. */
  function applyStyle(command: 'bold' | 'italic' | 'underline' | 'strikeThrough') {
    const el = editRef.current
    if (!el) return
    el.focus()
    const sel = window.getSelection()
    if (savedRange.current && sel) {
      sel.removeAllRanges()
      sel.addRange(savedRange.current)
    }
    if (!sel || sel.isCollapsed) return
    document.execCommand('styleWithCSS', false, 'false') // emit <b>/<i>, not CSS
    document.execCommand(command)
    syncEditing()
  }

  const selectedBox = slide?.boxes.find((b) => b.id === selectedId) ?? null
  const selectedImage = slide?.images?.find((im) => im.id === selectedId) ?? null
  const selectedTable = slide?.tables?.find((t) => t.id === selectedId) ?? null
  const selectedEl: Rect | null = selectedBox ?? selectedImage ?? selectedTable ?? null
  const selectedKind = selectedBox ? 'テキストボックス' : selectedImage ? '画像' : '表'

  function changeFontSize(delta: number) {
    if (selectedTable) {
      patchTable(selectedTable.id, { fontSize: clamp(selectedTable.fontSize + delta, 8, 240) })
      return
    }
    if (!selectedBox) return
    // While editing with a non-empty selection, resize only the selected text.
    if (editingId === selectedBox.id && applyFontDeltaToSelection(delta)) return
    // Otherwise resize the whole box (base size + any explicitly-sized runs).
    patchBox(selectedBox.id, {
      fontSize: clamp(selectedBox.fontSize + delta, 8, 240),
      runs: selectedBox.runs.map((r) => (r.fontSize ? { ...r, fontSize: clamp(r.fontSize + delta, 8, 240) } : r)),
    })
  }

  function applyFontDeltaToSelection(delta: number): boolean {
    const el = editRef.current
    if (!el || !selectedBox) return false
    el.focus()
    const sel = window.getSelection()
    if (savedRange.current && sel) {
      sel.removeAllRanges()
      sel.addRange(savedRange.current)
    }
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return false
    const range = sel.getRangeAt(0)
    if (!el.contains(range.commonAncestorContainer)) return false

    const ppiNow = ppiRef.current
    const startFs = effectiveFs(range.startContainer, el, selectedBox.fontSize)

    // Wrap the selection: bump already-sized runs relative to themselves, and
    // set the wrapper's size for text that used the box's base size.
    const frag = range.extractContents()
    frag.querySelectorAll('[data-fs]').forEach((node) => {
      const n = node as HTMLElement
      setFs(n, clamp((Number(n.dataset.fs) || selectedBox.fontSize) + delta, 8, 240), ppiNow)
    })
    const wrapper = document.createElement('span')
    setFs(wrapper, clamp(startFs + delta, 8, 240), ppiNow)
    wrapper.appendChild(frag)
    range.insertNode(wrapper)

    const nr = document.createRange()
    nr.selectNodeContents(wrapper)
    sel.removeAllRanges()
    sel.addRange(nr)
    savedRange.current = nr.cloneRange()
    syncEditing()
    return true
  }

  if (!slide) return null

  return (
    <div className="veditor">
      <div className="vtoolbar">
        <div className="vtoolbar-row">
        <div className="vgroup">
          <button onClick={addSlide} data-tip="スライドを追加">＋スライド</button>
        </div>

        <div className="vgroup">
          <span className="vtipwrap" data-tip="元に戻す (Ctrl/⌘+Z)">
            <button className="vicon" onClick={onUndo} disabled={!canUndo} aria-label="元に戻す">
              <UndoIcon />
            </button>
          </span>
          <span className="vtipwrap" data-tip="やり直す (Ctrl/⌘+Shift+Z)">
            <button className="vicon" onClick={onRedo} disabled={!canRedo} aria-label="やり直す">
              <RedoIcon />
            </button>
          </span>
        </div>

        <div className="vgroup">
          <button onClick={addBox} data-tip="テキストボックスを追加">＋テキストボックス</button>
          <button onClick={addTable} data-tip="表を追加">＋表</button>
          <div className="vimg-wrap" ref={imgWrapRef}>
            <button onClick={() => setImgMenuOpen((o) => !o)} data-tip="添付した画像をスライドに挿入" aria-haspopup="menu" aria-expanded={imgMenuOpen}>
              🖼 画像を挿入
            </button>
            {imgMenuOpen && (
              <div className="vimg-menu" role="menu">
                {Object.keys(images).length === 0 ? (
                  <div className="vimg-empty">左の Markdown パネルの「🖼 画像」から画像を読み込むと挿入できます</div>
                ) : (
                  Object.keys(images).map((name) => (
                    <button key={name} type="button" role="menuitem" className="vimg-item" onMouseDown={(e) => e.preventDefault()} onClick={() => insertImage(images[name])}>
                      <img src={images[name]} alt="" />
                      <span>{name}</span>
                    </button>
                  ))
                )}
              </div>
            )}
          </div>
        </div>

        <div className="vgroup vgrow">
          <button className="vregen" onClick={onRegenerate} data-tip="現在のMarkdownからスライドを作り直す（編集内容は破棄）">
            Markdownから作り直す
          </button>
        </div>
        </div>

        {/* Context row: fixed-height so selecting/editing never reflows the stage below. */}
        <div className="vtoolbar-row vtoolbar-context">
        {!selectedEl && !editingId && (
          <span className="vctx-hint">要素を選択すると書式ツールが表示されます（ダブルクリックで文字編集）</span>
        )}
        {selectedEl && (
          <div className="vgroup">
            {(selectedBox || selectedTable) && (
              <>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => changeFontSize(-2)} data-tip={selectedTable ? '文字を小さく' : '文字を小さく（範囲選択中は選択部分のみ）'}>
                  A−
                </button>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => changeFontSize(2)} data-tip={selectedTable ? '文字を大きく' : '文字を大きく（範囲選択中は選択部分のみ）'}>
                  A＋
                </button>
              </>
            )}
            {selectedBox && (
              <select
                className="vfont"
                value={selectedBox.fontFamily || ''}
                onChange={(e) => patchBox(selectedBox.id, { fontFamily: e.target.value || undefined })}
                data-tip="フォント"
                aria-label="フォント"
              >
                {FONTS.map((f) => (
                  <option key={f.value} value={f.value} style={{ fontFamily: f.value || 'sans-serif' }}>
                    {f.label}
                  </option>
                ))}
              </select>
            )}
            {selectedBox && (
              <span className="vtipwrap" data-tip="行間">
                <select
                  className="vfont"
                  aria-label="行間"
                  value={String(boxLineHeight(selectedBox))}
                  onChange={(e) => patchBox(selectedBox.id, { lineHeight: Number(e.target.value) })}
                >
                  {/* A pre box defaults to 1.25, which isn't in the list — keep it selectable. */}
                  {!LINE_HEIGHT_OPTIONS.includes(boxLineHeight(selectedBox)) && (
                    <option value={String(boxLineHeight(selectedBox))}>行間 {boxLineHeight(selectedBox)}</option>
                  )}
                  {LINE_HEIGHT_OPTIONS.map((v) => (
                    <option key={v} value={String(v)}>
                      行間 {v.toFixed(2).replace(/0$/, '').replace(/\.$/, '')}
                    </option>
                  ))}
                </select>
              </span>
            )}
            {selectedBox &&
              (['left', 'center', 'right'] as const).map((a) => (
                <button
                  key={a}
                  onMouseDown={(e) => e.preventDefault()}
                  className={`vicon${selectedBox.align === a ? ' active' : ''}`}
                  onClick={() => patchBox(selectedBox.id, { align: a })}
                  data-tip={ALIGN_LABEL[a]}
                  aria-label={ALIGN_LABEL[a]}
                >
                  <AlignIcon dir={a} />
                </button>
              ))}
            {selectedTable && (
              <>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => addTableRow(selectedTable)} data-tip="行を追加（末尾）">
                  ＋行
                </button>
                <button
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => deleteTableRow(selectedTable)}
                  disabled={selectedTable.rows.length <= 1}
                  data-tip="末尾の行を削除"
                >
                  −行
                </button>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => addTableCol(selectedTable)} data-tip="列を追加（右端）">
                  ＋列
                </button>
                <button
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => deleteTableCol(selectedTable)}
                  disabled={tableCols(selectedTable) <= 1}
                  data-tip="右端の列を削除"
                >
                  −列
                </button>
              </>
            )}
            <button
              onClick={deleteSelected}
              data-tip={`選択している${selectedKind}を削除（Backspace / Delete でも削除できます）`}
            >
              🗑 選択している{selectedKind}を削除
            </button>
          </div>
        )}

        {editingId && (
          <div className="vgroup">
            <button
              className="vstyle"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyStyle('bold')}
              data-tip="太字（選択した文字）"
              aria-label="太字"
            >
              <b>B</b>
            </button>
            <button
              className="vstyle vstyle-i"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyStyle('italic')}
              data-tip="斜体（選択した文字）"
              aria-label="斜体"
            >
              <i>I</i>
            </button>
            <button
              className="vstyle"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyStyle('underline')}
              data-tip="下線（選択した文字）"
              aria-label="下線"
            >
              <u>U</u>
            </button>
            <button
              className="vstyle"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyStyle('strikeThrough')}
              data-tip="取り消し線（選択した文字）"
              aria-label="取り消し線"
            >
              <s>S</s>
            </button>
          </div>
        )}

        {editingId && (
          <div className="vgroup vcolors">
            <span>文字色</span>
            {SWATCHES.map((c) => (
              <button
                key={c}
                className="swatch"
                style={{ background: `#${c}`, borderColor: c === 'FFFFFF' ? '#ccc' : `#${c}` }}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => applyColor(c)}
                data-tip={`#${c}`}
                aria-label={`文字色 #${c}`}
              />
            ))}
            <input type="color" onChange={(e) => applyColor(e.target.value.slice(1))} title="カスタム色" />
          </div>
        )}

        </div>
      </div>

      <div className="vbody">
        <div className="vrail" ref={railRef} aria-label="スライド一覧">
          {deck.slides.map((s, i) => (
            <SlideThumb
              key={s.id}
              slide={s}
              index={i}
              active={i === slideIndex}
              onSelect={() => selectSlide(i)}
              onDelete={deck.slides.length > 1 ? () => deleteSlideAt(i) : undefined}
              onDuplicate={() => duplicateSlideAt(i)}
              onDragStart={() => (dragSlideRef.current = i)}
              onDropOn={() => dropSlideOn(i)}
            />
          ))}
          <button className="vaddslide" onClick={addSlide} title="スライドを追加">
            ＋ スライド
          </button>
        </div>

        <div className="vstage-wrap">
        <div
          ref={stageRef}
          className="vstage"
          style={{ aspectRatio: `${SLIDE_W} / ${SLIDE_H}`, background: `#${slide.background}` }}
          onPointerDown={(e) => {
            if (e.target === e.currentTarget) {
              if (editingId) stopEditing()
              setEditingCell(null)
              setSelectedId(null)
            }
          }}
        >
          {guides.v != null && <div className="vguide vguide-v" style={{ left: guides.v * ppi }} aria-hidden />}
          {guides.h != null && <div className="vguide vguide-h" style={{ top: guides.h * ppi }} aria-hidden />}
          {(slide.tables ?? []).map((tb) => {
            const fractions = tableColFractions(tb)
            const rowFractions = tableRowFractions(tb)
            return (
            <div
              key={tb.id}
              className={`vtable${selectedId === tb.id ? ' selected' : ''}`}
              style={{ left: tb.x * ppi, top: tb.y * ppi, width: tb.w * ppi, height: tb.h * ppi, fontSize: (tb.fontSize * ppi) / 72 }}
              onPointerDown={(e) => startDrag(tb, 'move', e)}
            >
              <table className="vtable-grid">
                <colgroup>
                  {fractions.map((f, i) => (
                    <col key={i} style={{ width: `${f * 100}%` }} />
                  ))}
                </colgroup>
                <tbody>
                  {tb.rows.map((row, r) => (
                    <tr key={r} style={{ height: `${rowFractions[r] * 100}%` }}>
                      {row.map((cell, c) => {
                        const cls = tb.header && r === 0 ? 'vth' : undefined
                        const editing = editingCell?.id === tb.id && editingCell.r === r && editingCell.c === c
                        return editing ? (
                          <EditableCell
                            key={c}
                            initial={cell}
                            className={cls}
                            onCommit={(text) => commitCell(tb.id, r, c, text)}
                            onEnd={() => setEditingCell(null)}
                          />
                        ) : (
                          <td
                            key={c}
                            className={cls}
                            onDoubleClick={(e) => {
                              e.stopPropagation()
                              startCellEdit(tb.id, r, c)
                            }}
                          >
                            {cell || ' '}
                          </td>
                        )
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
              {selectedId === tb.id &&
                fractions.slice(0, -1).map((_, i) => {
                  const boundary = fractions.slice(0, i + 1).reduce((a, b) => a + b, 0)
                  return (
                    <div
                      key={`col-${i}`}
                      className="vcolresize"
                      style={{ left: boundary * tb.w * ppi }}
                      onPointerDown={(e) => startBandDrag(tb, 'col', i, e)}
                    />
                  )
                })}
              {selectedId === tb.id &&
                rowFractions.slice(0, -1).map((_, i) => {
                  const boundary = rowFractions.slice(0, i + 1).reduce((a, b) => a + b, 0)
                  return (
                    <div
                      key={`row-${i}`}
                      className="vrowresize"
                      style={{ top: boundary * tb.h * ppi }}
                      onPointerDown={(e) => startBandDrag(tb, 'row', i, e)}
                    />
                  )
                })}
            </div>
            )
          })}
          {(slide.images ?? []).map((im) => (
            <img
              key={im.id}
              className={`vimg${selectedId === im.id ? ' selected' : ''}`}
              src={im.src}
              alt=""
              draggable={false}
              style={{ left: im.x * ppi, top: im.y * ppi, width: im.w * ppi, height: im.h * ppi }}
              onPointerDown={(e) => startDrag(im, 'move', e)}
            />
          ))}
          {slide.boxes.map((box) => {
            const style: CSSProperties = {
              left: box.x * ppi,
              top: box.y * ppi,
              width: box.w * ppi,
              height: box.h * ppi,
              fontSize: (box.fontSize * ppi) / 72,
              textAlign: box.align,
              color: box.color ? `#${box.color}` : undefined,
              fontFamily: box.fontFamily || undefined,
              lineHeight: boxLineHeight(box),
            }
            if (editingId === box.id) {
              return (
                <EditableBox
                  key={box.id}
                  box={box}
                  style={style}
                  ppi={ppi}
                  editRef={editRef}
                  onSync={syncEditing}
                  onCommit={stopEditing}
                />
              )
            }
            return (
              <div
                key={box.id}
                className={`vbox${selectedId === box.id ? ' selected' : ''}${box.pre ? ' pre' : ''}`}
                style={style}
                onPointerDown={(e) => startDrag(box, 'move', e)}
                onDoubleClick={(e) => {
                  e.stopPropagation()
                  setEditingCell(null)
                  setSelectedId(box.id)
                  setEditingId(box.id)
                }}
                dangerouslySetInnerHTML={{ __html: runsToHtml(box.runs, ppi) || '&nbsp;' }}
              />
            )
          })}
          {selectedEl && editingId !== selectedEl.id && (
            <div
              className="vresize"
              style={{ left: (selectedEl.x + selectedEl.w) * ppi - 7, top: (selectedEl.y + selectedEl.h) * ppi - 7 }}
              onPointerDown={(e) => startDrag(selectedEl, 'resize', e)}
            />
          )}
        </div>
        </div>
      </div>

      <p className="vhint">
        ↑↓←→ でスライド切替・ドラッグで移動・角をドラッグでリサイズ・ダブルクリックで文字編集・編集中に文字を選択して色変更・選択中に Backspace / Delete で削除
      </p>
    </div>
  )
}

interface SlideThumbProps {
  slide: Slide
  index: number
  active: boolean
  onSelect: () => void
  onDelete?: () => void
  onDuplicate: () => void
  onDragStart: () => void
  onDropOn: () => void
}

/** Read-only miniature of a slide for the left rail (draggable to reorder). */
function SlideThumb({ slide, index, active, onSelect, onDelete, onDuplicate, onDragStart, onDropOn }: SlideThumbProps) {
  const width = 150
  const ppi = width / SLIDE_W
  return (
    <div
      className={`vthumb${active ? ' active' : ''}`}
      onClick={onSelect}
      draggable
      onDragStart={(e) => {
        e.dataTransfer.effectAllowed = 'move'
        onDragStart()
      }}
      onDragOver={(e) => {
        e.preventDefault()
        e.dataTransfer.dropEffect = 'move'
      }}
      onDrop={(e) => {
        e.preventDefault()
        onDropOn()
      }}
      title="ドラッグで並べ替え"
    >
      <span className="vthumb-num">{index + 1}</span>
      <div
        className="vthumb-stage"
        style={{ width, height: (width * SLIDE_H) / SLIDE_W, background: `#${slide.background}` }}
      >
        {(slide.images ?? []).map((im) => (
          <img
            key={im.id}
            className="vthumb-img"
            src={im.src}
            alt=""
            style={{ left: im.x * ppi, top: im.y * ppi, width: im.w * ppi, height: im.h * ppi }}
          />
        ))}
        {slide.boxes.map((box) => (
          <div
            key={box.id}
            className={`vthumb-box${box.pre ? ' pre' : ''}`}
            style={{
              left: box.x * ppi,
              top: box.y * ppi,
              width: box.w * ppi,
              height: box.h * ppi,
              fontSize: (box.fontSize * ppi) / 72,
              textAlign: box.align,
              color: box.color ? `#${box.color}` : undefined,
              fontFamily: box.fontFamily || undefined,
              lineHeight: boxLineHeight(box),
            }}
            dangerouslySetInnerHTML={{ __html: runsToHtml(box.runs, ppi) }}
          />
        ))}
        {(slide.tables ?? []).map((tb) => (
          <div
            key={tb.id}
            className="vthumb-table"
            style={{ left: tb.x * ppi, top: tb.y * ppi, width: tb.w * ppi, height: tb.h * ppi, fontSize: (tb.fontSize * ppi) / 72 }}
          >
            <table>
              <colgroup>
                {tableColFractions(tb).map((f, i) => (
                  <col key={i} style={{ width: `${f * 100}%` }} />
                ))}
              </colgroup>
              <tbody>
                {tb.rows.map((row, r) => (
                  <tr key={r} style={{ height: `${tableRowFractions(tb)[r] * 100}%` }}>
                    {row.map((cell, c) => (
                      <td key={c} className={tb.header && r === 0 ? 'vth' : undefined}>
                        {cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ))}
      </div>
      <button
        className="vthumb-dup"
        onClick={(e) => {
          e.stopPropagation()
          onDuplicate()
        }}
        title="スライドを複製"
      >
        ⧉
      </button>
      {onDelete && (
        <button
          className="vthumb-del"
          onClick={(e) => {
            e.stopPropagation()
            onDelete()
          }}
          title="スライドを削除"
        >
          ×
        </button>
      )}
    </div>
  )
}

interface EditableBoxProps {
  box: Box
  style: CSSProperties
  ppi: number
  editRef: MutableRefObject<HTMLDivElement | null>
  onSync: () => void
  onCommit: () => void
}

function EditableBox({ box, style, ppi, editRef, onSync, onCommit }: EditableBoxProps) {
  const ref = useRef<HTMLDivElement>(null)
  const syncRef = useRef(onSync)
  syncRef.current = onSync

  useEffect(() => {
    const el = ref.current
    if (!el) return
    editRef.current = el
    el.innerHTML = runsToHtml(box.runs, ppi)
    el.focus()
    const range = document.createRange()
    range.selectNodeContents(el)
    range.collapse(false)
    const sel = window.getSelection()
    sel?.removeAllRanges()
    sel?.addRange(range)
    return () => {
      // Commit the in-progress text before the box is torn down (e.g. tab switch).
      syncRef.current()
      if (editRef.current === el) editRef.current = null
    }
    // Set initial HTML once when editing starts.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  return (
    <div
      ref={ref}
      className={`vbox editing${box.pre ? ' pre' : ''}`}
      style={style}
      contentEditable
      suppressContentEditableWarning
      onPointerDown={(e) => e.stopPropagation()}
      onBlur={onSync}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.preventDefault()
          onCommit()
        }
      }}
    />
  )
}

interface EditableCellProps {
  initial: string
  className?: string
  onCommit: (text: string) => void
  onEnd: () => void
}

/** A single table cell edited as plain text (commits on blur / unmount). */
function EditableCell({ initial, className, onCommit, onEnd }: EditableCellProps) {
  const ref = useRef<HTMLTableCellElement>(null)
  const onCommitRef = useRef(onCommit)
  onCommitRef.current = onCommit

  useEffect(() => {
    const el = ref.current
    if (!el) return
    el.textContent = initial
    el.focus()
    const range = document.createRange()
    range.selectNodeContents(el)
    range.collapse(false)
    const sel = window.getSelection()
    sel?.removeAllRanges()
    sel?.addRange(range)
    return () => onCommitRef.current(el.textContent ?? '')
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  return (
    <td
      ref={ref}
      className={className}
      contentEditable
      suppressContentEditableWarning
      onPointerDown={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
      onBlur={(e) => {
        onCommitRef.current(e.currentTarget.textContent ?? '')
        onEnd()
      }}
      onKeyDown={(e) => {
        if (e.key === 'Escape' || (e.key === 'Enter' && !e.shiftKey)) {
          e.preventDefault()
          ;(e.currentTarget as HTMLElement).blur()
        }
      }}
    />
  )
}

function clamp(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, n))
}

/** Store a font size (points) on an element as data-fs plus a scaled px style. */
function setFs(el: HTMLElement, pt: number, ppi: number): void {
  el.dataset.fs = String(pt)
  el.style.fontSize = `${((pt * ppi) / 72).toFixed(2)}px`
}

/** Effective font size (points) at a node: nearest ancestor data-fs, else base. */
function effectiveFs(node: Node, root: HTMLElement, base: number): number {
  let el: HTMLElement | null = node.nodeType === Node.TEXT_NODE ? node.parentElement : (node as HTMLElement)
  while (el && el !== root) {
    if (el.dataset?.fs) return Number(el.dataset.fs)
    el = el.parentElement
  }
  return base
}
