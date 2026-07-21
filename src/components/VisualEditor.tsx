import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type MutableRefObject,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import { SLIDE_W, SLIDE_H, newBox, newSlide, type Box, type Deck, type Slide } from '../lib/deck'
import { runsToHtml, htmlToRuns } from '../lib/richText'

/** Anything positioned on a slide (a text box or an image). */
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
}

const SWATCHES = ['000000', 'E03131', '1971C2', '2F9E44', 'F08C00', '7048E8', '868E96', 'FFFFFF']

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

export default function VisualEditor({ deck, onChange, onRegenerate, onUndo, onRedo, canUndo, canRedo }: Props) {
  const [si, setSi] = useState(0)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [ppi, setPpi] = useState(88) // pixels per inch of the stage

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
    }
  >(null)
  const keyCounterRef = useRef(0)

  const slideIndex = Math.min(si, deck.slides.length - 1)
  const slide = deck.slides[slideIndex]
  const siRef = useRef(slideIndex)
  siRef.current = slideIndex
  const selectedIdRef = useRef(selectedId)
  selectedIdRef.current = selectedId
  const editingIdRef = useRef(editingId)
  editingIdRef.current = editingId

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
      const d = dragRef.current
      if (!d) return
      const dx = (e.clientX - d.sx) / ppiRef.current
      const dy = (e.clientY - d.sy) / ppiRef.current
      if (d.mode === 'move') {
        patchElement(d.id, { x: clamp(d.orig.x + dx, 0, SLIDE_W - 0.2), y: clamp(d.orig.y + dy, 0, SLIDE_H - 0.2) }, d.key)
      } else {
        patchElement(d.id, { w: clamp(d.orig.w + dx, 0.3, SLIDE_W), h: clamp(d.orig.h + dy, 0.3, SLIDE_H) }, d.key)
      }
    }
    const onUp = () => {
      dragRef.current = null
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
  /** Move/resize a box or an image (whichever matches the id). */
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
            },
      ),
      coalesceKey,
    )
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
    if (editingId === el.id) return
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
  function deleteSelected() {
    if (!selectedId) return
    const id = selectedId
    patchSlide({
      boxes: slide.boxes.filter((b) => b.id !== id),
      images: (slide.images ?? []).filter((im) => im.id !== id),
    })
    setSelectedId(null)
  }
  function selectSlide(index: number) {
    if (editingId) stopEditing()
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

  const selectedBox = slide?.boxes.find((b) => b.id === selectedId) ?? null
  const selectedImage = slide?.images?.find((im) => im.id === selectedId) ?? null
  const selectedEl: Rect | null = selectedBox ?? selectedImage ?? null

  function changeFontSize(delta: number) {
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
        </div>

        {selectedEl && (
          <div className="vgroup">
            {selectedBox && (
              <>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => changeFontSize(-2)} data-tip="文字を小さく（範囲選択中は選択部分のみ）">
                  A−
                </button>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => changeFontSize(2)} data-tip="文字を大きく（範囲選択中は選択部分のみ）">
                  A＋
                </button>
                {(['left', 'center', 'right'] as const).map((a) => (
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
              </>
            )}
            <button
              onClick={deleteSelected}
              data-tip={
                selectedBox
                  ? '選択しているテキストボックスを削除（Backspace / Delete でも削除できます）'
                  : '選択している画像を削除（Backspace / Delete でも削除できます）'
              }
            >
              🗑 {selectedBox ? '選択しているテキストボックスを削除' : '選択している画像を削除'}
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

        <div className="vgroup vgrow">
          <button className="vregen" onClick={onRegenerate} data-tip="現在のMarkdownからスライドを作り直す（編集内容は破棄）">
            Markdownから作り直す
          </button>
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
              setSelectedId(null)
            }
          }}
        >
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
                className={`vbox${selectedId === box.id ? ' selected' : ''}`}
                style={style}
                onPointerDown={(e) => startDrag(box, 'move', e)}
                onDoubleClick={(e) => {
                  e.stopPropagation()
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
}

/** Read-only miniature of a slide for the left rail. */
function SlideThumb({ slide, index, active, onSelect, onDelete }: SlideThumbProps) {
  const width = 150
  const ppi = width / SLIDE_W
  return (
    <div className={`vthumb${active ? ' active' : ''}`} onClick={onSelect}>
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
            className="vthumb-box"
            style={{
              left: box.x * ppi,
              top: box.y * ppi,
              width: box.w * ppi,
              height: box.h * ppi,
              fontSize: (box.fontSize * ppi) / 72,
              textAlign: box.align,
              color: box.color ? `#${box.color}` : undefined,
            }}
            dangerouslySetInnerHTML={{ __html: runsToHtml(box.runs, ppi) }}
          />
        ))}
      </div>
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
      className="vbox editing"
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
