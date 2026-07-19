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

interface Props {
  deck: Deck
  onChange: (deck: Deck) => void
  onRegenerate: () => void
}

const SWATCHES = ['000000', 'E03131', '1971C2', '2F9E44', 'F08C00', '7048E8', '868E96', 'FFFFFF']

export default function VisualEditor({ deck, onChange, onRegenerate }: Props) {
  const [si, setSi] = useState(0)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [ppi, setPpi] = useState(88) // pixels per inch of the stage

  const stageRef = useRef<HTMLDivElement>(null)
  const editRef = useRef<HTMLDivElement | null>(null)
  const savedRange = useRef<Range | null>(null)

  // Mirror latest values into refs for the persistent pointer listeners.
  const deckRef = useRef(deck)
  deckRef.current = deck
  const ppiRef = useRef(ppi)
  ppiRef.current = ppi
  const dragRef = useRef<null | { id: string; mode: 'move' | 'resize'; sx: number; sy: number; orig: Box }>(null)

  const slideIndex = Math.min(si, deck.slides.length - 1)
  const slide = deck.slides[slideIndex]
  const siRef = useRef(slideIndex)
  siRef.current = slideIndex

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
        patchBox(d.id, {
          x: clamp(d.orig.x + dx, 0, SLIDE_W - 0.2),
          y: clamp(d.orig.y + dy, 0, SLIDE_H - 0.2),
        })
      } else {
        patchBox(d.id, {
          w: clamp(d.orig.w + dx, 0.4, SLIDE_W),
          h: clamp(d.orig.h + dy, 0.3, SLIDE_H),
        })
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
  function commit(slides: Deck['slides']) {
    onChange({ slides })
  }
  function patchBox(id: string, patch: Partial<Box>) {
    const d = deckRef.current
    commit(
      d.slides.map((s, i) =>
        i !== siRef.current ? s : { ...s, boxes: s.boxes.map((b) => (b.id === id ? { ...b, ...patch } : b)) },
      ),
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

  function startDrag(box: Box, mode: 'move' | 'resize', e: ReactPointerEvent) {
    if (editingId === box.id) return
    e.stopPropagation()
    if (editingId) stopEditing()
    setSelectedId(box.id)
    dragRef.current = { id: box.id, mode, sx: e.clientX, sy: e.clientY, orig: box }
  }

  function addBox() {
    const box = newBox({ x: 2, y: 2, w: 4, h: 1.2, runs: [{ text: 'テキスト' }] })
    commit(deck.slides.map((s, i) => (i !== slideIndex ? s : { ...s, boxes: [...s.boxes, box] })))
    setSelectedId(box.id)
  }
  function deleteBox() {
    if (!selectedId) return
    patchSlide({ boxes: slide.boxes.filter((b) => b.id !== selectedId) })
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

  const selected = slide?.boxes.find((b) => b.id === selectedId) ?? null

  if (!slide) return null

  return (
    <div className="veditor">
      <div className="vtoolbar">
        <div className="vgroup">
          <label className="vfield" title="背景色">
            背景
            <input
              type="color"
              value={`#${slide.background}`}
              onChange={(e) => patchSlide({ background: e.target.value.slice(1).toUpperCase() })}
            />
          </label>
          <button onClick={addBox} title="テキストボックスを追加">＋ボックス</button>
        </div>

        {selected && (
          <div className="vgroup">
            <button onMouseDown={(e) => e.preventDefault()} onClick={() => patchBox(selected.id, { fontSize: Math.max(8, selected.fontSize - 2) })} title="文字を小さく">
              A−
            </button>
            <button onMouseDown={(e) => e.preventDefault()} onClick={() => patchBox(selected.id, { fontSize: selected.fontSize + 2 })} title="文字を大きく">
              A＋
            </button>
            {(['left', 'center', 'right'] as const).map((a) => (
              <button
                key={a}
                onMouseDown={(e) => e.preventDefault()}
                className={selected.align === a ? 'active' : ''}
                onClick={() => patchBox(selected.id, { align: a })}
                title={`揃え: ${a}`}
              >
                {a === 'left' ? '⤙' : a === 'center' ? '≡' : '⤚'}
              </button>
            ))}
            <button onClick={deleteBox} title="ボックスを削除">🗑ボックス</button>
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
                title={`#${c}`}
              />
            ))}
            <input type="color" onChange={(e) => applyColor(e.target.value.slice(1))} title="カスタム色" />
          </div>
        )}

        <div className="vgroup vgrow">
          <button className="vregen" onClick={onRegenerate} title="現在のMarkdownからスライドを作り直す（編集内容は破棄）">
            Markdownから作り直す
          </button>
        </div>
      </div>

      <div className="vbody">
        <div className="vrail" aria-label="スライド一覧">
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
          {slide.boxes.map((box) => {
            const style: CSSProperties = {
              left: box.x * ppi,
              top: box.y * ppi,
              width: box.w * ppi,
              height: box.h * ppi,
              fontSize: (box.fontSize * ppi) / 72,
              textAlign: box.align,
            }
            if (editingId === box.id) {
              return <EditableBox key={box.id} box={box} style={style} editRef={editRef} onCommit={stopEditing} />
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
                dangerouslySetInnerHTML={{ __html: runsToHtml(box.runs) || '&nbsp;' }}
              />
            )
          })}
          {selected && editingId !== selected.id && (
            <div
              className="vresize"
              style={{ left: (selected.x + selected.w) * ppi - 7, top: (selected.y + selected.h) * ppi - 7 }}
              onPointerDown={(e) => startDrag(selected, 'resize', e)}
            />
          )}
        </div>
        </div>
      </div>

      <p className="vhint">
        ドラッグで移動・角をドラッグでリサイズ・ダブルクリックで文字編集・編集中に文字を選択して色変更
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
            }}
            dangerouslySetInnerHTML={{ __html: runsToHtml(box.runs) }}
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
  editRef: MutableRefObject<HTMLDivElement | null>
  onCommit: () => void
}

function EditableBox({ box, style, editRef, onCommit }: EditableBoxProps) {
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const el = ref.current
    if (!el) return
    editRef.current = el
    el.innerHTML = runsToHtml(box.runs)
    el.focus()
    const range = document.createRange()
    range.selectNodeContents(el)
    range.collapse(false)
    const sel = window.getSelection()
    sel?.removeAllRanges()
    sel?.addRange(range)
    return () => {
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
