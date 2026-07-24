import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { Slide } from '../lib/deck'
import { fillSlideElement, DECK_PX, SLIDE_BASE_STYLE, SLIDE_CODE_CSS } from '../lib/rasterize'

interface Props {
  slides: Slide[]
  startIndex?: number
  onClose: () => void
}

/**
 * PowerPoint-style full-screen slideshow. Renders the deck live (crisp, vector-ish) by
 * reusing the same slide rendering as the rasterizer/export, scaled to fit the viewport.
 * Advance with click / →/Space, go back with ←, exit with Esc. Enters OS fullscreen when
 * allowed; either way the fixed overlay covers the window.
 */
export default function Slideshow({ slides, startIndex = 0, onClose }: Props) {
  const [index, setIndex] = useState(Math.min(Math.max(0, startIndex), Math.max(0, slides.length - 1)))
  const [scale, setScale] = useState(1)
  const [controls, setControls] = useState(true)
  const rootRef = useRef<HTMLDivElement>(null)
  const slideRef = useRef<HTMLDivElement>(null)
  const hideTimer = useRef<number | null>(null)

  const total = slides.length
  const next = useCallback(() => setIndex((i) => Math.min(total - 1, i + 1)), [total])
  const prev = useCallback(() => setIndex((i) => Math.max(0, i - 1)), [])

  // Render the current slide into the (natively sized) stage element.
  useLayoutEffect(() => {
    const el = slideRef.current
    if (!el || !slides[index]) return
    el.style.cssText = SLIDE_BASE_STYLE
    fillSlideElement(el, slides[index])
  }, [index, slides])

  // Fit the 1280×720 stage to the viewport.
  useLayoutEffect(() => {
    const fit = () => setScale(Math.min(window.innerWidth / DECK_PX.w, window.innerHeight / DECK_PX.h))
    fit()
    window.addEventListener('resize', fit)
    return () => window.removeEventListener('resize', fit)
  }, [])

  // Keyboard navigation.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      switch (e.key) {
        case 'ArrowRight':
        case 'ArrowDown':
        case ' ':
        case 'Enter':
        case 'PageDown':
          e.preventDefault()
          next()
          break
        case 'ArrowLeft':
        case 'ArrowUp':
        case 'PageUp':
        case 'Backspace':
          e.preventDefault()
          prev()
          break
        case 'Home':
          e.preventDefault()
          setIndex(0)
          break
        case 'End':
          e.preventDefault()
          setIndex(total - 1)
          break
        case 'Escape':
          e.preventDefault()
          onClose()
          break
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [next, prev, total, onClose])

  // Enter fullscreen if allowed; closing/leaving fullscreen ends the show.
  useEffect(() => {
    const root = rootRef.current
    root?.requestFullscreen?.().catch(() => {})
    const onFs = () => {
      if (!document.fullscreenElement) onClose()
    }
    document.addEventListener('fullscreenchange', onFs)
    return () => {
      document.removeEventListener('fullscreenchange', onFs)
      if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {})
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Auto-hide the controls a moment after the pointer stops moving.
  const bumpControls = useCallback(() => {
    setControls(true)
    if (hideTimer.current) window.clearTimeout(hideTimer.current)
    hideTimer.current = window.setTimeout(() => setControls(false), 2500)
  }, [])
  useEffect(() => {
    bumpControls()
    return () => {
      if (hideTimer.current) window.clearTimeout(hideTimer.current)
    }
  }, [bumpControls, index])

  const stop = (e: React.MouseEvent) => e.stopPropagation()

  return (
    <div
      ref={rootRef}
      className={`slideshow${controls ? '' : ' hide-cursor'}`}
      onClick={next}
      onMouseMove={bumpControls}
      role="presentation"
    >
      <style>{SLIDE_CODE_CSS}</style>
      <div className="show-stage" style={{ width: DECK_PX.w, height: DECK_PX.h, transform: `scale(${scale})` }}>
        <div ref={slideRef} />
      </div>

      <div className={`show-controls${controls ? '' : ' hidden'}`} onClick={stop}>
        <button className="show-btn" onClick={prev} disabled={index === 0} aria-label="前のスライド" title="前へ (←)">
          ‹
        </button>
        <span className="show-count">
          {index + 1} / {total}
        </span>
        <button className="show-btn" onClick={next} disabled={index === total - 1} aria-label="次のスライド" title="次へ (→)">
          ›
        </button>
        <button className="show-btn show-exit" onClick={onClose} aria-label="終了" title="終了 (Esc)">
          ✕
        </button>
      </div>
    </div>
  )
}
