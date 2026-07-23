import * as htmlToImage from 'html-to-image'
import { renderSlides } from './marp'
import { SLIDE_W, SLIDE_H, tableColFractions, tableRowFractions, type Deck, type TableEl } from './deck'
import { runsToHtml } from './richText'
import type { DocBox } from './docBox'

export interface PngSlide {
  data: string
  w: number
  h: number
}

export type ProgressFn = (done: number, total: number) => void

export interface RasterizeOptions {
  pixelRatio?: number
  /** Encode as JPEG instead of PNG (much smaller for PDFs; slides are opaque). */
  jpeg?: boolean
  onProgress?: ProgressFn
}

async function toDataUrl(el: HTMLElement, w: number, h: number, pixelRatio: number, jpeg: boolean): Promise<string> {
  const common = { width: w, height: h, pixelRatio, cacheBust: true }
  // JPEG has no alpha, so it needs an opaque backdrop. html-to-image's
  // `backgroundColor` both fills the canvas AND is written as an inline style on
  // the captured node, so it must be the element's OWN background — otherwise it
  // overrides a custom slide background (e.g. a color set in the visual editor).
  return jpeg
    ? htmlToImage.toJpeg(el, { ...common, quality: 0.92, backgroundColor: effectiveBg(el) })
    : htmlToImage.toPng(el, common)
}

/** The element's computed background color, or white if it's transparent. */
function effectiveBg(el: HTMLElement): string {
  const c = getComputedStyle(el).backgroundColor
  if (!c || c === 'transparent' || c === 'rgba(0, 0, 0, 0)') return '#ffffff'
  return c
}

/** Marp renders each slide at 1280x720; SLIDE_W inches maps to 1280px. */
const PX_PER_IN = 1280 / SLIDE_W
const DECK_W = 1280
const DECK_H = Math.round((1280 * SLIDE_H) / SLIDE_W)

/** Rasterize each Marp slide (rendered from Markdown) to an image. */
export async function rasterizeMarkdown(markdown: string, options: RasterizeOptions = {}): Promise<PngSlide[]> {
  const { pixelRatio = 2, jpeg = false, onProgress } = options
  const { slides, css } = renderSlides(markdown)
  if (slides.length === 0) return []

  const stage = createStage()
  const style = document.createElement('style')
  style.textContent = css
  stage.appendChild(style)
  const marpit = document.createElement('div')
  marpit.className = 'marpit'
  stage.appendChild(marpit)
  document.body.appendChild(stage)

  try {
    await documentFontsReady()
    const out: PngSlide[] = []
    for (let i = 0; i < slides.length; i++) {
      marpit.innerHTML = stripScripts(slides[i])
      const section = marpit.querySelector('section') as HTMLElement | null
      if (!section) continue
      await waitForImages(section)
      const w = section.offsetWidth || 1280
      const h = section.offsetHeight || 720
      out.push({ data: await toDataUrl(section, w, h, pixelRatio, jpeg), w, h })
      onProgress?.(i + 1, slides.length)
    }
    return out
  } finally {
    document.body.removeChild(stage)
  }
}

/** Rasterize the visual deck (boxes + images) to an image per slide. */
export async function rasterizeDeck(deck: Deck, options: RasterizeOptions = {}): Promise<PngSlide[]> {
  const { pixelRatio = 2, jpeg = false, onProgress } = options
  if (deck.slides.length === 0) return []

  const stage = createStage()
  const style = document.createElement('style')
  style.textContent = 'code{font-family:ui-monospace,Menlo,Consolas,monospace}'
  stage.appendChild(style)
  const slide = document.createElement('div')
  slide.style.cssText = `position:relative;width:${DECK_W}px;height:${DECK_H}px;overflow:hidden;font-family:Arial,"Noto Sans CJK JP","Yu Gothic",sans-serif;`
  stage.appendChild(slide)
  document.body.appendChild(stage)

  try {
    await documentFontsReady()
    const out: PngSlide[] = []
    for (let i = 0; i < deck.slides.length; i++) {
      const s = deck.slides[i]
      slide.style.background = `#${s.background || 'FFFFFF'}`
      slide.innerHTML = ''
      for (const tb of s.tables ?? []) {
        const wrap = document.createElement('div')
        wrap.style.cssText = pos(tb) + `overflow:hidden;box-sizing:border-box;font-size:${(tb.fontSize * PX_PER_IN) / 72}px;`
        wrap.innerHTML = tableHtml(tb)
        slide.appendChild(wrap)
      }
      for (const im of s.images ?? []) {
        const img = document.createElement('img')
        img.src = im.src
        img.style.cssText = pos(im) + 'object-fit:fill;'
        slide.appendChild(img)
      }
      for (const box of s.boxes) {
        const el = document.createElement('div')
        const fontFamily = box.pre
          ? 'ui-monospace,Menlo,Consolas,monospace'
          : box.fontFamily
            ? `"${box.fontFamily}"`
            : 'inherit'
        el.style.cssText =
          pos(box) +
          `overflow:hidden;padding:4px 6px;box-sizing:border-box;word-break:break-word;` +
          `white-space:${box.pre ? 'pre' : 'pre-wrap'};line-height:${box.pre ? 1.25 : 1.3};` +
          `font-size:${(box.fontSize * PX_PER_IN) / 72}px;text-align:${box.align};font-family:${fontFamily};` +
          `color:${box.color ? `#${box.color}` : '#111'};`
        el.innerHTML = runsToHtml(box.runs, PX_PER_IN) || '&nbsp;'
        slide.appendChild(el)
      }
      await waitForImages(slide)
      out.push({ data: await toDataUrl(slide, DECK_W, DECK_H, pixelRatio, jpeg), w: DECK_W, h: DECK_H })
      onProgress?.(i + 1, deck.slides.length)
    }
    return out
  } finally {
    document.body.removeChild(stage)
  }
}

/** The Docdown page sheet width in px (matches the editor's max-width). */
const DOC_W = 760

/**
 * Rasterize the Docdown document (flowing HTML + floating text boxes) into one tall
 * image, styled to match the on-screen page but without editor chrome. The caller
 * paginates it into PDF pages. Relies on the app's `.doc-page` CSS being loaded.
 */
export async function rasterizeDocument(html: string, boxes: DocBox[], options: RasterizeOptions = {}): Promise<PngSlide> {
  const { pixelRatio = 2, jpeg = true } = options
  const stage = createStage()
  const page = document.createElement('div')
  page.className = 'doc-page'
  // Fixed width, no shadow/border/rounding — a clean printable sheet.
  page.style.cssText = `position:relative;width:${DOC_W}px;max-width:none;margin:0;box-shadow:none;border:none;border-radius:0;`
  page.innerHTML = html
  for (const b of boxes) {
    const el = document.createElement('div')
    el.style.cssText =
      `position:absolute;left:${b.x}px;top:${b.y}px;width:${b.w}px;height:${b.h}px;` +
      `overflow:hidden;box-sizing:border-box;padding:8px 10px;line-height:1.6;`
    el.innerHTML = b.html || ''
    page.appendChild(el)
  }
  stage.appendChild(page)
  document.body.appendChild(stage)

  try {
    await documentFontsReady()
    await waitForImages(page)
    const w = page.offsetWidth || DOC_W
    const h = page.offsetHeight || DOC_W
    return { data: await toDataUrl(page, w, h, pixelRatio, jpeg), w, h }
  } finally {
    document.body.removeChild(stage)
  }
}

function pos(r: { x: number; y: number; w: number; h: number }): string {
  return `position:absolute;left:${r.x * PX_PER_IN}px;top:${r.y * PX_PER_IN}px;width:${r.w * PX_PER_IN}px;height:${r.h * PX_PER_IN}px;`
}

/** Static HTML for a table (mirrors the visual editor's table rendering). */
function tableHtml(tb: TableEl): string {
  const fr = tableColFractions(tb)
  const rowFr = tableRowFractions(tb)
  const cols = fr.length
  const colGroup = fr.map((f) => `<col style="width:${(f * 100).toFixed(3)}%">`).join('')
  const body = tb.rows
    .map((row, r) => {
      const header = tb.header && r === 0
      const cells = Array.from({ length: cols }, (_, c) => {
        const cell = escapeHtml(row[c] ?? '') || '&nbsp;'
        const style =
          `border:1px solid #cfd6e0;padding:2px 6px;color:#111;vertical-align:middle;overflow:hidden;` +
          `word-break:break-word;line-height:1.25;${header ? 'font-weight:700;background:#eef2f7;' : ''}`
        return `<td style="${style}">${cell}</td>`
      }).join('')
      return `<tr style="height:${(rowFr[r] * 100).toFixed(3)}%">${cells}</tr>`
    })
    .join('')
  return `<table style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed;background:#fff"><colgroup>${colGroup}</colgroup><tbody>${body}</tbody></table>`
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function createStage(): HTMLDivElement {
  const stage = document.createElement('div')
  stage.setAttribute('aria-hidden', 'true')
  stage.style.cssText = 'position:fixed;top:0;left:-100000px;pointer-events:none;'
  return stage
}

function stripScripts(html: string): string {
  return html.replace(/<script[\s\S]*?<\/script>/gi, '')
}

async function waitForImages(root: HTMLElement): Promise<void> {
  const imgs = Array.from(root.querySelectorAll('img'))
  await Promise.all(
    imgs.map(
      (img) =>
        new Promise<void>((resolve) => {
          if (img.complete && img.naturalWidth > 0) return resolve()
          const done = () => resolve()
          img.addEventListener('load', done, { once: true })
          img.addEventListener('error', done, { once: true })
          setTimeout(done, 2000)
        }),
    ),
  )
}

async function documentFontsReady(): Promise<void> {
  try {
    await document.fonts?.ready
  } catch {
    /* fonts API unavailable — proceed anyway */
  }
}
