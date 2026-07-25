import * as htmlToImage from 'html-to-image'
import { renderSlides } from './marp'
import { type Deck } from './deck'
import { fillSlideElement, DECK_PX, SLIDE_BASE_STYLE, SLIDE_CODE_CSS } from './slideRender'

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
  style.textContent = SLIDE_CODE_CSS
  stage.appendChild(style)
  const slide = document.createElement('div')
  slide.style.cssText = SLIDE_BASE_STYLE
  stage.appendChild(slide)
  document.body.appendChild(stage)

  try {
    await documentFontsReady()
    const out: PngSlide[] = []
    for (let i = 0; i < deck.slides.length; i++) {
      fillSlideElement(slide, deck.slides[i])
      await waitForImages(slide)
      out.push({ data: await toDataUrl(slide, DECK_PX.w, DECK_PX.h, pixelRatio, jpeg), w: DECK_PX.w, h: DECK_PX.h })
      onProgress?.(i + 1, deck.slides.length)
    }
    return out
  } finally {
    document.body.removeChild(stage)
  }
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
