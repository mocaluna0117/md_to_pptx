import * as htmlToImage from 'html-to-image'
import PptxGen from 'pptxgenjs'
import { renderSlides } from './marp'

/** Slide width in inches; height is derived from the rendered aspect ratio. */
const SLIDE_WIDTH_IN = 10

export interface ExportOptions {
  fileName?: string
  /** Rasterization scale. 2 = crisp on HiDPI displays; higher means larger files. */
  pixelRatio?: number
  onProgress?: (done: number, total: number) => void
}

/**
 * Convert Marp Markdown into a .pptx and trigger a browser download.
 * Each slide is rendered to a PNG and placed full-bleed onto a slide, so the
 * result looks exactly like Marp's output (images are not editable text).
 */
export async function exportPptx(markdown: string, options: ExportOptions = {}): Promise<void> {
  const { fileName = 'slides.pptx', pixelRatio = 2, onProgress } = options

  const { slides, css } = renderSlides(markdown)
  if (slides.length === 0) {
    throw new Error('スライドが見つかりませんでした。Markdown を入力してください。')
  }

  // Offscreen stage that lays out one slide at a time at its native pixel size.
  const stage = document.createElement('div')
  stage.setAttribute('aria-hidden', 'true')
  stage.style.cssText = 'position:fixed;top:0;left:-100000px;pointer-events:none;'

  const style = document.createElement('style')
  style.textContent = css
  stage.appendChild(style)

  const marpit = document.createElement('div')
  marpit.className = 'marpit'
  stage.appendChild(marpit)

  document.body.appendChild(stage)

  try {
    await documentFontsReady()

    const images: Array<{ data: string; w: number; h: number }> = []
    for (let i = 0; i < slides.length; i++) {
      marpit.innerHTML = stripScripts(slides[i])
      const section = marpit.querySelector('section') as HTMLElement | null
      if (!section) continue

      const w = section.offsetWidth || 1280
      const h = section.offsetHeight || 720
      const data = await htmlToImage.toPng(section, { width: w, height: h, pixelRatio, cacheBust: true })
      images.push({ data, w, h })
      onProgress?.(i + 1, slides.length)
    }

    if (images.length === 0) {
      throw new Error('スライドの描画に失敗しました。')
    }

    const pptx = new PptxGen()
    const aspect = images[0].h / images[0].w
    const slideH = Number((SLIDE_WIDTH_IN * aspect).toFixed(4))
    pptx.defineLayout({ name: 'MARP', width: SLIDE_WIDTH_IN, height: slideH })
    pptx.layout = 'MARP'

    for (const img of images) {
      const slide = pptx.addSlide()
      slide.addImage({ data: img.data, x: 0, y: 0, w: SLIDE_WIDTH_IN, h: slideH })
    }

    await pptx.writeFile({ fileName: ensureExt(fileName) })
  } finally {
    document.body.removeChild(stage)
  }
}

/** Scripts injected via innerHTML never execute, but strip them so they can't linger. */
function stripScripts(html: string): string {
  return html.replace(/<script[\s\S]*?<\/script>/gi, '')
}

function ensureExt(name: string): string {
  const trimmed = name.trim() || 'slides'
  return /\.pptx$/i.test(trimmed) ? trimmed : `${trimmed}.pptx`
}

async function documentFontsReady(): Promise<void> {
  try {
    await document.fonts?.ready
  } catch {
    /* fonts API unavailable — proceed anyway */
  }
}
