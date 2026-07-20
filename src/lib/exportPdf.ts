import { jsPDF } from 'jspdf'
import { rasterizeDeck, rasterizeMarkdown, type PngSlide } from './rasterize'
import type { Deck } from './deck'

const PAGE_WIDTH_IN = 10

export interface ExportOptions {
  fileName?: string
  pixelRatio?: number
  onProgress?: (done: number, total: number) => void
}

/** Rasterize Marp Markdown slides into a PDF (one image per page). */
export async function exportMarkdownToPdf(markdown: string, options: ExportOptions = {}): Promise<void> {
  const { fileName = 'slides.pdf', pixelRatio = 2, onProgress } = options
  const images = await rasterizeMarkdown(markdown, { pixelRatio, jpeg: true, onProgress })
  if (images.length === 0) {
    throw new Error('スライドが見つかりませんでした。Markdown を入力してください。')
  }
  slidesToPdf(images, fileName)
}

/** Rasterize the visual deck into a PDF (one image per page). */
export async function exportDeckToPdf(deck: Deck, options: ExportOptions = {}): Promise<void> {
  const { fileName = 'slides.pdf', pixelRatio = 2, onProgress } = options
  const images = await rasterizeDeck(deck, { pixelRatio, jpeg: true, onProgress })
  if (images.length === 0) {
    throw new Error('スライドがありません。')
  }
  slidesToPdf(images, fileName)
}

function slidesToPdf(images: PngSlide[], fileName: string): void {
  const aspect = images[0].h / images[0].w
  const wIn = PAGE_WIDTH_IN
  const hIn = Number((PAGE_WIDTH_IN * aspect).toFixed(4))
  const orientation = hIn <= wIn ? 'landscape' : 'portrait'

  const pdf = new jsPDF({ orientation, unit: 'in', format: [wIn, hIn] })
  images.forEach((img, i) => {
    if (i > 0) pdf.addPage([wIn, hIn], orientation)
    pdf.addImage(img.data, 'JPEG', 0, 0, wIn, hIn)
  })
  pdf.save(ensureExt(fileName))
}

function ensureExt(name: string): string {
  const trimmed = name.trim() || 'slides'
  return /\.pdf$/i.test(trimmed) ? trimmed : `${trimmed}.pdf`
}
