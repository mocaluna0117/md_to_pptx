import { jsPDF } from 'jspdf'
import { rasterizeDeck, rasterizeMarkdown, rasterizeDocument, type PngSlide } from './rasterize'
import type { Deck } from './deck'
import type { DocBox } from './docBox'

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

/**
 * Rasterize the Docdown document (flowing HTML + floating boxes) into a multi-page
 * A4 PDF. The tall sheet image is placed once per page at a shifting negative offset
 * so each page shows the next slice (image-based, like the deck/slide PDFs).
 */
export async function exportHtmlToPdf(html: string, boxes: DocBox[] = [], options: ExportOptions = {}): Promise<void> {
  // A document is mostly sharp-edged text on white, so render at a high pixel ratio
  // and encode as lossless PNG — JPEG compression blurs/rings text edges.
  const { fileName = 'document.pdf', pixelRatio = 3 } = options
  const img = await rasterizeDocument(html, boxes, { pixelRatio, jpeg: false })

  const pdf = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' })
  const pageW = pdf.internal.pageSize.getWidth()
  const pageH = pdf.internal.pageSize.getHeight()
  const imgH = pageW * (img.h / img.w)

  // A shared alias makes jsPDF embed the (identical) page image only once, even when
  // it is placed on several pages, keeping multi-page PDFs small.
  const alias = 'docpage'
  let position = 0
  let heightLeft = imgH
  pdf.addImage(img.data, 'PNG', 0, position, pageW, imgH, alias, 'FAST')
  heightLeft -= pageH
  while (heightLeft > 0) {
    position -= pageH
    pdf.addPage()
    pdf.addImage(img.data, 'PNG', 0, position, pageW, imgH, alias, 'FAST')
    heightLeft -= pageH
  }
  pdf.save(ensureExt(fileName))
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
