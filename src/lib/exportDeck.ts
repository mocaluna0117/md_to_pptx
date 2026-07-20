import PptxGen from 'pptxgenjs'
import { SLIDE_W, SLIDE_H, type Box, type Deck, type TextRun } from './deck'

export interface ExportOptions {
  fileName?: string
  onProgress?: (done: number, total: number) => void
}

/** Export the visual deck to an editable .pptx (each box → a native text box). */
export async function exportDeckToPptx(deck: Deck, options: ExportOptions = {}): Promise<void> {
  const { fileName = 'slides.pptx', onProgress } = options
  if (deck.slides.length === 0) {
    throw new Error('スライドがありません。')
  }

  const pptx = new PptxGen()
  pptx.defineLayout({ name: 'W16x9', width: SLIDE_W, height: SLIDE_H })
  pptx.layout = 'W16x9'

  deck.slides.forEach((slide, idx) => {
    const s = pptx.addSlide()
    s.background = { color: slide.background || 'FFFFFF' }
    for (const box of slide.boxes) {
      s.addText(runsToTextProps(box.runs), {
        x: clamp(box.x, 0, SLIDE_W),
        y: clamp(box.y, 0, SLIDE_H),
        w: clamp(box.w, 0.2, SLIDE_W),
        h: clamp(box.h, 0.2, SLIDE_H),
        fontSize: box.fontSize,
        color: box.color || undefined,
        align: box.align,
        valign: 'top',
        fontFace: 'Arial',
      })
    }
    onProgress?.(idx + 1, deck.slides.length)
  })

  await pptx.writeFile({ fileName: ensureExt(fileName) })
}

/** Convert box runs (which may contain "\n") into pptxgenjs paragraph runs. */
function runsToTextProps(runs: Box['runs']): PptxGen.TextProps[] {
  const pieces: Array<{ run: TextRun; text: string; endsLine: boolean }> = []
  for (const run of runs) {
    const parts = run.text.split('\n')
    parts.forEach((text, i) => pieces.push({ run, text, endsLine: i < parts.length - 1 }))
  }

  const props: PptxGen.TextProps[] = pieces.map((p, i) => ({
    text: p.text,
    options: {
      bold: p.run.bold,
      italic: p.run.italic,
      color: p.run.color || undefined,
      fontSize: p.run.fontSize || undefined,
      fontFace: p.run.code ? 'Courier New' : undefined,
      breakLine: p.endsLine || i === pieces.length - 1,
    },
  }))

  return props.length ? props : [{ text: ' ' }]
}

function clamp(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, n))
}

function ensureExt(name: string): string {
  const trimmed = name.trim() || 'slides'
  return /\.pptx$/i.test(trimmed) ? trimmed : `${trimmed}.pptx`
}
