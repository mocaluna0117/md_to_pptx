import PptxGen from 'pptxgenjs'
import { SLIDE_W, SLIDE_H, tableColFractions, tableRowFractions, type Box, type Deck, type TextRun } from './deck'

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
    for (const im of slide.images ?? []) {
      const pos = {
        x: clamp(im.x, 0, SLIDE_W),
        y: clamp(im.y, 0, SLIDE_H),
        w: clamp(im.w, 0.1, SLIDE_W),
        h: clamp(im.h, 0.1, SLIDE_H),
      }
      s.addImage(im.src.startsWith('data:') ? { data: im.src, ...pos } : { path: im.src, ...pos })
    }
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
    for (const tb of slide.tables ?? []) {
      const cols = Math.max(1, ...tb.rows.map((r) => r.length))
      const rows: PptxGen.TableRow[] = tb.rows.map((row, r) => {
        const isHeader = tb.header && r === 0
        return Array.from({ length: cols }, (_, c) => ({
          text: row[c] ?? '',
          options: isHeader ? { bold: true, fill: { color: 'EEF2F7' } } : {},
        }))
      })
      const w = clamp(tb.w, 0.5, SLIDE_W)
      const h = clamp(tb.h, 0.3, SLIDE_H)
      const fr = tableColFractions(tb)
      const rowFr = tableRowFractions(tb)
      s.addTable(rows, {
        x: clamp(tb.x, 0, SLIDE_W),
        y: clamp(tb.y, 0, SLIDE_H),
        w,
        h,
        colW: fr.map((f) => f * w),
        rowH: rowFr.map((f) => f * h),
        fontSize: tb.fontSize,
        fontFace: 'Arial',
        color: '111111',
        valign: 'middle',
        border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
        autoPage: false,
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
