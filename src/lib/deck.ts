import { parseSlides, type Para, type Run } from './markdownModel'
import { parseFrontMatter } from './frontmatter'

/** Slide dimensions in inches (16:9). All box coordinates are in inches. */
export const SLIDE_W = 10
export const SLIDE_H = 5.625

/** A styled span of text inside a box. `color` is bare hex (e.g. "FF0000"). */
export interface TextRun {
  text: string
  bold?: boolean
  italic?: boolean
  code?: boolean
  color?: string
  /** Per-run font size in points; falls back to the box's fontSize when unset. */
  fontSize?: number
}

export interface Box {
  id: string
  x: number
  y: number
  w: number
  h: number
  /** Font size in points. */
  fontSize: number
  align: 'left' | 'center' | 'right'
  /** Base text color (bare hex); individual runs may override it. */
  color?: string
  runs: TextRun[]
}

export interface Slide {
  id: string
  /** Bare hex background, e.g. "FFFFFF". */
  background: string
  boxes: Box[]
}

export interface Deck {
  slides: Slide[]
}

export function genId(): string {
  return typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2)
}

export function newBox(partial: Partial<Box> = {}): Box {
  return {
    id: genId(),
    x: 0.7,
    y: 0.7,
    w: 4,
    h: 1,
    fontSize: 18,
    align: 'left',
    runs: [{ text: 'テキスト' }],
    ...partial,
  }
}

export function newSlide(background = 'FFFFFF'): Slide {
  return { id: genId(), background, boxes: [] }
}

/** Convert "#RRGGBB" | "RRGGBB" | "rgb(r, g, b)" to bare uppercase hex, or null. */
export function toHex(input?: string | null): string | null {
  if (!input) return null
  const s = input.trim()
  const hex = s.match(/^#?([0-9a-fA-F]{6})$/)
  if (hex) return hex[1].toUpperCase()
  const rgb = s.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i)
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]]
      .map((n) => Math.max(0, Math.min(255, Number(n))).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  }
  return null
}

/** Build an editable deck from Marp Markdown (used to bootstrap the visual editor). */
export function deckFromMarkdown(markdown: string): Deck {
  const bg = toHex(parseFrontMatter(markdown).data.backgroundColor) ?? 'FFFFFF'
  const models = parseSlides(markdown)
  const source = models.length ? models : [{ paras: [], images: [], tables: [] }]

  const slides = source.map((model) => {
    const slide = newSlide(bg)
    const { title, body } = splitTitle(model.paras)
    let y = 0.5

    if (title) {
      slide.boxes.push(
        newBox({
          x: 0.5,
          y,
          w: SLIDE_W - 1,
          h: 1,
          fontSize: 32,
          runs: toRuns(title.runs, { bold: true }),
        }),
      )
      y += 1.15
    }

    const bodyRuns = parasToRuns(body)
    if (bodyRuns.length > 0) {
      slide.boxes.push(
        newBox({
          x: 0.5,
          y,
          w: SLIDE_W - 1,
          h: Math.max(1, SLIDE_H - y - 0.4),
          fontSize: 18,
          runs: bodyRuns,
        }),
      )
    }

    return slide
  })

  return { slides }
}

function splitTitle(paras: Para[]): { title: Para | null; body: Para[] } {
  let idx = paras.findIndex((p) => p.kind === 'h1')
  if (idx < 0) idx = paras.findIndex((p) => p.kind === 'h2')
  if (idx < 0) return { title: null, body: paras }
  return { title: paras[idx], body: paras.filter((_, i) => i !== idx) }
}

function toRuns(runs: Run[], extra: Partial<TextRun> = {}): TextRun[] {
  const out = runs
    .filter((r) => r.text.length)
    .map((r) => ({ text: r.text, bold: r.bold, italic: r.italic, code: r.code, ...extra }))
  return out.length ? out : [{ text: ' ', ...extra }]
}

/** Flatten body paragraphs into a single box's runs, with newlines between them. */
function parasToRuns(paras: Para[]): TextRun[] {
  const out: TextRun[] = []
  paras.forEach((p, i) => {
    if (i > 0) out.push({ text: '\n' })
    const prefix = p.kind === 'li' ? `${'  '.repeat(p.indent)}${p.ordered ? '1. ' : '• '}` : ''
    if (prefix) out.push({ text: prefix })
    const bold = p.kind.startsWith('h')
    const italic = p.kind === 'quote'
    for (const r of p.runs) {
      if (!r.text) continue
      out.push({ text: r.text, bold: r.bold || bold, italic: r.italic || italic, code: r.code || p.kind === 'code' })
    }
  })
  return out
}
