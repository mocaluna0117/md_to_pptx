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
  /** Preformatted (code block): preserve whitespace, monospace, no wrapping. */
  pre?: boolean
  runs: TextRun[]
}

/** An image placed on a slide. `src` is a data URI or URL. */
export interface ImageEl {
  id: string
  x: number
  y: number
  w: number
  h: number
  src: string
}

/** A table placed on a slide. `rows[r][c]` is plain cell text. */
export interface TableEl {
  id: string
  x: number
  y: number
  w: number
  h: number
  rows: string[][]
  /** First row is a header row (rendered bold / shaded). */
  header: boolean
  /** Font size in points. */
  fontSize: number
  /** Per-column width fractions (sum ≈ 1); falls back to equal columns. */
  colFr?: number[]
  /** Per-row height fractions (sum ≈ 1); falls back to equal rows. */
  rowFr?: number[]
}

/** Normalized per-column width fractions for a table (equal columns as fallback). */
export function tableColFractions(tb: TableEl): number[] {
  const cols = Math.max(1, ...tb.rows.map((r) => r.length))
  return normFractions(tb.colFr, cols)
}

/** Normalized per-row height fractions for a table (equal rows as fallback). */
export function tableRowFractions(tb: TableEl): number[] {
  return normFractions(tb.rowFr, Math.max(1, tb.rows.length))
}

function normFractions(fr: number[] | undefined, n: number): number[] {
  if (fr && fr.length === n) {
    const sum = fr.reduce((a, b) => a + b, 0)
    if (sum > 0) return fr.map((f) => f / sum)
  }
  return Array(n).fill(1 / n)
}

export interface Slide {
  id: string
  /** Bare hex background, e.g. "FFFFFF". */
  background: string
  boxes: Box[]
  images?: ImageEl[]
  tables?: TableEl[]
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
  return { id: genId(), background, boxes: [], images: [], tables: [] }
}

export function newTable(partial: Partial<TableEl> = {}): TableEl {
  return {
    id: genId(),
    x: 1,
    y: 1.5,
    w: 6,
    h: 1.8,
    header: true,
    fontSize: 14,
    rows: [
      ['列 1', '列 2', '列 3'],
      ['', '', ''],
      ['', '', ''],
    ],
    ...partial,
  }
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
