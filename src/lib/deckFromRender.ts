import { renderSlides } from './marp'
import {
  SLIDE_W,
  deckFromMarkdown,
  genId,
  toHex,
  type Box,
  type Deck,
  type ImageEl,
  type Slide,
  type TableEl,
  type TextRun,
} from './deck'

// Marp renders each slide at 1280x720; SLIDE_W inches maps to 1280px.
const PX_PER_IN = 1280 / SLIDE_W
const pxToIn = (px: number) => px / PX_PER_IN
const pxToPt = (px: number) => (px / PX_PER_IN) * 72
const round = (n: number) => Math.round(n * 1000) / 1000

const BLOCK_TAGS = new Set(['P', 'DIV', 'BLOCKQUOTE', 'PRE', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'FIGURE', 'HEADER', 'FOOTER'])

/**
 * Build an editable deck from the *rendered* Marp slides so themes / custom CSS
 * (backgrounds, colors, font sizes, bold, inline HTML) carry over to the boxes.
 * Falls back to the plain Markdown parse if rendering/measuring fails.
 */
export async function deckFromRenderedMarkdown(markdown: string): Promise<Deck> {
  try {
    const { slides, css } = renderSlides(markdown)
    if (slides.length === 0) return { slides: [emptySlide()] }

    const stage = document.createElement('div')
    stage.setAttribute('aria-hidden', 'true')
    stage.style.cssText = 'position:fixed;top:0;left:-100000px;'
    const styleEl = document.createElement('style')
    styleEl.textContent = css
    stage.appendChild(styleEl)
    const marpit = document.createElement('div')
    marpit.className = 'marpit'
    stage.appendChild(marpit)
    document.body.appendChild(stage)

    try {
      await documentFontsReady()
      const out: Slide[] = []
      for (const html of slides) {
        marpit.innerHTML = html.replace(/<script[\s\S]*?<\/script>/gi, '')
        const section = marpit.querySelector('section')
        if (section) {
          await waitForImages(section as HTMLElement)
          out.push(slideFromSection(section as HTMLElement))
        }
      }
      return { slides: out.length ? out : [emptySlide()] }
    } finally {
      document.body.removeChild(stage)
    }
  } catch {
    return deckFromMarkdown(markdown)
  }
}

function slideFromSection(section: HTMLElement): Slide {
  const secRect = section.getBoundingClientRect()
  const background = toHex(getComputedStyle(section).backgroundColor) ?? 'FFFFFF'
  const boxes: Box[] = []
  for (const child of Array.from(section.children)) {
    const box = boxFromBlock(child as HTMLElement, secRect)
    if (box) boxes.push(box)
  }
  return {
    id: genId(),
    background,
    boxes,
    images: imagesFromSection(section, secRect),
    tables: tablesFromSection(section, secRect),
  }
}

/** Extract each rendered <table> as an editable TableEl (rows of plain text). */
function tablesFromSection(section: HTMLElement, secRect: DOMRect): TableEl[] {
  const out: TableEl[] = []
  for (const table of Array.from(section.querySelectorAll('table'))) {
    const rect = table.getBoundingClientRect()
    if (rect.width < 1 || rect.height < 1) continue

    const cellText = (c: Element) => (c.textContent ?? '').replace(/\s+/g, ' ').trim()
    const rows: string[][] = []
    const headCells = Array.from(table.querySelectorAll(':scope > thead th, :scope > thead td'))
    const header = headCells.length > 0
    if (header) rows.push(headCells.map(cellText))
    for (const tr of Array.from(table.querySelectorAll(':scope > tbody > tr, :scope > tr'))) {
      const cells = Array.from(tr.children).map(cellText)
      if (cells.length) rows.push(cells)
    }
    if (rows.length === 0) continue

    const trEls = Array.from(table.querySelectorAll(':scope > thead > tr, :scope > tbody > tr, :scope > tr'))

    // Column width fractions from the rendered first row.
    let colFr: number[] | undefined
    if (trEls[0]) {
      const widths = Array.from(trEls[0].children).map((c) => c.getBoundingClientRect().width)
      const sum = widths.reduce((a, b) => a + b, 0)
      if (sum > 0) colFr = widths.map((w) => round(w / sum))
    }
    // Row height fractions from the rendered rows.
    let rowFr: number[] | undefined
    if (trEls.length === rows.length) {
      const heights = trEls.map((tr) => tr.getBoundingClientRect().height)
      const sum = heights.reduce((a, b) => a + b, 0)
      if (sum > 0) rowFr = heights.map((h) => round(h / sum))
    }

    const fontSize = round(pxToPt(parseFloat(getComputedStyle(table).fontSize) || 18))
    out.push({
      id: genId(),
      x: round(pxToIn(rect.left - secRect.left)),
      y: round(pxToIn(rect.top - secRect.top)),
      w: round(pxToIn(rect.width)),
      h: round(pxToIn(rect.height)),
      rows,
      header,
      fontSize,
      colFr,
      rowFr,
    })
  }
  return out
}

function imagesFromSection(section: HTMLElement, secRect: DOMRect): ImageEl[] {
  const out: ImageEl[] = []
  for (const img of Array.from(section.querySelectorAll('img'))) {
    const src = img.getAttribute('src') || img.currentSrc || ''
    if (!src) continue
    const rect = img.getBoundingClientRect()
    if (rect.width < 1 || rect.height < 1) continue
    out.push({
      id: genId(),
      x: round(pxToIn(rect.left - secRect.left)),
      y: round(pxToIn(rect.top - secRect.top)),
      w: round(pxToIn(rect.width)),
      h: round(pxToIn(rect.height)),
      src,
    })
  }
  return out
}

async function waitForImages(root: HTMLElement): Promise<void> {
  const imgs = Array.from(root.querySelectorAll('img'))
  await Promise.all(
    imgs.map(
      (img) =>
        new Promise<void>((resolve) => {
          if (img.complete && img.naturalWidth > 0) {
            resolve()
            return
          }
          const done = () => resolve()
          img.addEventListener('load', done, { once: true })
          img.addEventListener('error', done, { once: true })
          setTimeout(done, 2000)
        }),
    ),
  )
}

function boxFromBlock(el: HTMLElement, secRect: DOMRect): Box | null {
  // Tables are handled separately as editable TableEls, not text boxes.
  if (el.tagName === 'HR' || el.tagName === 'TABLE') return null
  const rect = el.getBoundingClientRect()
  if (rect.width < 1 || rect.height < 1) return null

  const pos = {
    x: round(pxToIn(rect.left - secRect.left)),
    y: round(pxToIn(rect.top - secRect.top)),
    w: round(pxToIn(rect.width)),
    h: round(pxToIn(rect.height)),
  }

  // Code block: keep whitespace/newlines verbatim, monospace, no wrapping.
  if (el.tagName === 'PRE') {
    const raw = (el.textContent ?? '').replace(/\n+$/, '')
    if (!raw.trim()) return null
    const pcs = getComputedStyle(el)
    return {
      id: genId(),
      ...pos,
      fontSize: round(pxToPt(parseFloat(pcs.fontSize) || 20)),
      align: 'left',
      color: toHex(pcs.color) ?? '111111',
      pre: true,
      runs: [{ text: raw, code: true }],
    }
  }

  const cs = getComputedStyle(el)
  const base = { color: toHex(cs.color) ?? '000000', size: round(pxToPt(parseFloat(cs.fontSize) || 24)) }
  const runs = extractRuns(el, base)
  if (runs.length === 0) return null // e.g. image-only blocks (not yet supported)

  const align = cs.textAlign === 'center' ? 'center' : cs.textAlign === 'right' ? 'right' : 'left'
  return {
    id: genId(),
    ...pos,
    fontSize: base.size,
    align,
    color: base.color,
    runs,
  }
}

interface Base {
  color: string
  size: number
}

function extractRuns(root: HTMLElement, base: Base): TextRun[] {
  const runs: TextRun[] = []

  const pushPlain = (text: string) => {
    if (text) runs.push({ text })
  }
  const ensureNL = () => {
    if (runs.length && !runs[runs.length - 1].text.endsWith('\n')) runs.push({ text: '\n' })
  }
  const pushStyled = (raw: string, parent: HTMLElement | null) => {
    const text = raw.replace(/\s+/g, ' ')
    if (!text) return
    const last = runs[runs.length - 1]
    if (text === ' ' && (!last || /[\s\n]$/.test(last.text))) return
    const cs = parent ? getComputedStyle(parent) : null
    const color = cs ? toHex(cs.color) ?? base.color : base.color
    const size = cs ? round(pxToPt(parseFloat(cs.fontSize) || 24)) : base.size
    const weight = cs ? parseInt(cs.fontWeight, 10) : 400
    runs.push({
      text,
      color: color !== base.color ? color : undefined,
      fontSize: size !== base.size ? size : undefined,
      bold: weight >= 600 || cs?.fontWeight === 'bold' || undefined,
      italic: cs?.fontStyle === 'italic' || undefined,
      code: !!parent?.closest('code, pre') || undefined,
    })
  }

  const walk = (node: Node, olCtr: { n: number } | null, depth: number) => {
    node.childNodes.forEach((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        pushStyled(child.nodeValue ?? '', child.parentElement)
        return
      }
      if (child.nodeType !== Node.ELEMENT_NODE) return
      const el = child as HTMLElement
      const tag = el.tagName

      if (tag === 'BR') {
        ensureNL()
        return
      }
      // A nested table is extracted as its own TableEl; don't inline its text.
      if (tag === 'TABLE') return
      if (tag === 'UL' || tag === 'OL') {
        ensureNL()
        walk(el, tag === 'OL' ? { n: 0 } : null, depth + 1)
        ensureNL()
        return
      }
      if (tag === 'LI') {
        ensureNL()
        pushPlain('  '.repeat(Math.max(0, depth - 1)))
        pushPlain(olCtr ? `${(olCtr.n += 1)}. ` : '• ')
        walk(el, null, depth)
        ensureNL()
        return
      }
      if (BLOCK_TAGS.has(tag)) {
        ensureNL()
        walk(el, olCtr, depth)
        ensureNL()
        return
      }
      walk(el, olCtr, depth)
    })
  }

  walk(root, null, 0)
  return mergeRuns(trimNewlines(runs))
}

function trimNewlines(runs: TextRun[]): TextRun[] {
  const out = runs.slice()
  while (out.length && out[0].text.trim() === '' && out[0].text.includes('\n')) out.shift()
  while (out.length && out[out.length - 1].text.trim() === '' && out[out.length - 1].text.includes('\n')) out.pop()
  return out
}

function mergeRuns(runs: TextRun[]): TextRun[] {
  const out: TextRun[] = []
  for (const r of runs) {
    const prev = out[out.length - 1]
    if (
      prev &&
      !!prev.bold === !!r.bold &&
      !!prev.italic === !!r.italic &&
      !!prev.code === !!r.code &&
      prev.color === r.color &&
      prev.fontSize === r.fontSize
    ) {
      prev.text += r.text
    } else if (r.text) {
      out.push({ ...r })
    }
  }
  return out
}

function emptySlide(): Slide {
  return { id: genId(), background: 'FFFFFF', boxes: [] }
}

async function documentFontsReady(): Promise<void> {
  try {
    await document.fonts?.ready
  } catch {
    /* ignore */
  }
}
