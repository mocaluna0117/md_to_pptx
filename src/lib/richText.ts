import { toHex, type TextRun } from './deck'

/** Serialize runs to HTML for a contentEditable box. Newlines become <br>. */
export function runsToHtml(runs: TextRun[]): string {
  if (runs.length === 0) return ''
  return runs
    .map((r) => {
      let html = escapeHtml(r.text).replace(/\n/g, '<br>')
      if (r.code) html = `<code>${html}</code>`
      if (r.italic) html = `<i>${html}</i>`
      if (r.bold) html = `<b>${html}</b>`
      if (r.color) html = `<span style="color:#${r.color}">${html}</span>`
      return html
    })
    .join('')
}

interface Ctx {
  bold?: boolean
  italic?: boolean
  code?: boolean
  color?: string
}

/** Parse the HTML produced by a contentEditable box back into styled runs. */
export function htmlToRuns(root: Node): TextRun[] {
  const runs: TextRun[] = []
  walk(root, {}, runs)
  // Drop a single trailing newline that block elements tend to add.
  if (runs.length && runs[runs.length - 1].text.endsWith('\n')) {
    runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\n$/, '')
  }
  return merge(runs.filter((r) => r.text.length > 0))
}

function walk(node: Node, ctx: Ctx, runs: TextRun[]): void {
  node.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      push(runs, child.nodeValue ?? '', ctx)
      return
    }
    if (child.nodeType !== Node.ELEMENT_NODE) return

    const el = child as HTMLElement
    const tag = el.tagName

    if (tag === 'BR') {
      push(runs, '\n', ctx)
      return
    }

    const isBlock = tag === 'DIV' || tag === 'P'
    if (isBlock && runs.length && !runs[runs.length - 1].text.endsWith('\n')) {
      push(runs, '\n', ctx)
    }

    const next: Ctx = { ...ctx }
    if (tag === 'B' || tag === 'STRONG') next.bold = true
    if (tag === 'I' || tag === 'EM') next.italic = true
    if (tag === 'CODE') next.code = true
    const color = colorOf(el)
    if (color) next.color = color

    walk(el, next, runs)
  })
}

function colorOf(el: HTMLElement): string | undefined {
  const css = el.style?.color
  if (css) return toHex(css) ?? undefined
  if (el.tagName === 'FONT') {
    const attr = el.getAttribute('color')
    if (attr) return toHex(attr) ?? undefined
  }
  return undefined
}

function push(runs: TextRun[], text: string, ctx: Ctx): void {
  if (!text) return
  runs.push({ text, bold: ctx.bold, italic: ctx.italic, code: ctx.code, color: ctx.color })
}

function merge(runs: TextRun[]): TextRun[] {
  const out: TextRun[] = []
  for (const r of runs) {
    const prev = out[out.length - 1]
    if (
      prev &&
      !!prev.bold === !!r.bold &&
      !!prev.italic === !!r.italic &&
      !!prev.code === !!r.code &&
      prev.color === r.color
    ) {
      prev.text += r.text
    } else {
      out.push({ ...r })
    }
  }
  return out
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}
