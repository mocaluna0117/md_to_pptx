import MarkdownIt from 'markdown-it'
import markdownItCjkFriendly from 'markdown-it-cjk-friendly'
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  LevelFormat,
  ExternalHyperlink,
  Textbox,
} from 'docx'
import { toHex } from './deck'
import type { DocBox } from './docBox'

const md = new MarkdownIt({ html: true, linkify: true, breaks: false }).use(markdownItCjkFriendly)

type MdToken = ReturnType<typeof md.parse>[number]
type Block = Paragraph | Table
type InlineChild = TextRun | ImageRun | ExternalHyperlink
interface ResolvedImage {
  data: Uint8Array
  width: number
  height: number
}
type ImageMap = Map<string, ResolvedImage>

interface RunStyle {
  bold?: boolean
  italic?: boolean
  strike?: boolean
  code?: boolean
  color?: string
  underline?: boolean
  font?: string
  size?: number
}

export interface DocxOptions {
  fileName?: string
}

/** Convert Markdown to an editable .docx and trigger a download. */
export async function exportMarkdownToDocx(markdown: string, options: DocxOptions = {}): Promise<void> {
  const { fileName = 'document.docx' } = options
  const tokens = md.parse(markdown, {})
  const images = await resolveImages(markdown)
  const children = blocksFromTokens(tokens, { n: 0 }, images)
  await packAndDownload(children, fileName)
}

/**
 * Convert the edited WYSIWYG document (contentEditable HTML) to an editable .docx.
 * This is Docdown's primary export path: the visually edited document is the source
 * of truth, so we walk its DOM rather than re-parsing Markdown.
 *
 * `boxes` are free-floating text boxes layered over the document; each becomes a
 * Word text box anchored to the page at an absolute position.
 */
export async function exportHtmlToDocx(html: string, boxes: DocBox[] = [], options: DocxOptions = {}): Promise<void> {
  const { fileName = 'document.docx' } = options
  const root = document.createElement('div')
  root.innerHTML = html
  const boxRoots = boxes.map((b) => {
    const d = document.createElement('div')
    d.innerHTML = b.html
    return d
  })
  const srcs = [...collectImageSrcs(root), ...boxRoots.flatMap(collectImageSrcs)]
  const images = await resolveImageSrcs(srcs)
  const children = blocksFromDom(root, { n: 0 }, images)
  // Text boxes are prepended so their anchor paragraph sits on the first page,
  // letting page-relative absolute positions place them at the intended coords.
  const boxChildren = boxes.map((b, i) => textboxFromBox(b, boxRoots[i], images))
  await packAndDownload([...boxChildren, ...children], fileName)
}

const PX_TO_PT = 0.75 // 96dpi CSS pixels → points

function textboxFromBox(box: DocBox, root: HTMLElement, images: ImageMap): Textbox {
  return new Textbox({
    style: {
      position: 'absolute',
      positionHorizontalRelative: 'page',
      positionVerticalRelative: 'page',
      left: `${Math.round(box.x * PX_TO_PT)}pt`,
      top: `${Math.round(box.y * PX_TO_PT)}pt`,
      width: `${Math.round(box.w * PX_TO_PT)}pt`,
      height: `${Math.round(box.h * PX_TO_PT)}pt`,
    },
    children: runsFromBoxDom(root, images),
  })
}

/** Flatten a box's edited HTML into one paragraph's runs (block boundaries → line breaks). */
function runsFromBoxDom(root: HTMLElement, images: ImageMap): InlineChild[] {
  const nodes = Array.from(root.childNodes)
  const hasBlocks = nodes.some((n) => n.nodeType === Node.ELEMENT_NODE && !isInlineNode(n))
  if (!hasBlocks) return inlineFromNodes(nodes, images, {})

  const out: InlineChild[] = []
  let first = true
  for (const node of nodes) {
    if (node.nodeType === Node.ELEMENT_NODE && !isInlineNode(node)) {
      if (!first) out.push(new TextRun({ break: 1 }))
      for (const c of Array.from(node.childNodes)) walkInline(c, {}, images, out)
      first = false
    } else if (hasInlineContent(node)) {
      walkInline(node, {}, images, out)
      first = false
    }
  }
  return out.length ? out : [new TextRun('')]
}

/** Assemble the shared Document (numbering + default style) and trigger the download. */
async function packAndDownload(children: Block[], fileName: string): Promise<void> {
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: 'ol',
          levels: [0, 1, 2, 3, 4].map((level) => ({
            level,
            format: LevelFormat.DECIMAL,
            text: `%${level + 1}.`,
            alignment: AlignmentType.START,
            style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
          })),
        },
      ],
    },
    styles: {
      default: {
        document: { run: { font: 'Calibri', size: 22 } },
      },
    },
    sections: [{ children: children.length ? children : [new Paragraph('')] }],
  })

  const blob = await Packer.toBlob(doc)
  downloadBlob(blob, ensureExt(fileName))
}

/** Walk the flat token stream into a list of docx blocks (paragraphs / tables). */
function blocksFromTokens(tokens: MdToken[], oc: { n: number }, images: ImageMap): Block[] {
  const out: Block[] = []
  const listStack: { ordered: boolean; instance: number }[] = []
  let quoteDepth = 0
  let i = 0

  while (i < tokens.length) {
    const t = tokens[i]
    switch (t.type) {
      case 'heading_open': {
        const level = Number(t.tag.slice(1)) || 1
        const inline = tokens[i + 1]
        out.push(new Paragraph({ heading: headingLevel(level), children: inlineToRuns(inline?.children ?? [], images) }))
        i += 3
        break
      }
      case 'paragraph_open': {
        const inline = tokens[i + 1]
        const runs = inlineToRuns(inline?.children ?? [], images, quoteDepth > 0 ? { italic: true } : {})
        out.push(paragraphInContext(runs, listStack, quoteDepth))
        i += 3
        break
      }
      case 'bullet_list_open':
        listStack.push({ ordered: false, instance: 0 })
        i += 1
        break
      case 'ordered_list_open':
        listStack.push({ ordered: true, instance: (oc.n += 1) })
        i += 1
        break
      case 'bullet_list_close':
      case 'ordered_list_close':
        listStack.pop()
        i += 1
        break
      case 'blockquote_open':
        quoteDepth += 1
        i += 1
        break
      case 'blockquote_close':
        quoteDepth = Math.max(0, quoteDepth - 1)
        i += 1
        break
      case 'fence':
      case 'code_block': {
        out.push(codeBlockParagraph(t.content))
        i += 1
        break
      }
      case 'table_open': {
        const { table, next } = parseTable(tokens, i, images)
        out.push(table)
        i = next
        break
      }
      case 'html_block': {
        const runs = imgRunsFromHtml(t.content, images)
        if (runs.length) out.push(new Paragraph({ children: runs }))
        i += 1
        break
      }
      case 'hr':
        out.push(
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC', space: 1 } },
            children: [],
          }),
        )
        i += 1
        break
      default:
        i += 1
    }
  }
  return out
}

function paragraphInContext(
  runs: InlineChild[],
  listStack: { ordered: boolean; instance: number }[],
  quoteDepth: number,
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType],
): Paragraph {
  const ctx = listStack[listStack.length - 1]
  if (ctx) {
    const level = Math.min(4, listStack.length - 1)
    return new Paragraph({
      children: runs,
      alignment,
      ...(ctx.ordered
        ? { numbering: { reference: 'ol', level, instance: ctx.instance } }
        : { bullet: { level } }),
    })
  }
  if (quoteDepth > 0) {
    return new Paragraph({
      children: runs,
      alignment,
      indent: { left: 480 * quoteDepth },
      border: { left: { style: BorderStyle.SINGLE, size: 18, color: 'CBD5E1', space: 12 } },
      spacing: { before: 40, after: 40 },
    })
  }
  return new Paragraph({ children: runs, alignment, spacing: { after: 120 } })
}

function inlineToRuns(children: MdToken[], images: ImageMap, base: RunStyle = {}): (TextRun | ImageRun)[] {
  const runs: (TextRun | ImageRun)[] = []
  const stack: RunStyle[] = []
  let style: RunStyle = { ...base }

  for (const c of children) {
    switch (c.type) {
      case 'text':
        if (c.content) runs.push(makeRun(c.content, style))
        break
      case 'code_inline':
        runs.push(makeRun(c.content, { ...style, code: true }))
        break
      case 'strong_open':
        stack.push(style)
        style = { ...style, bold: true }
        break
      case 'em_open':
        stack.push(style)
        style = { ...style, italic: true }
        break
      case 's_open':
        stack.push(style)
        style = { ...style, strike: true }
        break
      case 'link_open':
        stack.push(style)
        style = { ...style, color: '2563EB', underline: true }
        break
      case 'strong_close':
      case 'em_close':
      case 's_close':
      case 'link_close':
        style = stack.pop() ?? { ...base }
        break
      case 'softbreak':
        runs.push(new TextRun({ text: ' ' }))
        break
      case 'hardbreak':
        runs.push(new TextRun({ break: 1 }))
        break
      case 'image': {
        const src = c.attrGet?.('src') ?? ''
        const img = src && images.get(src)
        if (img) runs.push(imageRun(img))
        else if (c.content) runs.push(makeRun(c.content, style))
        break
      }
      case 'html_inline': {
        if (/<br\s*\/?>/i.test(c.content)) runs.push(new TextRun({ break: 1 }))
        else runs.push(...imgRunsFromHtml(c.content, images))
        break
      }
      default:
        break
    }
  }
  return runs.length ? runs : [new TextRun('')]
}

function makeRun(text: string, s: RunStyle): TextRun {
  return new TextRun({
    text,
    bold: s.bold,
    italics: s.italic,
    strike: s.strike,
    color: s.color,
    underline: s.underline ? {} : undefined,
    font: s.code ? 'Consolas' : s.font,
    size: s.size ? Math.round(s.size * 2) : undefined, // docx size is in half-points
    shading: s.code ? { fill: 'F0F0F0' } : undefined,
  })
}

function codeBlockParagraph(content: string): Paragraph {
  const lines = content.replace(/\n+$/, '').split('\n')
  const runs: TextRun[] = lines.map(
    (line, idx) => new TextRun({ text: line.length ? line : ' ', font: 'Consolas', size: 20, break: idx === 0 ? 0 : 1 }),
  )
  return new Paragraph({
    children: runs,
    shading: { fill: 'F5F5F5' },
    spacing: { before: 100, after: 100 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'E2E8F0', space: 4 },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E2E8F0', space: 4 },
      left: { style: BorderStyle.SINGLE, size: 4, color: 'E2E8F0', space: 6 },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'E2E8F0', space: 6 },
    },
  })
}

function parseTable(tokens: MdToken[], start: number, images: ImageMap): { table: Table; next: number } {
  const rows: TableRow[] = []
  let cells: TableCell[] = []
  let inHeader = false
  let i = start + 1

  while (i < tokens.length && tokens[i].type !== 'table_close') {
    const t = tokens[i]
    if (t.type === 'thead_open') inHeader = true
    else if (t.type === 'thead_close') inHeader = false
    else if (t.type === 'tr_open') cells = []
    else if (t.type === 'tr_close') rows.push(new TableRow({ children: cells }))
    else if (t.type === 'th_open' || t.type === 'td_open') {
      const inline = tokens[i + 1]
      const runs = inlineToRuns(inline?.children ?? [], images, inHeader ? { bold: true } : {})
      cells.push(
        new TableCell({
          children: [new Paragraph({ children: runs })],
          shading: inHeader ? { fill: 'EEF2F7' } : undefined,
          margins: { top: 60, bottom: 60, left: 110, right: 110 },
        }),
      )
    }
    i += 1
  }

  const border = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' }
  const table = new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
  })
  return { table, next: i + 1 }
}

// ---- Edited-document (HTML DOM) → docx ----

const INLINE_TAGS = new Set([
  'A', 'B', 'STRONG', 'I', 'EM', 'S', 'STRIKE', 'DEL', 'U', 'INS', 'CODE', 'KBD', 'SAMP', 'TT',
  'SPAN', 'FONT', 'BR', 'IMG', 'MARK', 'SMALL', 'SUB', 'SUP', 'ABBR', 'CITE', 'Q', 'WBR', 'LABEL',
])

interface ListFrame {
  ordered: boolean
  instance: number
}
interface DomCtx {
  oc: { n: number }
  listStack: ListFrame[]
  quoteDepth: number
}

/** Walk an edited-document DOM into a list of docx blocks (paragraphs / tables). */
function blocksFromDom(root: HTMLElement, oc: { n: number }, images: ImageMap): Block[] {
  const out: Block[] = []
  walkBlocks(root, { oc, listStack: [], quoteDepth: 0 }, images, out)
  return out
}

function isInlineNode(node: Node): boolean {
  if (node.nodeType === Node.TEXT_NODE) return true
  if (node.nodeType !== Node.ELEMENT_NODE) return false
  return INLINE_TAGS.has((node as Element).tagName)
}

function walkBlocks(parent: Node, ctx: DomCtx, images: ImageMap, out: Block[]): void {
  let buffer: Node[] = []
  const flush = () => {
    if (buffer.some(hasInlineContent)) {
      out.push(paragraphInContext(inlineFromNodes(buffer, images, {}), ctx.listStack, ctx.quoteDepth))
    }
    buffer = []
  }

  for (const child of Array.from(parent.childNodes)) {
    if (isInlineNode(child)) {
      buffer.push(child)
      continue
    }
    flush()
    const el = child as HTMLElement
    const tag = el.tagName
    if (/^H[1-6]$/.test(tag)) {
      out.push(
        new Paragraph({
          heading: headingLevel(Number(tag[1])),
          alignment: alignOf(el),
          children: inlineFromNodes(Array.from(el.childNodes), images, {}),
        }),
      )
    } else if (tag === 'P' || tag === 'DIV') {
      out.push(paragraphInContext(inlineFromNodes(Array.from(el.childNodes), images, {}), ctx.listStack, ctx.quoteDepth, alignOf(el)))
    } else if (tag === 'BLOCKQUOTE') {
      walkBlocks(el, { ...ctx, quoteDepth: ctx.quoteDepth + 1 }, images, out)
    } else if (tag === 'UL' || tag === 'OL') {
      const ordered = tag === 'OL'
      const frame: ListFrame = { ordered, instance: ordered ? (ctx.oc.n += 1) : 0 }
      const inner: DomCtx = { ...ctx, listStack: [...ctx.listStack, frame] }
      for (const li of Array.from(el.children)) {
        if (li.tagName === 'LI') processListItem(li as HTMLElement, inner, images, out)
      }
    } else if (tag === 'TABLE') {
      out.push(tableFromDom(el, images))
    } else if (tag === 'PRE') {
      out.push(codeBlockParagraph(el.textContent ?? ''))
    } else if (tag === 'HR') {
      out.push(
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC', space: 1 } },
          children: [],
        }),
      )
    } else if (tag === 'FIGURE') {
      walkBlocks(el, ctx, images, out)
    } else {
      // Unknown block: descend so its contents aren't lost.
      walkBlocks(el, ctx, images, out)
    }
  }
  flush()
}

/** A list item becomes one list paragraph for its inline content, plus any nested lists. */
function processListItem(li: HTMLElement, ctx: DomCtx, images: ImageMap, out: Block[]): void {
  const inlineNodes: Node[] = []
  const nestedLists: HTMLElement[] = []
  for (const node of Array.from(li.childNodes)) {
    if (node.nodeType === Node.ELEMENT_NODE && ((node as Element).tagName === 'UL' || (node as Element).tagName === 'OL')) {
      nestedLists.push(node as HTMLElement)
    } else {
      inlineNodes.push(node)
    }
  }
  out.push(paragraphInContext(inlineFromNodes(inlineNodes, images, {}), ctx.listStack, ctx.quoteDepth))
  for (const list of nestedLists) {
    const ordered = list.tagName === 'OL'
    const frame: ListFrame = { ordered, instance: ordered ? (ctx.oc.n += 1) : 0 }
    const inner: DomCtx = { ...ctx, listStack: [...ctx.listStack, frame] }
    for (const nested of Array.from(list.children)) {
      if (nested.tagName === 'LI') processListItem(nested as HTMLElement, inner, images, out)
    }
  }
}

function tableFromDom(table: HTMLElement, images: ImageMap): Table {
  const rows: TableRow[] = []
  for (const tr of Array.from(table.querySelectorAll('tr'))) {
    const cells: TableCell[] = []
    for (const cell of Array.from(tr.children)) {
      if (cell.tagName !== 'TD' && cell.tagName !== 'TH') continue
      const header = cell.tagName === 'TH'
      cells.push(
        new TableCell({
          children: [
            new Paragraph({
              alignment: alignOf(cell as HTMLElement),
              children: inlineFromNodes(Array.from(cell.childNodes), images, header ? { bold: true } : {}),
            }),
          ],
          shading: header ? { fill: 'EEF2F7' } : undefined,
          margins: { top: 60, bottom: 60, left: 110, right: 110 },
        }),
      )
    }
    if (cells.length) rows.push(new TableRow({ children: cells }))
  }
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' }
  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
  })
}

/** Map an element's CSS text-align to a docx alignment (undefined = default/left). */
function alignOf(el: HTMLElement): (typeof AlignmentType)[keyof typeof AlignmentType] | undefined {
  switch (el.style.textAlign) {
    case 'center':
      return AlignmentType.CENTER
    case 'right':
      return AlignmentType.RIGHT
    case 'justify':
      return AlignmentType.JUSTIFIED
    case 'left':
      return AlignmentType.LEFT
    default:
      return undefined
  }
}

/** True if a node carries text or an image (used to skip whitespace-only buffers). */
function hasInlineContent(node: Node): boolean {
  if (node.nodeType === Node.TEXT_NODE) return (node.textContent ?? '').trim().length > 0
  if (node.nodeType !== Node.ELEMENT_NODE) return false
  const el = node as HTMLElement
  if (el.tagName === 'IMG' || el.tagName === 'BR') return true
  return (el.textContent ?? '').trim().length > 0 || !!el.querySelector('img')
}

function inlineFromNodes(nodes: Node[], images: ImageMap, base: RunStyle): InlineChild[] {
  const out: InlineChild[] = []
  for (const node of nodes) walkInline(node, base, images, out)
  return out.length ? out : [new TextRun('')]
}

function walkInline(node: Node, style: RunStyle, images: ImageMap, out: InlineChild[]): void {
  if (node.nodeType === Node.TEXT_NODE) {
    const raw = node.textContent ?? ''
    const text = style.code ? raw : raw.replace(/\s+/g, ' ')
    if (text) out.push(makeRun(text, style))
    return
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return
  const el = node as HTMLElement

  switch (el.tagName) {
    case 'BR':
      out.push(new TextRun({ break: 1 }))
      return
    case 'IMG': {
      const src = el.getAttribute('src') ?? ''
      const img = src && images.get(src)
      if (img) out.push(imageRun(img))
      else if (el.getAttribute('alt')) out.push(makeRun(el.getAttribute('alt') ?? '', style))
      return
    }
    case 'A': {
      const href = el.getAttribute('href')
      const runs: InlineChild[] = []
      for (const c of Array.from(el.childNodes)) walkInline(c, { ...style, color: '2563EB', underline: true }, images, runs)
      if (href) {
        out.push(new ExternalHyperlink({ link: href, children: runs.length ? runs : [makeRun(href, { ...style, color: '2563EB', underline: true })] }))
      } else {
        out.push(...runs)
      }
      return
    }
  }

  const next = { ...style, ...styleFromElement(el) }
  for (const c of Array.from(el.childNodes)) walkInline(c, next, images, out)
}

/** Derive run-style deltas from an inline element's tag and inline CSS. */
function styleFromElement(el: HTMLElement): RunStyle {
  const s: RunStyle = {}
  switch (el.tagName) {
    case 'B':
    case 'STRONG':
      s.bold = true
      break
    case 'I':
    case 'EM':
    case 'CITE':
      s.italic = true
      break
    case 'S':
    case 'STRIKE':
    case 'DEL':
      s.strike = true
      break
    case 'U':
    case 'INS':
      s.underline = true
      break
    case 'CODE':
    case 'KBD':
    case 'SAMP':
    case 'TT':
      s.code = true
      break
  }
  const fw = el.style.fontWeight
  if (fw === 'bold' || fw === 'bolder' || (/^\d+$/.test(fw) && Number(fw) >= 600)) s.bold = true
  if (el.style.fontStyle === 'italic') s.italic = true
  const deco = el.style.textDecorationLine || el.style.textDecoration
  if (deco.includes('underline')) s.underline = true
  if (deco.includes('line-through')) s.strike = true

  const color = el.style.color || (el.tagName === 'FONT' ? el.getAttribute('color') : null)
  const hex = color ? toHex(color) : null
  if (hex) s.color = hex

  const family = el.style.fontFamily || (el.tagName === 'FONT' ? el.getAttribute('face') : null)
  if (family) s.font = family.split(',')[0].replace(/['"]/g, '').trim()

  const fs = el.dataset?.fs
  if (fs && !Number.isNaN(Number(fs))) s.size = Number(fs)
  return s
}

// ---- Images ----

const MAX_IMG_W = 480

function imageRun(img: ResolvedImage): ImageRun {
  const scale = img.width > MAX_IMG_W ? MAX_IMG_W / img.width : 1
  return new ImageRun({
    data: img.data,
    type: 'png',
    transformation: { width: Math.max(1, Math.round(img.width * scale)), height: Math.max(1, Math.round(img.height * scale)) },
  })
}

function imgRunsFromHtml(html: string, images: ImageMap): ImageRun[] {
  const out: ImageRun[] = []
  for (const m of html.matchAll(/<img\b[^>]*\bsrc\s*=\s*["']([^"']+)["']/gi)) {
    const img = images.get(m[1])
    if (img) out.push(imageRun(img))
  }
  return out
}

/** Find every image src in the Markdown and load each into PNG bytes + size. */
async function resolveImages(markdown: string): Promise<ImageMap> {
  const srcs = new Set<string>()
  for (const m of markdown.matchAll(/!\[[^\]]*\]\(\s*<?([^)\s>]+)>?/g)) srcs.add(m[1])
  for (const m of markdown.matchAll(/<img\b[^>]*\bsrc\s*=\s*["']([^"']+)["']/gi)) srcs.add(m[1])
  return resolveImageSrcs(srcs)
}

/** Load a set of image srcs into PNG bytes + size (shared by the Markdown and HTML paths). */
async function resolveImageSrcs(srcs: Iterable<string>): Promise<ImageMap> {
  const map: ImageMap = new Map()
  await Promise.all(
    [...new Set(srcs)].map(async (src) => {
      const r = await loadImageAsPng(src)
      if (r) map.set(src, r)
    }),
  )
  return map
}

/** Collect every <img> src from an edited-document DOM. */
function collectImageSrcs(root: HTMLElement): string[] {
  return Array.from(root.querySelectorAll('img[src]'))
    .map((img) => img.getAttribute('src') || '')
    .filter(Boolean)
}

/** Load an image (data URI or CORS-friendly URL) and re-encode as PNG bytes. */
async function loadImageAsPng(src: string): Promise<ResolvedImage | null> {
  try {
    const img = new Image()
    img.crossOrigin = 'anonymous'
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve()
      img.onerror = () => reject(new Error('load failed'))
      img.src = src
    })
    const width = img.naturalWidth || 1
    const height = img.naturalHeight || 1
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.drawImage(img, 0, 0)
    const dataUrl = canvas.toDataURL('image/png') // throws if the source is cross-origin tainted
    return { data: base64ToBytes(dataUrl.split(',')[1] ?? ''), width, height }
  } catch {
    return null
  }
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return bytes
}

function headingLevel(n: number) {
  const levels = [
    HeadingLevel.HEADING_1,
    HeadingLevel.HEADING_2,
    HeadingLevel.HEADING_3,
    HeadingLevel.HEADING_4,
    HeadingLevel.HEADING_5,
    HeadingLevel.HEADING_6,
  ]
  return levels[Math.min(levels.length - 1, Math.max(0, n - 1))]
}

function downloadBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function ensureExt(name: string): string {
  const trimmed = name.trim() || 'document'
  return /\.docx$/i.test(trimmed) ? trimmed : `${trimmed}.docx`
}
