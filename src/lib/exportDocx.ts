import MarkdownIt from 'markdown-it'
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  LevelFormat,
} from 'docx'

const md = new MarkdownIt({ html: false, linkify: true, breaks: false })

type MdToken = ReturnType<typeof md.parse>[number]
type Block = Paragraph | Table

interface RunStyle {
  bold?: boolean
  italic?: boolean
  strike?: boolean
  code?: boolean
  color?: string
  underline?: boolean
}

export interface DocxOptions {
  fileName?: string
}

/** Convert Markdown to an editable .docx and trigger a download. */
export async function exportMarkdownToDocx(markdown: string, options: DocxOptions = {}): Promise<void> {
  const { fileName = 'document.docx' } = options
  const tokens = md.parse(markdown, {})
  const orderedCounter = { n: 0 }
  const children = blocksFromTokens(tokens, orderedCounter)

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
function blocksFromTokens(tokens: MdToken[], oc: { n: number }): Block[] {
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
        out.push(new Paragraph({ heading: headingLevel(level), children: inlineToRuns(inline?.children ?? []) }))
        i += 3
        break
      }
      case 'paragraph_open': {
        const inline = tokens[i + 1]
        const runs = inlineToRuns(inline?.children ?? [], quoteDepth > 0 ? { italic: true } : {})
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
        const { table, next } = parseTable(tokens, i)
        out.push(table)
        i = next
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
  runs: TextRun[],
  listStack: { ordered: boolean; instance: number }[],
  quoteDepth: number,
): Paragraph {
  const ctx = listStack[listStack.length - 1]
  if (ctx) {
    const level = Math.min(4, listStack.length - 1)
    return new Paragraph({
      children: runs,
      ...(ctx.ordered
        ? { numbering: { reference: 'ol', level, instance: ctx.instance } }
        : { bullet: { level } }),
    })
  }
  if (quoteDepth > 0) {
    return new Paragraph({
      children: runs,
      indent: { left: 480 * quoteDepth },
      border: { left: { style: BorderStyle.SINGLE, size: 18, color: 'CBD5E1', space: 12 } },
      spacing: { before: 40, after: 40 },
    })
  }
  return new Paragraph({ children: runs, spacing: { after: 120 } })
}

function inlineToRuns(children: MdToken[], base: RunStyle = {}): TextRun[] {
  const runs: TextRun[] = []
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
      case 'image':
        if (c.content) runs.push(makeRun(c.content, style))
        break
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
    font: s.code ? 'Consolas' : undefined,
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

function parseTable(tokens: MdToken[], start: number): { table: Table; next: number } {
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
      const runs = inlineToRuns(inline?.children ?? [], inHeader ? { bold: true } : {})
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
