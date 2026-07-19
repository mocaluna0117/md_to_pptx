import MarkdownIt from 'markdown-it'
import type Token from 'markdown-it/lib/token.mjs'

/** A styled span of text inside a paragraph. */
export interface Run {
  text: string
  bold?: boolean
  italic?: boolean
  strike?: boolean
  code?: boolean
  link?: string
}

export type ParaKind = 'h1' | 'h2' | 'h3' | 'h4' | 'p' | 'li' | 'code' | 'quote'

/** A block of text that flows in the body text box. */
export interface Para {
  kind: ParaKind
  runs: Run[]
  /** Nesting level for list items / blockquotes (0-based). */
  indent: number
  /** For list items: true = numbered, false = bulleted. */
  ordered?: boolean
}

export interface ImageBlock {
  src: string
  alt: string
}

export interface TableBlock {
  header: string[]
  rows: string[][]
}

export interface SlideModel {
  paras: Para[]
  images: ImageBlock[]
  tables: TableBlock[]
}

const md = new MarkdownIt({ html: false, linkify: true, breaks: false })

/** Split a Marp document into per-slide Markdown, respecting fenced code blocks. */
export function splitSlides(source: string): string[] {
  let src = source.replace(/^﻿/, '')

  // Strip a leading YAML front-matter block (Marp global directives).
  const fm = src.match(/^---[ \t]*\r?\n[\s\S]*?\r?\n---[ \t]*\r?\n?/)
  if (fm) src = src.slice(fm[0].length)

  // Strip HTML comments (Marp directives like <!-- _class: ... -->).
  src = src.replace(/<!--[\s\S]*?-->/g, '')

  const lines = src.split(/\r?\n/)
  const slides: string[][] = [[]]
  let fence: string | null = null

  for (const line of lines) {
    const fenceMatch = line.match(/^\s*(`{3,}|~{3,})/)
    if (fenceMatch) {
      const marker = fenceMatch[1][0].repeat(3)
      if (fence === null) fence = marker
      else if (line.trim().startsWith(fence)) fence = null
    }

    const isBreak = fence === null && /^ {0,3}(-{3,}|\*{3,}|_{3,})[ \t]*$/.test(line)
    if (isBreak) slides.push([])
    else slides[slides.length - 1].push(line)
  }

  return slides.map((s) => s.join('\n').trim()).filter((s) => s.length > 0)
}

/** Parse a full Marp document into slide models. */
export function parseSlides(source: string): SlideModel[] {
  return splitSlides(source).map(parseSlide)
}

interface TreeNode {
  type: string
  token: Token
  children: TreeNode[]
}

function toTree(tokens: Token[]): TreeNode {
  const root: TreeNode = { type: 'root', token: null as unknown as Token, children: [] }
  const stack: TreeNode[] = [root]
  for (const t of tokens) {
    if (t.nesting === 1) {
      const node: TreeNode = { type: t.type.replace(/_open$/, ''), token: t, children: [] }
      stack[stack.length - 1].children.push(node)
      stack.push(node)
    } else if (t.nesting === -1) {
      stack.pop()
    } else {
      stack[stack.length - 1].children.push({ type: t.type, token: t, children: [] })
    }
  }
  return root
}

function parseSlide(source: string): SlideModel {
  const tree = toTree(md.parse(source, {}))
  const model: SlideModel = { paras: [], images: [], tables: [] }
  walk(tree.children, model, { indent: 0, quote: false })
  return model
}

interface Ctx {
  indent: number
  quote: boolean
}

function walk(nodes: TreeNode[], model: SlideModel, ctx: Ctx): void {
  for (const node of nodes) {
    switch (node.type) {
      case 'heading': {
        const level = Math.min(4, Number(node.token.tag.slice(1)) || 1)
        model.paras.push({
          kind: `h${level}` as ParaKind,
          runs: inlineRuns(inlineChildren(node), model),
          indent: 0,
        })
        break
      }
      case 'paragraph': {
        const runs = inlineRuns(inlineChildren(node), model)
        if (runs.some((r) => r.text.trim())) {
          model.paras.push({ kind: ctx.quote ? 'quote' : 'p', runs, indent: ctx.indent })
        }
        break
      }
      case 'bullet_list':
      case 'ordered_list': {
        const ordered = node.type === 'ordered_list'
        for (const item of node.children) {
          if (item.type !== 'list_item') continue
          walkListItem(item, model, { ...ctx, indent: ctx.indent }, ordered)
        }
        break
      }
      case 'blockquote':
        walk(node.children, model, { indent: ctx.indent, quote: true })
        break
      case 'fence':
      case 'code_block': {
        const codeLines = node.token.content.replace(/\n$/, '').split('\n')
        for (const line of codeLines) {
          model.paras.push({ kind: 'code', runs: [{ text: line || ' ', code: true }], indent: ctx.indent })
        }
        break
      }
      case 'table':
        model.tables.push(parseTable(node))
        break
      default:
        if (node.children.length) walk(node.children, model, ctx)
    }
  }
}

function walkListItem(item: TreeNode, model: SlideModel, ctx: Ctx, ordered: boolean): void {
  for (const child of item.children) {
    if (child.type === 'paragraph') {
      model.paras.push({
        kind: 'li',
        runs: inlineRuns(inlineChildren(child), model),
        indent: ctx.indent,
        ordered,
      })
    } else if (child.type === 'bullet_list' || child.type === 'ordered_list') {
      const nestedOrdered = child.type === 'ordered_list'
      for (const nested of child.children) {
        if (nested.type === 'list_item') {
          walkListItem(nested, model, { ...ctx, indent: ctx.indent + 1 }, nestedOrdered)
        }
      }
    }
  }
}

function inlineChildren(node: TreeNode): Token[] {
  const inline = node.children.find((c) => c.type === 'inline')
  return inline?.token.children ?? []
}

function inlineRuns(children: Token[], model: SlideModel): Run[] {
  const runs: Run[] = []
  const style = { bold: 0, italic: 0, strike: 0 }
  let link: string | null = null

  const push = (text: string, extra: Partial<Run> = {}) => {
    if (!text) return
    runs.push({
      text,
      bold: style.bold > 0 || undefined,
      italic: style.italic > 0 || undefined,
      strike: style.strike > 0 || undefined,
      link: link ?? undefined,
      ...extra,
    })
  }

  for (const c of children) {
    switch (c.type) {
      case 'text':
        push(c.content)
        break
      case 'code_inline':
        push(c.content, { code: true })
        break
      case 'strong_open':
        style.bold++
        break
      case 'strong_close':
        style.bold--
        break
      case 'em_open':
        style.italic++
        break
      case 'em_close':
        style.italic--
        break
      case 's_open':
        style.strike++
        break
      case 's_close':
        style.strike--
        break
      case 'link_open':
        link = c.attrGet('href')
        break
      case 'link_close':
        link = null
        break
      case 'softbreak':
      case 'hardbreak':
        push(' ')
        break
      case 'image':
        model.images.push({ src: c.attrGet('src') ?? '', alt: c.content ?? '' })
        break
      default:
        break
    }
  }

  return mergeRuns(runs)
}

/** Collapse adjacent runs that share identical styling. */
function mergeRuns(runs: Run[]): Run[] {
  const out: Run[] = []
  for (const r of runs) {
    const prev = out[out.length - 1]
    if (
      prev &&
      !!prev.bold === !!r.bold &&
      !!prev.italic === !!r.italic &&
      !!prev.strike === !!r.strike &&
      !!prev.code === !!r.code &&
      prev.link === r.link
    ) {
      prev.text += r.text
    } else {
      out.push({ ...r })
    }
  }
  return out
}

function parseTable(node: TreeNode): TableBlock {
  const header: string[] = []
  const rows: string[][] = []
  for (const section of node.children) {
    for (const tr of section.children) {
      if (tr.type !== 'tr') continue
      const cells = tr.children
        .filter((c) => c.type === 'th' || c.type === 'td')
        .map((cell) => plainText(inlineChildren(cell)))
      if (section.type === 'thead') header.push(...cells)
      else rows.push(cells)
    }
  }
  return { header, rows }
}

function plainText(children: Token[]): string {
  return children
    .map((c) => (c.type === 'text' || c.type === 'code_inline' ? c.content : c.type === 'softbreak' ? ' ' : ''))
    .join('')
    .trim()
}
