/**
 * Minimal Marp front-matter helpers. Marp global directives live in a leading
 * YAML block delimited by `---`. We only need simple `key: value` scalar lines
 * (theme, paginate, backgroundColor, color, ...), preserving any other lines.
 */

const FM_RE = /^﻿?---[ \t]*\r?\n([\s\S]*?)\r?\n---[ \t]*(?:\r?\n|$)/

export interface FrontMatter {
  data: Record<string, string>
  /** True when the document already has a front-matter block. */
  hasBlock: boolean
}

export function parseFrontMatter(md: string): FrontMatter {
  const m = md.match(FM_RE)
  if (!m) return { data: {}, hasBlock: false }
  const data: Record<string, string> = {}
  for (const line of m[1].split(/\r?\n/)) {
    const km = line.match(/^([A-Za-z0-9_-]+)[ \t]*:[ \t]*(.*)$/)
    if (km) data[km[1]] = km[2].trim()
  }
  return { data, hasBlock: true }
}

/**
 * Set (or remove, when `value` is null) a front-matter key, returning the new
 * Markdown. Creates a front-matter block if none exists and always keeps
 * `marp: true` present.
 */
export function setFrontMatterKey(md: string, key: string, value: string | null): string {
  const m = md.match(FM_RE)

  if (!m) {
    if (value === null) return md
    const block = `---\nmarp: true\n${key}: ${value}\n---\n\n`
    return block + md
  }

  const body = md.slice(m[0].length)
  let lines = m[1].split(/\r?\n/)
  const keyRe = new RegExp(`^${escapeRe(key)}[ \\t]*:`)

  if (value === null) {
    lines = lines.filter((l) => !keyRe.test(l))
  } else if (lines.some((l) => keyRe.test(l))) {
    lines = lines.map((l) => (keyRe.test(l) ? `${key}: ${value}` : l))
  } else {
    lines.push(`${key}: ${value}`)
  }

  if (!lines.some((l) => /^marp[ \t]*:/.test(l))) lines.unshift('marp: true')

  const inner = lines.filter((l, i) => !(l.trim() === '' && i === lines.length - 1)).join('\n')
  return `---\n${inner}\n---\n${body.startsWith('\n') ? '' : '\n'}${body}`
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
