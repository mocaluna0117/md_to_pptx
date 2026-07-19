/**
 * Minimal Marp front-matter reader. Marp global directives live in a leading
 * YAML block delimited by `---`. We only need simple `key: value` scalar lines
 * (backgroundColor, color, ...) to carry style directives into the exporters.
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
