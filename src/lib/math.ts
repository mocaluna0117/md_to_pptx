import katex from 'katex'
import 'katex/dist/katex.min.css'
import { toPng } from 'html-to-image'

/**
 * Replace LaTeX math (`$…$` inline, `$$…$$` display) in Markdown with rendered
 * PNG `<img>` tags (data URIs), so both apps' existing image pipelines carry the
 * math through preview, slides, and export. Heights are in `em` so the math
 * scales with the surrounding text. Code spans / fences are left untouched.
 */
export async function mathToImages(markdown: string): Promise<string> {
  if (!markdown.includes('$')) return markdown

  // Protect code so `$` inside it isn't treated as math.
  const vault: string[] = []
  const stash = (m: string) => `@@MDCODE${vault.push(m) - 1}@@`
  let src = markdown
    .replace(/```[\s\S]*?```/g, stash)
    .replace(/~~~[\s\S]*?~~~/g, stash)
    .replace(/`[^`\n]+`/g, stash)

  // Collect math jobs, leaving indexed placeholders.
  const jobs: { latex: string; display: boolean }[] = []
  const mark = (latex: string, display: boolean) => `@@MDMATH${jobs.push({ latex, display }) - 1}@@`
  src = src.replace(/\$\$([\s\S]+?)\$\$/g, (_m, latex: string) => mark(latex, true))
  // Inline: no space just inside $…$, closing $ not followed by a digit (skips "$5 … $10").
  src = src.replace(/(?<![\\$])\$(?!\s)((?:[^$\n\\]|\\.)+?)(?<![\s\\])\$(?!\d)/g, (_m, latex: string) =>
    mark(latex, false),
  )

  if (jobs.length > 0) {
    const tags = await Promise.all(jobs.map(({ latex, display }) => renderMathTag(latex.trim(), display)))
    src = src.replace(/@@MDMATH(\d+)@@/g, (_m, i: string) => tags[Number(i)] ?? '')
  }

  // Restore protected code.
  return src.replace(/@@MDCODE(\d+)@@/g, (_m, i: string) => vault[Number(i)] ?? '')
}

async function renderMathTag(latex: string, display: boolean): Promise<string> {
  const rendered = await renderMathPng(latex, display)
  const alt = latex.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
  if (!rendered) return display ? `<pre>${alt}</pre>` : `<code>${alt}</code>`
  const em = (rendered.height / 16).toFixed(2)
  if (display) {
    return `<img src="${rendered.url}" alt="${alt}" style="display:block;margin:0.7em auto;max-width:100%;height:${em}em" />`
  }
  return `<img src="${rendered.url}" alt="${alt}" style="vertical-align:middle;height:${em}em" />`
}

interface RenderedMath {
  url: string
  width: number
  height: number
}

const cache = new Map<string, RenderedMath | null>()

/** Render one LaTeX string with KaTeX and rasterize it to a PNG data URI. */
async function renderMathPng(latex: string, display: boolean): Promise<RenderedMath | null> {
  const key = (display ? 'D:' : 'I:') + latex
  if (cache.has(key)) return cache.get(key) ?? null

  let html: string
  try {
    html = katex.renderToString(latex, { displayMode: display, throwOnError: false, output: 'html' })
  } catch {
    cache.set(key, null)
    return null
  }

  const wrap = document.createElement('div')
  wrap.style.cssText =
    'position:fixed;left:-99999px;top:0;font-size:16px;line-height:normal;color:#111;background:transparent;display:inline-block;padding:1px;white-space:nowrap;'
  // Drop KaTeX display margins so the rasterized box is tight.
  wrap.innerHTML = html.replace('class="katex-display"', 'class="katex-display" style="margin:0"')
  document.body.appendChild(wrap)
  try {
    await (document.fonts?.ready ?? Promise.resolve())
    const rect = wrap.getBoundingClientRect()
    const url = await toPng(wrap, { pixelRatio: 3, cacheBust: true })
    const res: RenderedMath = { url, width: rect.width, height: rect.height || 20 }
    cache.set(key, res)
    return res
  } catch {
    cache.set(key, null)
    return null
  } finally {
    document.body.removeChild(wrap)
  }
}
