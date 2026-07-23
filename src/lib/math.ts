/**
 * Replace LaTeX math (`$…$` inline, `$$…$$` display) in Markdown with rendered
 * PNG `<img>` tags (data URIs), so both apps' existing image pipelines carry the
 * math through preview, slides, and export. Math is rendered with MathJax to SVG
 * (glyphs embedded as paths — no font dependency) and rasterized to PNG. Heights
 * are in `em` so the math scales with the surrounding text.
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
  const em = (rendered.height / 16).toFixed(3)
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

// Lazily-initialized MathJax tex→svg-string converter (keeps it out of the main bundle).
let converter: Promise<(latex: string, display: boolean) => string> | null = null
function getConverter() {
  if (!converter) {
    converter = (async () => {
      const [{ mathjax }, { TeX }, { SVG }, { liteAdaptor }, { RegisterHTMLHandler }, { AllPackages }] =
        await Promise.all([
          import('mathjax-full/js/mathjax.js'),
          import('mathjax-full/js/input/tex.js'),
          import('mathjax-full/js/output/svg.js'),
          import('mathjax-full/js/adaptors/liteAdaptor.js'),
          import('mathjax-full/js/handlers/html.js'),
          import('mathjax-full/js/input/tex/AllPackages.js'),
        ])
      const adaptor = liteAdaptor()
      RegisterHTMLHandler(adaptor)
      const tex = new TeX({ packages: AllPackages })
      const out = new SVG({ fontCache: 'none' })
      const doc = mathjax.document('', { InputJax: tex, OutputJax: out })
      return (latex: string, display: boolean) => {
        const node = doc.convert(latex, { display })
        return adaptor.outerHTML(node)
      }
    })()
  }
  return converter
}

/** Render one LaTeX string with MathJax (SVG) and rasterize it to a PNG data URI. */
async function renderMathPng(latex: string, display: boolean): Promise<RenderedMath | null> {
  const key = (display ? 'D:' : 'I:') + latex
  if (cache.has(key)) return cache.get(key) ?? null

  let svgMarkup: string
  try {
    const tex2svg = await getConverter()
    const container = tex2svg(latex, display)
    const match = container.match(/<svg[\s\S]*?<\/svg>/)
    if (!match) throw new Error('no svg')
    svgMarkup = match[0]
  } catch {
    cache.set(key, null)
    return null
  }

  // Resolve the SVG's ex-based size to pixels using a 16px font context.
  const holder = document.createElement('div')
  holder.style.cssText = 'position:fixed;left:-99999px;top:0;font-size:16px;color:#111;'
  holder.innerHTML = svgMarkup
  document.body.appendChild(holder)
  try {
    const svgEl = holder.querySelector('svg') as SVGElement | null
    if (!svgEl) throw new Error('no svg element')
    const rect = svgEl.getBoundingClientRect()
    const width = Math.max(1, rect.width)
    const height = Math.max(1, rect.height)
    svgEl.setAttribute('width', String(width))
    svgEl.setAttribute('height', String(height))
    svgEl.setAttribute('style', 'color:#111')
    const serialized = new XMLSerializer().serializeToString(svgEl)
    const svgUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(serialized)
    const url = await svgToPng(svgUrl, width, height, 3)
    const res: RenderedMath = { url, width, height }
    cache.set(key, res)
    return res
  } catch {
    cache.set(key, null)
    return null
  } finally {
    document.body.removeChild(holder)
  }
}

async function svgToPng(svgUrl: string, width: number, height: number, scale: number): Promise<string> {
  const img = new Image()
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve()
    img.onerror = () => reject(new Error('svg load failed'))
    img.src = svgUrl
  })
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.round(width * scale))
  canvas.height = Math.max(1, Math.round(height * scale))
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('no canvas context')
  ctx.scale(scale, scale)
  ctx.drawImage(img, 0, 0, width, height)
  return canvas.toDataURL('image/png')
}
