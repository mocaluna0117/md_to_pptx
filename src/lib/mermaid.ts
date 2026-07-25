/**
 * Replace ```mermaid fenced blocks in Markdown with rendered PNG `<img>` tags
 * (data URIs), so diagrams ride the same image pipeline as attached images and
 * math through preview, slides, and every export. Mermaid is imported lazily
 * (it is heavy) and each diagram is cached by its source.
 */
export async function mermaidToImages(markdown: string): Promise<string> {
  if (!markdown.includes('```mermaid')) return markdown

  const fences: string[] = []
  const marked = markdown.replace(
    /```mermaid[^\n]*\n([\s\S]*?)```/g,
    (_m, code: string) => `@@MDMERMAID${fences.push(code.trim()) - 1}@@`,
  )
  const tags = await Promise.all(fences.map((code) => renderMermaidTag(code)))
  return marked.replace(/@@MDMERMAID(\d+)@@/g, (_m, i: string) => {
    const n = Number(i)
    // On a render failure, put the original fence back (shows as a code block).
    return tags[n] ?? '```mermaid\n' + fences[n] + '\n```'
  })
}

const cache = new Map<string, string | null>()
let seq = 0

// Lazily-initialized mermaid instance.
let mermaidMod: Promise<typeof import('mermaid').default> | null = null
function getMermaid() {
  if (!mermaidMod) {
    mermaidMod = import('mermaid').then(({ default: mermaid }) => {
      mermaid.initialize({
        startOnLoad: false,
        // strict also forces htmlLabels off, keeping the SVG canvas-rasterizable
        // (foreignObject labels would be dropped by canvas drawing).
        securityLevel: 'strict',
        theme: 'neutral',
        fontFamily: '"Helvetica Neue", Arial, "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif',
      })
      return mermaid
    })
  }
  return mermaidMod
}

async function renderMermaidTag(code: string): Promise<string | null> {
  if (cache.has(code)) return cache.get(code) ?? null
  let holder: HTMLDivElement | null = null
  try {
    const mermaid = await getMermaid()
    const { svg } = await mermaid.render(`mermaid-${(seq += 1)}`, code)

    // Measure the SVG at its natural size, then rasterize at 2x.
    holder = document.createElement('div')
    holder.style.cssText = 'position:fixed;left:-99999px;top:0;'
    holder.innerHTML = svg
    document.body.appendChild(holder)
    const svgEl = holder.querySelector('svg') as SVGElement | null
    if (!svgEl) throw new Error('no svg')
    svgEl.style.maxWidth = 'none'
    const rect = svgEl.getBoundingClientRect()
    const width = Math.max(1, Math.ceil(rect.width))
    const height = Math.max(1, Math.ceil(rect.height))
    svgEl.setAttribute('width', String(width))
    svgEl.setAttribute('height', String(height))
    const serialized = new XMLSerializer().serializeToString(svgEl)
    const svgUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(serialized)
    const png = await rasterize(svgUrl, width, height, 2)
    const tag = `<img src="${png}" alt="mermaid diagram" width="${width}" style="display:block;margin:0.7em auto;max-width:100%" />`
    cache.set(code, tag)
    return tag
  } catch {
    cache.set(code, null)
    return null
  } finally {
    if (holder) document.body.removeChild(holder)
  }
}

async function rasterize(svgUrl: string, width: number, height: number, scale: number): Promise<string> {
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
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.scale(scale, scale)
  ctx.drawImage(img, 0, 0, width, height)
  return canvas.toDataURL('image/png')
}
