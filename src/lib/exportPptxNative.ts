import PptxGen from 'pptxgenjs'
import { parseSlides, type Para, type Run, type SlideModel } from './markdownModel'

/** 16:9 slide in inches. */
const W = 10
const H = 5.625
const MARGIN = 0.5
const CONTENT_W = W - MARGIN * 2

export interface ExportOptions {
  fileName?: string
  onProgress?: (done: number, total: number) => void
}

interface LoadedImage {
  data: string
  w: number
  h: number
}

/**
 * Convert Marp Markdown into an **editable** .pptx: headings, paragraphs, lists,
 * code, quotes and tables become native PowerPoint text/table objects (not images),
 * so the text can be edited in PowerPoint. Visual fidelity to Marp themes is limited.
 */
export async function exportPptxNative(markdown: string, options: ExportOptions = {}): Promise<void> {
  const { fileName = 'slides.pptx', onProgress } = options
  const models = parseSlides(markdown)
  if (models.length === 0) {
    throw new Error('スライドが見つかりませんでした。Markdown を入力してください。')
  }

  // Preload every referenced image once (async), keyed by src.
  const cache = new Map<string, LoadedImage>()
  const srcs = [...new Set(models.flatMap((m) => m.images.map((i) => i.src)))]
  await Promise.all(
    srcs.map(async (src) => {
      const img = await loadImage(src)
      if (img) cache.set(src, img)
    }),
  )

  const pptx = new PptxGen()
  pptx.defineLayout({ name: 'W16x9', width: W, height: H })
  pptx.layout = 'W16x9'

  models.forEach((model, idx) => {
    const slide = pptx.addSlide()
    slide.background = { color: 'FFFFFF' }

    const { title, body } = extractTitle(model.paras)
    let y = MARGIN

    if (title) {
      slide.addText(titleProps(title.runs), {
        x: MARGIN,
        y,
        w: CONTENT_W,
        h: 0.9,
        valign: 'top',
        fontFace: 'Arial',
      })
      y += 1.0
    }

    const hasImages = model.images.some((i) => cache.has(i.src))
    const hasMedia = hasImages || model.tables.length > 0
    const mediaBand = hasMedia ? 2.4 : 0
    const bodyProps = buildBody(body)
    const bodyH = Math.max(0.4, H - MARGIN - y - mediaBand)

    if (bodyProps.length > 0) {
      slide.addText(bodyProps, { x: MARGIN, y, w: CONTENT_W, h: bodyH, valign: 'top', fontFace: 'Arial' })
    }

    if (hasMedia) {
      const top = bodyProps.length > 0 ? y + bodyH + 0.1 : y
      placeMedia(slide, model, cache, { x: MARGIN, y: top, w: CONTENT_W, h: H - MARGIN - top })
    }

    onProgress?.(idx + 1, models.length)
  })

  await pptx.writeFile({ fileName: ensureExt(fileName) })
}

function extractTitle(paras: Para[]): { title: Para | null; body: Para[] } {
  let idx = paras.findIndex((p) => p.kind === 'h1')
  if (idx < 0) idx = paras.findIndex((p) => p.kind === 'h2')
  if (idx < 0) return { title: null, body: paras }
  return { title: paras[idx], body: paras.filter((_, i) => i !== idx) }
}

function titleProps(runs: Run[]): PptxGen.TextProps[] {
  const base: PptxGen.TextPropsOptions = { fontSize: 30, bold: true, color: '12203A' }
  const rs = runs.length ? runs : [{ text: ' ' }]
  return rs.map((r, i) => ({
    text: r.text,
    options: {
      ...base,
      italic: r.italic,
      fontFace: r.code ? 'Courier New' : undefined,
      breakLine: i === rs.length - 1,
    },
  }))
}

function paraStyle(p: Para): PptxGen.TextPropsOptions {
  switch (p.kind) {
    case 'h1':
      return { fontSize: 26, bold: true, color: '12203A', paraSpaceBefore: 8, paraSpaceAfter: 4 }
    case 'h2':
      return { fontSize: 22, bold: true, color: '1A2233', paraSpaceBefore: 8, paraSpaceAfter: 4 }
    case 'h3':
      return { fontSize: 18, bold: true, color: '1A2233', paraSpaceBefore: 6, paraSpaceAfter: 2 }
    case 'h4':
      return { fontSize: 16, bold: true, color: '333333', paraSpaceBefore: 4, paraSpaceAfter: 2 }
    case 'li':
      return {
        fontSize: 16,
        color: '333333',
        bullet: p.ordered ? { type: 'number' } : true,
        indentLevel: p.indent,
        paraSpaceAfter: 2,
      }
    case 'code':
      return { fontSize: 13, fontFace: 'Courier New', color: '2E7D32', paraSpaceAfter: 0 }
    case 'quote':
      return { fontSize: 16, italic: true, color: '6B7280', indentLevel: Math.max(1, p.indent + 1), paraSpaceBefore: 4, paraSpaceAfter: 4 }
    default:
      return { fontSize: 16, color: '333333', paraSpaceAfter: 6 }
  }
}

function buildBody(paras: Para[]): PptxGen.TextProps[] {
  const out: PptxGen.TextProps[] = []
  for (const p of paras) {
    const base = paraStyle(p)
    const runs = p.runs.length ? p.runs : [{ text: ' ' } as Run]
    runs.forEach((r, i) => {
      out.push({
        text: r.text,
        options: {
          ...base,
          bold: r.bold ?? base.bold,
          italic: r.italic ?? base.italic,
          strike: r.strike,
          fontFace: r.code ? 'Courier New' : base.fontFace,
          color: r.code ? '2E7D32' : base.color,
          hyperlink: r.link ? { url: r.link, tooltip: r.link } : undefined,
          breakLine: i === runs.length - 1,
        },
      })
    })
  }
  return out
}

interface Region {
  x: number
  y: number
  w: number
  h: number
}

function placeMedia(slide: PptxGen.Slide, model: SlideModel, cache: Map<string, LoadedImage>, region: Region): void {
  let y = region.y

  for (const table of model.tables) {
    const rows: PptxGen.TableRow[] = []
    if (table.header.length) {
      rows.push(table.header.map((c) => ({ text: c, options: { bold: true, fill: { color: 'F1F3F5' } } })))
    }
    for (const r of table.rows) rows.push(r.map((c) => ({ text: c })))
    if (rows.length === 0) continue

    const rowH = 0.32
    slide.addTable(rows, {
      x: region.x,
      y,
      w: region.w,
      rowH,
      fontSize: 12,
      fontFace: 'Arial',
      color: '333333',
      valign: 'middle',
      border: { type: 'solid', pt: 1, color: 'D0D5DD' },
    })
    y += rowH * rows.length + 0.15
  }

  const imgs = model.images.map((i) => cache.get(i.src)).filter((v): v is LoadedImage => Boolean(v))
  const bandH = region.y + region.h - y
  if (imgs.length > 0 && bandH > 0.4) {
    layoutImagesRow(slide, imgs, { x: region.x, y, w: region.w, h: bandH })
  }
}

function layoutImagesRow(slide: PptxGen.Slide, imgs: LoadedImage[], r: Region): void {
  const gap = 0.15
  const gaps = gap * (imgs.length - 1)
  let h = r.h
  let widths = imgs.map((im) => h * (im.w / im.h))
  let total = widths.reduce((a, b) => a + b, 0)
  if (total + gaps > r.w) {
    h = (h * (r.w - gaps)) / total
    widths = imgs.map((im) => h * (im.w / im.h))
    total = widths.reduce((a, b) => a + b, 0)
  }
  let x = r.x + (r.w - total - gaps) / 2
  const y = r.y + (r.h - h) / 2
  imgs.forEach((im, i) => {
    slide.addImage({ data: im.data, x, y, w: widths[i], h })
    x += widths[i] + gap
  })
}

async function loadImage(src: string): Promise<LoadedImage | null> {
  try {
    const img = new Image()
    img.crossOrigin = 'anonymous'
    img.decoding = 'async'
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve()
      img.onerror = () => reject(new Error(`failed to load ${src}`))
      img.src = src
    })
    const w = img.naturalWidth || 800
    const h = img.naturalHeight || 600
    const canvas = document.createElement('canvas')
    canvas.width = w
    canvas.height = h
    const ctx = canvas.getContext('2d')
    if (!ctx) return src.startsWith('data:') ? { data: src, w, h } : null
    ctx.drawImage(img, 0, 0)
    return { data: canvas.toDataURL('image/png'), w, h }
  } catch {
    // Tainted canvas (cross-origin without CORS) or load failure.
    return src.startsWith('data:') ? { data: src, w: 800, h: 600 } : null
  }
}

function ensureExt(name: string): string {
  const trimmed = name.trim() || 'slides'
  return /\.pptx$/i.test(trimmed) ? trimmed : `${trimmed}.pptx`
}
