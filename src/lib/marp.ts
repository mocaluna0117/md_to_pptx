import { Marp } from '@marp-team/marp-core'

// Preview uses the default inline-SVG output: each slide is an <svg viewBox="0 0 1280 720">
// that scales responsively to its container — ideal for a live preview pane.
const previewMarp = new Marp()

// Export uses plain <section> output (no nested <foreignObject>), which rasterizes far
// more reliably with html-to-image. Slides stay at their native 1280x720 size.
const exportMarp = new Marp({ inlineSVG: false })

export interface PreviewResult {
  html: string
  css: string
}

export function renderPreview(markdown: string): PreviewResult {
  const { html, css } = previewMarp.render(markdown)
  return { html, css }
}

export interface SlidesResult {
  /** One HTML string per slide: `<section ...>...</section>` */
  slides: string[]
  /** Marpit stylesheet, scoped to `div.marpit > section` */
  css: string
}

export function renderSlides(markdown: string): SlidesResult {
  const { html, css } = exportMarp.render(markdown, { htmlAsArray: true })
  return { slides: html, css }
}
