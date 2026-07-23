import { Marp } from '@marp-team/marp-core'
import markdownItCjkFriendly from 'markdown-it-cjk-friendly'

// Preview uses the default inline-SVG output: each slide is an <svg viewBox="0 0 1280 720">
// that scales responsively to its container — ideal for a live preview pane.
// The CJK-friendly plugin fixes **bold**/*italic* not applying when a delimiter sits next
// to full-width punctuation (e.g. `…相関（r=0.7）**を`), a common CJK CommonMark pitfall.
const previewMarp = new Marp().use(markdownItCjkFriendly)

// Export uses plain <section> output (no nested <foreignObject>), which rasterizes far
// more reliably with html-to-image. Slides stay at their native 1280x720 size.
const exportMarp = new Marp({ inlineSVG: false }).use(markdownItCjkFriendly)

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
