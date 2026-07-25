import { SLIDE_W, SLIDE_H, boxLineHeight, tableColFractions, tableRowFractions, type Slide, type TableEl } from './deck'
import { runsToHtml } from './richText'

/** Marp renders each slide at 1280x720; SLIDE_W inches maps to 1280px. */
export const PX_PER_IN = 1280 / SLIDE_W
const DECK_W = 1280
const DECK_H = Math.round((1280 * SLIDE_H) / SLIDE_W)

/** Native slide pixel size (the coordinate space boxes are positioned in). */
export const DECK_PX = { w: DECK_W, h: DECK_H }
/** Base style for a slide element in its native 1280×720 coordinate space. */
export const SLIDE_BASE_STYLE = `position:relative;width:${DECK_W}px;height:${DECK_H}px;overflow:hidden;font-family:Arial,"Noto Sans CJK JP","Yu Gothic",sans-serif;`
/** Inline-code font, needed when a slide element is rendered live (not via rasterize's stage). */
export const SLIDE_CODE_CSS = 'code{font-family:ui-monospace,Menlo,Consolas,monospace}'

/** Populate a slide element with its tables, images and boxes (shared by rasterize + slideshow). */
export function fillSlideElement(slide: HTMLElement, s: Slide): void {
  slide.style.background = `#${s.background || 'FFFFFF'}`
  slide.innerHTML = ''
  for (const tb of s.tables ?? []) {
    const wrap = document.createElement('div')
    wrap.style.cssText = pos(tb) + `overflow:hidden;box-sizing:border-box;font-size:${(tb.fontSize * PX_PER_IN) / 72}px;`
    wrap.innerHTML = tableHtml(tb)
    slide.appendChild(wrap)
  }
  for (const im of s.images ?? []) {
    const img = document.createElement('img')
    img.src = im.src
    img.style.cssText = pos(im) + 'object-fit:fill;'
    slide.appendChild(img)
  }
  for (const box of s.boxes) {
    const el = document.createElement('div')
    const fontFamily = box.pre
      ? 'ui-monospace,Menlo,Consolas,monospace'
      : box.fontFamily
        ? `"${box.fontFamily}"`
        : 'inherit'
    el.style.cssText =
      pos(box) +
      `overflow:hidden;padding:4px 6px;box-sizing:border-box;word-break:break-word;` +
      `white-space:${box.pre ? 'pre' : 'pre-wrap'};line-height:${boxLineHeight(box)};` +
      `font-size:${(box.fontSize * PX_PER_IN) / 72}px;text-align:${box.align};font-family:${fontFamily};` +
      `color:${box.color ? `#${box.color}` : '#111'};`
    el.innerHTML = runsToHtml(box.runs, PX_PER_IN) || '&nbsp;'
    slide.appendChild(el)
  }
}

function pos(r: { x: number; y: number; w: number; h: number }): string {
  return `position:absolute;left:${r.x * PX_PER_IN}px;top:${r.y * PX_PER_IN}px;width:${r.w * PX_PER_IN}px;height:${r.h * PX_PER_IN}px;`
}

/** Static HTML for a table (mirrors the visual editor's table rendering). */
function tableHtml(tb: TableEl): string {
  const fr = tableColFractions(tb)
  const rowFr = tableRowFractions(tb)
  const cols = fr.length
  const colGroup = fr.map((f) => `<col style="width:${(f * 100).toFixed(3)}%">`).join('')
  const body = tb.rows
    .map((row, r) => {
      const header = tb.header && r === 0
      const cells = Array.from({ length: cols }, (_, c) => {
        const cell = escapeHtml(row[c] ?? '') || '&nbsp;'
        const style =
          `border:1px solid #cfd6e0;padding:2px 6px;color:#111;vertical-align:middle;overflow:hidden;` +
          `word-break:break-word;line-height:1.25;${header ? 'font-weight:700;background:#eef2f7;' : ''}`
        return `<td style="${style}">${cell}</td>`
      }).join('')
      return `<tr style="height:${(rowFr[r] * 100).toFixed(3)}%">${cells}</tr>`
    })
    .join('')
  return `<table style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed;background:#fff"><colgroup>${colGroup}</colgroup><tbody>${body}</tbody></table>`
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}
