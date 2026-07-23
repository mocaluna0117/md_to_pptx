/**
 * A free-floating text box layered over the flowing document (Docdown).
 * Coordinates are pixels relative to the page sheet's top-left; the .docx export
 * converts them to points and emits a Word text box anchored to the page.
 */
export interface DocBox {
  id: string
  x: number
  y: number
  w: number
  h: number
  /** Rich-text HTML of the box body (contentEditable). */
  html: string
}

let counter = 0

export function newDocBox(partial: Partial<DocBox> = {}): DocBox {
  counter += 1
  return {
    id: `box-${Date.now().toString(36)}-${counter}`,
    x: 56,
    y: 56,
    w: 260,
    h: 120,
    html: 'テキスト',
    ...partial,
  }
}
