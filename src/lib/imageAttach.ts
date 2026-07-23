/** Attached local images, keyed by file name (basename), value = data URI. */
export type AttachedImages = Record<string, string>

export const IMAGE_EXT = /\.(png|jpe?g|gif|webp|svg|bmp|avif)$/i

/** Last path segment (also URL-decoded), e.g. "images/fig1.png" -> "fig1.png". */
export function basename(path: string): string {
  const clean = path.trim().replace(/^<|>$/g, '')
  let decoded = clean
  try {
    decoded = decodeURIComponent(clean)
  } catch {
    /* keep raw */
  }
  return decoded.split(/[\\/]/).pop() || decoded
}

/** True when src points to an external file (not a data URI / absolute URL / blob). */
export function isRelativeRef(src: string): boolean {
  return !/^(data:|https?:|\/\/|blob:)/i.test(src.trim())
}

/**
 * Rewrite relative image references in Markdown to attached data URIs, matched by
 * file name. Covers both `![alt](path)` and `<img src="path">`. Absolute URLs and
 * existing data URIs are left untouched. The original Markdown is not mutated.
 */
export function resolveImagePaths(markdown: string, images: AttachedImages): string {
  if (Object.keys(images).length === 0) return markdown

  let out = markdown.replace(/(!\[[^\]]*\]\(\s*)(<[^>]+>|[^)\s]+)/g, (whole, pre: string, rawSrc: string) => {
    const src = rawSrc.replace(/^<|>$/g, '')
    if (!isRelativeRef(src)) return whole
    const data = images[basename(src)]
    return data ? pre + data : whole
  })

  out = out.replace(/(<img\b[^>]*?\bsrc\s*=\s*["'])([^"']+)(["'])/gi, (whole, pre: string, src: string, post: string) => {
    if (!isRelativeRef(src)) return whole
    const data = images[basename(src)]
    return data ? pre + data + post : whole
  })

  return out
}

/** Read image File objects into { fileName: dataURI }. */
export async function readImageFiles(files: File[]): Promise<AttachedImages> {
  const out: AttachedImages = {}
  await Promise.all(
    files.map(
      (f) =>
        new Promise<void>((resolve) => {
          const reader = new FileReader()
          reader.onload = () => {
            out[f.name] = String(reader.result)
            resolve()
          }
          reader.onerror = () => resolve()
          reader.readAsDataURL(f)
        }),
    ),
  )
  return out
}
