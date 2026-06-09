export const A4_WIDTH = 595
export const A4_HEIGHT = 842

export type Point = [number, number]
export type QuadPolygon = [Point, Point, Point, Point]

export type TextBox = {
  x: number
  y: number
  width: number
  height: number
}

export type NextTextBoxOptions = {
  cursor?: { x: number; y: number }
  pageWidth?: number
  marginLeft?: number
  marginRight?: number
  marginTop?: number
  fontSize?: number
  lineHeight?: number
  maxWidth?: number
  gap?: number
}

function finiteOr(value: number | undefined, fallback: number): number {
  return Number.isFinite(value) ? Number(value) : fallback
}

export function boxToPolygon(
  x: number,
  y: number,
  width: number,
  height: number
): QuadPolygon {
  return [
    [x, y],
    [x + width, y],
    [x + width, y + height],
    [x, y + height]
  ]
}

export function nextTextBox(
  text: string,
  options: NextTextBoxOptions = {}
): TextBox & { nextCursor: { x: number; y: number } } {
  const marginLeft = finiteOr(options.marginLeft, 40)
  const marginRight = finiteOr(options.marginRight, 40)
  const marginTop = finiteOr(options.marginTop, 40)
  const pageWidth = finiteOr(options.pageWidth, A4_WIDTH)
  const fontSize = finiteOr(options.fontSize, 12)
  const lineHeight = finiteOr(options.lineHeight, Math.round(fontSize * 1.2))
  const gap = finiteOr(options.gap, 4)
  const cursor = options.cursor ?? { x: marginLeft, y: marginTop }
  const availableWidth = Math.max(1, pageWidth - marginLeft - marginRight)
  const requestedWidth = Math.max(1, Math.ceil(text.length * fontSize * 0.55))
  const maxWidth = Math.min(
    finiteOr(options.maxWidth, availableWidth),
    availableWidth
  )
  const width = Math.min(requestedWidth, maxWidth)
  const startsNewLine = cursor.x > marginLeft && cursor.x + width > pageWidth - marginRight
  const x = startsNewLine ? marginLeft : cursor.x
  const y = startsNewLine ? cursor.y + lineHeight : cursor.y
  const nextX = x + width + gap
  const shouldWrapNext = nextX > pageWidth - marginRight

  return {
    x,
    y,
    width,
    height: lineHeight,
    nextCursor: {
      x: shouldWrapNext ? marginLeft : nextX,
      y: shouldWrapNext ? y + lineHeight : y
    }
  }
}
