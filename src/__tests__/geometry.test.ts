import { A4_HEIGHT, A4_WIDTH, boxToPolygon, nextTextBox } from '../geometry.js'

describe('geometry', () => {
  it('exports A4 size in points', () => {
    expect(A4_WIDTH).toBe(595)
    expect(A4_HEIGHT).toBe(842)
  })

  it('converts a box to a four-point polygon', () => {
    expect(boxToPolygon(10, 20, 30, 40)).toEqual([
      [10, 20],
      [40, 20],
      [40, 60],
      [10, 60]
    ])
  })

  it('creates deterministic text boxes and cursor positions', () => {
    const first = nextTextBox('Hello', {
      fontSize: 10,
      lineHeight: 14,
      cursor: { x: 40, y: 40 }
    })

    expect(first).toEqual({
      x: 40,
      y: 40,
      width: 28,
      height: 14,
      nextCursor: { x: 72, y: 40 }
    })
  })

  it('wraps text boxes when they exceed the usable line width', () => {
    const box = nextTextBox('wrapped', {
      pageWidth: 100,
      marginLeft: 10,
      marginRight: 10,
      fontSize: 10,
      lineHeight: 12,
      cursor: { x: 70, y: 24 }
    })

    expect(box.x).toBe(10)
    expect(box.y).toBe(36)
    expect(box.nextCursor.y).toBe(36)
  })
})
