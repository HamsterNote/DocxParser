import { jest } from '@jest/globals'
import {
  IntermediateDocument,
  IntermediateOutline,
  IntermediateOutlineDestType,
  IntermediatePage,
  IntermediatePageMap,
  IntermediateText,
  TextDir
} from '@hamster-note/types'
import { DocxDocument } from '../DocxDocument.js'
import { DocxPage } from '../DocxPage.js'

type FakeCanvas = {
  width: number
  height: number
  style: Record<string, string>
  getContext: () => { fillStyle: string; fillRect: jest.Mock } | null
}

function makeText(id: string, content: string): IntermediateText {
  return new IntermediateText({
    id,
    content,
    fontSize: 12,
    fontFamily: 'Arial',
    fontWeight: 400,
    italic: false,
    color: '#000000',
    polygon: [
      [0, 0],
      [60, 0],
      [60, 16],
      [0, 16]
    ],
    lineHeight: 16,
    ascent: 10,
    descent: 4,
    dir: TextDir.LTR,
    skew: 0,
    isEOL: true
  })
}

function makeDocument(): IntermediateDocument {
  const pagesMap = IntermediatePageMap.makeByInfoList([
    {
      id: 'page-1',
      pageNumber: 1,
      size: { x: 612, y: 792 },
      getData: async () =>
        new IntermediatePage({
          id: 'page-1',
          number: 1,
          width: 612,
          height: 792,
          texts: [makeText('text-1', 'First page')]
        })
    },
    {
      id: 'page-2',
      pageNumber: 2,
      size: { x: 300, y: 400 },
      getData: async () =>
        new IntermediatePage({
          id: 'page-2',
          number: 2,
          width: 300,
          height: 400,
          texts: [makeText('text-2', 'Second page')]
        })
    }
  ])

  return new IntermediateDocument({
    id: 'doc-1',
    title: 'DOCX API',
    pagesMap,
    outline: [
      new IntermediateOutline({
        ...makeText('outline-1', 'Heading'),
        dest: { targetType: IntermediateOutlineDestType.PAGE, pageId: 'page-1' }
      })
    ]
  })
}

describe('DocxDocument', () => {
  it('wraps all pages as DocxPage instances', async () => {
    const doc = new DocxDocument(makeDocument())
    const pages = await doc.getPages()

    expect(pages).toHaveLength(2)
    expect(pages[0]).toBeInstanceOf(DocxPage)
    expect(pages[0].getNumber()).toBe(1)
    expect(pages[1].getNumber()).toBe(2)
  })

  it('returns a single page by page number or undefined', async () => {
    const doc = new DocxDocument(makeDocument())

    await expect(doc.getPage(2)).resolves.toBeInstanceOf(DocxPage)
    await expect(doc.getPage(99)).resolves.toBeUndefined()
  })

  it('returns first outline item and metadata accessors', async () => {
    const intermediateDocument = makeDocument()
    const doc = new DocxDocument(intermediateDocument)
    const outline = await doc.getOutline()

    expect(outline?.content).toBe('Heading')
    expect(doc.getTitle()).toBe('DOCX API')
    expect(doc.getId()).toBe('doc-1')
    expect(doc.getIntermediateDocument()).toBe(intermediateDocument)
  })

  it('returns undefined outline when document has no outline', async () => {
    const doc = new DocxDocument(
      new IntermediateDocument({
        id: 'empty-outline',
        title: 'No Outline',
        pagesMap: IntermediatePageMap.makeByInfoList([])
      })
    )

    await expect(doc.getOutline()).resolves.toBeUndefined()
  })

  it('creates blank cover canvas fallback using first page size', async () => {
    const fillRect = jest.fn()
    const canvas: FakeCanvas = {
      width: 0,
      height: 0,
      style: {},
      getContext: () => ({ fillStyle: '', fillRect })
    }
    const previousDocument = globalThis.document

    Object.defineProperty(globalThis, 'document', {
      configurable: true,
      value: {
        createElement: (tagName: string) => {
          if (tagName !== 'canvas') throw new Error(tagName)
          return canvas
        }
      }
    })

    try {
      const cover = await new DocxDocument(makeDocument()).getCover()

      expect(cover).toBe(canvas)
      expect(canvas.width).toBe(612)
      expect(canvas.height).toBe(792)
      expect(fillRect).toHaveBeenCalledWith(0, 0, 612, 792)
    } finally {
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: previousDocument
      })
    }
  })
})
