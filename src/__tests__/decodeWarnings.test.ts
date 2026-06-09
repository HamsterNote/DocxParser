import {
  IntermediateDocument,
  IntermediateImage,
  IntermediatePage,
  IntermediatePageMap
} from '@hamster-note/types'
import { decodeIntermediateToDocx } from '../decode.js'
import { DOCX_PARSER_WARNING_CODES, createDocxParserWarning } from '../errors.js'
import { boxToPolygon } from '../geometry.js'
import { DocxParser } from '../index.js'

const makeImageDocument = (image: IntermediateImage): IntermediateDocument => {
  const page = new IntermediatePage({
    id: 'page-warning',
    number: 1,
    width: 595,
    height: 842,
    content: [image],
    paragraphs: []
  })

  return new IntermediateDocument({
    id: 'doc-warning',
    title: 'Decode warning fixture',
    pagesMap: IntermediatePageMap.makeByInfoList([
      {
        id: page.id,
        pageNumber: page.number,
        size: { x: page.width, y: page.height },
        getData: async () => page
      }
    ])
  })
}

const makeImage = (id: string, src: string): IntermediateImage =>
  new IntermediateImage({
    id,
    src,
    polygon: boxToPolygon(40, 40, 24, 24),
    opacity: 1
  })

describe('decode typed warnings', () => {
  it('unsupported image output format emits typed UNSUPPORTED_IMAGE_FORMAT warning', async () => {
    const warnings: ReturnType<typeof createDocxParserWarning>[] = []
    const document = makeImageDocument(makeImage('image-svg', 'data:image/svg+xml;base64,PHN2Zy8+'))

    const buffer = await decodeIntermediateToDocx(document, {
      onWarning: (warning) => warnings.push(warning)
    })

    expect(buffer).toBeInstanceOf(ArrayBuffer)
    expect(warnings).toEqual([
      expect.objectContaining({
        code: 'UNSUPPORTED_IMAGE_FORMAT',
        message: expect.stringContaining('unsupported DOCX output type image/svg+xml'),
        path: 'image:image-svg',
        feature: 'image/svg+xml'
      })
    ])
  })

  it('invalid image data URL emits typed MISSING_IMAGE warning', async () => {
    const warnings: ReturnType<typeof createDocxParserWarning>[] = []
    const document = makeImageDocument(makeImage('image-invalid', 'not-a-data-url'))

    const buffer = await decodeIntermediateToDocx(document, {
      onWarning: (warning) => warnings.push(warning)
    })

    expect(buffer).toBeInstanceOf(ArrayBuffer)
    expect(warnings).toEqual([
      expect.objectContaining({
        code: 'MISSING_IMAGE',
        message: expect.stringContaining('Skipped image image-invalid'),
        path: 'image:image-invalid',
        feature: 'image'
      })
    ])
  })

  it('DocxParser instance decode collects and exposes last decode warnings', async () => {
    const parser = new DocxParser()
    const document = makeImageDocument(makeImage('image-webp', 'data:image/webp;base64,UklGRg=='))

    const buffer = await parser.decode(document)
    const warnings = parser.getLastDecodeWarnings()

    expect(buffer).toBeInstanceOf(ArrayBuffer)
    expect(warnings).toEqual([
      expect.objectContaining({
        code: 'UNSUPPORTED_IMAGE_FORMAT',
        path: 'image:image-webp',
        feature: 'image/webp'
      })
    ])

    warnings.length = 0
    expect(parser.getLastDecodeWarnings()).toHaveLength(1)
  })

  it('complex OOXML exclusion warning codes exist and can be created', () => {
    const complexWarnings = [
      createDocxParserWarning({
        code: 'UNSUPPORTED_REVISION',
        message: 'Revision markup is excluded from v1 conversion.',
        feature: 'revision'
      }),
      createDocxParserWarning({
        code: 'UNSUPPORTED_EQUATION',
        message: 'Equation objects are excluded from v1 conversion.',
        feature: 'equation'
      }),
      createDocxParserWarning({
        code: 'UNSUPPORTED_SMART_ART',
        message: 'SmartArt objects are excluded from v1 conversion.',
        feature: 'smartArt'
      }),
      createDocxParserWarning({
        code: 'UNSUPPORTED_OLE',
        message: 'OLE embedded objects are excluded from v1 conversion.',
        feature: 'oleObject'
      })
    ]

    expect(DOCX_PARSER_WARNING_CODES).toEqual(
      expect.arrayContaining([
        'UNSUPPORTED_REVISION',
        'UNSUPPORTED_EQUATION',
        'UNSUPPORTED_SMART_ART',
        'UNSUPPORTED_OLE'
      ])
    )
    expect(complexWarnings.map((warning) => warning.code)).toEqual([
      'UNSUPPORTED_REVISION',
      'UNSUPPORTED_EQUATION',
      'UNSUPPORTED_SMART_ART',
      'UNSUPPORTED_OLE'
    ])
  })
})
