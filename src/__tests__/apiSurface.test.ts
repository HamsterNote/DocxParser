import { DocumentParser } from '@hamster-note/document-parser'
import { IntermediateDocument } from '@hamster-note/types'
import {
  createDocxParserWarning,
  DocxDocument,
  DocxPage,
  DocxParser,
  DocxParserError,
  DocxParserWarning,
  RenderViews,
  type DocxParserWarning as DocxParserWarningValue,
  type RenderOptions
} from '../index.js'
import { makeParagraphsFixture } from '../testUtils/fixtureFactory.js'
import { makeSimpleIntermediateDocument } from '../testUtils/intermediateTestUtils.js'

describe('DocxParser public API surface', () => {
  it('exports parser, wrappers, render types, errors, and warnings', () => {
    const renderOptions: RenderOptions = { scale: 1, views: [RenderViews.TEXT] }
    const warning: DocxParserWarningValue = createDocxParserWarning({
      code: 'PAGINATION_FALLBACK',
      message: 'fallback'
    })

    expect(DocxParser).toBeDefined()
    expect(DocxDocument).toBeDefined()
    expect(DocxPage).toBeDefined()
    expect(DocxParserWarning.codes).toContain('PAGINATION_FALLBACK')
    expect(RenderViews.TEXT).toBe('TEXT')
    expect(RenderViews.THUMBNAIL).toBe('THUMBNAIL')
    expect(renderOptions.views).toEqual(['TEXT'])
    expect(new DocxParserError('DOCX_ENCODE_FAILED', 'failed').code).toBe(
      'DOCX_ENCODE_FAILED'
    )
    expect(warning.code).toBe('PAGINATION_FALLBACK')
  })

  it('defines HtmlParser-compatible parser metadata', () => {
    const parser = new DocxParser()

    expect(parser).toBeInstanceOf(DocumentParser)
    expect(DocxParser.exts).toEqual(['docx'])
    expect(DocxParser.ext).toBe('docx')
  })

  it('static encode returns DocxDocument wrapper around parsed intermediate document', async () => {
    const docxBuffer = await makeParagraphsFixture()
    const doc = await DocxParser.encode(docxBuffer)

    expect(doc).toBeInstanceOf(DocxDocument)
    expect(doc.getIntermediateDocument()).toBeInstanceOf(IntermediateDocument)
    expect(typeof doc.getTitle()).toBe('string')
    expect(doc.getTitle().length).toBeGreaterThan(0)
  })

  it('instance encode delegates to encodeToIntermediate', async () => {
    const docxBuffer = await makeParagraphsFixture()
    const parser = new DocxParser()
    const intermediateDocument = await parser.encode(docxBuffer)

    expect(intermediateDocument).toBeInstanceOf(IntermediateDocument)
    expect(typeof intermediateDocument.title).toBe('string')
    expect(intermediateDocument.title.length).toBeGreaterThan(0)
  })

  it('static and instance decode return ArrayBuffer', async () => {
    const intermediateDocument = makeSimpleIntermediateDocument()
    const parser = new DocxParser()

    await expect(DocxParser.decode(intermediateDocument)).resolves.toBeInstanceOf(
      ArrayBuffer
    )
    await expect(parser.decode(intermediateDocument)).resolves.toBeInstanceOf(
      ArrayBuffer
    )
  })
})
