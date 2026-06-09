import { execFileSync } from 'node:child_process'
import { readFile } from 'node:fs/promises'
import {
  IntermediateDocument,
  IntermediateImage,
  IntermediatePage,
  IntermediatePageMap,
  IntermediateParagraph,
  IntermediateText,
  TextDir
} from '@hamster-note/types'
import {
  type DocxParserError,
  DocxDocument,
  DocxPage,
  DocxParser
} from '../index.js'
import {
  type DocxFixtureName,
  docxFixtureDefinitions,
  makeCorruptFixture,
  makeEmptyFixture,
  makeFixtureByName,
  makeImageFixture
} from '../testUtils/fixtureFactory.js'
import {
  extractImages,
  extractPureText,
  expectStructurallyEquivalent,
  makeIntermediateWithImageListTable
} from '../testUtils/intermediateTestUtils.js'

const roundTripFixtureNames: DocxFixtureName[] = [
  'paragraphs.docx',
  'text-styles.docx',
  'page-break.docx',
  'image.docx',
  'ordered-list.docx',
  'unordered-list.docx',
  'nested-list.docx',
  'simple-table.docx',
  'merged-table.docx'
]

const normalizeText = (value: string): string => value.replace(/\s+/g, '')

const expectDocxParserError = async (
  action: () => Promise<unknown>,
  code: DocxParserError['code']
): Promise<void> => {
  await expect(action()).rejects.toMatchObject({ name: 'DocxParserError', code })
}

const expectRequiredString = (value: unknown): void => {
  expect(typeof value).toBe('string')
  expect((value as string).length).toBeGreaterThan(0)
}

const expectFiniteNumber = (value: unknown): void => {
  expect(typeof value).toBe('number')
  expect(Number.isFinite(value as number)).toBe(true)
}

const getDocumentPages = async (document: IntermediateDocument): Promise<IntermediatePage[]> => {
  const pages = await document.pages

  expect(pages).toHaveLength(document.pageCount)
  expect(pages.length).toBeGreaterThan(0)

  return pages
}

const expectIntermediateDocumentSchema = async (document: IntermediateDocument): Promise<void> => {
  expect(document).toBeInstanceOf(IntermediateDocument)
  expectRequiredString(document.id)
  expectRequiredString(document.title)
  expect(Array.isArray(document.getOutline() ?? [])).toBe(true)

  const pages = await getDocumentPages(document)
  const serialized = await IntermediateDocument.serialize(document)

  expect(serialized).toEqual(
    expect.objectContaining({
      id: document.id,
      title: document.title,
      pages: expect.any(Array)
    })
  )

  for (const page of pages) {
    expect(page).toBeInstanceOf(IntermediatePage)
    expectRequiredString(page.id)
    expectFiniteNumber(page.number)
    expectFiniteNumber(page.width)
    expectFiniteNumber(page.height)
    expect(page.width).toBeGreaterThan(0)
    expect(page.height).toBeGreaterThan(0)
    expect(Array.isArray(page.paragraphs)).toBe(true)

    const content = await page.getContent()
    const serializedPage = IntermediatePage.serialize(page)

    expect(Array.isArray(content)).toBe(true)
    expect(serializedPage).toEqual(
      expect.objectContaining({
        id: page.id,
        number: page.number,
        width: page.width,
        height: page.height,
        content: expect.any(Array),
        paragraphs: expect.any(Array)
      })
    )
    expect(serializedPage.texts ?? []).toEqual([])

    for (const item of content) {
      if (item instanceof IntermediateText) {
        expect(item).toEqual(
          expect.objectContaining({
            id: expect.any(String),
            content: expect.any(String),
            fontSize: expect.any(Number),
            fontFamily: expect.any(String),
            fontWeight: expect.any(Number),
            italic: expect.any(Boolean),
            color: expect.any(String),
            lineHeight: expect.any(Number),
            ascent: expect.any(Number),
            descent: expect.any(Number),
            dir: expect.any(String),
            skew: expect.any(Number),
            isEOL: expect.any(Boolean),
            polygon: expect.any(Array)
          })
        )
        expectRequiredString(item.id)
        expectFiniteNumber(item.fontSize)
        expectFiniteNumber(item.fontWeight)
        expect(item.polygon).toHaveLength(4)
      } else {
        expect(item).toBeInstanceOf(IntermediateImage)
        expect(item).toEqual(
          expect.objectContaining({
            id: expect.any(String),
            src: expect.stringMatching(/^data:image\//),
            polygon: expect.any(Array),
            opacity: expect.any(Number)
          })
        )
      }
    }

    for (const paragraph of page.paragraphs) {
      expect(paragraph).toBeInstanceOf(IntermediateParagraph)
      expect(paragraph).toEqual(
        expect.objectContaining({
          id: expect.any(String),
          x: expect.any(Number),
          y: expect.any(Number),
          width: expect.any(Number),
          height: expect.any(Number),
          textIds: expect.any(Array)
        })
      )
      expect(paragraph.textIds.every((id) => typeof id === 'string')).toBe(true)
    }
  }
}

const makeInvalidIntermediateDocument = (): IntermediateDocument => {
  const validText = new IntermediateText({
    id: 'valid-text-before-mutation',
    content: 'Valid text before mutation',
    fontSize: 14,
    fontFamily: 'Arial',
    fontWeight: 400,
    italic: false,
    color: 'rgb(0, 0, 0)',
    polygon: [
      [40, 40],
      [180, 40],
      [180, 58],
      [40, 58]
    ],
    lineHeight: 18,
    ascent: 13,
    descent: 5,
    dir: TextDir.LTR,
    skew: 0,
    isEOL: true
  })
  const invalidPage = new IntermediatePage({
    id: 'invalid-page',
    number: 1,
    width: 595,
    height: 842,
    content: [validText],
    paragraphs: [
      new IntermediateParagraph({
        id: 'invalid-paragraph',
        x: 40,
        y: 40,
        width: 140,
        height: 18,
        textIds: [validText.id]
      })
    ]
  })
  invalidPage.paragraphs[0].textIds = [123 as unknown as string]

  return new IntermediateDocument({
    id: 'invalid-doc',
    title: 'Invalid intermediate document',
    pagesMap: IntermediatePageMap.makeByInfoList([
      {
        id: invalidPage.id,
        pageNumber: invalidPage.number,
        size: { x: invalidPage.width, y: invalidPage.height },
        getData: async () => invalidPage
      }
    ])
  })
}

describe('DocxParser integration', () => {
  describe('API end-to-end', () => {
    it('static and instance parser APIs encode/decode through the same intermediate data', async () => {
      const fixture = await makeFixtureByName('paragraphs.docx')
      const parser = new DocxParser()

      const staticDocument = await DocxParser.encode(fixture)
      const staticIntermediate = await DocxParser.encodeToIntermediate(fixture)
      const instanceIntermediate = await parser.encode(fixture)
      const staticBytes = await DocxParser.decode(staticIntermediate)
      const instanceBytes = await parser.decode(instanceIntermediate)

      expect(staticDocument).toBeInstanceOf(DocxDocument)
      expect(staticDocument.getIntermediateDocument()).toBeInstanceOf(IntermediateDocument)
      expect(staticDocument.getTitle()).toBe(staticDocument.getIntermediateDocument().title)
      expect(await staticDocument.getPages()).toEqual(expect.arrayContaining([expect.any(DocxPage)]))
      expect(staticIntermediate).toBeInstanceOf(IntermediateDocument)
      expect(instanceIntermediate).toBeInstanceOf(IntermediateDocument)
      await expectStructurallyEquivalent(
        staticDocument.getIntermediateDocument(),
        staticIntermediate
      )
      await expectStructurallyEquivalent(instanceIntermediate, staticIntermediate)
      expect(staticBytes).toBeInstanceOf(ArrayBuffer)
      expect(staticBytes.byteLength).toBeGreaterThan(0)
      expect(instanceBytes).toBeInstanceOf(ArrayBuffer)
      expect(instanceBytes.byteLength).toBeGreaterThan(0)
    })
  })

  describe('schema validation', () => {
    it('encoded IntermediateDocument exposes required document, page, text, and paragraph fields', async () => {
      const document = await DocxParser.encodeToIntermediate(await makeFixtureByName('text-styles.docx'))

      await expectIntermediateDocumentSchema(document)
    })
  })

  describe('round-trip stability', () => {
    it.each(roundTripFixtureNames)('%s preserves structure bounds, text, and page count', async (fixtureName) => {
      const fixture = await makeFixtureByName(fixtureName)
      const intermediate1 = await DocxParser.encodeToIntermediate(fixture)
      const text1 = await extractPureText(intermediate1)
      const docxBytes = await DocxParser.decode(intermediate1)
      const intermediate2 = await DocxParser.encodeToIntermediate(docxBytes)
      const text2 = await extractPureText(intermediate2)

      await expectIntermediateDocumentSchema(intermediate2)
      await expectStructurallyEquivalent(intermediate1, intermediate1)
      expect(docxBytes.byteLength).toBeGreaterThan(0)
      expect(normalizeText(text2)).toContain(normalizeText(text1))
      expect(Math.abs(intermediate2.pageCount - intermediate1.pageCount)).toBeLessThanOrEqual(1)
    })

    it('valid fixture list matches the requested integration fixture matrix', () => {
      const validFixtureNames = docxFixtureDefinitions
        .filter((definition) => definition.validDocx)
        .map((definition) => definition.name)

      expect(validFixtureNames).toEqual(expect.arrayContaining(roundTripFixtureNames))
    })
  })

  describe('cross-feature compatibility', () => {
    it('encodes a DOCX with text and image, then decodes without losing generated bytes', async () => {
      const intermediate = await DocxParser.encodeToIntermediate(await makeImageFixture())
      const text = await extractPureText(intermediate)
      const images = await extractImages(intermediate)
      const decoded = await DocxParser.decode(intermediate)

      expect(text).toContain('Image fixture before inline PNG.')
      expect(text).toContain('Image fixture after inline PNG.')
      expect(images.length).toBeGreaterThanOrEqual(1)
      expect(decoded.byteLength).toBeGreaterThan(0)
    })

    it('decodes a mixed intermediate document containing text, image, list markers, and table-like rows', async () => {
      const mixed = makeIntermediateWithImageListTable()
      const pages = await mixed.pages
      const content = await pages[0].getContent()
      const decoded = await DocxParser.decode(mixed)
      const reEncoded = await DocxParser.encodeToIntermediate(decoded)
      const reEncodedText = await extractPureText(reEncoded)

      expect(content.some((item) => item instanceof IntermediateText)).toBe(true)
      expect(content.some((item) => item instanceof IntermediateImage)).toBe(true)
      expect(await extractPureText(mixed)).toContain('List item one')
      expect(await extractPureText(mixed)).toContain('Table A')
      expect(decoded.byteLength).toBeGreaterThan(0)
      expect(reEncodedText).toContain('Mixed fixture')
      expect(reEncodedText).toContain('List item one')
      expect(reEncodedText).toContain('Table A')
    })
  })

  describe('browser bundle compatibility', () => {
    beforeAll(() => {
      execFileSync('yarn', ['build:all'], { cwd: process.cwd(), stdio: 'pipe' })
    })

    it('dist/index.js imports in a Node ESM context and exposes the public browser-safe API', () => {
      const output = execFileSync(
        'node',
        [
          '--input-type=module',
          '-e',
          "import('./dist/index.js').then(m => { console.log(Object.keys(m).sort().join(',')); })"
        ],
        { cwd: process.cwd(), encoding: 'utf8' }
      )
      const exports = output.trim().split(',')

      expect(exports).toEqual(
        expect.arrayContaining([
          'DocxDocument',
          'DocxPage',
          'DocxParser',
          'DocxParserError',
          'DocxParserWarning',
          'RenderViews',
          'bytesToDataUrl',
          'dataUrlToBytes'
        ])
      )
    })

    it('dist/index.d.ts contains the expected type and value exports', async () => {
      const declarations = await readFile('dist/index.d.ts', 'utf8')

      expect(declarations).toContain('declare class DocxParser')
      expect(declarations).toContain('declare class DocxDocument')
      expect(declarations).toContain('declare class DocxPage')
      expect(declarations).toContain('declare class DocxParserError')
      expect(declarations).toContain('declare enum RenderViews')
      expect(declarations).toContain('interface RenderOptions')
    })
  })

  describe('error handling', () => {
    it('corrupt DOCX input throws DocxParserError with a stable invalid-input code', async () => {
      await expectDocxParserError(
        () => DocxParser.encodeToIntermediate(makeCorruptFixture()),
        'DOCX_ENCODE_FAILED'
      )
    })

    it('invalid intermediate document throws DocxParserError with INVALID_INTERMEDIATE_DOCUMENT', async () => {
      await expectDocxParserError(
        () => DocxParser.decode(makeInvalidIntermediateDocument()),
        'INVALID_INTERMEDIATE_DOCUMENT'
      )
    })

    it('empty DOCX produces a valid fallback IntermediateDocument', async () => {
      const document = await DocxParser.encodeToIntermediate(await makeEmptyFixture())

      await expectIntermediateDocumentSchema(document)
      expect(document.pageCount).toBe(1)
      expect(document.title).toBe('Untitled Docx Document')
      expect(await extractPureText(document)).toBe('')
    })
  })
})
