import {
  IntermediateDocument,
  IntermediatePage,
  IntermediatePageMap,
  IntermediateParagraph,
  IntermediateText,
  TextDir
} from '@hamster-note/types'
import JSZip from 'jszip'
import mammoth from 'mammoth'
import { decodeIntermediateToDocx } from '../decode.js'
import { encodeDocxToIntermediate } from '../encode.js'
import { A4_HEIGHT, A4_WIDTH, boxToPolygon } from '../geometry.js'
import { makeFixtureByName } from '../testUtils/fixtureFactory.js'
import {
  extractImages,
  extractPureText,
  makeIntermediateWithImageListTable
} from '../testUtils/intermediateTestUtils.js'

const toMammothInput = (buffer: ArrayBuffer) => ({ buffer: Buffer.from(buffer) })

const normalizeTextContent = (value: string): string => value.replace(/\s+/g, '')

const docxXml = async (buffer: ArrayBuffer, path: string): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer)
  const file = zip.file(path)
  expect(file).toBeTruthy()

  return file?.async('string') ?? ''
}

const docxMediaEntries = async (buffer: ArrayBuffer): Promise<string[]> => {
  const zip = await JSZip.loadAsync(buffer)
  return Object.keys(zip.files).filter((path) => path.startsWith('word/media/'))
}

const text = (id: string, content: string, x: number, y: number, width = content.length * 7) =>
  new IntermediateText({
    id,
    content,
    fontSize: 14,
    fontFamily: 'Arial',
    fontWeight: 400,
    italic: false,
    color: 'rgb(0, 0, 0)',
    polygon: boxToPolygon(x, y, width, 18),
    lineHeight: 18,
    ascent: 13,
    descent: 5,
    vertical: false,
    dir: TextDir.LTR,
    skew: 0,
    isEOL: true
  })

const documentFromPages = (id: string, title: string, pages: IntermediatePage[]): IntermediateDocument =>
  new IntermediateDocument({
    id,
    title,
    pagesMap: IntermediatePageMap.makeByInfoList(
      pages.map((page) => ({
        id: page.id,
        pageNumber: page.number,
        size: { x: page.width, y: page.height },
        getData: async () => page
      }))
    )
  })

const makeListIntermediate = (): IntermediateDocument => {
  const items = [
    text('list-1', '1. Ordered item one', 24, 40, 130),
    text('list-2', '2. Ordered item two', 24, 68, 130),
    text('list-3', '• Bullet item one', 48, 96, 110)
  ]

  const page = new IntermediatePage({
    id: 'page-list',
    number: 1,
    width: A4_WIDTH,
    height: A4_HEIGHT,
    content: items,
    paragraphs: items.map(
      (item, index) =>
        new IntermediateParagraph({
          id: `paragraph-list-${index}`,
          x: index === 2 ? 48 : 24,
          y: 40 + index * 28,
          width: 140,
          height: 18,
          textIds: [item.id]
        })
    )
  })

  return documentFromPages('doc-list', 'List decode fixture', [page])
}

const makeTableIntermediate = (): IntermediateDocument => {
  const cells = [
    text('table-a', 'Header A', 40, 48, 55),
    text('table-b', 'Header B', 160, 48, 55),
    text('table-c', 'Cell A1', 40, 78, 45),
    text('table-d', 'Cell B1', 160, 78, 45)
  ]

  const page = new IntermediatePage({
    id: 'page-table',
    number: 1,
    width: A4_WIDTH,
    height: A4_HEIGHT,
    content: cells,
    paragraphs: cells.map(
      (cell, index) =>
        new IntermediateParagraph({
          id: `paragraph-table-${index}`,
          x: index % 2 === 0 ? 40 : 160,
          y: index < 2 ? 48 : 78,
          width: 70,
          height: 18,
          textIds: [cell.id]
        })
    )
  })

  return documentFromPages('doc-table', 'Table decode fixture', [page])
}

const makeMultiPageIntermediate = (): IntermediateDocument => {
  const first = text('page-one-text', 'Page one content.', 40, 40)
  const second = text('page-two-text', 'Page two content.', 40, 40)
  const pages = [
    new IntermediatePage({
      id: 'page-one',
      number: 1,
      width: A4_WIDTH,
      height: A4_HEIGHT,
      content: [first],
      paragraphs: [
        new IntermediateParagraph({
          id: 'paragraph-one',
          x: 40,
          y: 40,
          width: 110,
          height: 18,
          textIds: [first.id]
        })
      ]
    }),
    new IntermediatePage({
      id: 'page-two',
      number: 2,
      width: A4_WIDTH,
      height: A4_HEIGHT,
      content: [second],
      paragraphs: [
        new IntermediateParagraph({
          id: 'paragraph-two',
          x: 40,
          y: 40,
          width: 110,
          height: 18,
          textIds: [second.id]
        })
      ]
    })
  ]

  return documentFromPages('doc-pages', 'Pagination decode fixture', pages)
}

describe('decode rich content DOCX generation', () => {
  it('writes IntermediateImage content as embedded DOCX media', async () => {
    const intermediateDocument = makeIntermediateWithImageListTable()
    const buffer = await decodeIntermediateToDocx(intermediateDocument)
    const images = await extractImages(await encodeDocxToIntermediate(buffer).then((result) => result.document))

    expect(await docxMediaEntries(buffer)).toEqual(expect.arrayContaining([expect.stringMatching(/\.png$/)]))
    expect(images).toHaveLength(1)
    expect(images[0].src).toMatch(/^data:image\/png;base64,/)
  })

  it('turns marker-prefixed paragraphs into DOCX numbering and strips marker text', async () => {
    const buffer = await decodeIntermediateToDocx(makeListIntermediate())
    const documentXml = await docxXml(buffer, 'word/document.xml')
    const numberingXml = await docxXml(buffer, 'word/numbering.xml')
    const rawText = await mammoth.extractRawText(toMammothInput(buffer))

    expect(documentXml).toContain('w:numPr')
    expect(numberingXml).toContain('decimal')
    expect(numberingXml).toContain('bullet')
    expect(rawText.value).toContain('Ordered item one')
    expect(rawText.value).toContain('Bullet item one')
    expect(rawText.value).not.toContain('1. Ordered item one')
    expect(rawText.value).not.toContain('• Bullet item one')
  })

  it('groups same-row table-style text into visible DOCX tables', async () => {
    const buffer = await decodeIntermediateToDocx(makeTableIntermediate())
    const documentXml = await docxXml(buffer, 'word/document.xml')
    const rawText = await mammoth.extractRawText(toMammothInput(buffer))

    expect(documentXml).toContain('<w:tbl>')
    expect(documentXml.match(/<w:tr>/g)).toHaveLength(2)
    expect(documentXml.match(/<w:tc>/g)).toHaveLength(4)
    expect(documentXml).toContain('<w:tblBorders>')
    expect(rawText.value).toContain('Header A')
    expect(rawText.value).toContain('Cell B1')
  })

  it('inserts page breaks between Intermediate pages and re-encode preserves page text', async () => {
    const buffer = await decodeIntermediateToDocx(makeMultiPageIntermediate())
    const documentXml = await docxXml(buffer, 'word/document.xml')
    const reEncoded = await encodeDocxToIntermediate(buffer)

    expect(documentXml.match(/w:type="page"/g)).toHaveLength(1)
    expect(reEncoded.document.pageCount).toBe(2)
    expect(await extractPureText(reEncoded.document)).toContain('Page one content.')
    expect(await extractPureText(reEncoded.document)).toContain('Page two content.')
  })

  it('round-trips fixture DOCX through encode → decode → re-encode preserving text content', async () => {
    const original = await encodeDocxToIntermediate(await makeFixtureByName('simple-table.docx'))
    const expectedText = await extractPureText(original.document)
    const decodedBuffer = await decodeIntermediateToDocx(original.document)
    const reEncoded = await encodeDocxToIntermediate(decodedBuffer)

    expect(normalizeTextContent(await extractPureText(reEncoded.document))).toBe(
      normalizeTextContent(expectedText)
    )
  })
})
