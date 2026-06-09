import {
  IntermediateDocument,
  IntermediateImage,
  IntermediatePage,
  IntermediatePageMap,
  IntermediateParagraph,
  IntermediateText,
  TextDir
} from '@hamster-note/types'
import { A4_HEIGHT, A4_WIDTH, boxToPolygon } from '../geometry.js'

type StructuralOptions = {
  ignoreIds?: boolean
}

const normalizeId = (value: string, ignoreIds: boolean): string =>
  ignoreIds ? '<id>' : value

const serializePage = async (page: IntermediatePage, ignoreIds: boolean) => {
  const serialized = IntermediatePage.serialize(page)
  const content = await page.getContent()

  return {
    ...serialized,
    id: normalizeId(serialized.id, ignoreIds),
    content: content.map((item) => {
      if (item instanceof IntermediateText) {
        return {
          ...IntermediateText.serialize(item),
          id: normalizeId(item.id, ignoreIds)
        }
      }

      return {
        ...IntermediateImage.serialize(item),
        id: normalizeId(item.id, ignoreIds)
      }
    }),
    paragraphs: serialized.paragraphs?.map((paragraph) => ({
      ...paragraph,
      id: normalizeId(paragraph.id, ignoreIds),
      textIds: ignoreIds ? paragraph.textIds.map(() => '<id>') : paragraph.textIds
    })),
    thumbnail: serialized.thumbnail
      ? {
          ...serialized.thumbnail,
          id: normalizeId(serialized.thumbnail.id, ignoreIds)
        }
      : undefined,
    texts: undefined
  }
}

export const normalizeIntermediateDocument = async (
  doc: IntermediateDocument,
  options: StructuralOptions = {}
) => {
  const ignoreIds = options.ignoreIds ?? true
  const pages = await doc.pages
  const normalizedPages = await Promise.all(
    pages.map((page) => serializePage(page, ignoreIds))
  )

  return {
    id: normalizeId(doc.id, ignoreIds),
    title: doc.title,
    outline: doc.getOutline(),
    pages: normalizedPages
  }
}

export const extractPureText = async (doc: IntermediateDocument): Promise<string> => {
  const pages = await doc.pages
  const content = await Promise.all(pages.map((page) => page.getContent()))

  return content
    .flat()
    .filter((item): item is IntermediateText => item instanceof IntermediateText)
    .map((text) => text.content)
    .join('')
}

export const extractImages = async (
  doc: IntermediateDocument
): Promise<IntermediateImage[]> => {
  const pages = await doc.pages
  const content = await Promise.all(pages.map((page) => page.getContent()))

  return content
    .flat()
    .filter((item): item is IntermediateImage => item instanceof IntermediateImage)
}

export const expectStructurallyEquivalent = async (
  actual: IntermediateDocument,
  expected: IntermediateDocument,
  options: StructuralOptions = {}
): Promise<void> => {
  await expect(normalizeIntermediateDocument(actual, options)).resolves.toEqual(
    await normalizeIntermediateDocument(expected, options)
  )
}

const makeText = (
  id: string,
  content: string,
  x: number,
  y: number,
  width = content.length * 7
): IntermediateText =>
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

const makeDocument = (
  id: string,
  title: string,
  page: IntermediatePage
): IntermediateDocument =>
  new IntermediateDocument({
    id,
    title,
    pagesMap: IntermediatePageMap.makeByInfoList([
      {
        id: page.id,
        pageNumber: page.number,
        size: { x: page.width, y: page.height },
        getData: async () => page
      }
    ])
  })

export const makeSimpleIntermediateDocument = (): IntermediateDocument => {
  const title = makeText('text-title', 'Simple fixture title', 40, 40, 120)
  const body = makeText('text-body', 'Simple fixture body', 40, 70, 110)
  const page = new IntermediatePage({
    id: 'page-1',
    number: 1,
    width: A4_WIDTH,
    height: A4_HEIGHT,
    content: [title, body],
    paragraphs: [
      new IntermediateParagraph({
        id: 'paragraph-1',
        x: 40,
        y: 40,
        width: 120,
        height: 18,
        textIds: [title.id]
      }),
      new IntermediateParagraph({
        id: 'paragraph-2',
        x: 40,
        y: 70,
        width: 110,
        height: 18,
        textIds: [body.id]
      })
    ]
  })

  return makeDocument('doc-simple', 'Simple intermediate fixture', page)
}

export const makeIntermediateWithImageListTable = (): IntermediateDocument => {
  const heading = makeText('text-heading', 'Mixed fixture', 40, 40, 90)
  const listOne = makeText('text-list-1', 'List item one', 60, 78, 80)
  const listTwo = makeText('text-list-2', 'List item two', 60, 104, 80)
  const tableA = makeText('text-table-a', 'Table A', 40, 148, 55)
  const tableB = makeText('text-table-b', 'Table B', 140, 148, 55)
  const image = new IntermediateImage({
    id: 'image-1',
    src: 'data:image/png;base64,iVBORw0KGgo=',
    polygon: boxToPolygon(220, 40, 24, 24),
    opacity: 1
  })
  const page = new IntermediatePage({
    id: 'page-1',
    number: 1,
    width: A4_WIDTH,
    height: A4_HEIGHT,
    content: [heading, image, listOne, listTwo, tableA, tableB],
    paragraphs: [
      new IntermediateParagraph({
        id: 'paragraph-heading',
        x: 40,
        y: 40,
        width: 90,
        height: 18,
        textIds: [heading.id]
      }),
      new IntermediateParagraph({
        id: 'paragraph-list',
        x: 60,
        y: 78,
        width: 90,
        height: 44,
        textIds: [listOne.id, listTwo.id]
      }),
      new IntermediateParagraph({
        id: 'paragraph-table',
        x: 40,
        y: 148,
        width: 160,
        height: 18,
        textIds: [tableA.id, tableB.id]
      })
    ]
  })

  return makeDocument('doc-mixed', 'Mixed intermediate fixture', page)
}
