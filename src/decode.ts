import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  PageBreak,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun
} from 'docx'
import {
  IntermediateDocument,
  IntermediateImage,
  IntermediatePage,
  IntermediateParagraph,
  IntermediateText
} from '@hamster-note/types'
import { dataUrlToBytes } from './dataUrl.js'
import {
  createDocxParserWarning,
  DocxParserError,
  type DocxParserWarning
} from './errors.js'
import {
  intermediateParagraphToDocxParagraphOptions,
  intermediateTextToDocxTextRunOptions,
  type DocxTextRunStyleOptions
} from './styleMapping.js'

export type DecodeIntermediateToDocxOptions = {
  creator?: string
  lastModifiedBy?: string
  onWarning?: (warning: DocxParserWarning) => void
}

type TextRunStyle = Omit<DocxTextRunStyleOptions, 'color'> & {
  color?: string
}

type DocxDocumentChild = Paragraph | Table

export type DetectedListMarker = {
  type: 'ordered' | 'unordered' | null
  level: number
  marker: string
  content: string
}

type PageRenderItem =
  | {
      kind: 'paragraph'
      paragraph?: IntermediateParagraph
      texts: IntermediateText[]
      x: number
      y: number
    }
  | {
      kind: 'image'
      image: IntermediateImage
      x: number
      y: number
    }

type TableCellCandidate = {
  text: string
  x: number
}

type TableRowCandidate = {
  cells: TableCellCandidate[]
  nextIndex: number
}

const alignmentByIntermediateAlignment = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  both: AlignmentType.JUSTIFIED
} as const

const invalidIntermediateDocument = (message: string): DocxParserError =>
  new DocxParserError('INVALID_INTERMEDIATE_DOCUMENT', message)

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function assertString(value: unknown, path: string): asserts value is string {
  if (typeof value !== 'string') {
    throw invalidIntermediateDocument(`${path} must be a string`)
  }
}

function assertFiniteNumber(value: unknown, path: string): asserts value is number {
  if (!isFiniteNumber(value)) {
    throw invalidIntermediateDocument(`${path} must be a finite number`)
  }
}

function assertIntermediateText(value: unknown, path: string): asserts value is IntermediateText {
  if (!(value instanceof IntermediateText)) {
    throw invalidIntermediateDocument(`${path} must be an IntermediateText`)
  }

  assertString(value.id, `${path}.id`)
  assertString(value.content, `${path}.content`)
  assertFiniteNumber(value.fontSize, `${path}.fontSize`)
  assertString(value.fontFamily, `${path}.fontFamily`)
  assertFiniteNumber(value.fontWeight, `${path}.fontWeight`)
  if (typeof value.italic !== 'boolean') {
    throw invalidIntermediateDocument(`${path}.italic must be a boolean`)
  }
  assertString(value.color, `${path}.color`)
}

function assertPolygon(value: unknown, path: string): void {
  if (!Array.isArray(value) || value.length !== 4) {
    throw invalidIntermediateDocument(`${path} must be a four-point polygon`)
  }

  value.forEach((point, pointIndex) => {
    if (!Array.isArray(point) || point.length !== 2) {
      throw invalidIntermediateDocument(`${path}[${pointIndex}] must be a coordinate pair`)
    }
    assertFiniteNumber(point[0], `${path}[${pointIndex}][0]`)
    assertFiniteNumber(point[1], `${path}[${pointIndex}][1]`)
  })
}

function assertIntermediateImage(value: unknown, path: string): asserts value is IntermediateImage {
  if (!(value instanceof IntermediateImage)) {
    throw invalidIntermediateDocument(`${path} must be an IntermediateImage`)
  }

  assertString(value.id, `${path}.id`)
  assertString(value.src, `${path}.src`)
  assertPolygon(value.polygon, `${path}.polygon`)
  assertFiniteNumber(value.opacity, `${path}.opacity`)
}

function assertIntermediateParagraph(
  value: unknown,
  path: string
): asserts value is IntermediateParagraph {
  if (!(value instanceof IntermediateParagraph)) {
    throw invalidIntermediateDocument(`${path} must be an IntermediateParagraph`)
  }

  assertString(value.id, `${path}.id`)
  assertFiniteNumber(value.x, `${path}.x`)
  assertFiniteNumber(value.y, `${path}.y`)
  assertFiniteNumber(value.width, `${path}.width`)
  assertFiniteNumber(value.height, `${path}.height`)
  if (!Array.isArray(value.textIds) || !value.textIds.every((id) => typeof id === 'string')) {
    throw invalidIntermediateDocument(`${path}.textIds must be a string array`)
  }
}

function assertIntermediatePage(value: unknown, path: string): asserts value is IntermediatePage {
  if (!(value instanceof IntermediatePage)) {
    throw invalidIntermediateDocument(`${path} must be an IntermediatePage`)
  }

  assertString(value.id, `${path}.id`)
  assertFiniteNumber(value.number, `${path}.number`)
  assertFiniteNumber(value.width, `${path}.width`)
  assertFiniteNumber(value.height, `${path}.height`)
  if (!Array.isArray(value.paragraphs)) {
    throw invalidIntermediateDocument(`${path}.paragraphs must be an array`)
  }
}

export async function validateIntermediateDocument(
  document: IntermediateDocument
): Promise<IntermediatePage[]> {
  if (!(document instanceof IntermediateDocument)) {
    throw invalidIntermediateDocument('document must be an IntermediateDocument')
  }

  assertString(document.id, 'document.id')
  assertString(document.title, 'document.title')

  const pages = await document.pages
  if (!Array.isArray(pages)) {
    throw invalidIntermediateDocument('document.pages must resolve to an array')
  }

  for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
    const page = pages[pageIndex]
    const pagePath = `document.pages[${pageIndex}]`
    assertIntermediatePage(page, pagePath)

    const content = await page.getContent()
    if (!Array.isArray(content)) {
      throw invalidIntermediateDocument(`${pagePath}.content must resolve to an array`)
    }

    content.forEach((item, contentIndex) => {
      if (item instanceof IntermediateText) {
        assertIntermediateText(item, `${pagePath}.content[${contentIndex}]`)
      }
      if (item instanceof IntermediateImage) {
        assertIntermediateImage(item, `${pagePath}.content[${contentIndex}]`)
      }
    })

    page.paragraphs.forEach((paragraph, paragraphIndex) => {
      assertIntermediateParagraph(paragraph, `${pagePath}.paragraphs[${paragraphIndex}]`)
    })
  }

  return [...pages].sort((a, b) => a.number - b.number)
}

function normalizeDocxColor(color: string | undefined): string | undefined {
  if (!color) return undefined

  const hex = /^#?([0-9a-f]{6})$/i.exec(color)
  if (hex) return hex[1].toUpperCase()

  const rgb = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/i.exec(color)
  if (!rgb) return undefined

  const channels = rgb.slice(1).map((channel) => Number(channel))
  if (channels.some((channel) => channel < 0 || channel > 255)) return undefined

  return channels.map((channel) => channel.toString(16).padStart(2, '0')).join('').toUpperCase()
}

export function createTextRunFromIntermediateText(text: IntermediateText): TextRun {
  const mappedStyle = intermediateTextToDocxTextRunOptions(text)
  const style: TextRunStyle = {
    ...mappedStyle,
    color: normalizeDocxColor(mappedStyle.color ?? text.color)
  }

  return new TextRun({
    text: text.content,
    ...style
  })
}

function createTextRunFromIntermediateTextContent(
  text: IntermediateText,
  content: string
): TextRun {
  const mappedStyle = intermediateTextToDocxTextRunOptions(text)
  const style: TextRunStyle = {
    ...mappedStyle,
    color: normalizeDocxColor(mappedStyle.color ?? text.color)
  }

  return new TextRun({
    text: content,
    ...style
  })
}

function getPolygonBounds(polygon: IntermediateImage['polygon'] | IntermediateText['polygon']) {
  const xs = polygon.map(([x]) => x)
  const ys = polygon.map(([, y]) => y)
  const minX = Math.min(...xs)
  const maxX = Math.max(...xs)
  const minY = Math.min(...ys)
  const maxY = Math.max(...ys)

  return {
    x: minX,
    y: minY,
    width: Math.max(1, maxX - minX),
    height: Math.max(1, maxY - minY)
  }
}

function clampListLevel(level: number): number {
  return Math.max(0, Math.min(8, Math.floor(level)))
}

export function detectListMarker(text: string): DetectedListMarker {
  const ordered = /^(\s*)(\d+[.)])\s+(.+)$/.exec(text)
  if (ordered) {
    return {
      type: 'ordered',
      level: clampListLevel(ordered[1].length / 2),
      marker: ordered[2],
      content: ordered[3]
    }
  }

  const unordered = /^(\s*)([•●○◦*-])\s+(.+)$/.exec(text)
  if (unordered) {
    return {
      type: 'unordered',
      level: clampListLevel(unordered[1].length / 2),
      marker: unordered[2],
      content: unordered[3]
    }
  }

  return { type: null, level: 0, marker: '', content: text }
}

function detectParagraphListMarker(
  texts: IntermediateText[],
  paragraph?: IntermediateParagraph
): DetectedListMarker {
  const joinedText = texts.map((text) => text.content).join('')
  const marker = detectListMarker(joinedText)
  if (!marker.type) return marker

  const sourceX = paragraph?.x ?? getPolygonBounds(texts[0].polygon).x
  return {
    ...marker,
    level: clampListLevel(sourceX / 24)
  }
}

function createTextRunsForListParagraph(
  texts: IntermediateText[],
  marker: DetectedListMarker
): TextRun[] {
  let remainingPrefixLength = marker.marker.length

  return texts.flatMap((text) => {
    let content = text.content
    if (remainingPrefixLength > 0) {
      const trimLength = Math.min(remainingPrefixLength, content.length)
      content = content.slice(trimLength)
      remainingPrefixLength -= trimLength
    }

    if (remainingPrefixLength === 0) {
      content = content.replace(/^\s+/, '')
      remainingPrefixLength = -1
    }

    return content ? [createTextRunFromIntermediateTextContent(text, content)] : []
  })
}

function getParagraphHeading(children: IntermediateText[]) {
  if (children.some((text) => text.fontSize >= 28 && text.fontWeight >= 600)) {
    return HeadingLevel.TITLE
  }
  if (children.some((text) => text.fontSize >= 24 && text.fontWeight >= 600)) {
    return HeadingLevel.HEADING_1
  }
  if (children.some((text) => text.fontSize >= 20 && text.fontWeight >= 600)) {
    return HeadingLevel.HEADING_2
  }
  if (children.some((text) => text.fontSize >= 16 && text.fontWeight >= 600)) {
    return HeadingLevel.HEADING_3
  }

  return undefined
}

export function createParagraphFromIntermediateTexts(
  texts: IntermediateText[],
  paragraph?: IntermediateParagraph
): Paragraph {
  const paragraphOptions = intermediateParagraphToDocxParagraphOptions(paragraph)
  const alignment = paragraphOptions.alignment
    ? alignmentByIntermediateAlignment[paragraphOptions.alignment]
    : undefined
  const listMarker = detectParagraphListMarker(texts, paragraph)
  const numbering = listMarker.type
    ? {
        reference: listMarker.type === 'ordered' ? 'hamster-note-ordered-list' : 'hamster-note-unordered-list',
        level: listMarker.level
      }
    : undefined

  return new Paragraph({
    children: listMarker.type
      ? createTextRunsForListParagraph(texts, listMarker)
      : texts.map(createTextRunFromIntermediateText),
    heading: getParagraphHeading(texts),
    alignment,
    numbering,
    spacing: paragraphOptions.spacing
  })
}

function dataUrlContentTypeToDocxImageType(
  contentType: string
): 'png' | 'jpg' | 'gif' | undefined {
  if (contentType === 'image/png') return 'png'
  if (contentType === 'image/jpeg') return 'jpg'
  if (contentType === 'image/gif') return 'gif'
  return undefined
}

function warnSkippedRichContent(
  options: DecodeIntermediateToDocxOptions,
  warning: DocxParserWarning
): void {
  options.onWarning?.(createDocxParserWarning(warning))
}

function getUnsupportedImageContentType(error: unknown): string | undefined {
  if (!(error instanceof DocxParserError) || error.code !== 'UNSUPPORTED_DOCX_TYPE') {
    return undefined
  }

  const match = /Unsupported image content type: (.+)$/i.exec(error.message)
  return match?.[1]
}

function createParagraphFromIntermediateImage(
  image: IntermediateImage,
  options: DecodeIntermediateToDocxOptions
): Paragraph | undefined {
  try {
    const { bytes, contentType } = dataUrlToBytes(image.src)
    const type = dataUrlContentTypeToDocxImageType(contentType)
    if (!type) {
      warnSkippedRichContent(options, {
        code: 'UNSUPPORTED_IMAGE_FORMAT',
        message: `Skipped image ${image.id}: unsupported DOCX output type ${contentType}`,
        path: `image:${image.id}`,
        feature: contentType
      })
      return undefined
    }

    const bounds = getPolygonBounds(image.polygon)
    return new Paragraph({
      children: [
        new ImageRun({
          type,
          data: bytes,
          transformation: {
            width: Math.round(bounds.width),
            height: Math.round(bounds.height)
          },
          altText: {
            title: image.id,
            description: `Intermediate image ${image.id}`,
            name: image.id
          }
        })
      ]
    })
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    const unsupportedContentType = getUnsupportedImageContentType(error)
    if (unsupportedContentType) {
      warnSkippedRichContent(options, {
        code: 'UNSUPPORTED_IMAGE_FORMAT',
        message: `Skipped image ${image.id}: unsupported DOCX output type ${unsupportedContentType}`,
        path: `image:${image.id}`,
        feature: unsupportedContentType.toLowerCase()
      })
      return undefined
    }

    warnSkippedRichContent(options, {
      code: 'MISSING_IMAGE',
      message: `Skipped image ${image.id}: ${detail}`,
      path: `image:${image.id}`,
      feature: 'image'
    })
    return undefined
  }
}

function textX(text: IntermediateText): number {
  return getPolygonBounds(text.polygon).x
}

function paragraphCellText(texts: IntermediateText[]): string {
  return texts.map((text) => text.content).join('').trim()
}

function getSeparatedTextCells(item: Extract<PageRenderItem, { kind: 'paragraph' }>): TableCellCandidate[] {
  const sortedTexts = [...item.texts].sort((left, right) => textX(left) - textX(right))
  if (sortedTexts.length < 2) return []

  const hasCellGap = sortedTexts.some((text, index) => {
    if (index === 0) return false
    const previousBounds = getPolygonBounds(sortedTexts[index - 1].polygon)
    return textX(text) - (previousBounds.x + previousBounds.width) >= 24
  })

  if (!hasCellGap) return []

  return sortedTexts.map((text) => ({ text: text.content.trim(), x: textX(text) }))
}

function isListParagraphItem(item: PageRenderItem): boolean {
  return item.kind === 'paragraph' && Boolean(detectParagraphListMarker(item.texts, item.paragraph).type)
}

function collectTableRowCandidate(items: PageRenderItem[], startIndex: number): TableRowCandidate | undefined {
  const first = items[startIndex]
  if (!first || first.kind !== 'paragraph' || isListParagraphItem(first)) return undefined

  const separatedCells = getSeparatedTextCells(first)
  if (separatedCells.length >= 2) {
    return { cells: separatedCells, nextIndex: startIndex + 1 }
  }

  const rowItems: Extract<PageRenderItem, { kind: 'paragraph' }>[] = [first]
  let nextIndex = startIndex + 1
  while (nextIndex < items.length) {
    const candidate = items[nextIndex]
    if (
      !candidate ||
      candidate.kind !== 'paragraph' ||
      isListParagraphItem(candidate) ||
      Math.abs(candidate.y - first.y) > 6
    ) {
      break
    }

    rowItems.push(candidate)
    nextIndex += 1
  }

  if (rowItems.length < 2) return undefined

  return {
    cells: rowItems
      .map((item) => ({ text: paragraphCellText(item.texts), x: item.x }))
      .sort((left, right) => left.x - right.x),
    nextIndex
  }
}

const visibleTableBorder = { style: BorderStyle.SINGLE, size: 1, color: 'BFBFBF' } as const

function createTableFromRows(rows: TableCellCandidate[][]): Table {
  return new Table({
    width: { size: 100, type: 'pct' },
    borders: {
      top: visibleTableBorder,
      bottom: visibleTableBorder,
      left: visibleTableBorder,
      right: visibleTableBorder,
      insideHorizontal: visibleTableBorder,
      insideVertical: visibleTableBorder
    },
    rows: rows.map(
      (row) =>
        new TableRow({
          children: row.map(
            (cell) =>
              new TableCell({
                children: [new Paragraph(cell.text || ' ')],
                margins: { top: 80, bottom: 80, left: 120, right: 120 }
              })
          )
        })
    )
  })
}

function createRenderItemsFromPageContent(
  page: IntermediatePage,
  content: Array<IntermediateText | IntermediateImage>
): PageRenderItem[] {
  const textById = new Map<string, IntermediateText>()
  const referencedTextIds = new Set<string>()

  content.forEach((item) => {
    if (item instanceof IntermediateText) {
      textById.set(item.id, item)
    }
  })

  const items: PageRenderItem[] = []

  page.paragraphs.forEach((paragraph) => {
    const texts = paragraph.textIds.flatMap((textId) => {
      const text = textById.get(textId)
      if (!text) return []
      referencedTextIds.add(textId)
      return [text]
    })

    if (texts.length > 0) {
      items.push({
        kind: 'paragraph',
        paragraph,
        texts,
        x: paragraph.x,
        y: paragraph.y
      })
    }
  })

  content.forEach((item) => {
    if (item instanceof IntermediateText && !referencedTextIds.has(item.id)) {
      const bounds = getPolygonBounds(item.polygon)
      items.push({ kind: 'paragraph', texts: [item], x: bounds.x, y: bounds.y })
    }
    if (item instanceof IntermediateImage) {
      const bounds = getPolygonBounds(item.polygon)
      items.push({ kind: 'image', image: item, x: bounds.x, y: bounds.y })
    }
  })

  return items.sort((left, right) => left.y - right.y || left.x - right.x)
}

async function createChildrenFromPage(
  page: IntermediatePage,
  options: DecodeIntermediateToDocxOptions
): Promise<DocxDocumentChild[]> {
  const content = await page.getContent()
  const items = createRenderItemsFromPageContent(page, content)
  const children: DocxDocumentChild[] = []

  for (let index = 0; index < items.length; index += 1) {
    const item = items[index]
    if (item.kind === 'image') {
      const imageParagraph = createParagraphFromIntermediateImage(item.image, options)
      if (imageParagraph) children.push(imageParagraph)
      continue
    }

    const firstRow = collectTableRowCandidate(items, index)
    if (firstRow) {
      const rows = [firstRow.cells]
      let nextIndex = firstRow.nextIndex

      while (nextIndex < items.length) {
        const nextRow = collectTableRowCandidate(items, nextIndex)
        if (!nextRow || nextRow.cells.length !== firstRow.cells.length) break
        rows.push(nextRow.cells)
        nextIndex = nextRow.nextIndex
      }

      children.push(createTableFromRows(rows))
      index = nextIndex - 1
      continue
    }

    children.push(createParagraphFromIntermediateTexts(item.texts, item.paragraph))
  }

  return children
}

export async function intermediateDocumentToDocxDocument(
  intermediateDocument: IntermediateDocument,
  options: DecodeIntermediateToDocxOptions = {}
): Promise<Document> {
  const pages = await validateIntermediateDocument(intermediateDocument)
  const children: DocxDocumentChild[] = []

  for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
    children.push(...(await createChildrenFromPage(pages[pageIndex], options)))

    if (pageIndex < pages.length - 1) {
      children.push(new Paragraph({ children: [new PageBreak()] }))
    }
  }

  return new Document({
    title: intermediateDocument.title,
    creator: options.creator ?? '@hamster-note/docx-parser',
    lastModifiedBy: options.lastModifiedBy ?? options.creator ?? '@hamster-note/docx-parser',
    numbering: {
      config: [
        {
          reference: 'hamster-note-ordered-list',
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: LevelFormat.DECIMAL,
            text: `%${level + 1}.`,
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720 + level * 360, hanging: 260 } } }
          }))
        },
        {
          reference: 'hamster-note-unordered-list',
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: LevelFormat.BULLET,
            text: level % 2 === 0 ? '•' : '◦',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720 + level * 360, hanging: 260 } } }
          }))
        }
      ]
    },
    sections: [
      {
        properties: {},
        children
      }
    ]
  })
}

export async function packDocxDocumentToArrayBuffer(document: Document): Promise<ArrayBuffer> {
  if (typeof Packer.toArrayBuffer === 'function') {
    return Packer.toArrayBuffer(document)
  }

  const blob = await Packer.toBlob(document)
  return blob.arrayBuffer()
}

export async function decodeIntermediateToDocx(
  document: IntermediateDocument,
  options: DecodeIntermediateToDocxOptions = {}
): Promise<ArrayBuffer> {
  const docxDocument = await intermediateDocumentToDocxDocument(document, options)
  return packDocxDocumentToArrayBuffer(docxDocument)
}
