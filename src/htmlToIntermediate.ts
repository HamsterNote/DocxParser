import type { AnyNode, Element, Text } from 'domhandler'
import * as htmlparser2 from 'htmlparser2'
import {
  IntermediateDocument,
  IntermediateImage,
  IntermediateOutline,
  IntermediateOutlineDestType,
  IntermediatePage,
  IntermediatePageMap,
  IntermediateParagraph,
  IntermediateText,
  TextDir
} from '@hamster-note/types'
import { A4_HEIGHT, A4_WIDTH, boxToPolygon, nextTextBox } from './geometry.js'
import { mammothClassNameToIntermediateStyle } from './styleMapping.js'
import type { DocxParserWarning } from './errors.js'

export const DOCX_PAGE_BREAK_CLASS = 'docx-page-break'

type TextStyle = {
  fontSize: number
  fontFamily: string
  fontWeight: number
  italic: boolean
  color: string
}

type ParagraphBlock = {
  tagName: string
  text: string
  runs: TextRun[]
  layout?: BlockLayout
}

type ImageBlock = {
  src: string
}

type TextRun = {
  content: string
  style: TextStyle
}

type ContentBlock = ParagraphBlock | ImageBlock

type PageBlock = ContentBlock | { pageBreak: true }

type BlockLayout = {
  indentLevel?: number
  tableCell?: {
    rowIndex: number
    colIndex: number
    columnCount: number
    isLastInRow: boolean
  }
}

type ListContext = {
  type: 'ol' | 'ul'
  level: number
  counter: number
}

type CollectState = {
  tableMergeBestEffort: boolean
}

export type HtmlToIntermediateOptions = {
  documentId: string
  sourceHash: string
  fallbackWarnings?: DocxParserWarning[]
  coreTitle?: string
}

const blockTags = new Set([
  'p',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'pre',
  'blockquote'
])

const defaultTextStyle: TextStyle = {
  fontSize: 12,
  fontFamily: 'Arial',
  fontWeight: 400,
  italic: false,
  color: '#000000'
}

const FLOW_LEFT = 40
const FLOW_TOP = 40
const FLOW_GAP = 10
const IMAGE_FLOW_WIDTH = 320
const IMAGE_FLOW_HEIGHT = 180
const imageTagNames = new Set(['img'])

function isElement(node: AnyNode): node is Element {
  return node.type === 'tag'
}

function isText(node: AnyNode): node is Text {
  return node.type === 'text'
}

function isImageBlock(block: ContentBlock): block is ImageBlock {
  return 'src' in block
}

function classesOf(element: Element): string[] {
  return (element.attribs.class ?? '').split(/\s+/).filter(Boolean)
}

function hasClass(element: Element, className: string): boolean {
  return classesOf(element).includes(className)
}

function normalizeSpace(text: string): string {
  return text.replace(/\s+/g, ' ')
}

function parseSpan(value: string | undefined): number {
  const span = Number(value)
  return Number.isInteger(span) && span > 1 ? span : 1
}

function normalizeColor(value: string | undefined): string | undefined {
  if (!value) return undefined
  const color = value.trim()
  if (/^#[0-9a-f]{3}$/i.test(color)) {
    return `#${color.slice(1).split('').map((char) => `${char}${char}`).join('')}`.toUpperCase()
  }
  if (/^#[0-9a-f]{6}$/i.test(color)) return color.toUpperCase()
  const rgb = color.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i)
  if (!rgb) return undefined
  return `#${rgb.slice(1).map((part) => Number(part).toString(16).padStart(2, '0')).join('')}`.toUpperCase()
}

function mergeStyle(parent: TextStyle, element: Element): TextStyle {
  const style = { ...parent }
  const tagName = element.name.toLowerCase()

  if (tagName === 'strong' || tagName === 'b') style.fontWeight = 700
  if (tagName === 'th') style.fontWeight = 700
  if (tagName === 'em' || tagName === 'i') style.italic = true
  if (tagName === 'code') style.fontFamily = 'monospace'
  if (tagName === 'h1') Object.assign(style, { fontSize: 24, fontWeight: 700 })
  if (tagName === 'h2') Object.assign(style, { fontSize: 20, fontWeight: 700 })
  if (tagName === 'h3') Object.assign(style, { fontSize: 16, fontWeight: 700 })
  if (tagName === 'h4') Object.assign(style, { fontSize: 14, fontWeight: 700 })
  if (tagName === 'h5' || tagName === 'h6') style.fontWeight = 700

  for (const className of classesOf(element)) {
    const fontSize = className.match(/^docx-font-size-(\d+)$/)?.[1]
    const fontFamily = className.match(/^docx-font-family-(.+)$/)?.[1]
    if (fontSize) style.fontSize = Number(fontSize)
    if (fontFamily) style.fontFamily = fontFamily.replace(/-/g, ' ')

    const mapped = mammothClassNameToIntermediateStyle(className)?.intermediateText
    if (mapped?.fontSize) style.fontSize = mapped.fontSize
    if (mapped?.fontFamily) style.fontFamily = mapped.fontFamily
    if (mapped?.fontWeight) style.fontWeight = mapped.fontWeight
    if (typeof mapped?.italic === 'boolean') style.italic = mapped.italic
    if (mapped?.color) style.color = mapped.color
  }

  const inlineStyle = element.attribs.style ?? ''
  const color = normalizeColor(inlineStyle.match(/(?:^|;)\s*color\s*:\s*([^;]+)/i)?.[1])
  const fontSize = inlineStyle.match(/(?:^|;)\s*font-size\s*:\s*(\d+(?:\.\d+)?)px/i)?.[1]
  const fontFamily = inlineStyle.match(/(?:^|;)\s*font-family\s*:\s*([^;]+)/i)?.[1]

  if (color) style.color = color
  if (fontSize) style.fontSize = Number(fontSize)
  if (fontFamily) style.fontFamily = fontFamily.replace(/["']/g, '').split(',')[0].trim()

  return style
}

function textDirection(text: string): TextDir {
  return /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]/.test(text) ? TextDir.RTL : TextDir.LTR
}

function collectRuns(
  node: AnyNode,
  inheritedStyle: TextStyle,
  runs: TextRun[],
  options: { skipNestedLists?: boolean } = {}
): void {
  if (isText(node)) {
    const content = normalizeSpace(node.data)
    if (content.trim()) runs.push({ content, style: inheritedStyle })
    return
  }

  if (!isElement(node)) return
  if (hasClass(node, DOCX_PAGE_BREAK_CLASS)) return

  const tagName = node.name.toLowerCase()
  if (options.skipNestedLists && (tagName === 'ol' || tagName === 'ul')) return

  const style = mergeStyle(inheritedStyle, node)
  if (tagName === 'br') {
    runs.push({ content: '\n', style })
    return
  }

  for (const child of node.children) collectRuns(child, style, runs, options)
}

function findChildElements(element: Element, tagName: string): Element[] {
  return element.children.filter(
    (child): child is Element => isElement(child) && child.name.toLowerCase() === tagName
  )
}

function findDescendantElements(element: Element, tagNames: Set<string>): Element[] {
  const matches: Element[] = []

  for (const child of element.children) {
    if (!isElement(child)) continue

    if (tagNames.has(child.name.toLowerCase())) matches.push(child)
    matches.push(...findDescendantElements(child, tagNames))
  }

  return matches
}

function collectListItem(
  element: Element,
  blocks: PageBlock[],
  listStack: ListContext[],
  state: CollectState
): void {
  const currentList = listStack[listStack.length - 1]
  const prefix = currentList?.type === 'ol' ? `${currentList.counter++}.` : '•'
  const level = currentList?.level ?? 0
  const style = mergeStyle(defaultTextStyle, element)
  const runs: TextRun[] = [{ content: `${prefix} `, style: defaultTextStyle }]

  collectRuns(element, style, runs, { skipNestedLists: true })

  const text = runs.map((run) => run.content).join('').trim()
  if (text) blocks.push({ tagName: 'li', text, runs, layout: { indentLevel: level } })

  for (const child of element.children) {
    if (!isElement(child)) continue
    const tagName = child.name.toLowerCase()
    if (tagName === 'ol' || tagName === 'ul') collectList(child, blocks, listStack, state)
  }
}

function collectList(
  element: Element,
  blocks: PageBlock[],
  listStack: ListContext[],
  state: CollectState
): void {
  const tagName = element.name.toLowerCase() as 'ol' | 'ul'
  const listContext: ListContext = {
    type: tagName,
    level: listStack.length,
    counter: 1
  }

  for (const child of element.children) {
    if (!isElement(child)) continue
    if (child.name.toLowerCase() === 'li') {
      collectListItem(child, blocks, [...listStack, listContext], state)
    }
  }
}

function collectTable(element: Element, blocks: PageBlock[], state: CollectState): void {
  const rows = findDescendantElements(element, new Set(['tr']))
  const occupied = new Set<string>()
  const pendingRowspans = new Map<number, number>()
  const rowCells: ParagraphBlock[][] = []
  let maxColumns = 0

  rows.forEach((row, rowIndex) => {
    const cells = findChildElements(row, 'td').concat(findChildElements(row, 'th'))
    const currentRow: ParagraphBlock[] = []
    let colIndex = 0

    for (const [column, remainingRows] of [...pendingRowspans.entries()]) {
      if (remainingRows <= 1) pendingRowspans.delete(column)
      else pendingRowspans.set(column, remainingRows - 1)
    }

    for (const cell of cells) {
      while (occupied.has(`${rowIndex}:${colIndex}`)) colIndex++

      const colspan = parseSpan(cell.attribs.colspan)
      const rowspan = parseSpan(cell.attribs.rowspan)
      if (colspan > 1 || rowspan > 1) state.tableMergeBestEffort = true

      for (let rowOffset = 0; rowOffset < rowspan; rowOffset++) {
        for (let colOffset = 0; colOffset < colspan; colOffset++) {
          occupied.add(`${rowIndex + rowOffset}:${colIndex + colOffset}`)
        }
      }
      if (rowspan > 1) pendingRowspans.set(colIndex, rowspan)

      const runs: TextRun[] = []
      collectRuns(cell, mergeStyle(defaultTextStyle, cell), runs)
      const text = runs.map((run) => run.content).join('').trim()

      if (text) {
        currentRow.push({
          tagName: cell.name.toLowerCase(),
          text,
          runs,
          layout: {
            tableCell: {
              rowIndex,
              colIndex,
              columnCount: 1,
              isLastInRow: false
            }
          }
        })
      }

      colIndex += colspan
      maxColumns = Math.max(maxColumns, colIndex)
    }

    rowCells.push(currentRow)
  })

  const columnCount = Math.max(1, maxColumns)
  for (const row of rowCells) {
    row.forEach((block, index) => {
      if (block.layout?.tableCell) {
        block.layout.tableCell.columnCount = columnCount
        block.layout.tableCell.isLastInRow = index === row.length - 1
      }
      blocks.push(block)
    })
  }
}

function collectImageBlocks(element: Element, blocks: PageBlock[]): void {
  for (const image of findDescendantElements(element, imageTagNames)) {
    const src = image.attribs.src?.trim()
    if (src) blocks.push({ src })
  }
}

function collectBlocks(
  nodes: AnyNode[],
  blocks: PageBlock[],
  listStack: ListContext[] = [],
  state: CollectState = { tableMergeBestEffort: false }
): void {
  for (const node of nodes) {
    if (!isElement(node)) continue

    const tagName = node.name.toLowerCase()
    if (hasClass(node, DOCX_PAGE_BREAK_CLASS)) {
      blocks.push({ pageBreak: true })
      continue
    }

    if (tagName === 'ol' || tagName === 'ul') {
      collectList(node, blocks, listStack, state)
      continue
    }

    if (tagName === 'table') {
      collectTable(node, blocks, state)
      continue
    }

    if (tagName === 'img') {
      const src = node.attribs.src?.trim()
      if (src) blocks.push({ src })
      continue
    }

    if (tagName === 'li' && listStack.length > 0) {
      collectListItem(node, blocks, listStack, state)
      continue
    }

    if (blockTags.has(tagName)) {
      const runs: TextRun[] = []
      collectRuns(node, mergeStyle(defaultTextStyle, node), runs)
      const text = runs.map((run) => run.content).join('').trim()
      if (text) blocks.push({ tagName, text, runs })
      collectImageBlocks(node, blocks)
      continue
    }

    collectBlocks(node.children, blocks, listStack, state)
  }
}

function firstTitleCandidate(blocks: ParagraphBlock[], coreTitle?: string): string {
  const normalizedCoreTitle = coreTitle?.trim()
  if (normalizedCoreTitle) return normalizedCoreTitle

  return (
    blocks.find((block) => /^h[1-6]$/.test(block.tagName))?.text ||
    blocks.find((block) => block.text)?.text ||
    'Untitled Docx Document'
  )
}

function splitPages(blocks: PageBlock[]): ContentBlock[][] {
  const pages: ContentBlock[][] = [[]]

  for (const block of blocks) {
    if ('pageBreak' in block) {
      if (pages[pages.length - 1].length > 0) pages.push([])
      continue
    }
    pages[pages.length - 1].push(block)
  }

  return pages.filter((page, index) => page.length > 0 || index === 0)
}

function makeTextId(sourceHash: string, pageNumber: number, textIndex: number): string {
  return `docx-${sourceHash}-p${pageNumber}-t${textIndex}`
}

function makeParagraphId(sourceHash: string, pageNumber: number, paragraphIndex: number): string {
  return `docx-${sourceHash}-p${pageNumber}-para${paragraphIndex}`
}

function makeImageId(sourceHash: string, pageNumber: number, imageIndex: number): string {
  return `docx-${sourceHash}-p${pageNumber}-img${imageIndex}`
}

function imageWidthForPage(pageWidth: number): number {
  return Math.min(IMAGE_FLOW_WIDTH, Math.max(1, pageWidth - FLOW_LEFT * 2))
}

function materializePage(
  sourceHash: string,
  pageNumber: number,
  pageBlocks: ContentBlock[]
): { page: IntermediatePage; outline: IntermediateOutline[] } {
  const pageId = `docx-${sourceHash}-page-${pageNumber}`
  const content: Array<IntermediateText | IntermediateImage> = []
  const paragraphs: IntermediateParagraph[] = []
  const outline: IntermediateOutline[] = []
  let cursor = { x: FLOW_LEFT, y: FLOW_TOP }
  let textIndex = 0
  let imageIndex = 0

  pageBlocks.forEach((block, paragraphIndex) => {
    if (isImageBlock(block)) {
      const width = imageWidthForPage(A4_WIDTH)

      content.push(
        new IntermediateImage({
          id: makeImageId(sourceHash, pageNumber, imageIndex++),
          src: block.src,
          polygon: boxToPolygon(FLOW_LEFT, cursor.y, width, IMAGE_FLOW_HEIGHT),
          opacity: 1
        })
      )
      cursor = { x: FLOW_LEFT, y: cursor.y + IMAGE_FLOW_HEIGHT + FLOW_GAP }
      return
    }

    const tableCell = block.layout?.tableCell
    const indent = (block.layout?.indentLevel ?? 0) * 24
    const tableCellWidth = tableCell
      ? Math.max(48, Math.floor((A4_WIDTH - FLOW_LEFT * 2) / tableCell.columnCount))
      : 0
    const blockMarginLeft = tableCell ? FLOW_LEFT + tableCell.colIndex * tableCellWidth : FLOW_LEFT + indent
    const blockMaxWidth = tableCell ? tableCellWidth : A4_WIDTH - blockMarginLeft - FLOW_LEFT
    const blockCursor = tableCell
      ? { x: blockMarginLeft, y: cursor.y }
      : { x: blockMarginLeft, y: cursor.y }
    const textIds: string[] = []
    let paragraphMinX = Number.POSITIVE_INFINITY
    let paragraphMinY = Number.POSITIVE_INFINITY
    let paragraphMaxX = blockMarginLeft
    let paragraphMaxY = cursor.y
    let runCursor = blockCursor

    block.runs.forEach((run, runIndex) => {
      const normalizedContent = run.content.replace(/\n/g, ' ').trim()
      if (!normalizedContent) return

      const lineHeight = run.style.fontSize * 1.2
      const box = nextTextBox(normalizedContent, {
        cursor: runCursor,
        fontSize: run.style.fontSize,
        lineHeight,
        pageWidth: A4_WIDTH,
        marginLeft: blockMarginLeft,
        maxWidth: blockMaxWidth
      })
      runCursor = box.nextCursor

      const text = new IntermediateText({
        id: makeTextId(sourceHash, pageNumber, textIndex++),
        content: normalizedContent,
        fontSize: run.style.fontSize,
        fontFamily: run.style.fontFamily,
        fontWeight: run.style.fontWeight,
        italic: run.style.italic,
        color: run.style.color,
        polygon: boxToPolygon(box.x, box.y, box.width, box.height),
        lineHeight,
        ascent: run.style.fontSize * 0.8,
        descent: run.style.fontSize * 0.2,
        vertical: false,
        dir: textDirection(normalizedContent),
        skew: 0,
        isEOL: runIndex === block.runs.length - 1
      })

      content.push(text)
      textIds.push(text.id)
      paragraphMinX = Math.min(paragraphMinX, box.x)
      paragraphMinY = Math.min(paragraphMinY, box.y)
      paragraphMaxX = Math.max(paragraphMaxX, box.x + box.width)
      paragraphMaxY = Math.max(paragraphMaxY, box.y + box.height)
    })

    if (!textIds.length) return

    paragraphs.push(
      new IntermediateParagraph({
        id: makeParagraphId(sourceHash, pageNumber, paragraphIndex),
        x: paragraphMinX,
        y: paragraphMinY,
        width: paragraphMaxX - paragraphMinX,
        height: paragraphMaxY - paragraphMinY,
        textIds
      })
    )

    if (/^h[1-6]$/.test(block.tagName)) {
      const headingText = content.find(
        (item): item is IntermediateText => item instanceof IntermediateText && item.id === textIds[0]
      )
      if (headingText) {
        outline.push(
          new IntermediateOutline({
            ...IntermediateText.serialize(headingText),
            content: block.text,
            dest: { targetType: IntermediateOutlineDestType.PAGE, pageId }
          })
        )
      }
    }

    if (tableCell && !tableCell.isLastInRow) {
      cursor = { x: FLOW_LEFT, y: Math.max(cursor.y, paragraphMaxY) }
    } else {
      cursor = { x: FLOW_LEFT, y: Math.max(cursor.y, paragraphMaxY) + (tableCell ? 6 : FLOW_GAP) }
    }
  })

  return {
    page: new IntermediatePage({
      id: pageId,
      number: pageNumber,
      width: A4_WIDTH,
      height: A4_HEIGHT,
      content,
      paragraphs
    }),
    outline
  }
}

export function htmlToIntermediateDocument(
  html: string,
  options: HtmlToIntermediateOptions
): { document: IntermediateDocument; warnings: DocxParserWarning[] } {
  const dom = htmlparser2.parseDocument(html, { decodeEntities: true })
  const blocks: PageBlock[] = []
  const collectState: CollectState = { tableMergeBestEffort: false }
  collectBlocks(dom.children, blocks, [], collectState)

  const hadExplicitPageBreak = blocks.some((block) => 'pageBreak' in block)
  const pageBlocks = splitPages(blocks)
  const materialized = pageBlocks.map((page, index) =>
    materializePage(options.sourceHash, index + 1, page)
  )
  const pages = materialized.map((entry) => entry.page)
  const allParagraphBlocks = blocks.filter(
    (block): block is ParagraphBlock => !('pageBreak' in block) && !isImageBlock(block)
  )
  const warnings = [...(options.fallbackWarnings ?? [])]

  if (!hadExplicitPageBreak) {
    warnings.push({
      code: 'PAGINATION_FALLBACK',
      message: 'DOCX has no explicit page break markers; emitted one A4 fallback page.',
      feature: 'pagination'
    })
  }

  if (collectState.tableMergeBestEffort) {
    warnings.push({
      code: 'TABLE_MERGE_BEST_EFFORT',
      message: 'DOCX table contains merged cells; flattened visible cell text best-effort.',
      feature: 'table-merge'
    })
  }

  return {
    document: new IntermediateDocument({
      id: options.documentId,
      title: firstTitleCandidate(allParagraphBlocks, options.coreTitle),
      outline: materialized.flatMap((entry) => entry.outline),
      pagesMap: IntermediatePageMap.makeByInfoList(
        pages.map((page) => ({
          id: page.id,
          pageNumber: page.number,
          size: { x: page.width, y: page.height },
          getData: async () => page
        }))
      )
    }),
    warnings
  }
}
