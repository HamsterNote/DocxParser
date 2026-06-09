import type { ParserInput } from '@hamster-note/document-parser'
import type { IntermediateDocument } from '@hamster-note/types'
import mammoth from 'mammoth/mammoth.browser.js'
import {
  bytesToDataUrl,
  createUnsupportedImageFormatWarning,
  isSupportedImageContentType
} from './dataUrl.js'
import { DocxParserError, type DocxParserWarning } from './errors.js'
import { normalizeParserInput } from './input.js'
import {
  DOCX_PAGE_BREAK_CLASS,
  htmlToIntermediateDocument
} from './htmlToIntermediate.js'

type MammothMessage = {
  type: string
  message: string
  error?: unknown
}

type MammothImage = {
  contentType: string
  read: () => Promise<ArrayBuffer | ArrayBufferView>
}

type MammothConvertOptions = NonNullable<Parameters<typeof mammoth.convertToHtml>[1]>
type MammothImageConverter = NonNullable<MammothConvertOptions['convertImage']>

export type EncodeDocxOptions = {
  mammothOptions?: Record<string, unknown>
  title?: string
}

export type EncodeDocxResult = {
  document: IntermediateDocument
  warnings: DocxParserWarning[]
}

const PAGE_BREAK_STYLE_NAME = 'HamsterNote Page Break'
const FONT_SIZE_STYLE_PREFIX = 'HamsterNote Font Size '
const FONT_FAMILY_STYLE_PREFIX = 'HamsterNote Font Family '
const FONT_STYLE_PREFIX = 'HamsterNote Font '
const knownFontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36]
const knownFontFamilies = ['Arial', 'Calibri', 'Courier New', 'Georgia', 'Times New Roman', 'SimSun']

const DEFAULT_STYLE_MAP = [
  "p[style-name='Title'] => p.docx-title:fresh",
  "p[style-name='Heading 1'] => h1.docx-heading-1:fresh",
  "p[style-name='Heading 2'] => h2.docx-heading-2:fresh",
  "p[style-name='Heading 3'] => h3.docx-heading-3:fresh",
  "p[style-name='Heading 4'] => h4:fresh",
  "p[style-name='Heading 5'] => h5:fresh",
  "p[style-name='Heading 6'] => h6:fresh",
  "p[style-name='Quote'] => blockquote.docx-quote:fresh",
  "p[style-name='Code'] => pre.docx-code:fresh",
  `p[style-name='${PAGE_BREAK_STYLE_NAME}'] => p.${DOCX_PAGE_BREAK_CLASS}:fresh`,
  'b => strong.docx-strong',
  'i => em.docx-emphasis',
  'u => span.docx-underline',
  'strike => s'
]

for (const size of knownFontSizes) {
  DEFAULT_STYLE_MAP.push(
    `r[style-name='${FONT_SIZE_STYLE_PREFIX}${size}'] => span.docx-font-size-${size}`
  )
}

for (const family of knownFontFamilies) {
  DEFAULT_STYLE_MAP.push(
    `r[style-name='${FONT_FAMILY_STYLE_PREFIX}${family}'] => span.docx-font-family-${family.replace(/\s+/g, '-')}`
  )
}

for (const size of knownFontSizes) {
  for (const family of knownFontFamilies) {
    DEFAULT_STYLE_MAP.push(
      `r[style-name='${FONT_STYLE_PREFIX}${size} ${family}'] => span.docx-font-size-${size}.docx-font-family-${family.replace(/\s+/g, '-')}`
    )
  }
}

function hashArrayBuffer(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer)
  let hash = 0x811c9dc5

  for (const byte of bytes) {
    hash ^= byte
    hash = Math.imul(hash, 0x01000193) >>> 0
  }

  return hash.toString(16).padStart(8, '0').slice(0, 8)
}

function warningCodeForMammothMessage(message: MammothMessage): DocxParserWarning['code'] {
  const normalized = message.message.toLowerCase()
  if (normalized.includes('comment')) return 'UNSUPPORTED_COMMENT'
  if (normalized.includes('footnote') || normalized.includes('endnote')) return 'UNSUPPORTED_FOOTNOTE'
  if (normalized.includes('field')) return 'UNSUPPORTED_FIELD'
  if (normalized.includes('image')) return 'MISSING_IMAGE'
  return 'UNSUPPORTED_FIELD'
}

function mammothMessagesToWarnings(messages: MammothMessage[]): DocxParserWarning[] {
  return messages.map((message) => ({
    code: warningCodeForMammothMessage(message),
    message: message.message,
    feature: `mammoth:${message.type}`
  }))
}

function bytesFromImageReadResult(result: ArrayBuffer | ArrayBufferView): Uint8Array {
  if (result instanceof ArrayBuffer) return new Uint8Array(result)

  return new Uint8Array(result.buffer, result.byteOffset, result.byteLength)
}

function createMissingImageWarning(contentType: string, cause: unknown): DocxParserWarning {
  const reason = cause instanceof Error ? `: ${cause.message}` : ''

  return {
    code: 'MISSING_IMAGE',
    message: `Unable to read DOCX embedded image${reason}`,
    feature: contentType
  }
}

function createImageConverter(warnings: DocxParserWarning[]): MammothImageConverter {
  return mammoth.images.imgElement(async (image: MammothImage) => {
    const contentType = image.contentType

    if (!isSupportedImageContentType(contentType)) {
      warnings.push(createUnsupportedImageFormatWarning(contentType))
      return { src: '' }
    }

    try {
      const bytes = bytesFromImageReadResult(await image.read())
      return { src: bytesToDataUrl(bytes, contentType) }
    } catch (error) {
      warnings.push(createMissingImageWarning(contentType, error))
      return { src: '' }
    }
  })
}

function containsPageBreak(element: unknown): boolean {
  const candidate = element as { breakType?: string; children?: unknown[] }
  return (
    candidate.breakType === 'page' ||
    (Array.isArray(candidate.children) && candidate.children.some(containsPageBreak))
  )
}

function markPageBreakParagraphs(element: unknown): unknown {
  const candidate = element as {
    type?: string
    children?: unknown[]
    styleId?: string
    styleName?: string
  }

  if (Array.isArray(candidate.children)) {
    candidate.children = candidate.children.map(markPageBreakParagraphs)
  }

  if (candidate.type !== 'paragraph' || !Array.isArray(candidate.children)) return candidate

  const hasPageBreak = candidate.children.some(containsPageBreak)

  if (!hasPageBreak) return candidate

  return {
    ...candidate,
    styleId: 'HamsterNotePageBreak',
    styleName: PAGE_BREAK_STYLE_NAME,
    children: [
      {
        type: 'run',
        children: [{ type: 'text', value: 'DOCX_PAGE_BREAK' }]
      }
    ]
  }
}

function markRunStyles(element: unknown): unknown {
  const candidate = element as {
    type?: string
    children?: unknown[]
    styleId?: string
    styleName?: string
    font?: string | null
    fontSize?: number | null
  }

  if (Array.isArray(candidate.children)) {
    candidate.children = candidate.children.map(markRunStyles)
  }

  if (candidate.type !== 'run') return candidate
  if (candidate.fontSize && candidate.font) {
    return {
      ...candidate,
      styleId: `HamsterNoteFont${candidate.fontSize}${candidate.font.replace(/\s+/g, '')}`,
      styleName: `${FONT_STYLE_PREFIX}${candidate.fontSize} ${candidate.font}`
    }
  }
  if (candidate.fontSize) {
    return {
      ...candidate,
      styleId: `HamsterNoteFontSize${candidate.fontSize}`,
      styleName: `${FONT_SIZE_STYLE_PREFIX}${candidate.fontSize}`
    }
  }
  if (candidate.font) {
    return {
      ...candidate,
      styleId: `HamsterNoteFontFamily${candidate.font.replace(/\s+/g, '')}`,
      styleName: `${FONT_FAMILY_STYLE_PREFIX}${candidate.font}`
    }
  }

  return candidate
}

function transformDocumentForEncode(element: unknown): unknown {
  return markRunStyles(markPageBreakParagraphs(element))
}

export async function encodeDocxToIntermediate(
  input: ParserInput,
  options: EncodeDocxOptions = {}
): Promise<EncodeDocxResult> {
  const arrayBuffer = await normalizeParserInput(input)
  const sourceHash = hashArrayBuffer(arrayBuffer)
  const imageWarnings: DocxParserWarning[] = []

  try {
    const result = await mammoth.convertToHtml(
      { arrayBuffer },
      {
        includeDefaultStyleMap: true,
        ignoreEmptyParagraphs: false,
        styleMap: DEFAULT_STYLE_MAP,
        convertImage: createImageConverter(imageWarnings),
        transformDocument: transformDocumentForEncode,
        ...(options.mammothOptions ?? {})
      }
    )

    return htmlToIntermediateDocument(result.value, {
      documentId: `docx-${sourceHash}`,
      sourceHash,
      coreTitle: options.title,
      fallbackWarnings: [
        ...imageWarnings,
        ...mammothMessagesToWarnings(result.messages as MammothMessage[])
      ]
    })
  } catch (error) {
    throw new DocxParserError('DOCX_ENCODE_FAILED', 'Failed to encode DOCX to IntermediateDocument', {
      cause: error
    })
  }
}

export const encodeDocxStyleMap = DEFAULT_STYLE_MAP
export { htmlToIntermediateDocument }
