import type {
  IntermediateParagraph,
  IntermediateText
} from '@hamster-note/types'

export type IntermediateTextStyleFields = Pick<
  IntermediateText,
  'fontSize' | 'fontFamily' | 'fontWeight' | 'italic' | 'color' | 'lineHeight'
>

export type IntermediateParagraphStyleFields = Partial<
  Pick<IntermediateParagraph, 'x' | 'y' | 'width' | 'height'>
> & {
  alignment?: 'left' | 'center' | 'right' | 'both'
  spacingAfter?: number
}

export type MammothHtmlStyleMapping = {
  className: string
  intermediateText?: Partial<IntermediateTextStyleFields>
  intermediateParagraph?: IntermediateParagraphStyleFields
}

export type DocxTextRunStyleOptions = {
  bold?: boolean
  italics?: boolean
  color?: string
  font?: string
  size?: number
  style?: string
}

export type DocxParagraphStyleOptions = {
  style?: string
  alignment?: 'left' | 'center' | 'right' | 'both'
  spacing?: { after?: number; line?: number }
}

export type IntermediateToDocxStyleMapping = {
  name: string
  whenText: Partial<IntermediateTextStyleFields>
  textRun: DocxTextRunStyleOptions
  paragraph?: DocxParagraphStyleOptions
}

export const MAMMOTH_HTML_TO_INTERMEDIATE_STYLE_MAP: Record<
  string,
  MammothHtmlStyleMapping
> = {
  'docx-title': {
    className: 'docx-title',
    intermediateText: { fontSize: 28, fontWeight: 700 },
    intermediateParagraph: { spacingAfter: 240 }
  },
  'docx-heading-1': {
    className: 'docx-heading-1',
    intermediateText: { fontSize: 24, fontWeight: 700 },
    intermediateParagraph: { spacingAfter: 180 }
  },
  'docx-heading-2': {
    className: 'docx-heading-2',
    intermediateText: { fontSize: 20, fontWeight: 700 },
    intermediateParagraph: { spacingAfter: 160 }
  },
  'docx-heading-3': {
    className: 'docx-heading-3',
    intermediateText: { fontSize: 16, fontWeight: 700 },
    intermediateParagraph: { spacingAfter: 120 }
  },
  'docx-strong': {
    className: 'docx-strong',
    intermediateText: { fontWeight: 700 }
  },
  'docx-emphasis': {
    className: 'docx-emphasis',
    intermediateText: { italic: true }
  },
  'docx-code': {
    className: 'docx-code',
    intermediateText: { fontFamily: 'monospace', fontSize: 11 }
  },
  'docx-quote': {
    className: 'docx-quote',
    intermediateText: { color: '#666666', italic: true },
    intermediateParagraph: { spacingAfter: 120 }
  }
}

export const INTERMEDIATE_TO_DOCX_STYLE_MAP: Record<
  string,
  IntermediateToDocxStyleMapping
> = {
  title: {
    name: 'title',
    whenText: { fontSize: 28, fontWeight: 700 },
    textRun: { bold: true, size: 56, style: 'Title' },
    paragraph: { style: 'Title', spacing: { after: 240 } }
  },
  heading1: {
    name: 'heading1',
    whenText: { fontSize: 24, fontWeight: 700 },
    textRun: { bold: true, size: 48, style: 'Heading1' },
    paragraph: { style: 'Heading1', spacing: { after: 180 } }
  },
  heading2: {
    name: 'heading2',
    whenText: { fontSize: 20, fontWeight: 700 },
    textRun: { bold: true, size: 40, style: 'Heading2' },
    paragraph: { style: 'Heading2', spacing: { after: 160 } }
  },
  heading3: {
    name: 'heading3',
    whenText: { fontSize: 16, fontWeight: 700 },
    textRun: { bold: true, size: 32, style: 'Heading3' },
    paragraph: { style: 'Heading3', spacing: { after: 120 } }
  },
  strong: {
    name: 'strong',
    whenText: { fontWeight: 700 },
    textRun: { bold: true }
  },
  emphasis: {
    name: 'emphasis',
    whenText: { italic: true },
    textRun: { italics: true }
  },
  code: {
    name: 'code',
    whenText: { fontFamily: 'monospace' },
    textRun: { font: 'Courier New' }
  }
}

function normalizeHexColor(color: string | undefined): string | undefined {
  if (!color) return undefined
  return color.replace(/^#/, '').toUpperCase()
}

export function mammothClassNameToIntermediateStyle(
  className: string
): MammothHtmlStyleMapping | undefined {
  return MAMMOTH_HTML_TO_INTERMEDIATE_STYLE_MAP[className]
}

export function intermediateTextToDocxTextRunOptions(
  text: Pick<
    IntermediateText,
    'fontSize' | 'fontFamily' | 'fontWeight' | 'italic' | 'color'
  >
): DocxTextRunStyleOptions {
  return {
    bold: text.fontWeight >= 600 || undefined,
    italics: text.italic || undefined,
    color: normalizeHexColor(text.color),
    font: text.fontFamily || undefined,
    size: Number.isFinite(text.fontSize) ? Math.round(text.fontSize * 2) : undefined
  }
}

export function intermediateParagraphToDocxParagraphOptions(
  paragraph: IntermediateParagraphStyleFields = {}
): DocxParagraphStyleOptions {
  return {
    alignment: paragraph.alignment,
    spacing: {
      after: paragraph.spacingAfter,
      line: paragraph.height
    }
  }
}
