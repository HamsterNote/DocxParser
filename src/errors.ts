export const DOCX_PARSER_ERROR_CODES = [
  'INVALID_DOCX_INPUT',
  'CORRUPT_DOCX_ZIP',
  'PASSWORD_PROTECTED_DOCX',
  'UNSUPPORTED_DOCX_TYPE',
  'DOCX_ENCODE_FAILED',
  'DOCX_DECODE_FAILED',
  'INVALID_INTERMEDIATE_DOCUMENT'
] as const

export type DocxParserErrorCode = (typeof DOCX_PARSER_ERROR_CODES)[number]

export class DocxParserError extends Error {
  readonly code: DocxParserErrorCode

  constructor(
    code: DocxParserErrorCode,
    message: string,
    options?: { cause?: unknown }
  ) {
    super(message)
    this.name = 'DocxParserError'
    this.code = code

    if (options && 'cause' in options) {
      ;(this as Error & { cause?: unknown }).cause = options.cause
    }
  }
}

export const DOCX_PARSER_WARNING_CODES = [
  'UNSUPPORTED_REVISION',
  'UNSUPPORTED_COMMENT',
  'UNSUPPORTED_FOOTNOTE',
  'UNSUPPORTED_FIELD',
  'UNSUPPORTED_EQUATION',
  'UNSUPPORTED_SMART_ART',
  'UNSUPPORTED_OLE',
  'MISSING_IMAGE',
  'UNSUPPORTED_IMAGE_FORMAT',
  'TABLE_MERGE_BEST_EFFORT',
  'PAGINATION_FALLBACK'
] as const

export type DocxParserWarningCode = (typeof DOCX_PARSER_WARNING_CODES)[number]

type WarningBase<TCode extends DocxParserWarningCode> = {
  code: TCode
  message: string
  path?: string
  feature?: string
}

export type DocxParserWarning = {
  [Code in DocxParserWarningCode]: WarningBase<Code>
}[DocxParserWarningCode]

export const DocxParserWarning = Object.freeze({
  codes: DOCX_PARSER_WARNING_CODES
})

export function createDocxParserWarning(
  warning: DocxParserWarning
): DocxParserWarning {
  return warning
}
