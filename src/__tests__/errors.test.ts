import {
  DOCX_PARSER_ERROR_CODES,
  DOCX_PARSER_WARNING_CODES,
  DocxParserError,
  createDocxParserWarning,
  type DocxParserWarning
} from '../errors.js'

describe('errors', () => {
  it('exposes all required error codes', () => {
    expect(DOCX_PARSER_ERROR_CODES).toEqual([
      'INVALID_DOCX_INPUT',
      'CORRUPT_DOCX_ZIP',
      'PASSWORD_PROTECTED_DOCX',
      'UNSUPPORTED_DOCX_TYPE',
      'DOCX_ENCODE_FAILED',
      'DOCX_DECODE_FAILED',
      'INVALID_INTERMEDIATE_DOCUMENT'
    ])
  })

  it('creates DocxParserError with code and cause', () => {
    const cause = new Error('zip failed')
    const error = new DocxParserError('CORRUPT_DOCX_ZIP', 'bad zip', {
      cause
    })

    expect(error).toBeInstanceOf(Error)
    expect(error.name).toBe('DocxParserError')
    expect(error.code).toBe('CORRUPT_DOCX_ZIP')
    expect(error.message).toBe('bad zip')
    expect((error as Error & { cause?: unknown }).cause).toBe(cause)
  })

  it('exposes all required warning codes', () => {
    expect(DOCX_PARSER_WARNING_CODES).toEqual([
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
    ])
  })

  it('keeps DocxParserWarning discriminated by code', () => {
    const warning: DocxParserWarning = createDocxParserWarning({
      code: 'UNSUPPORTED_FIELD',
      message: 'Field code was skipped',
      path: 'word/document.xml',
      feature: 'field'
    })

    expect(warning.code).toBe('UNSUPPORTED_FIELD')
    expect(warning.path).toBe('word/document.xml')
    expect(warning.feature).toBe('field')
  })
})
