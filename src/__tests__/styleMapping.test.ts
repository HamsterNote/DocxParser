import {
  INTERMEDIATE_TO_DOCX_STYLE_MAP,
  MAMMOTH_HTML_TO_INTERMEDIATE_STYLE_MAP,
  intermediateParagraphToDocxParagraphOptions,
  intermediateTextToDocxTextRunOptions,
  mammothClassNameToIntermediateStyle
} from '../styleMapping.js'

describe('styleMapping', () => {
  it('maps mammoth HTML classes to Intermediate text fields', () => {
    expect(MAMMOTH_HTML_TO_INTERMEDIATE_STYLE_MAP['docx-heading-1']).toMatchObject({
      className: 'docx-heading-1',
      intermediateText: { fontSize: 24, fontWeight: 700 }
    })
    expect(mammothClassNameToIntermediateStyle('docx-emphasis')).toMatchObject({
      intermediateText: { italic: true }
    })
    expect(mammothClassNameToIntermediateStyle('unknown')).toBeUndefined()
  })

  it('maps Intermediate style families to docx style options', () => {
    expect(INTERMEDIATE_TO_DOCX_STYLE_MAP.heading1.textRun).toMatchObject({
      bold: true,
      size: 48,
      style: 'Heading1'
    })
    expect(INTERMEDIATE_TO_DOCX_STYLE_MAP.emphasis.textRun).toEqual({
      italics: true
    })
  })

  it('converts IntermediateText-like values to docx TextRun options', () => {
    expect(
      intermediateTextToDocxTextRunOptions({
        fontSize: 13,
        fontFamily: 'Arial',
        fontWeight: 700,
        italic: true,
        color: '#aabbcc'
      })
    ).toEqual({
      bold: true,
      italics: true,
      color: 'AABBCC',
      font: 'Arial',
      size: 26
    })
  })

  it('omits falsey TextRun flags while preserving color conversion', () => {
    expect(
      intermediateTextToDocxTextRunOptions({
        fontSize: 12,
        fontFamily: '',
        fontWeight: 400,
        italic: false,
        color: '#000000'
      })
    ).toEqual({
      bold: undefined,
      italics: undefined,
      color: '000000',
      font: undefined,
      size: 24
    })
  })

  it('converts paragraph layout hints to docx paragraph options', () => {
    expect(
      intermediateParagraphToDocxParagraphOptions({
        alignment: 'center',
        spacingAfter: 120,
        height: 240
      })
    ).toEqual({
      alignment: 'center',
      spacing: { after: 120, line: 240 }
    })
  })
})
