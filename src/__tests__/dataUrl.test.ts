import {
  bytesToDataUrl,
  createUnsupportedImageFormatWarning,
  dataUrlToBytes,
  isSupportedImageContentType,
  isUnsupportedVectorImageContentType
} from '../dataUrl.js'

describe('dataUrl', () => {
  it.each([
    ['image/png', [137, 80, 78, 71]],
    ['image/jpeg', [255, 216, 255]],
    ['image/gif', [71, 73, 70]],
    ['image/svg+xml', Array.from(new TextEncoder().encode('<svg/>'))]
  ])('round-trips %s bytes', (contentType, rawBytes) => {
    const bytes = new Uint8Array(rawBytes)
    const dataUrl = bytesToDataUrl(bytes, contentType)
    const decoded = dataUrlToBytes(dataUrl)

    expect(dataUrl.startsWith(`data:${contentType};base64,`)).toBe(true)
    expect(decoded.contentType).toBe(contentType)
    expect(Array.from(decoded.bytes)).toEqual(Array.from(bytes))
  })

  it('decodes percent-encoded SVG data URLs', () => {
    const decoded = dataUrlToBytes('data:image/svg+xml,%3Csvg%3Eok%3C%2Fsvg%3E')

    expect(decoded.contentType).toBe('image/svg+xml')
    expect(new TextDecoder().decode(decoded.bytes)).toBe('<svg>ok</svg>')
  })

  it('rejects unsupported image content types', () => {
    expect(() => bytesToDataUrl(new Uint8Array([1]), 'image/webp')).toThrow(
      'Unsupported image content type'
    )
    expect(() => dataUrlToBytes('data:image/webp;base64,AQ==')).toThrow(
      'Unsupported image content type'
    )
  })

  it('rejects malformed data URLs', () => {
    expect(() => dataUrlToBytes('not-a-data-url')).toThrow('Invalid image data URL')
  })

  it('classifies EMF and WMF as unsupported vector image warnings', () => {
    expect(isSupportedImageContentType('image/png')).toBe(true)
    expect(isUnsupportedVectorImageContentType('image/x-emf')).toBe(true)
    expect(isUnsupportedVectorImageContentType('image/wmf')).toBe(true)
    expect(createUnsupportedImageFormatWarning('image/x-emf')).toMatchObject({
      code: 'UNSUPPORTED_IMAGE_FORMAT',
      feature: 'image/x-emf'
    })
  })
})
