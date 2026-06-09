import { DocxParserError, type DocxParserWarning } from './errors.js'

export const SUPPORTED_IMAGE_CONTENT_TYPES = [
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/svg+xml'
] as const

export type SupportedImageContentType =
  (typeof SUPPORTED_IMAGE_CONTENT_TYPES)[number]

const UNSUPPORTED_VECTOR_CONTENT_TYPES = new Set([
  'image/emf',
  'image/x-emf',
  'image/wmf',
  'image/x-wmf'
])

function normalizeContentType(contentType: string): string {
  return contentType.trim().toLowerCase()
}

export function isSupportedImageContentType(
  contentType: string
): contentType is SupportedImageContentType {
  return SUPPORTED_IMAGE_CONTENT_TYPES.includes(
    normalizeContentType(contentType) as SupportedImageContentType
  )
}

export function createUnsupportedImageFormatWarning(
  contentType: string,
  path?: string
): DocxParserWarning {
  return {
    code: 'UNSUPPORTED_IMAGE_FORMAT',
    message: `Unsupported DOCX image format: ${contentType}`,
    path,
    feature: normalizeContentType(contentType)
  }
}

export function isUnsupportedVectorImageContentType(contentType: string): boolean {
  return UNSUPPORTED_VECTOR_CONTENT_TYPES.has(normalizeContentType(contentType))
}

function bytesToBinaryString(bytes: Uint8Array): string {
  let result = ''
  const chunkSize = 0x8000

  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    const chunk = bytes.subarray(offset, offset + chunkSize)
    result += String.fromCharCode(...chunk)
  }

  return result
}

function binaryStringToBytes(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length)

  for (let index = 0; index < binary.length; index++) {
    bytes[index] = binary.charCodeAt(index)
  }

  return bytes
}

function encodeBase64(bytes: Uint8Array): string {
  return btoa(bytesToBinaryString(bytes))
}

function decodeBase64(base64: string): Uint8Array {
  return binaryStringToBytes(atob(base64))
}

function decodePercentData(data: string): Uint8Array {
  return new TextEncoder().encode(decodeURIComponent(data))
}

export function bytesToDataUrl(
  bytes: Uint8Array,
  contentType: string
): string {
  const normalizedContentType = normalizeContentType(contentType)

  if (!isSupportedImageContentType(normalizedContentType)) {
    throw new DocxParserError(
      'UNSUPPORTED_DOCX_TYPE',
      `Unsupported image content type: ${contentType}`
    )
  }

  return `data:${normalizedContentType};base64,${encodeBase64(bytes)}`
}

export function dataUrlToBytes(src: string): {
  bytes: Uint8Array
  contentType: string
} {
  const match = /^data:([^;,]+)((?:;[^,]*)?),(.*)$/i.exec(src)
  if (!match) {
    throw new DocxParserError('INVALID_DOCX_INPUT', 'Invalid image data URL')
  }

  const [, rawContentType, parameters, data] = match
  const contentType = normalizeContentType(rawContentType)
  if (!isSupportedImageContentType(contentType)) {
    throw new DocxParserError(
      'UNSUPPORTED_DOCX_TYPE',
      `Unsupported image content type: ${rawContentType}`
    )
  }

  const isBase64 = /(?:^|;)base64(?:;|$)/i.test(parameters)
  return {
    bytes: isBase64 ? decodeBase64(data) : decodePercentData(data),
    contentType
  }
}
