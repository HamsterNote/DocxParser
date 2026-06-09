import type { ParserInput } from '@hamster-note/document-parser'
import { DocxParserError } from './errors.js'

function isBlobLike(input: ParserInput): input is Blob {
  return (
    typeof Blob !== 'undefined' &&
    input instanceof Blob &&
    typeof input.arrayBuffer === 'function'
  )
}

function cloneArrayBuffer(buffer: ArrayBuffer): ArrayBuffer {
  return buffer.slice(0)
}

function arrayBufferViewToExactBuffer(view: ArrayBufferView): ArrayBuffer {
  const bytes = new Uint8Array(view.buffer, view.byteOffset, view.byteLength)
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy.buffer
}

function assertDocxZipMagic(buffer: ArrayBuffer): void {
  if (buffer.byteLength === 0) {
    throw new DocxParserError('INVALID_DOCX_INPUT', 'DOCX input is empty')
  }

  const bytes = new Uint8Array(buffer)
  if (bytes.length < 2 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new DocxParserError(
      'INVALID_DOCX_INPUT',
      'DOCX input does not start with ZIP magic bytes'
    )
  }
}

export async function normalizeParserInput(input: ParserInput): Promise<ArrayBuffer> {
  let buffer: ArrayBuffer

  if (input instanceof ArrayBuffer) {
    buffer = cloneArrayBuffer(input)
  } else if (ArrayBuffer.isView(input)) {
    buffer = arrayBufferViewToExactBuffer(input)
  } else if (isBlobLike(input)) {
    buffer = await input.arrayBuffer()
  } else {
    throw new DocxParserError(
      'INVALID_DOCX_INPUT',
      'DOCX input must be an ArrayBuffer, ArrayBufferView, or Blob'
    )
  }

  assertDocxZipMagic(buffer)
  return buffer
}
