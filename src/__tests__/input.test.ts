import { DocxParserError } from '../errors.js'
import { normalizeParserInput } from '../input.js'

const zipBytes = (): Uint8Array<ArrayBuffer> => {
  const buffer = new ArrayBuffer(6)
  const bytes = new Uint8Array(buffer)
  bytes.set([0x50, 0x4b, 0x03, 0x04, 1, 2])
  return bytes
}

describe('normalizeParserInput', () => {
  it('accepts ArrayBuffer and returns a cloned ArrayBuffer', async () => {
    const bytes = zipBytes()
    const normalized = await normalizeParserInput(bytes.buffer)

    expect(normalized).not.toBe(bytes.buffer)
    expect(Array.from(new Uint8Array(normalized))).toEqual(Array.from(bytes))
  })

  it('accepts ArrayBufferView and slices to the view range', async () => {
    const source = new Uint8Array([9, 9, 0x50, 0x4b, 0x03, 0x04, 8])
    const view = source.subarray(2, 6)
    const normalized = await normalizeParserInput(view)

    expect(Array.from(new Uint8Array(normalized))).toEqual([
      0x50, 0x4b, 0x03, 0x04
    ])
  })

  it('accepts Blob input', async () => {
    const blob = new Blob([zipBytes()], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    })
    const normalized = await normalizeParserInput(blob)

    expect(Array.from(new Uint8Array(normalized))).toEqual(Array.from(zipBytes()))
  })

  it('rejects empty buffers with INVALID_DOCX_INPUT', async () => {
    await expect(normalizeParserInput(new ArrayBuffer(0))).rejects.toMatchObject({
      code: 'INVALID_DOCX_INPUT'
    })
  })

  it('rejects non-ZIP magic bytes with INVALID_DOCX_INPUT', async () => {
    await expect(
      normalizeParserInput(new Uint8Array([0x7b, 0x22, 0x61]).buffer)
    ).rejects.toBeInstanceOf(DocxParserError)
    await expect(
      normalizeParserInput(new Uint8Array([0x7b, 0x22, 0x61]).buffer)
    ).rejects.toMatchObject({ code: 'INVALID_DOCX_INPUT' })
  })
})
