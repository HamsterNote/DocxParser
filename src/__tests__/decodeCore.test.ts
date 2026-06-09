import type { IntermediateDocument } from '@hamster-note/types'
import mammoth from 'mammoth'
import { decodeIntermediateToDocx } from '../decode.js'
import type { DocxParserError } from '../errors.js'
import {
  extractPureText,
  makeSimpleIntermediateDocument,
  normalizeIntermediateDocument
} from '../testUtils/intermediateTestUtils.js'

const toMammothInput = (buffer: ArrayBuffer) => ({ buffer: Buffer.from(buffer) })

const normalizeTextContent = (value: string): string => value.replace(/\s+/g, '')

describe('decodeIntermediateToDocx', () => {
  it('returns a non-empty ArrayBuffer for a simple IntermediateDocument', async () => {
    const intermediateDocument = makeSimpleIntermediateDocument()
    const buffer = await decodeIntermediateToDocx(intermediateDocument)

    expect(buffer).toBeInstanceOf(ArrayBuffer)
    expect(buffer.byteLength).toBeGreaterThan(0)
  })

  it('re-encodes generated DOCX text equivalent to the normalized IntermediateDocument text', async () => {
    const intermediateDocument = makeSimpleIntermediateDocument()
    const expectedText = await extractPureText(intermediateDocument)
    const buffer = await decodeIntermediateToDocx(intermediateDocument)
    const rawText = await mammoth.extractRawText(toMammothInput(buffer))

    expect(rawText.messages).toEqual([])
    expect(normalizeTextContent(rawText.value)).toBe(normalizeTextContent(expectedText))
  })

  it('does not mutate the input IntermediateDocument while generating DOCX bytes', async () => {
    const intermediateDocument = makeSimpleIntermediateDocument()
    const before = await normalizeIntermediateDocument(intermediateDocument, {
      ignoreIds: false
    })

    await decodeIntermediateToDocx(intermediateDocument)

    await expect(
      normalizeIntermediateDocument(intermediateDocument, { ignoreIds: false })
    ).resolves.toEqual(before)
  })

  it('throws INVALID_INTERMEDIATE_DOCUMENT for invalid input structure', async () => {
    const invalidDocument = { id: 'invalid', title: 123 } as unknown as IntermediateDocument

    await expect(decodeIntermediateToDocx(invalidDocument)).rejects.toMatchObject({
      name: 'DocxParserError',
      code: 'INVALID_INTERMEDIATE_DOCUMENT'
    } satisfies Partial<DocxParserError>)
  })
})
