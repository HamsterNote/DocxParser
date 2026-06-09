import {
  DocumentParser,
  type ParserInput
} from '@hamster-note/document-parser'
import type { IntermediateDocument } from '@hamster-note/types'
import { decodeIntermediateToDocx } from './decode.js'
import { DocxDocument } from './DocxDocument.js'
import { encodeDocxToIntermediate } from './encode.js'
import type { DocxParserWarning } from './errors.js'

export class DocxParser extends DocumentParser {
  static readonly exts = ['docx'] as const
  static readonly ext = 'docx'

  private lastDecodeWarnings: DocxParserWarning[] = []

  async encode(input: ParserInput): Promise<IntermediateDocument> {
    return DocxParser.encodeToIntermediate(input)
  }

  async decode(intermediateDocument: IntermediateDocument): Promise<ArrayBuffer> {
    this.lastDecodeWarnings = []
    return decodeIntermediateToDocx(intermediateDocument, {
      onWarning: (warning) => {
        this.lastDecodeWarnings.push(warning)
      }
    })
  }

  getLastDecodeWarnings(): DocxParserWarning[] {
    return [...this.lastDecodeWarnings]
  }

  static async encode(fileOrBuffer: ParserInput): Promise<DocxDocument> {
    const intermediateDocument = await DocxParser.encodeToIntermediate(fileOrBuffer)
    return new DocxDocument(intermediateDocument)
  }

  static async encodeToIntermediate(
    fileOrBuffer: ParserInput
  ): Promise<IntermediateDocument> {
    const result = await encodeDocxToIntermediate(fileOrBuffer)
    return result.document
  }

  static async decode(intermediateDocument: IntermediateDocument): Promise<ArrayBuffer> {
    return decodeIntermediateToDocx(intermediateDocument)
  }
}

export { DocxDocument } from './DocxDocument.js'
export { DocxPage, RenderViews, type RenderOptions } from './DocxPage.js'
export * from './dataUrl.js'
export * from './devLog.js'
export * from './errors.js'
export * from './geometry.js'
export * from './input.js'
export * from './styleMapping.js'
