import type { IntermediateDocument, IntermediateOutline } from '@hamster-note/types'
import { DocxPage } from './DocxPage.js'

export class DocxDocument {
  constructor(private readonly intermediateDocument: IntermediateDocument) {}

  async getPages(): Promise<DocxPage[]> {
    const pages = await this.intermediateDocument.pages
    return pages.map((page) => new DocxPage(page))
  }

  async getPage(pageNumber: number): Promise<DocxPage | undefined> {
    const pagePromise = this.intermediateDocument.getPageByPageNumber(pageNumber)
    if (!pagePromise) return undefined

    const page = await pagePromise
    return page ? new DocxPage(page) : undefined
  }

  async getOutline(): Promise<IntermediateOutline | undefined> {
    const outline = this.intermediateDocument.getOutline()
    if (!outline || outline.length === 0) return undefined

    return outline[0]
  }

  async getCover(): Promise<HTMLCanvasElement | HTMLImageElement> {
    const cover = await this.intermediateDocument.getCover()

    if (cover?.src) {
      const img = createImageElement()
      if (img) {
        return new Promise((resolve) => {
          img.onload = () => resolve(img)
          img.onerror = () => resolve(this.createBlankCoverCanvas())
          img.src = cover.src
        })
      }
    }

    return this.createBlankCoverCanvas()
  }

  getTitle(): string {
    return this.intermediateDocument.title
  }

  getId(): string {
    return this.intermediateDocument.id
  }

  getIntermediateDocument(): IntermediateDocument {
    return this.intermediateDocument
  }

  private createBlankCoverCanvas(): HTMLCanvasElement {
    const doc = resolveDocument()
    const canvas = doc.createElement('canvas')
    const firstPageSize = this.intermediateDocument.getPageSizeByPageNumber(1)

    canvas.width = firstPageSize?.x ?? 800
    canvas.height = firstPageSize?.y ?? 1000

    const ctx = canvas.getContext('2d')
    if (ctx) {
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, canvas.width, canvas.height)
    }

    return canvas
  }
}

function resolveDocument(): Document {
  if (typeof document === 'undefined') {
    throw new Error('DocxDocument.getCover requires a document context')
  }

  return document
}

function createImageElement(): HTMLImageElement | undefined {
  if (typeof Image !== 'undefined') return new Image()

  if (typeof document !== 'undefined') {
    return document.createElement('img')
  }

  return undefined
}
