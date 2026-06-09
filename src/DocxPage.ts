import type {
  IntermediateContent,
  IntermediateImage,
  IntermediatePage,
  IntermediateText,
  Number2
} from '@hamster-note/types'

export enum RenderViews {
  TEXT = 'TEXT',
  THUMBNAIL = 'THUMBNAIL'
}

export interface RenderOptions {
  scale?: number
  views?: RenderViews[]
}

type DocxPageSize = Number2 & { width: number; height: number }

export class DocxPage {
  constructor(private readonly intermediatePage: IntermediatePage) {}

  getNumber(): number {
    return this.intermediatePage.number
  }

  getSize(scale = 1): DocxPageSize {
    const width = this.intermediatePage.width * scale
    const height = this.intermediatePage.height * scale

    return { x: width, y: height, width, height }
  }

  getPureText(): string {
    return this.intermediatePage.content
      .filter(isIntermediateText)
      .map((text) => text.content)
      .join('\n')
  }

  async render(container: HTMLElement, options?: RenderOptions): Promise<void> {
    const scale = options?.scale ?? 1
    const views = options?.views ?? [RenderViews.TEXT, RenderViews.THUMBNAIL]
    const ownerDocument = container.ownerDocument ?? globalThis.document

    if (!ownerDocument) {
      throw new Error('DocxPage.render requires a document context')
    }

    container.innerHTML = ''
    container.style.position = 'relative'
    container.style.overflow = 'hidden'
    container.style.width = `${this.intermediatePage.width * scale}px`
    container.style.height = `${this.intermediatePage.height * scale}px`

    if (views.includes(RenderViews.THUMBNAIL)) {
      const thumbnail = await this.intermediatePage.getThumbnail(0.3)
      if (thumbnail?.src) {
        container.style.backgroundImage = `url('${thumbnail.src}')`
        container.style.backgroundRepeat = 'no-repeat'
        container.style.backgroundPosition = 'top center'
        container.style.backgroundSize = 'contain'
      }
    }

    if (views.includes(RenderViews.TEXT)) {
      const content = await this.intermediatePage.getContent()
      const layer = ownerDocument.createElement('div')
      layer.style.position = 'absolute'
      layer.style.top = '0'
      layer.style.left = '0'
      layer.style.width = '100%'
      layer.style.height = '100%'

      content.forEach((item) => {
        if (isIntermediateText(item)) {
          layer.appendChild(renderTextSpan(ownerDocument, item, scale))
          return
        }

        if (isIntermediateImage(item)) {
          layer.appendChild(renderImage(ownerDocument, item, scale))
        }
      })

      container.appendChild(layer)
    }
  }

  getIntermediatePage(): IntermediatePage {
    return this.intermediatePage
  }
}

function isIntermediateText(content: IntermediateContent): content is IntermediateText {
  return (
    'content' in content &&
    typeof content.content === 'string' &&
    'fontSize' in content &&
    'polygon' in content
  )
}

function isIntermediateImage(content: IntermediateContent): content is IntermediateImage {
  return 'src' in content && typeof content.src === 'string' && 'polygon' in content
}

function getPolygonBox(content: Pick<IntermediateText | IntermediateImage, 'polygon'>): {
  x: number
  y: number
  width: number
  height: number
} {
  const xs = content.polygon.map(([x]) => x)
  const ys = content.polygon.map(([, y]) => y)
  const minX = Math.min(...xs)
  const minY = Math.min(...ys)
  const maxX = Math.max(...xs)
  const maxY = Math.max(...ys)

  return {
    x: minX,
    y: minY,
    width: Math.max(0, maxX - minX),
    height: Math.max(0, maxY - minY)
  }
}

function applyBoxStyle(
  element: HTMLElement,
  box: ReturnType<typeof getPolygonBox>,
  scale: number
): void {
  element.style.position = 'absolute'
  element.style.left = `${box.x * scale}px`
  element.style.top = `${box.y * scale}px`
  element.style.width = `${box.width * scale}px`
  element.style.height = `${box.height * scale}px`
}

function renderTextSpan(
  ownerDocument: Document,
  text: IntermediateText,
  scale: number
): HTMLSpanElement {
  const span = ownerDocument.createElement('span')
  const box = getPolygonBox(text)

  span.className = 'hamster-note-text'
  span.id = text.id
  span.textContent = text.content
  applyBoxStyle(span, box, scale)
  span.style.fontSize = `${text.fontSize * scale}px`
  span.style.fontFamily = text.fontFamily
  span.style.fontWeight = String(text.fontWeight)
  span.style.fontStyle = text.italic ? 'italic' : 'normal'
  span.style.color = text.color
  span.style.lineHeight = `${text.lineHeight * scale}px`
  span.style.whiteSpace = 'pre-wrap'
  span.style.transformOrigin = '0 0'
  if (typeof text.opacity === 'number') span.style.opacity = String(text.opacity)

  return span
}

function renderImage(
  ownerDocument: Document,
  image: IntermediateImage,
  scale: number
): HTMLImageElement {
  const img = ownerDocument.createElement('img')
  const box = getPolygonBox(image)

  img.className = 'hamster-note-image'
  img.id = image.id
  img.src = image.src
  applyBoxStyle(img, box, scale)
  img.style.objectFit = 'contain'
  img.style.opacity = String(image.opacity)

  return img
}
