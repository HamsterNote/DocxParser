import {
  IntermediateImage,
  IntermediatePage,
  IntermediateText,
  TextDir
} from '@hamster-note/types'
import { DocxPage, RenderViews } from '../DocxPage.js'

type FakeElement = {
  tagName: string
  className: string
  id: string
  textContent: string | null
  innerHTML: string
  src: string
  style: Record<string, string>
  children: FakeElement[]
  ownerDocument?: FakeDocument
  appendChild: (child: FakeElement) => FakeElement
}

type FakeDocument = {
  createElement: (tagName: string) => FakeElement
}

function createFakeDocument(): FakeDocument {
  const doc: FakeDocument = {
    createElement: (tagName: string) => createFakeElement(tagName, doc)
  }
  return doc
}

function createFakeElement(tagName: string, ownerDocument: FakeDocument): FakeElement {
  return {
    tagName,
    className: '',
    id: '',
    textContent: null,
    innerHTML: '',
    src: '',
    style: {},
    children: [],
    ownerDocument,
    appendChild(child: FakeElement) {
      this.children.push(child)
      return child
    }
  }
}

function makeText(id: string, content: string): IntermediateText {
  return new IntermediateText({
    id,
    content,
    fontSize: 10,
    fontFamily: 'Arial',
    fontWeight: 700,
    italic: true,
    color: '#123456',
    polygon: [
      [10, 20],
      [110, 20],
      [110, 40],
      [10, 40]
    ],
    lineHeight: 14,
    ascent: 9,
    descent: 3,
    dir: TextDir.LTR,
    opacity: 0.8,
    skew: 0,
    isEOL: true
  })
}

function makeImage(): IntermediateImage {
  return new IntermediateImage({
    id: 'image-1',
    src: 'data:image/png;base64,abc',
    polygon: [
      [20, 50],
      [120, 50],
      [120, 100],
      [20, 100]
    ],
    opacity: 0.5
  })
}

function makePage(): IntermediatePage {
  const page = new IntermediatePage({
    id: 'page-1',
    number: 3,
    width: 200,
    height: 100,
    content: [makeText('text-1', 'Hello'), makeImage(), makeText('text-2', 'World')]
  })

  page.setGetThumbnail(async () => makeImage())
  return page
}

describe('DocxPage', () => {
  it('exposes number, scaled size, pure text, and underlying page', () => {
    const intermediatePage = makePage()
    const page = new DocxPage(intermediatePage)

    expect(page.getNumber()).toBe(3)
    expect(page.getSize()).toEqual({ x: 200, y: 100, width: 200, height: 100 })
    expect(page.getSize(2)).toEqual({ x: 400, y: 200, width: 400, height: 200 })
    expect(page.getPureText()).toBe('Hello\nWorld')
    expect(page.getIntermediatePage()).toBe(intermediatePage)
  })

  it('renders thumbnail background, text spans, and images', async () => {
    const doc = createFakeDocument()
    const container = createFakeElement('div', doc)

    await new DocxPage(makePage()).render(container as unknown as HTMLElement, {
      scale: 2,
      views: [RenderViews.THUMBNAIL, RenderViews.TEXT]
    })

    expect(container.innerHTML).toBe('')
    expect(container.style.position).toBe('relative')
    expect(container.style.width).toBe('400px')
    expect(container.style.height).toBe('200px')
    expect(container.style.backgroundImage).toBe("url('data:image/png;base64,abc')")

    const layer = container.children[0]
    expect(layer.style.position).toBe('absolute')
    expect(layer.children).toHaveLength(3)

    const firstText = layer.children[0]
    expect(firstText.className).toBe('hamster-note-text')
    expect(firstText.id).toBe('text-1')
    expect(firstText.textContent).toBe('Hello')
    expect(firstText.style.left).toBe('20px')
    expect(firstText.style.top).toBe('40px')
    expect(firstText.style.fontSize).toBe('20px')
    expect(firstText.style.fontStyle).toBe('italic')
    expect(firstText.style.opacity).toBe('0.8')

    const image = layer.children[1]
    expect(image.className).toBe('hamster-note-image')
    expect(image.id).toBe('image-1')
    expect(image.src).toBe('data:image/png;base64,abc')
    expect(image.style.objectFit).toBe('contain')
    expect(image.style.opacity).toBe('0.5')
  })

  it('can render text-only layer without thumbnail background', async () => {
    const doc = createFakeDocument()
    const container = createFakeElement('div', doc)

    await new DocxPage(makePage()).render(container as unknown as HTMLElement, {
      views: [RenderViews.TEXT]
    })

    expect(container.style.backgroundImage).toBeUndefined()
    expect(container.children[0].children[0].className).toBe('hamster-note-text')
  })

  it('throws when render has no document context', async () => {
    const container = createFakeElement('div', createFakeDocument())
    container.ownerDocument = undefined
    const previousDocument = globalThis.document

    Object.defineProperty(globalThis, 'document', {
      configurable: true,
      value: undefined
    })

    try {
      await expect(
        new DocxPage(makePage()).render(container as unknown as HTMLElement)
      ).rejects.toThrow('DocxPage.render requires a document context')
    } finally {
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: previousDocument
      })
    }
  })
})
