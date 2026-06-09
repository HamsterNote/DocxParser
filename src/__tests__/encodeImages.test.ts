import { IntermediateImage } from '@hamster-note/types'
import mammoth from 'mammoth/mammoth.browser.js'
import { encodeDocxToIntermediate } from '../encode.js'
import { boxToPolygon } from '../geometry.js'
import { makeFixtureByName } from '../testUtils/fixtureFactory.js'
import { extractImages, extractPureText } from '../testUtils/intermediateTestUtils.js'

type MammothImageForTest = {
  contentType: string
  read: () => Promise<Uint8Array>
}

type MammothImageAttributes = {
  src: string
}

type MammothImageHandler = (image: MammothImageForTest) => Promise<MammothImageAttributes>

describe('encodeDocxToIntermediate image handling', () => {
  it('image.docx 应该把内嵌 PNG 转成 data URL IntermediateImage', async () => {
    const result = await encodeDocxToIntermediate(await makeFixtureByName('image.docx'))
    const images = await extractImages(result.document)

    expect(images.length).toBeGreaterThanOrEqual(1)
    expect(images[0]).toEqual(
      expect.objectContaining({
        id: expect.stringMatching(/^docx-[0-9a-f]{8}-p1-img0$/),
        src: expect.stringMatching(/^data:image\/png;base64,/),
        polygon: boxToPolygon(40, 40 + 12 * 1.2 + 10, 320, 180),
        opacity: 1
      })
    )
    expect(await extractPureText(result.document)).toContain('Image fixture before inline PNG.')
    expect(result.warnings.map((warning) => warning.code)).not.toContain('UNSUPPORTED_IMAGE_FORMAT')
  })

  it('image-only.docx 应该生成只有图片内容的 IntermediateDocument 且不因无文本崩溃', async () => {
    const result = await encodeDocxToIntermediate(await makeFixtureByName('image-only.docx'))
    const [page] = await result.document.pages
    const content = await page.getContent()

    expect(result.document.pageCount).toBe(1)
    expect(result.document.title).toBe('Untitled Docx Document')
    expect(await extractPureText(result.document)).toBe('')
    expect(content).toHaveLength(1)
    expect(content[0]).toBeInstanceOf(IntermediateImage)
    expect((content[0] as IntermediateImage).src).toMatch(/^data:image\/png;base64,/)
  })

  it('不支持的图片格式应该跳过图片并返回 UNSUPPORTED_IMAGE_FORMAT warning', async () => {
    const originalImgElement = mammoth.images.imgElement

    mammoth.images.imgElement = ((handler: MammothImageHandler) =>
      originalImgElement(async () =>
        handler({
          contentType: 'image/emf',
          read: async () => new Uint8Array([1, 2, 3])
        })
      )) as typeof mammoth.images.imgElement

    try {
      const result = await encodeDocxToIntermediate(await makeFixtureByName('image.docx'))

      expect(await extractImages(result.document)).toHaveLength(0)
      expect(result.warnings).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            code: 'UNSUPPORTED_IMAGE_FORMAT',
            feature: 'image/emf'
          })
        ])
      )
    } finally {
      mammoth.images.imgElement = originalImgElement
    }
  })

  it('图片读取失败应该跳过图片并返回 MISSING_IMAGE warning', async () => {
    const originalImgElement = mammoth.images.imgElement

    mammoth.images.imgElement = ((handler: MammothImageHandler) =>
      originalImgElement(async () =>
        handler({
          contentType: 'image/png',
          read: async () => {
            throw new Error('fixture image missing')
          }
        })
      )) as typeof mammoth.images.imgElement

    try {
      const result = await encodeDocxToIntermediate(await makeFixtureByName('image.docx'))

      expect(await extractImages(result.document)).toHaveLength(0)
      expect(result.warnings).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            code: 'MISSING_IMAGE',
            message: expect.stringContaining('fixture image missing'),
            feature: 'image/png'
          })
        ])
      )
    } finally {
      mammoth.images.imgElement = originalImgElement
    }
  })
})
