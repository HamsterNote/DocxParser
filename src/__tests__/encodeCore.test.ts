import { Document, Packer, Paragraph } from 'docx'
import { IntermediateOutlineDestType, IntermediateText } from '@hamster-note/types'
import { encodeDocxToIntermediate } from '../encode.js'
import { makeFixtureByName } from '../testUtils/fixtureFactory.js'
import { extractPureText } from '../testUtils/intermediateTestUtils.js'

const pageTexts = async (result: Awaited<ReturnType<typeof encodeDocxToIntermediate>>) => {
  const pages = await result.document.pages
  return Promise.all(
    pages.map(async (page) =>
      (await page.getContent())
        .filter((item): item is IntermediateText => item instanceof IntermediateText)
        .map((text) => text.content)
        .join('')
    )
  )
}

const makeUnknownStyleDocx = async (): Promise<ArrayBuffer> => {
  const doc = new Document({
    styles: {
      paragraphStyles: [
        {
          id: 'MysteryStyle',
          name: 'Mystery Style',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true
        }
      ]
    },
    sections: [
      {
        children: [new Paragraph({ text: 'Mystery styled text.', style: 'MysteryStyle' })]
      }
    ]
  })

  return Packer.toArrayBuffer(doc)
}

describe('encodeDocxToIntermediate', () => {
  it('paragraphs.docx 应该生成有标题、文本、大纲和稳定 docx hash ID 的 IntermediateDocument', async () => {
    const input = await makeFixtureByName('paragraphs.docx')
    const { document, warnings } = await encodeDocxToIntermediate(input)
    const text = await extractPureText(document)
    const outline = document.getOutline()

    expect(document.id).toMatch(/^docx-[0-9a-f]{8}$/)
    expect(document.pageCount).toBeGreaterThanOrEqual(1)
    expect(document.title).toBe('Fixture Paragraphs')
    expect(text).toContain('Fixture Paragraphs')
    expect(text).toContain('First self-authored paragraph.')
    expect(outline).toHaveLength(1)
    expect(outline?.[0].content).toBe('Fixture Paragraphs')
    expect(outline?.[0].dest.targetType).toBe(IntermediateOutlineDestType.PAGE)
    expect(warnings.map((warning) => warning.code)).toContain('PAGINATION_FALLBACK')
  })

  it('page-break.docx 应该按显式分页拆成至少两页且不加入 fallback 警告', async () => {
    const result = await encodeDocxToIntermediate(await makeFixtureByName('page-break.docx'))
    const texts = await pageTexts(result)

    expect(result.document.pageCount).toBeGreaterThanOrEqual(2)
    expect(texts[0]).toContain('Page one content.')
    expect(texts[1]).toContain('Page two content after explicit break.')
    expect(result.warnings.map((warning) => warning.code)).not.toContain('PAGINATION_FALLBACK')
  })

  it('无显式分页的 DOCX 应该生成 exactly 1 个 A4 fallback page 并报告 PAGINATION_FALLBACK', async () => {
    const result = await encodeDocxToIntermediate(await makeFixtureByName('paragraphs.docx'))
    const [page] = await result.document.pages

    expect(result.document.pageCount).toBe(1)
    expect(page.width).toBe(595)
    expect(page.height).toBe(842)
    expect(result.warnings).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'PAGINATION_FALLBACK', feature: 'pagination' })
      ])
    )
  })

  it('text-styles.docx 应该保留 mammoth 可暴露的 bold/italic/font-size/font-family 核心文本样式', async () => {
    const result = await encodeDocxToIntermediate(await makeFixtureByName('text-styles.docx'))
    const [page] = await result.document.pages
    const texts = (await page.getContent()).filter(
      (item): item is IntermediateText => item instanceof IntermediateText
    )

    expect(texts.find((text) => text.content.includes('Bold sample'))?.fontWeight).toBe(700)
    expect(texts.find((text) => text.content.includes('Italic sample'))?.italic).toBe(true)
    expect(texts.find((text) => text.content.includes('Large blue serif sample'))).toEqual(
      expect.objectContaining({ fontSize: 16, fontFamily: 'Georgia', color: '#000000' })
    )
  })

  it('mammoth conversion messages 应该作为 DocxParserWarning[] 保留', async () => {
    const result = await encodeDocxToIntermediate(await makeUnknownStyleDocx())

    expect(await extractPureText(result.document)).toContain('Mystery styled text.')
    expect(result.warnings).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'UNSUPPORTED_FIELD',
          message: expect.stringContaining('Unrecognised paragraph style')
        })
      ])
    )
  })
})
