import { IntermediatePage, IntermediateText } from '@hamster-note/types'
import mammoth from 'mammoth/mammoth.browser.js'
import { htmlToIntermediateDocument } from '../htmlToIntermediate.js'
import { makeFixtureByName } from '../testUtils/fixtureFactory.js'

const pageTexts = async (fixtureName: Parameters<typeof makeFixtureByName>[0]) => {
  const html = await mammoth.convertToHtml({ arrayBuffer: await makeFixtureByName(fixtureName) })
  const result = htmlToIntermediateDocument(html.value, {
    documentId: `docx-${fixtureName}`,
    sourceHash: fixtureName.replace(/\W/g, '')
  })
  const [page] = await result.document.pages

  return (await page.getContent()).filter(
    (item): item is IntermediateText => item instanceof IntermediateText
  )
}

const pageTextLayout = async (fixtureName: Parameters<typeof makeFixtureByName>[0]) => {
  const html = await mammoth.convertToHtml({ arrayBuffer: await makeFixtureByName(fixtureName) })
  const result = htmlToIntermediateDocument(html.value, {
    documentId: `docx-${fixtureName}`,
    sourceHash: fixtureName.replace(/\W/g, '')
  })
  const [page] = await result.document.pages
  const texts = (await page.getContent()).filter(
    (item): item is IntermediateText => item instanceof IntermediateText
  )

  return { texts, paragraphs: IntermediatePage.serialize(page).paragraphs ?? [] }
}

const joinedContent = (texts: IntermediateText[]): string =>
  texts.map((text) => text.content).join('')

const paragraphXForContent = (
  texts: IntermediateText[],
  paragraphs: NonNullable<ReturnType<typeof IntermediatePage.serialize>['paragraphs']>,
  content: string
): number => {
  const text = texts.find((item) => item.content.includes(content))
  const paragraph = paragraphs.find((item) => text && item.textIds.includes(text.id))
  expect(paragraph).toBeDefined()

  return paragraph?.x ?? 0
}

describe('encodeDocxToIntermediate list mapping', () => {
  it('ordered-list.docx 应该按顺序保留 1. 和 2. 编号 marker', async () => {
    const text = joinedContent(await pageTexts('ordered-list.docx'))

    expect(text.indexOf('1.')).toBeLessThan(text.indexOf('Ordered item one'))
    expect(text.indexOf('2.')).toBeLessThan(text.indexOf('Ordered item two'))
    expect(text.indexOf('1.')).toBeLessThan(text.indexOf('2.'))
  })

  it('unordered-list.docx 应该保留 bullet marker 和列表文本', async () => {
    const text = joinedContent(await pageTexts('unordered-list.docx'))

    expect(text).toContain('•')
    expect(text).toContain('Bullet item one')
    expect(text).toContain('Bullet item two')
  })

  it('nested-list.docx 的子项 x 坐标应该比父项至少多 20pt', async () => {
    const { texts, paragraphs } = await pageTextLayout('nested-list.docx')
    const parentX = paragraphXForContent(texts, paragraphs, 'Nested parent one')
    const childX = paragraphXForContent(texts, paragraphs, 'Nested child one')

    expect(childX - parentX).toBeGreaterThanOrEqual(20)
  })
})
