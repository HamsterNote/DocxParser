import mammoth from 'mammoth/mammoth.browser.js'
import { htmlToIntermediateDocument } from '../htmlToIntermediate.js'
import { makeFixtureByName } from '../testUtils/fixtureFactory.js'
import { extractPureText } from '../testUtils/intermediateTestUtils.js'

const mapFixture = async (fixtureName: Parameters<typeof makeFixtureByName>[0]) => {
  const html = await mammoth.convertToHtml({ arrayBuffer: await makeFixtureByName(fixtureName) })

  return htmlToIntermediateDocument(html.value, {
    documentId: `docx-${fixtureName}`,
    sourceHash: fixtureName.replace(/\W/g, '')
  })
}

describe('encodeDocxToIntermediate table mapping', () => {
  it('simple-table.docx 应该按 row-major 顺序保留所有单元格文本', async () => {
    const result = await mapFixture('simple-table.docx')
    const text = await extractPureText(result.document)
    const expectedOrder = ['Header A', 'Header B', 'Cell A1', 'Cell B1']

    for (const expectedText of expectedOrder) expect(text).toContain(expectedText)
    expect(expectedOrder.map((expectedText) => text.indexOf(expectedText))).toEqual(
      [...expectedOrder]
        .map((expectedText) => text.indexOf(expectedText))
        .sort((left, right) => left - right)
    )
  })

  it('merged-table.docx 应该发出 TABLE_MERGE_BEST_EFFORT 且保留可见文本', async () => {
    const result = await mapFixture('merged-table.docx')
    const text = await extractPureText(result.document)

    expect(result.warnings).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: 'TABLE_MERGE_BEST_EFFORT' })])
    )
    expect(text).toContain('Merged header across two columns')
    expect(text).toContain('Merged row start')
    expect(text).toContain('Right cell one')
    expect(text).toContain('Right cell two')
  })
})
