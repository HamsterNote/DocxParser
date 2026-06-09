import mammoth from 'mammoth'
import {
  docxFixtureDefinitions,
  MAX_FIXTURE_BYTES,
  makeFixtureByName
} from '../testUtils/fixtureFactory.js'

const byteLength = (buffer: ArrayBuffer): number => buffer.byteLength

const toMammothInput = (buffer: ArrayBuffer) => ({ buffer: Buffer.from(buffer) })

const startsWithZipMagic = (buffer: ArrayBuffer): boolean => {
  const bytes = new Uint8Array(buffer, 0, Math.min(buffer.byteLength, 4))
  return bytes[0] === 80 && bytes[1] === 75
}

describe('fixtureFactory', () => {
  it('应该覆盖所有要求的 fixture 名称', () => {
    expect(docxFixtureDefinitions.map((fixture) => fixture.name)).toEqual([
      'paragraphs.docx',
      'text-styles.docx',
      'page-break.docx',
      'image.docx',
      'image-only.docx',
      'ordered-list.docx',
      'unordered-list.docx',
      'nested-list.docx',
      'simple-table.docx',
      'merged-table.docx',
      'empty.docx',
      'corrupt.docx',
      'not-docx.bin'
    ])
  })

  it.each(docxFixtureDefinitions.filter((fixture) => fixture.validDocx))(
    '$name 应该生成 mammoth 可打开的小型 DOCX',
    async (fixture) => {
      const buffer = await fixture.factory()
      const rawText = await mammoth.extractRawText(toMammothInput(buffer))
      const html = await mammoth.convertToHtml(toMammothInput(buffer))

      expect(byteLength(buffer)).toBeGreaterThan(0)
      expect(byteLength(buffer)).toBeLessThanOrEqual(MAX_FIXTURE_BYTES)
      expect(startsWithZipMagic(buffer)).toBe(true)
      expect(rawText.messages).toEqual([])
      expect(html.messages).toEqual([])

      for (const expectedText of fixture.expectedText) {
        expect(rawText.value).toContain(expectedText)
      }
    }
  )

  it('image.docx 应该包含 mammoth 可转换的内联图片', async () => {
    const buffer = await makeFixtureByName('image.docx')
    const html = await mammoth.convertToHtml(toMammothInput(buffer))

    expect(html.value).toContain('<img')
    expect(html.value).toContain('data:image/png;base64,')
  })

  it('empty.docx 应该生成可打开但无正文文本的 DOCX', async () => {
    const buffer = await makeFixtureByName('empty.docx')
    const rawText = await mammoth.extractRawText(toMammothInput(buffer))

    expect(byteLength(buffer)).toBeLessThanOrEqual(MAX_FIXTURE_BYTES)
    expect(rawText.value.trim()).toBe('')
  })

  it.each(docxFixtureDefinitions.filter((fixture) => !fixture.validDocx))(
    '$name 应该生成小型无效输入供错误路径测试',
    async (fixture) => {
      const buffer = await fixture.factory()

      expect(byteLength(buffer)).toBeGreaterThan(0)
      expect(byteLength(buffer)).toBeLessThanOrEqual(MAX_FIXTURE_BYTES)
      await expect(mammoth.extractRawText(toMammothInput(buffer))).rejects.toThrow()
    }
  )

  it('corrupt.docx 应该保留 ZIP 魔数但内容损坏', async () => {
    const buffer = await makeFixtureByName('corrupt.docx')

    expect(startsWithZipMagic(buffer)).toBe(true)
    await expect(mammoth.convertToHtml(toMammothInput(buffer))).rejects.toThrow()
  })

  it('not-docx.bin 应该不是 ZIP/DOCX 输入', async () => {
    const buffer = await makeFixtureByName('not-docx.bin')

    expect(startsWithZipMagic(buffer)).toBe(false)
    await expect(mammoth.convertToHtml(toMammothInput(buffer))).rejects.toThrow()
  })
})
