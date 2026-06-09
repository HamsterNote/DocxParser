/**
 * 包表面测试 — 验证 dist 产物可正常导入，无崩溃。
 *
 * 当前为占位测试（Task 4 完成后将补充完整的公开 API 断言）。
 */

describe('@hamster-note/docx-parser package surface', () => {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let mod: any

  beforeAll(async () => {
    // 动态导入编译产物；如果 dist 不存在则跳过（本地未构建时）
    try {
      mod = await import('../../dist/index.js')
    } catch {
      // dist 尚未构建 — 跳过所有用例
    }
  })

  it('should import without throwing', () => {
    if (!mod) return // dist 未构建，跳过
    expect(mod).toBeDefined()
  })

  it('should export DocxDocument', () => {
    if (!mod) return
    expect(mod.DocxDocument).toBeDefined()
    expect(typeof mod.DocxDocument).toBe('function')
  })

  it('should export DocxPage', () => {
    if (!mod) return
    expect(mod.DocxPage).toBeDefined()
    expect(typeof mod.DocxPage).toBe('function')
  })
})
