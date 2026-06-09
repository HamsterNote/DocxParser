import {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalMergeType,
  WidthType
} from 'docx'

export const MAX_FIXTURE_BYTES = 50 * 1024

export type DocxFixtureName =
  | 'paragraphs.docx'
  | 'text-styles.docx'
  | 'page-break.docx'
  | 'image.docx'
  | 'image-only.docx'
  | 'ordered-list.docx'
  | 'unordered-list.docx'
  | 'nested-list.docx'
  | 'simple-table.docx'
  | 'merged-table.docx'
  | 'empty.docx'
  | 'corrupt.docx'
  | 'not-docx.bin'

export type FixtureFactory = () => Promise<ArrayBuffer> | ArrayBuffer

export type DocxFixtureDefinition = {
  name: DocxFixtureName
  expectedText: readonly string[]
  validDocx: boolean
  factory: FixtureFactory
}

const transparentPng1x1 = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0,
  1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 10, 73, 68,
  65, 84, 120, 156, 99, 0, 1, 0, 0, 5, 0, 1, 13, 10, 45, 180, 0, 0, 0, 0,
  73, 69, 78, 68, 174, 66, 96, 130
])

const tableWidth = {
  size: 100,
  type: WidthType.PERCENTAGE
} as const

const packDocument = async (children: Array<Paragraph | Table>): Promise<ArrayBuffer> => {
  const doc = new Document({
    creator: 'HamsterNote DocxParser tests',
    title: 'Self-authored DocxParser fixture',
    numbering: {
      config: [
        {
          reference: 'ordered-list',
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: '%1.',
              alignment: AlignmentType.LEFT
            }
          ]
        },
        {
          reference: 'unordered-list',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT
            }
          ]
        },
        {
          reference: 'nested-list',
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: '%1.',
              alignment: AlignmentType.LEFT
            },
            {
              level: 1,
              format: LevelFormat.BULLET,
              text: '◦',
              alignment: AlignmentType.LEFT
            }
          ]
        }
      ]
    },
    sections: [
      {
        properties: {},
        children
      }
    ]
  })

  return Packer.toArrayBuffer(doc)
}

const cell = (text: string, options: Partial<ConstructorParameters<typeof TableCell>[0]> = {}) =>
  new TableCell({
    width: { size: 2400, type: WidthType.DXA },
    ...options,
    children: [new Paragraph(text)]
  })

export const makeParagraphsFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph({ text: 'Fixture Paragraphs', heading: HeadingLevel.HEADING_1 }),
    new Paragraph('First self-authored paragraph.'),
    new Paragraph('Second self-authored paragraph.')
  ])

export const makeTextStylesFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph({
      children: [
        new TextRun({ text: 'Bold sample', bold: true }),
        new TextRun(' '),
        new TextRun({ text: 'Italic sample', italics: true }),
        new TextRun(' '),
        new TextRun({ text: 'Large blue serif sample', size: 32, color: '1F4E79', font: 'Georgia' }),
        new TextRun(' '),
        new TextRun({ text: '中文样式文本', font: { eastAsia: 'SimSun' } })
      ]
    })
  ])

export const makePageBreakFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph('Page one content.'),
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph('Page two content after explicit break.')
  ])

export const makeImageFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph('Image fixture before inline PNG.'),
    new Paragraph({
      children: [
        new ImageRun({
          type: 'png',
          data: transparentPng1x1,
          transformation: { width: 16, height: 16 },
          altText: {
            title: 'Self-authored transparent PNG',
            description: 'A generated one pixel PNG used by DocxParser tests',
            name: 'fixture-transparent-png'
          }
        })
      ]
    }),
    new Paragraph('Image fixture after inline PNG.')
  ])

export const makeImageOnlyFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph({
      children: [
        new ImageRun({
          type: 'png',
          data: transparentPng1x1,
          transformation: { width: 16, height: 16 },
          altText: {
            title: 'Self-authored image only PNG',
            description: 'A generated one pixel PNG without surrounding text',
            name: 'fixture-image-only-transparent-png'
          }
        })
      ]
    })
  ])

export const makeOrderedListFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph({ text: 'Ordered item one', numbering: { reference: 'ordered-list', level: 0 } }),
    new Paragraph({ text: 'Ordered item two', numbering: { reference: 'ordered-list', level: 0 } }),
    new Paragraph({ text: 'Ordered item three', numbering: { reference: 'ordered-list', level: 0 } })
  ])

export const makeUnorderedListFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph({ text: 'Bullet item one', numbering: { reference: 'unordered-list', level: 0 } }),
    new Paragraph({ text: 'Bullet item two', numbering: { reference: 'unordered-list', level: 0 } }),
    new Paragraph({ text: 'Bullet item three', numbering: { reference: 'unordered-list', level: 0 } })
  ])

export const makeNestedListFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph({ text: 'Nested parent one', numbering: { reference: 'nested-list', level: 0 } }),
    new Paragraph({ text: 'Nested child one', numbering: { reference: 'nested-list', level: 1 } }),
    new Paragraph({ text: 'Nested child two', numbering: { reference: 'nested-list', level: 1 } }),
    new Paragraph({ text: 'Nested parent two', numbering: { reference: 'nested-list', level: 0 } })
  ])

export const makeSimpleTableFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph('Simple table fixture.'),
    new Table({
      width: tableWidth,
      rows: [
        new TableRow({ children: [cell('Header A'), cell('Header B')] }),
        new TableRow({ children: [cell('Cell A1'), cell('Cell B1')] })
      ]
    })
  ])

export const makeMergedTableFixture = (): Promise<ArrayBuffer> =>
  packDocument([
    new Paragraph('Merged table fixture.'),
    new Table({
      width: tableWidth,
      rows: [
        new TableRow({
          children: [cell('Merged header across two columns', { columnSpan: 2 })]
        }),
        new TableRow({
          children: [
            cell('Merged row start', { verticalMerge: VerticalMergeType.RESTART }),
            cell('Right cell one')
          ]
        }),
        new TableRow({
          children: [
            cell('', { verticalMerge: VerticalMergeType.CONTINUE }),
            cell('Right cell two')
          ]
        })
      ]
    })
  ])

export const makeEmptyFixture = (): Promise<ArrayBuffer> => packDocument([])

export const makeCorruptFixture = (): ArrayBuffer =>
  new Uint8Array([80, 75, 3, 4, 20, 0, 0, 0, 8, 0, 72, 78, 111, 116, 65, 68, 111, 99, 120]).buffer

export const makeNotDocxFixture = (): ArrayBuffer =>
  new TextEncoder().encode('This is self-authored plain text, not a DOCX ZIP package.').buffer

export const docxFixtureDefinitions: readonly DocxFixtureDefinition[] = [
  {
    name: 'paragraphs.docx',
    expectedText: ['Fixture Paragraphs', 'First self-authored paragraph.', 'Second self-authored paragraph.'],
    validDocx: true,
    factory: makeParagraphsFixture
  },
  {
    name: 'text-styles.docx',
    expectedText: ['Bold sample', 'Italic sample', 'Large blue serif sample', '中文样式文本'],
    validDocx: true,
    factory: makeTextStylesFixture
  },
  {
    name: 'page-break.docx',
    expectedText: ['Page one content.', 'Page two content after explicit break.'],
    validDocx: true,
    factory: makePageBreakFixture
  },
  {
    name: 'image.docx',
    expectedText: ['Image fixture before inline PNG.', 'Image fixture after inline PNG.'],
    validDocx: true,
    factory: makeImageFixture
  },
  {
    name: 'image-only.docx',
    expectedText: [],
    validDocx: true,
    factory: makeImageOnlyFixture
  },
  {
    name: 'ordered-list.docx',
    expectedText: ['Ordered item one', 'Ordered item two', 'Ordered item three'],
    validDocx: true,
    factory: makeOrderedListFixture
  },
  {
    name: 'unordered-list.docx',
    expectedText: ['Bullet item one', 'Bullet item two', 'Bullet item three'],
    validDocx: true,
    factory: makeUnorderedListFixture
  },
  {
    name: 'nested-list.docx',
    expectedText: ['Nested parent one', 'Nested child one', 'Nested child two', 'Nested parent two'],
    validDocx: true,
    factory: makeNestedListFixture
  },
  {
    name: 'simple-table.docx',
    expectedText: ['Simple table fixture.', 'Header A', 'Header B', 'Cell A1', 'Cell B1'],
    validDocx: true,
    factory: makeSimpleTableFixture
  },
  {
    name: 'merged-table.docx',
    expectedText: ['Merged table fixture.', 'Merged header across two columns', 'Merged row start', 'Right cell one', 'Right cell two'],
    validDocx: true,
    factory: makeMergedTableFixture
  },
  {
    name: 'empty.docx',
    expectedText: [],
    validDocx: true,
    factory: makeEmptyFixture
  },
  {
    name: 'corrupt.docx',
    expectedText: [],
    validDocx: false,
    factory: makeCorruptFixture
  },
  {
    name: 'not-docx.bin',
    expectedText: [],
    validDocx: false,
    factory: makeNotDocxFixture
  }
]

export const makeFixtureByName = (name: DocxFixtureName): Promise<ArrayBuffer> | ArrayBuffer => {
  const fixture = docxFixtureDefinitions.find((definition) => definition.name === name)

  if (!fixture) {
    throw new Error(`Unknown DOCX fixture: ${name}`)
  }

  return fixture.factory()
}
