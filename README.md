# @hamster-note/docx-parser

DOCX parser for the Hamster Note ecosystem.

`@hamster-note/docx-parser` converts DOCX files into Hamster Note's intermediate document model and can generate DOCX output from an `IntermediateDocument`.

## Installation

```bash
yarn add @hamster-note/docx-parser
```

## Basic usage

```typescript
import { DocxParser } from '@hamster-note/docx-parser'

// Encode DOCX to IntermediateDocument
const doc = await DocxParser.encode(docxBuffer)
const pages = await doc.getPages()

// Decode IntermediateDocument to DOCX
const docxBuffer = await DocxParser.decode(intermediateDocument)
```

## API overview

### `DocxParser`

Main parser entry point. It exposes static encode/decode helpers for common package usage and supports the parser metadata expected by the Hamster Note document-parser layer.

- `DocxParser.encode(input)` reads DOCX input and returns a `DocxDocument` wrapper.
- `DocxParser.decode(intermediateDocument)` writes an `IntermediateDocument` back to DOCX bytes.

### `DocxDocument`

Wrapper around a Hamster Note `IntermediateDocument`.

- `getPages()` returns parsed pages as `DocxPage` objects.
- `getPage(pageNumber)` returns a single page wrapper when available.
- `getTitle()`, `getId()`, `getOutline()`, and `getCover()` expose document metadata helpers.

### `DocxPage`

Wrapper around an `IntermediatePage` with page-level helpers.

- `getNumber()` and `getSize(scale?)` expose page metadata.
- `getPureText()` returns page text content.
- `render(container, options?)` renders a lightweight content layer for preview use.

## Features

- Paragraph extraction and generation
- Text style preservation for common bold, italic, underline, strike, font, size, and color data
- Embedded image handling for supported image formats
- Ordered and unordered list mapping
- Best-effort table flattening and reconstruction
- Pagination support with page-break handling and fallback pagination

## License

MIT
