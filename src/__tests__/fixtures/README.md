# DOCX fixtures

These fixtures are self-authored for `@hamster-note/docx-parser` tests and are safe to commit.

The project prefers generated fixtures over committed binary DOCX files. `src/testUtils/fixtureFactory.ts` uses the `docx` library to build small ArrayBuffers programmatically, including valid DOCX cases plus intentionally invalid `corrupt.docx` and `not-docx.bin` buffers.

If a future test needs a committed binary fixture, keep it under 50KB and document why generation is not practical.
