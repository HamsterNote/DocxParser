const DOCX_PARSER_PREFIX = '[DocxParser]'

function isProduction(): boolean {
  return typeof process !== 'undefined' && process.env?.NODE_ENV === 'production'
}

export function devConsoleLog(label: string, data?: unknown): void {
  if (isProduction()) return

  // eslint-disable-next-line no-console
  console.log(`${DOCX_PARSER_PREFIX} ${label}`, data !== undefined ? data : '')
}

export function devConsoleWarn(label: string, data?: unknown): void {
  if (isProduction()) return

  // eslint-disable-next-line no-console
  console.warn(`${DOCX_PARSER_PREFIX} ⚠️ ${label}`, data !== undefined ? data : '')
}

export function devConsoleError(label: string, data?: unknown): void {
  if (isProduction()) return

  // eslint-disable-next-line no-console
  console.error(`${DOCX_PARSER_PREFIX} ❌ ${label}`, data !== undefined ? data : '')
}
