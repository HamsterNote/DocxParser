import { DocxParser } from '../dist/index.js'

const input = document.querySelector('#docx-input')
const dropzone = document.querySelector('.dropzone')
const encodeButton = document.querySelector('#encode-button')
const decodeButton = document.querySelector('#decode-button')
const encodeStatus = document.querySelector('#encode-status')
const decodeStatus = document.querySelector('#decode-status')
const downloadLink = document.querySelector('#download-link')
const preview = document.querySelector('#preview')
const summaryPill = document.querySelector('#summary-pill')
const metricTitle = document.querySelector('#metric-title')
const metricId = document.querySelector('#metric-id')
const metricPages = document.querySelector('#metric-pages')
const metricSize = document.querySelector('#metric-size')

let selectedFile
let encodedDocument
let activeObjectUrl

function setStatus(element, message, state = 'idle') {
  element.textContent = message
  element.className = `status ${state}`
}

function formatBytes(bytes) {
  if (!bytes) return '-'
  if (bytes < 1024) return `${bytes} B`
  return `${(bytes / 1024).toFixed(1)} KB`
}

function resetDecodeResult() {
  decodeButton.disabled = true
  downloadLink.hidden = true
  downloadLink.removeAttribute('href')
  metricSize.textContent = '-'

  if (activeObjectUrl) {
    URL.revokeObjectURL(activeObjectUrl)
    activeObjectUrl = undefined
  }
}

function selectFile(file) {
  selectedFile = file
  encodedDocument = undefined
  encodeButton.disabled = !file
  resetDecodeResult()

  if (!file) {
    setStatus(encodeStatus, '等待上传 DOCX。')
    return
  }

  setStatus(encodeStatus, `已选择 ${file.name} (${formatBytes(file.size)})。`)
  setStatus(decodeStatus, '完成 encode 后即可 decode。')
}

async function summarizeDocument(docxDocument) {
  const pages = await docxDocument.getPages()
  const lines = []

  for (const page of pages) {
    const text = await page.getPureText()
    lines.push(`Page ${page.getNumber()}: ${text.trim() || '[empty page]'}`)
  }

  return { pages, text: lines.join('\n\n') }
}

async function encodeSelectedFile() {
  if (!selectedFile) return

  encodeButton.disabled = true
  resetDecodeResult()
  setStatus(encodeStatus, 'Encoding DOCX -> IntermediateDocument ...')

  try {
    encodedDocument = await DocxParser.encode(selectedFile)
    const intermediate = encodedDocument.getIntermediateDocument()
    const summary = await summarizeDocument(encodedDocument)

    metricTitle.textContent = encodedDocument.getTitle()
    metricId.textContent = encodedDocument.getId()
    metricPages.textContent = String(summary.pages.length)
    summaryPill.textContent = 'Encoded'
    preview.textContent = summary.text || '文档没有可提取文本，但已成功 encode。'
    decodeButton.disabled = false
    setStatus(encodeStatus, `Encode 完成：${intermediate.pageCount} 页。`, 'success')
    setStatus(decodeStatus, '可以开始 decode 生成 DOCX。')
  } catch (error) {
    encodedDocument = undefined
    summaryPill.textContent = 'Encode failed'
    preview.textContent = error instanceof Error ? error.stack || error.message : String(error)
    setStatus(encodeStatus, `Encode 失败：${error instanceof Error ? error.message : String(error)}`, 'error')
  } finally {
    encodeButton.disabled = !selectedFile
  }
}

async function decodeCurrentDocument() {
  if (!encodedDocument) return

  decodeButton.disabled = true
  setStatus(decodeStatus, 'Decoding IntermediateDocument -> DOCX ...')

  try {
    const buffer = await DocxParser.decode(encodedDocument.getIntermediateDocument())
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    })

    if (activeObjectUrl) URL.revokeObjectURL(activeObjectUrl)
    activeObjectUrl = URL.createObjectURL(blob)
    downloadLink.href = activeObjectUrl
    downloadLink.download = `decoded-${selectedFile?.name || 'hamster-note-demo.docx'}`
    downloadLink.hidden = false
    metricSize.textContent = formatBytes(buffer.byteLength)
    summaryPill.textContent = 'Decoded'
    setStatus(decodeStatus, `Decode 完成：生成 ${formatBytes(buffer.byteLength)} DOCX。`, 'success')
  } catch (error) {
    setStatus(decodeStatus, `Decode 失败：${error instanceof Error ? error.message : String(error)}`, 'error')
  } finally {
    decodeButton.disabled = !encodedDocument
  }
}

input.addEventListener('change', () => selectFile(input.files?.[0]))
encodeButton.addEventListener('click', encodeSelectedFile)
decodeButton.addEventListener('click', decodeCurrentDocument)

for (const eventName of ['dragenter', 'dragover']) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault()
    dropzone.classList.add('is-dragging')
  })
}

for (const eventName of ['dragleave', 'drop']) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault()
    dropzone.classList.remove('is-dragging')
  })
}

dropzone.addEventListener('drop', (event) => {
  const file = event.dataTransfer?.files?.[0]
  if (file) selectFile(file)
})
