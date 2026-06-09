import { createReadStream, existsSync } from 'node:fs'
import { stat } from 'node:fs/promises'
import { createServer } from 'node:http'
import { networkInterfaces } from 'node:os'
import { extname, join, normalize, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const PORT = 8875
const HOST = '0.0.0.0'
const ROOT = resolve(fileURLToPath(new URL('..', import.meta.url)))

const contentTypes = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
}

function getLanUrls() {
  const urls = []

  // 绑定 0.0.0.0 后，局域网设备需要用本机网卡 IP 访问。
  for (const addresses of Object.values(networkInterfaces())) {
    for (const address of addresses ?? []) {
      if (address.family === 'IPv4' && !address.internal) {
        urls.push(`http://${address.address}:${PORT}`)
      }
    }
  }

  return urls
}

function resolveRequestPath(url) {
  const pathname = decodeURIComponent(new URL(url, `http://localhost:${PORT}`).pathname)
  const normalized = normalize(pathname).replace(/^\/+/, '')
  const candidate = resolve(ROOT, normalized || 'demo/index.html')

  if (!candidate.startsWith(ROOT)) return undefined
  if (existsSync(candidate)) return candidate

  const indexCandidate = join(candidate, 'index.html')
  return existsSync(indexCandidate) ? indexCandidate : undefined
}

const server = createServer(async (request, response) => {
  const filePath = resolveRequestPath(request.url ?? '/')

  if (!filePath) {
    response.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' })
    response.end('Not found')
    return
  }

  try {
    const fileStat = await stat(filePath)
    if (!fileStat.isFile()) throw new Error('Requested path is not a file')

    response.writeHead(200, {
      'content-length': fileStat.size,
      'content-type': contentTypes[extname(filePath)] ?? 'application/octet-stream'
    })
    createReadStream(filePath).pipe(response)
  } catch (error) {
    response.writeHead(500, { 'content-type': 'text/plain; charset=utf-8' })
    response.end(error instanceof Error ? error.message : String(error))
  }
})

server.listen(PORT, HOST, () => {
  console.log(`DocxParser demo running locally at http://localhost:${PORT}`)

  const lanUrls = getLanUrls()
  if (lanUrls.length > 0) {
    console.log('LAN access:')
    for (const url of lanUrls) console.log(`  ${url}`)
  }
})
