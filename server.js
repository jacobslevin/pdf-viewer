const http = require("http")
const fs = require("fs")
const path = require("path")

const PORT = Number(process.env.PORT || 4173)
const ROOT = __dirname

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".pdf": "application/pdf",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml"
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  })
  response.end(JSON.stringify(payload))
}

function serveFile(response, filePath) {
  const extension = path.extname(filePath).toLowerCase()
  const contentType = MIME_TYPES[extension] || "application/octet-stream"

  fs.readFile(filePath, (error, content) => {
    if (error) {
      sendJson(response, 404, { error: "Not found" })
      return
    }

    response.writeHead(200, { "Content-Type": contentType })
    response.end(content)
  })
}

function resolveStaticPath(requestUrl) {
  const rawPath = requestUrl === "/" ? "/index.html" : requestUrl
  const pathname = decodeURIComponent(rawPath.split("?")[0])
  const resolvedPath = path.normalize(path.join(ROOT, pathname))
  if (!resolvedPath.startsWith(ROOT)) return null
  return resolvedPath
}

async function handleVisionRequest(request, response) {
  if (!process.env.OPENAI_API_KEY) {
    sendJson(response, 500, { error: "OPENAI_API_KEY is not set on the local server." })
    return
  }

  let rawBody = ""
  request.on("data", (chunk) => {
    rawBody += chunk
  })

  request.on("end", async () => {
    try {
      const body = JSON.parse(rawBody || "{}")
      const upstreamResponse = await fetch("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${process.env.OPENAI_API_KEY}`
        },
        body: JSON.stringify({
          model: body.model || "gpt-4.1",
          input: [
            {
              role: "user",
              content: [
                { type: "input_text", text: body.prompt || "" },
                { type: "input_image", image_url: body.imageUrl || "" }
              ]
            }
          ]
        })
      })

      const payload = await upstreamResponse.text()
      response.writeHead(upstreamResponse.status, {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
      })
      response.end(payload)
    } catch (error) {
      sendJson(response, 500, { error: error instanceof Error ? error.message : "Vision proxy failed." })
    }
  })
}

const server = http.createServer((request, response) => {
  if (!request.url) {
    sendJson(response, 400, { error: "Missing request URL." })
    return
  }

  if (request.method === "GET" && request.url.startsWith("/api/vision-status")) {
    sendJson(response, 200, { configured: Boolean(process.env.OPENAI_API_KEY) })
    return
  }

  if (request.method === "POST" && request.url.startsWith("/api/vision")) {
    handleVisionRequest(request, response)
    return
  }

  if (request.method !== "GET" && request.method !== "HEAD") {
    sendJson(response, 405, { error: "Method not allowed." })
    return
  }

  const filePath = resolveStaticPath(request.url)
  if (!filePath) {
    sendJson(response, 403, { error: "Forbidden path." })
    return
  }

  serveFile(response, filePath)
})

server.listen(PORT, () => {
  process.stdout.write(`Assisted Spec Capture POC server running at http://localhost:${PORT}\n`)
  process.stdout.write(`OPENAI_API_KEY configured: ${process.env.OPENAI_API_KEY ? "yes" : "no"}\n`)
})
