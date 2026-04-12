const http = require("http")
const fs = require("fs")
const path = require("path")

loadLocalEnv()

const PORT = Number(process.env.PORT || 4173)
const HOST = process.env.HOST || "0.0.0.0"
const UPSTREAM_FETCH_TIMEOUT_MS = 20000
const WEBSITE_PDF_DISCOVERY_TIMEOUT_MS = 8000
const RELATED_RESOURCE_PAGE_LIMIT = 3
const RELATED_RESOURCE_FETCH_TIMEOUT_MS = 8000
const OPENAI_API_TIMEOUT_MS = 40000
const ATTRIBUTE_ANSWER_MODEL = process.env.OPENAI_ATTRIBUTE_MODEL || "gpt-4.1"
const AMBIGUITY_CANDIDATE_LIMIT = 24
const AI_HTML_CHAR_LIMIT = 60000
const AI_TEXT_CHAR_LIMIT = 24000
const BROWSER_LIKE_FETCH_HEADERS = {
  "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
  "Accept-Language": "en-US,en;q=0.9",
  "Cache-Control": "no-cache",
  "Pragma": "no-cache"
}
const ROOT = __dirname
const LIVE_WATCH_FILES = [
  path.join(ROOT, "styles.css")
]

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

function loadLocalEnv() {
  const envPath = path.join(__dirname, ".env")
  if (!fs.existsSync(envPath)) return

  const envContent = fs.readFileSync(envPath, "utf8")
  for (const rawLine of envContent.split(/\r?\n/)) {
    const line = rawLine.trim()
    if (!line || line.startsWith("#")) continue

    const separatorIndex = line.indexOf("=")
    if (separatorIndex <= 0) continue

    const key = line.slice(0, separatorIndex).trim()
    if (!key || process.env[key]) continue

    let value = line.slice(separatorIndex + 1).trim()
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1)
    }

    process.env[key] = value
  }
}

function stripHtmlTags(value) {
  return String(value || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\s+/g, " ")
    .trim()
}

function normalizeWhitespace(value) {
  return String(value || "").replace(/\s+/g, " ").trim()
}

function collectMatches(html, regex, limit = 20) {
  const results = []
  for (const match of String(html || "").matchAll(regex)) {
    const cleaned = stripHtmlTags(match[1] || "")
    if (!cleaned) continue
    if (!results.includes(cleaned)) {
      results.push(cleaned)
    }
    if (results.length >= limit) break
  }
  return results
}

function decodeHtmlEntities(value) {
  return String(value || "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
}

function getAttributeValue(tagHtml, attributeName) {
  const match = String(tagHtml || "").match(new RegExp(`${attributeName}=["']([\\s\\S]*?)["']`, "i"))
  return decodeHtmlEntities(match?.[1] || "").trim()
}

function isGenericResourceLabel(label) {
  const normalized = stripHtmlTags(label).toLowerCase()
  if (!normalized) return true
  return (
    normalized === "download"
    || normalized === "downloads"
    || normalized === "resources"
    || normalized === "resource"
    || normalized === "open"
    || normalized === "view"
    || normalized === "learn more"
    || normalized.includes("download hi-res image")
    || normalized.includes("hi-res image")
    || normalized.includes("customize & price")
  )
}

function getResourceLabelFromHref(href) {
  try {
    const parsed = new URL(href, "https://placeholder.local")
    const rawName = decodeURIComponent(parsed.pathname.split("/").pop() || "")
      .replace(/\.[a-z0-9]+$/i, "")
      .replace(/[-_]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
    return rawName
  } catch {
    return ""
  }
}

function isDocumentHref(href) {
  return /\.(pdf|docx?|xlsx?|zip)(?:[\?#]|$)/i.test(String(href || ""))
}

function isPdfHref(href) {
  return /\.pdf(?:[\?#]|$)/i.test(String(href || ""))
}

function isImageHref(href) {
  return /\.(png|jpe?g|gif|webp|svg)(?:[\?#]|$)/i.test(String(href || ""))
}

function getResourcePriorityScore(item) {
  const href = String(item?.href || "").toLowerCase()
  const label = String(item?.label || "").toLowerCase()
  let score = 0
  if (isDocumentHref(href)) score += 100
  if (isPdfHref(href)) score += 20
  if (/\/documents\/|\/downloadslisting\/|\/downloads\//i.test(href)) score += 12
  if (/price|pricing|price book|product sheet|literature|spec/i.test(label)) score += 10
  if (/revit|autocad|cad|sketchup|model/i.test(label)) score += 4
  if (/resources\/materials|\/pro-resources\/?$/i.test(href)) score -= 20
  return score
}

function resolveHttpUrl(rawHref, baseUrl) {
  try {
    const resolved = new URL(String(rawHref || "").trim(), baseUrl)
    if (!/^https?:$/i.test(resolved.protocol)) return ""
    return resolved.toString()
  } catch {
    return ""
  }
}

function titleCaseLabel(value) {
  return String(value || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ")
}

function normalizeResourceLabel(value) {
  return String(value || "")
    .replace(/\[[^\]]*]/g, " ")
    .replace(/[–—]/g, "-")
    .replace(/[^a-z0-9]+/gi, " ")
    .trim()
    .toLowerCase()
}

function extractAnchorResources(html) {
  return [...String(html || "").matchAll(/(<a[^>]+href=["']([^"']+)["'][^>]*>)([\s\S]*?)<\/a>/gi)]
    .map((match) => {
      const fullTag = String(match[1] || "")
      const href = String(match[2] || "").trim()
      const innerLabel = stripHtmlTags(match[3] || "")
      const titleLabel = getAttributeValue(fullTag, "title")
      const ariaLabel = getAttributeValue(fullTag, "aria-label")
      const fallbackHrefLabel = getResourceLabelFromHref(href)
      const rawLabel = [innerLabel, titleLabel, ariaLabel].find((value) => !isGenericResourceLabel(value)) || fallbackHrefLabel || innerLabel || titleLabel || ariaLabel
      const label = titleCaseLabel(rawLabel)
      return { href, label }
    })
    .filter((item) => {
      if (!item.href) return false
      if (isImageHref(item.href)) return false
      if (isPdfHref(item.href)) return true
      return false
    })
    .filter((item) => !isGenericResourceLabel(item.label))
}

function extractGotoResources(html) {
  return [...String(html || "").matchAll(/<([a-z0-9]+)[^>]*\bgoto=["']([^"']+\.pdf(?:[\?#][^"']*)?\s*)["'][^>]*>/gi)]
    .map((match) => {
      const fullTag = String(match[0] || "")
      const href = String(match[2] || "").trim()
      const startIndex = Number(match.index || 0)
      const contextHtml = String(html || "").slice(startIndex, startIndex + 1600)
      const titleLabel = getAttributeValue(fullTag, "title")
      const ariaLabel = getAttributeValue(fullTag, "aria-label")
      const headingLabel = collectMatches(contextHtml, /<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>/gi, 1)[0] || ""
      const paragraphLabel = collectMatches(contextHtml, /<p[^>]*>([\s\S]*?)<\/p>/gi, 3).find((value) => !isGenericResourceLabel(value)) || ""
      const innerLabel = stripHtmlTags(contextHtml).slice(0, 240)
      const fallbackHrefLabel = getResourceLabelFromHref(href)
      const rawLabel = [headingLabel, titleLabel, ariaLabel, paragraphLabel, innerLabel].find((value) => !isGenericResourceLabel(value)) || fallbackHrefLabel || headingLabel || paragraphLabel || innerLabel || titleLabel || ariaLabel
      const label = titleCaseLabel(rawLabel)
      return { href, label }
    })
    .filter((item) => item.href && isPdfHref(item.href) && !isImageHref(item.href))
    .filter((item) => !isGenericResourceLabel(item.label))
}

function extractLikelyResourcePageLinks(html, baseUrl) {
  const pageLinks = [...String(html || "").matchAll(/(<a[^>]+href=["']([^"']+)["'][^>]*>)([\s\S]*?)<\/a>/gi)]
    .map((match) => {
      const fullTag = String(match[1] || "")
      const href = String(match[2] || "").trim()
      const innerLabel = stripHtmlTags(match[3] || "")
      const titleLabel = getAttributeValue(fullTag, "title")
      const ariaLabel = getAttributeValue(fullTag, "aria-label")
      const label = normalizeWhitespace(innerLabel || titleLabel || ariaLabel || getResourceLabelFromHref(href))
      const resolvedHref = resolveHttpUrl(href, baseUrl)
      return { href: resolvedHref, label }
    })
    .filter((item) => item.href && item.label && !isDocumentHref(item.href) && !isImageHref(item.href))

  let baseOrigin = ""
  try {
    baseOrigin = new URL(baseUrl).origin
  } catch {}

  const filtered = pageLinks.filter((item) => {
    const combined = `${item.label} ${item.href}`.toLowerCase()
    if (!/pro[-\s]?resources|resources|resource library|downloads?|pricing|price book|literature|documents?|specs?|materials\/resources/i.test(combined)) {
      return false
    }
    try {
      const targetOrigin = new URL(item.href).origin
      if (baseOrigin && targetOrigin !== baseOrigin) return false
    } catch {
      return false
    }
    return true
  })

  const deduped = []
  const seen = new Set()
  filtered.forEach((item) => {
    const key = item.href.toLowerCase()
    if (seen.has(key)) return
    seen.add(key)
    deduped.push(item)
  })

  const basePath = (() => {
    try {
      return new URL(baseUrl).pathname.toLowerCase()
    } catch {
      return ""
    }
  })()

  const scored = deduped
    .map((item) => {
      const href = item.href.toLowerCase()
      const label = item.label.toLowerCase()
      let score = 0
      if (/\/pro-resources\/?$/i.test(href)) score += 80
      if (/pro[-\s]?resources/.test(label)) score += 50
      if (/pricing|price book|literature|downloads?/.test(`${label} ${href}`)) score += 25
      if (/\/products\//.test(href)) score += 12
      if (/\/resources\/materials\/resources\/?$/i.test(href)) score -= 10
      if (basePath && href.includes(basePath.replace(/\/+$/, ""))) score += 18
      return {
        ...item,
        score
      }
    })
    .sort((left, right) => right.score - left.score)

  return scored.slice(0, RELATED_RESOURCE_PAGE_LIMIT).map(({ href, label }) => ({ href, label }))
}

function buildAdditionalResourcePageCandidates(baseUrl, html) {
  const candidates = extractLikelyResourcePageLinks(html, baseUrl)

  try {
    const parsed = new URL(baseUrl)
    const pathname = parsed.pathname || "/"
    if (/\/products\//i.test(pathname) && !/\/pro-resources\/?$/i.test(pathname)) {
      const proResourcesUrl = new URL(`${pathname.replace(/\/?$/, "/")}pro-resources/`, parsed.origin).toString()
      if (!candidates.some((item) => String(item?.href || "").toLowerCase() === proResourcesUrl.toLowerCase())) {
        candidates.unshift({
          href: proResourcesUrl,
          label: "Pro Resources"
        })
      }
    }
  } catch {
    // Ignore URL parsing issues and fall back to extracted links only.
  }

  const deduped = []
  const seen = new Set()
  candidates.forEach((item) => {
    const href = normalizeWhitespace(item?.href || "")
    if (!href) return
    const key = href.toLowerCase()
    if (seen.has(key)) return
    seen.add(key)
    deduped.push({
      href,
      label: normalizeWhitespace(item?.label || "") || getResourceLabelFromHref(href) || "Resources"
    })
  })
  return deduped.slice(0, RELATED_RESOURCE_PAGE_LIMIT)
}

function mergeResourceLists(...lists) {
  const resourceMap = new Map()
  lists.flat().forEach((item) => {
    const href = normalizeWhitespace(item?.href || "")
    if (!href || !isPdfHref(href)) return
    const label = normalizeWhitespace(item?.label || "")
    const existing = resourceMap.get(href)
    if (!existing) {
      resourceMap.set(href, { href, label: label || getResourceLabelFromHref(href) || "Resource" })
      return
    }
    const existingScore = normalizeResourceLabel(existing.label).length
    const nextScore = normalizeResourceLabel(label).length
    if (nextScore > existingScore) {
      resourceMap.set(href, { href, label: label || existing.label })
    }
  })
  return [...resourceMap.values()]
    .sort((left, right) => getResourcePriorityScore(right) - getResourcePriorityScore(left))
    .slice(0, 24)
}

function extractDownloadListingResources(html) {
  const resources = []
  const liMatches = [...String(html || "").matchAll(/<li[^>]*downloadFile=["']([^"']+)["'][^>]*>([\s\S]*?)<\/li>/gi)]

  liMatches.forEach((match) => {
    const href = normalizeWhitespace(match[1] || "")
    if (!href || isImageHref(href) || !isPdfHref(href)) return
    const liHtml = String(match[2] || "")
    const fileLabel =
      collectMatches(liHtml, /<span[^>]*class=["'][^"']*filename[^"']*["'][^>]*>([\s\S]*?)<\/span>/gi, 1)[0]
      || collectMatches(liHtml, /<a[^>]*>([\s\S]*?)<\/a>/gi, 1)[0]
      || getResourceLabelFromHref(href)
      || "Resource"
    const label = titleCaseLabel(String(fileLabel || "").replace(/\[[^\]]*]/g, " ").trim())
    resources.push({ href, label })
  })

  return resources
}

function summarizeHtmlDocument(html, url) {
  const cleanedText = stripHtmlTags(html)
  const title = collectMatches(html, /<title[^>]*>([\s\S]*?)<\/title>/gi, 1)[0] || ""
  const description =
    collectMatches(html, /<meta[^>]+name=["']description["'][^>]+content=["']([\s\S]*?)["'][^>]*>/gi, 1)[0]
    || collectMatches(html, /<meta[^>]+property=["']og:description["'][^>]+content=["']([\s\S]*?)["'][^>]*>/gi, 1)[0]
    || ""
  const headings = [
    ...collectMatches(html, /<h1[^>]*>([\s\S]*?)<\/h1>/gi, 4),
    ...collectMatches(html, /<h2[^>]*>([\s\S]*?)<\/h2>/gi, 6)
  ].slice(0, 6)
  const paragraphs = collectMatches(html, /<p[^>]*>([\s\S]*?)<\/p>/gi, 20)
    .filter((text) => text.length >= 60)
    .slice(0, 4)
  const bullets = [
    ...collectMatches(html, /<li[^>]*>([\s\S]*?)<\/li>/gi, 24),
    ...collectMatches(html, /<dt[^>]*>([\s\S]*?)<\/dt>/gi, 12)
  ]
    .filter((text) => text.length >= 8 && text.length <= 180)
    .slice(0, 6)
  const anchorResources = extractAnchorResources(html)
  const gotoResources = extractGotoResources(html)
  const downloadListingResources = extractDownloadListingResources(html)
  const resources = mergeResourceLists(downloadListingResources, anchorResources, gotoResources)

  return {
    url,
    title,
    description,
    headings,
    paragraphs,
    bullets,
    resources,
    cleanedText: cleanedText.slice(0, 24000)
  }
}

function extractDirectWebsitePdfResources(html) {
  const anchorResources = extractAnchorResources(html)
  const gotoResources = extractGotoResources(html)
  const downloadListingResources = extractDownloadListingResources(html)
  return mergeResourceLists(downloadListingResources, anchorResources, gotoResources)
    .filter((item) => /price|spec/i.test(`${item.label || ""} ${item.href || ""}`))
}

async function extractDirectWebsitePdfResourcesFromResponse(upstreamResponse) {
  const responseBody = upstreamResponse?.body
  if (!responseBody || typeof responseBody.getReader !== "function") {
    return extractDirectWebsitePdfResources(await upstreamResponse.text())
  }

  const reader = responseBody.getReader()
  const decoder = new TextDecoder()
  let html = ""
  const MAX_SCAN_CHARS = 750_000
  const MIN_EARLY_EXIT_SCAN_CHARS = 24_000

  try {
    while (html.length < MAX_SCAN_CHARS) {
      const { done, value } = await reader.read()
      if (done) break
      html += decoder.decode(value, { stream: true })

      if (html.length >= MIN_EARLY_EXIT_SCAN_CHARS) {
        const resources = extractDirectWebsitePdfResources(html)
        if (resources.length) {
          try {
            await reader.cancel()
          } catch {
            // Ignore cancellation failures; we already have enough to continue.
          }
          return resources
        }
      }
    }

    html += decoder.decode()
    return extractDirectWebsitePdfResources(html)
  } finally {
    try {
      reader.releaseLock()
    } catch {
      // Ignore release issues from partially consumed bodies.
    }
  }
}

function sanitizeHtmlForAi(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/\sdata-[a-z0-9-]+=(["'])[\s\S]*?\1/gi, "")
    .replace(/\son[a-z]+=(["'])[\s\S]*?\1/gi, "")
    .replace(/\sclass=(["'])[\s\S]*?\1/gi, "")
    .replace(/\sid=(["'])[\s\S]*?\1/gi, "")
    .replace(/\sstyle=(["'])[\s\S]*?\1/gi, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, AI_HTML_CHAR_LIMIT)
}

function extractHtmlEvidenceBlocks(html) {
  const blocks = []
  const pushBlock = (kind, content) => {
    const cleaned = normalizeWhitespace(stripHtmlTags(content))
    if (!cleaned || cleaned.length < 12) return
    blocks.push({ kind, text: cleaned })
  }

  for (const match of String(html || "").matchAll(/<tr\b[\s\S]*?<\/tr>/gi)) {
    pushBlock("table_row", match[0])
    if (blocks.length >= 80) break
  }
  if (blocks.length < 40) {
    for (const match of String(html || "").matchAll(/<li\b[\s\S]*?<\/li>/gi)) {
      pushBlock("list_item", match[0])
      if (blocks.length >= 80) break
    }
  }
  if (blocks.length < 40) {
    for (const match of String(html || "").matchAll(/<div\b[\s\S]*?<\/div>/gi)) {
      const raw = String(match[0] || "")
      if (!/\b(?:112\d{2}|113\d{2}|122\d{2}|123\d{2}|model|chair|stool|base|arms?|back)\b/i.test(raw)) continue
      pushBlock("div_block", raw)
      if (blocks.length >= 80) break
    }
  }

  const deduped = []
  const seen = new Set()
  blocks.forEach((block) => {
    const key = normalizeToken(block.text)
    if (!key || seen.has(key)) return
    seen.add(key)
    deduped.push(block)
  })
  return deduped.slice(0, 60)
}

function normalizeAttributeSchema(rawSchema) {
  const schema = Array.isArray(rawSchema) ? rawSchema : []
  return schema
    .map((item, index) => {
      if (typeof item === "string") {
        const label = normalizeWhitespace(item)
        if (!label) return null
        return {
          id: label.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || `attribute-${index + 1}`,
          label,
          required: true
        }
      }

      const label = normalizeWhitespace(item?.label || item?.name || item?.id || "")
      if (!label) return null
      return {
        id: normalizeWhitespace(item?.id || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || label.toLowerCase().replace(/[^a-z0-9]+/g, "-"),
        label,
        required: item?.required !== false
      }
    })
    .filter(Boolean)
}

function normalizeToken(value) {
  return normalizeWhitespace(value).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim()
}

function tokenize(value) {
  return normalizeToken(value).split(/\s+/).filter((token) => token.length >= 2)
}

function getProductMatchScore(summary, productName) {
  const expected = normalizeWhitespace(productName)
  if (!expected) return 0

  const haystacks = [
    summary?.title || "",
    ...(Array.isArray(summary?.headings) ? summary.headings : []),
    summary?.description || "",
    summary?.cleanedText || ""
  ]
    .map((value) => normalizeToken(value))
    .filter(Boolean)

  if (!haystacks.length) return 0

  const expectedTokenList = tokenize(expected)
  if (!expectedTokenList.length) return 0

  const exactName = normalizeToken(expected)
  const combined = haystacks.join(" ")
  if (combined.includes(exactName)) return 1

  const tokenHits = expectedTokenList.filter((token) => combined.includes(token)).length
  return Math.max(0, Math.min(1, tokenHits / expectedTokenList.length))
}

function getModelCodeMatches(text) {
  return [...new Set(
    [...String(text || "").matchAll(/\b(?:[A-Z]{1,5}[- ]?\d{2,5}[A-Z0-9-]*|\d{4,6})\b/g)]
      .map((match) => normalizeWhitespace(match[0] || ""))
      .filter(Boolean)
  )]
}

function getClarificationValue(userInputs = {}) {
  return normalizeWhitespace(
    userInputs?.model_number
    || userInputs?.modelNumber
    || userInputs?.variant_focus
    || userInputs?.variantFocus
    || userInputs?.product_line
    || userInputs?.productLine
    || ""
  )
}

function splitCleanedTextIntoSnippets(cleanedText) {
  return String(cleanedText || "")
    .split(/(?<=[\.\!\?])\s+|\n+/)
    .map((item) => normalizeWhitespace(item))
    .filter((item) => item.length >= 10 && item.length <= 220)
    .slice(0, 80)
}

function buildWebsiteEvidenceCorpus(summary) {
  const entries = []
  const push = (source, text) => {
    const normalized = normalizeWhitespace(text)
    if (!normalized) return
    entries.push({ source, text: normalized })
  }

  push("title", summary?.title)
  push("description", summary?.description)
  ;(summary?.headings || []).forEach((item) => push("heading", item))
  ;(summary?.paragraphs || []).forEach((item) => push("paragraph", item))
  ;(summary?.bullets || []).forEach((item) => push("bullet", item))
  splitCleanedTextIntoSnippets(summary?.cleanedText || "").forEach((item) => push("body", item))

  const deduped = []
  const seen = new Set()
  entries.forEach((entry) => {
    const key = normalizeToken(entry.text)
    if (!key || seen.has(key)) return
    seen.add(key)
    deduped.push(entry)
  })
  return deduped
}

function extractResponseText(payload) {
  if (!payload) return ""
  if (typeof payload.output_text === "string" && payload.output_text.trim()) return payload.output_text.trim()

  const collected = []
  const outputs = Array.isArray(payload.output) ? payload.output : []
  outputs.forEach((item) => {
    const contents = Array.isArray(item?.content) ? item.content : []
    contents.forEach((entry) => {
      if (typeof entry?.text === "string" && entry.text.trim()) {
        collected.push(entry.text.trim())
      }
    })
  })
  return collected.join("\n").trim()
}

function parseJsonPayload(rawText) {
  const text = normalizeWhitespace(rawText)
  if (!text) return null

  const directAttempt = (() => {
    try {
      return JSON.parse(text)
    } catch {
      return null
    }
  })()
  if (directAttempt) return directAttempt

  const fencedMatch = rawText.match(/```(?:json)?\s*([\s\S]*?)```/i)
  if (fencedMatch?.[1]) {
    try {
      return JSON.parse(fencedMatch[1].trim())
    } catch {}
  }

  const firstBrace = rawText.indexOf("{")
  const lastBrace = rawText.lastIndexOf("}")
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    try {
      return JSON.parse(rawText.slice(firstBrace, lastBrace + 1))
    } catch {}
  }

  return null
}

function buildAttributeAnswerPrompt(summary, rawHtml, productName, schema, userInputs = {}) {
  const normalizedSchema = normalizeAttributeSchema(schema)
  const resources = Array.isArray(summary?.resources) ? summary.resources.slice(0, 12) : []
  const headings = Array.isArray(summary?.headings) ? summary.headings.slice(0, 10) : []
  const bullets = Array.isArray(summary?.bullets) ? summary.bullets.slice(0, 16) : []
  const paragraphs = Array.isArray(summary?.paragraphs) ? summary.paragraphs.slice(0, 8) : []
  const cleanedText = String(summary?.cleanedText || "").slice(0, AI_TEXT_CHAR_LIMIT)
  const sanitizedHtml = sanitizeHtmlForAi(rawHtml)
  const htmlEvidenceBlocks = extractHtmlEvidenceBlocks(rawHtml)

  return [
    "You are an extraction engine for product research.",
    "Your job is to classify a webpage and fill structured product attributes.",
    "You are not a summarizer.",
    "",
    "Classify the page into exactly one of these states:",
    '1. FOUND: The page clearly maps to one identifiable product/model and the attributes can be safely extracted for that entity.',
    '2. NOT_FOUND: The page does not contain enough information to complete the requested attributes.',
    '3. AMBIGUOUS: The page contains relevant product data, but it spans multiple products, models, or configurations and cannot be safely tied to one entity.',
    "",
    "Critical rules:",
    "- If the page includes data for multiple models/configurations and the target model is unclear, outcome must be AMBIGUOUS.",
    "- Do not guess, average, merge, or mix values across models.",
    "- Only mark an attribute as filled when the value is explicitly supported and attributable to one entity.",
    "- When outcome is AMBIGUOUS, the message should ask for clarification before extraction proceeds.",
    "- When outcome is AMBIGUOUS, include as many concrete product/model candidates as are visible and useful, not just a tiny sample.",
    "- Use the sanitized HTML and extracted row/list blocks as the primary evidence for counting variants and identifying concrete model rows.",
    "- When outcome is NOT_FOUND, the message should say the webpage does not appear to contain enough information.",
    "- Always include the requested product name in the output.",
    "- Evidence snippets must be short excerpts copied from the provided content.",
    "- Return JSON only. No markdown fences.",
    "",
    "Required JSON shape:",
    "{",
    '  "outcome": "FOUND | NOT_FOUND | AMBIGUOUS",',
    '  "product_name": "string",',
    '  "message": "string",',
    '  "product_match": { "score": 0.0, "status": "strong | possible | weak" },',
    '  "attributes": [',
    '    {',
    '      "attribute_id": "string",',
    '      "label": "string",',
    '      "status": "filled | not_found | ambiguous",',
    '      "value": "string or null",',
    '      "confidence": 0.0,',
    '      "evidence_snippet": "string"',
    '    }',
    '  ],',
    '  "ambiguity": {',
    '    "reasons": ["string"],',
    '    "product_candidates": [',
    '      { "label": "string", "description": "short differentiator vs the nearby alternatives" }',
    '    ],',
    '    "narrowing_groups": [',
    '      { "label": "string", "options": ["string", "string"] }',
    '    ],',
    '    "required_user_inputs": ["model_number", "variant_focus"],',
    '    "action_prompts": ["string"]',
    '  },',
    '  "not_found": {',
    '    "reason": "string"',
    '  }',
    "}",
    "",
    `Requested product name: ${normalizeWhitespace(productName) || "Unknown"}`,
    `User clarification hints: ${JSON.stringify(userInputs || {})}`,
    `Target URL: ${summary?.url || ""}`,
    `Attribute schema: ${JSON.stringify(normalizedSchema)}`,
    `Page title: ${summary?.title || ""}`,
    `Meta description: ${summary?.description || ""}`,
    `Headings: ${JSON.stringify(headings)}`,
    `Bullets: ${JSON.stringify(bullets)}`,
    `Paragraphs: ${JSON.stringify(paragraphs)}`,
    `Linked resources: ${JSON.stringify(resources)}`,
    `Extracted html evidence blocks: ${JSON.stringify(htmlEvidenceBlocks)}`,
    `Visible text: ${cleanedText || ""}`,
    `Sanitized HTML excerpt: ${sanitizedHtml || ""}`
  ].join("\n")
}

function normalizeAiOutcome(value) {
  const normalized = normalizeWhitespace(value).toUpperCase()
  if (normalized === "FOUND") return "completed"
  if (normalized === "AMBIGUOUS") return "ambiguous"
  return "needs_pdf_escalation"
}

function normalizeAiProductMatch(payload) {
  const score = Number(payload?.score || 0)
  const status = normalizeWhitespace(payload?.status || "").toLowerCase()
  return {
    score: Number.isFinite(score) ? Math.max(0, Math.min(1, score)) : 0,
    status: status === "strong" || status === "possible" || status === "weak" ? status : "weak"
  }
}

function normalizeAiAttributes(rawAttributes, schema) {
  const byId = new Map((Array.isArray(rawAttributes) ? rawAttributes : []).map((item) => [normalizeWhitespace(item?.attribute_id || ""), item]))
  return normalizeAttributeSchema(schema).map((attribute) => {
    const item = byId.get(attribute.id) || byId.get(attribute.label) || {}
    const status = normalizeWhitespace(item?.status || "").toLowerCase()
    return {
      attributeId: attribute.id,
      label: attribute.label,
      value: normalizeWhitespace(item?.value || "") || null,
      confidence: Math.max(0, Math.min(1, Number(item?.confidence || 0))),
      status: status === "filled" || status === "ambiguous" ? status : "not_found",
      sourceStage: "webpage",
      evidence: normalizeWhitespace(item?.evidence_snippet || "")
        ? [{
            sourceType: "webpage",
            snippet: normalizeWhitespace(item.evidence_snippet)
          }]
        : []
    }
  })
}

function buildFallbackNarrowingGroups(candidates) {
  const items = Array.isArray(candidates) ? candidates : []
  const combinedText = items
    .map((candidate) => typeof candidate === "string" ? candidate : `${candidate?.label || ""} ${candidate?.description || ""}`)
    .join("\n")
    .toLowerCase()

  const groups = []
  const hasPair = (left, right) => combinedText.includes(left) && combinedText.includes(right)

  if (hasPair("with arms", "without arms") || hasPair("with arms", "armless") || hasPair("arms", "armless")) {
    groups.push({ label: "Arms", options: ["With Arms", "Without Arms"] })
  }

  const baseOptions = []
  ;[
    "wire base",
    "chrome base",
    "wood base",
    "four leg",
    "sled base",
    "caster base",
    "swivel base"
  ].forEach((option) => {
    if (combinedText.includes(option)) baseOptions.push(option.replace(/\b\w/g, (char) => char.toUpperCase()))
  })
  if (baseOptions.length >= 2) {
    groups.push({ label: "Base", options: [...new Set(baseOptions)].slice(0, 6) })
  }

  const backOptions = []
  ;["low back", "mid back", "high back"].forEach((option) => {
    if (combinedText.includes(option)) backOptions.push(option.replace(/\b\w/g, (char) => char.toUpperCase()))
  })
  if (backOptions.length >= 2) {
    groups.push({ label: "Back", options: [...new Set(backOptions)] })
  }

  return groups.slice(0, 6)
}

function buildFallbackNarrowingGroupsFromSummary(summary) {
  const combinedText = normalizeToken([
    summary?.title || "",
    ...(Array.isArray(summary?.headings) ? summary.headings : []),
    ...(Array.isArray(summary?.bullets) ? summary.bullets : []),
    ...(Array.isArray(summary?.paragraphs) ? summary.paragraphs : []),
    summary?.cleanedText || ""
  ].join(" "))

  const groups = []
  const includes = (value) => combinedText.includes(normalizeToken(value))

  if ((includes("with arms") || includes("arms")) && (includes("without arms") || includes("armless"))) {
    groups.push({ label: "Arms", options: ["With Arms", "Without Arms"] })
  }

  const baseOptions = []
  ;["Nylon Base", "Aluminum Base", "Chrome Base", "Wire Base", "Jury Base", "Caster Base", "Swivel Base"].forEach((option) => {
    if (includes(option)) baseOptions.push(option)
  })
  if (baseOptions.length >= 2) groups.push({ label: "Base", options: [...new Set(baseOptions)].slice(0, 8) })

  const backOptions = []
  ;["Mesh Back", "Upholstered Back", "Mid Back", "High Back", "Low Back"].forEach((option) => {
    if (includes(option)) backOptions.push(option)
  })
  if (backOptions.length >= 2) groups.push({ label: "Back", options: [...new Set(backOptions)].slice(0, 8) })

  return groups.slice(0, 6)
}

function normalizeProductCandidates(rawCandidates, fallbackCandidates = []) {
  const normalized = []
  const pushCandidate = (candidate) => {
    if (!candidate) return
    if (typeof candidate === "string") {
      const label = normalizeWhitespace(candidate)
      if (!label) return
      normalized.push({ label, description: "" })
      return
    }
    const label = normalizeWhitespace(candidate?.label || candidate?.id || "")
    if (!label) return
    normalized.push({
      label,
      description: normalizeWhitespace(candidate?.description || candidate?.blurb || "")
    })
  }

  ;(Array.isArray(rawCandidates) ? rawCandidates : []).forEach(pushCandidate)
  ;(Array.isArray(fallbackCandidates) ? fallbackCandidates : []).forEach((candidate) => pushCandidate(candidate))

  const seen = new Set()
  return normalized.filter((candidate) => {
    const key = normalizeToken(candidate.label)
    if (!key || seen.has(key)) return false
    if (isNonProductCandidate(candidate.label, candidate.description)) return false
    seen.add(key)
    return true
  }).slice(0, AMBIGUITY_CANDIDATE_LIMIT)
}

function isNonProductCandidate(label, description = "") {
  const combined = normalizeToken(`${label} ${description}`)
  if (!combined) return true
  if (/^(iso\s*\d+|files?\s*\d+|ca\s*\d+)$/.test(combined)) return true
  if (/\b(iso|leed|bifma|greenguard|certificate|certification|files?)\b/.test(combined)) return true
  return false
}

function normalizeSelectionTerms(userInputs = {}) {
  const narrowingSelections = Object.values(userInputs?.narrowing_selections || userInputs?.narrowingSelections || {})
    .map((item) => normalizeWhitespace(item))
    .filter(Boolean)
  const variantFocus = normalizeWhitespace(userInputs?.variant_focus || userInputs?.variantFocus || "")
  return [...new Set([...narrowingSelections, variantFocus].filter(Boolean))]
}

function candidateMatchesSelections(candidate, selectionTerms) {
  const selections = Array.isArray(selectionTerms) ? selectionTerms.map((item) => normalizeToken(item)).filter(Boolean) : []
  if (!selections.length) return true
  const haystack = normalizeToken(`${candidate?.label || ""} ${candidate?.description || ""}`)
  return selections.every((selection) => {
    if (selection === normalizeToken("Without Arms")) {
      return /\b(without arms|armless)\b/.test(haystack)
    }
    if (selection === normalizeToken("With Arms")) {
      return /\bwith arms\b/.test(haystack) && !/\bwithout arms|armless\b/.test(haystack)
    }
    return haystack.includes(selection)
  })
}

function normalizeNarrowingGroups(rawGroups, fallbackCandidates = [], summary = null) {
  const groups = []
  ;(Array.isArray(rawGroups) ? rawGroups : []).forEach((group) => {
    const label = normalizeWhitespace(group?.label || "")
    const options = [...new Set((Array.isArray(group?.options) ? group.options : [])
      .map((item) => normalizeWhitespace(item))
      .filter(Boolean))]
    if (!label || options.length < 2) return
    groups.push({ label, options: options.slice(0, 16) })
  })

  if (!groups.length) {
    const summaryGroups = buildFallbackNarrowingGroupsFromSummary(summary)
    return summaryGroups.length ? summaryGroups : buildFallbackNarrowingGroups(fallbackCandidates)
  }

  return groups.slice(0, 8)
}

async function requestOpenAiAttributeAnswer(summary, rawHtml, productName, schema, userInputs = {}) {
  if (!process.env.OPENAI_API_KEY) {
    throw new Error("OPENAI_API_KEY is required for AI-backed attribute extraction.")
  }

  const prompt = buildAttributeAnswerPrompt(summary, rawHtml, productName, schema, userInputs)
  const upstreamResponse = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    signal: AbortSignal.timeout(OPENAI_API_TIMEOUT_MS),
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`
    },
    body: JSON.stringify({
      model: ATTRIBUTE_ANSWER_MODEL,
      input: [
        {
          role: "user",
          content: [
            { type: "input_text", text: prompt }
          ]
        }
      ]
    })
  })

  const rawPayload = await upstreamResponse.text()
  if (!upstreamResponse.ok) {
    throw new Error(`OpenAI attribute extraction failed (${upstreamResponse.status}): ${rawPayload.slice(0, 300)}`)
  }

  let payload = null
  try {
    payload = JSON.parse(rawPayload)
  } catch {
    throw new Error("OpenAI attribute extraction returned unreadable JSON.")
  }

  const responseText = extractResponseText(payload)
  const parsed = parseJsonPayload(responseText)
  if (!parsed) {
    throw new Error("OpenAI attribute extraction did not return parseable JSON.")
  }
  return parsed
}

function detectProductCandidates(summary, rawHtml = "") {
  const candidates = new Set()
  const headings = Array.isArray(summary?.headings) ? summary.headings : []
  headings.forEach((heading) => {
    const normalized = normalizeWhitespace(heading)
    if (!normalized) return
    if (/\b(chair|stool|bench|table|ottoman|lounge|settee|sofa)\b/i.test(normalized) && normalized.split(/\s+/).length >= 2) {
      candidates.add(normalized)
    }
  })

  const rawSources = [
    summary?.cleanedText || "",
    stripHtmlTags(String(rawHtml || "")).slice(0, AI_HTML_CHAR_LIMIT)
  ]
  rawSources.forEach((source) => {
    getModelCodeMatches(source).slice(0, AMBIGUITY_CANDIDATE_LIMIT * 2).forEach((item) => candidates.add(item))
  })
  return [...candidates].slice(0, AMBIGUITY_CANDIDATE_LIMIT)
}

function detectAmbiguity(summary, productName, userInputs = {}) {
  const combinedText = [
    summary?.title || "",
    ...(Array.isArray(summary?.headings) ? summary.headings : []),
    ...(Array.isArray(summary?.paragraphs) ? summary.paragraphs : []),
    ...(Array.isArray(summary?.bullets) ? summary.bullets : []),
    summary?.cleanedText || ""
  ]
    .map((item) => normalizeWhitespace(item))
    .filter(Boolean)
    .join("\n")

  const lowered = combinedText.toLowerCase()
  const titleAndHeadings = [summary?.title || "", ...(summary?.headings || [])].join(" ").toLowerCase()
  const modelCodes = getModelCodeMatches(combinedText)
  const productCandidates = detectProductCandidates(summary)
  const clarification = getClarificationValue(userInputs)
  const clarificationTokens = tokenize(clarification)

  let score = 0
  const reasons = []

  if (/\b(family|series|collection|line)\b/.test(titleAndHeadings)) {
    score += 2
    reasons.push("The page is labeled like a product family or series.")
  }
  if (/\b(models?|configurations?|variants?|multiple options|available in)\b/.test(lowered)) {
    score += 2
    reasons.push("The content describes multiple models or configurations.")
  }
  if (modelCodes.length >= 2) {
    score += 3
    reasons.push("Multiple model codes appear on the page.")
  }
  if (productCandidates.length >= 2) {
    score += 2
    reasons.push("Multiple product candidates appear in the page headings or labels.")
  }
  if (/\bchoose\b|\bselect\b|\boption\b|\bupholstery\b|\bbase finish\b|\bframe finish\b/.test(lowered)) {
    score += 1
    reasons.push("The specs vary by selectable options.")
  }

  if (clarificationTokens.length) {
    const combinedNormalized = normalizeToken(combinedText)
    const clarificationHits = clarificationTokens.filter((token) => combinedNormalized.includes(token)).length
    if (clarificationHits >= Math.max(1, clarificationTokens.length - 1)) {
      score -= 3
    }
  }

  return {
    isAmbiguous: score >= 3,
    reasons: [...new Set(reasons)].slice(0, 4),
    productCandidates,
    requiredUserInputs: ["model_number", "variant_focus"],
    clarificationPrompt: "This page contains multiple products or configurations. I need a bit more direction to pull the right information.",
    actionPrompts: [
      "Can you confirm the model number?",
      "Do you know which specific product or variation you're looking for?",
      "Select or describe the option you want me to focus on."
    ]
  }
}

function getAttributeSearchTerms(attribute) {
  const label = normalizeWhitespace(attribute?.label || "")
  const normalized = normalizeToken(label)
  const synonyms = new Set([label, normalized])

  const synonymMap = [
    { test: /\bwidth\b/i, terms: ["width", "overall width", "w"] },
    { test: /\blength\b/i, terms: ["length", "overall length", "l"] },
    { test: /\bdepth\b/i, terms: ["depth", "overall depth", "d"] },
    { test: /\bheight\b/i, terms: ["height", "overall height", "h"] },
    { test: /\bseat height\b/i, terms: ["seat height", "sh"] },
    { test: /\barm height\b/i, terms: ["arm height", "ah"] },
    { test: /\bmaterial\b/i, terms: ["material", "materials", "construction"] },
    { test: /\bfinish\b/i, terms: ["finish", "finishes", "surface finish"] },
    { test: /\bapplication\b/i, terms: ["application", "use", "recommended use"] },
    { test: /\bthickness\b/i, terms: ["thickness", "gauge"] },
    { test: /\bupholstery\b/i, terms: ["upholstery", "fabric", "leather"] },
    { test: /\bbase\b/i, terms: ["base", "legs", "base finish"] }
  ]

  synonymMap.forEach((entry) => {
    if (entry.test.test(label)) {
      entry.terms.forEach((term) => synonyms.add(term))
    }
  })

  return [...synonyms]
    .map((item) => normalizeWhitespace(item))
    .filter(Boolean)
}

function extractDimensionLikeValue(text, terms) {
  for (const term of terms) {
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    const pattern = new RegExp(`\\b${escaped}\\b\\s*(?:[:\\-]|is|are)?\\s*([0-9][^,;\\|]{0,40}?(?:mm|cm|m|inches|inch|in\\.?|ft|\"))`, "i")
    const match = String(text || "").match(pattern)
    if (match?.[1]) return normalizeWhitespace(match[1])
  }
  return ""
}

function extractLabeledValue(text, terms) {
  for (const term of terms) {
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    const patterns = [
      new RegExp(`\\b${escaped}\\b\\s*(?:[:\\-]|is|are)\\s*([^.;\\n\\|]{2,100})`, "i"),
      new RegExp(`\\b${escaped}\\b\\s+([^.;\\n\\|]{2,80})`, "i")
    ]
    for (const pattern of patterns) {
      const match = String(text || "").match(pattern)
      if (!match?.[1]) continue
      const value = normalizeWhitespace(match[1])
        .replace(/^(available in|includes?|with)\s+/i, "")
        .replace(/\s+(available|shown|only)$/i, "")
      if (value && value.length <= 100) return value
    }
  }
  return ""
}

function classifyEvidenceQuality(source) {
  if (source === "title") return 0.72
  if (source === "heading") return 0.78
  if (source === "bullet") return 0.82
  if (source === "paragraph") return 0.76
  if (source === "description") return 0.65
  return 0.58
}

function extractAttributeValueFromCorpus(attribute, corpus, context = {}) {
  const terms = getAttributeSearchTerms(attribute)
  const normalizedLabel = normalizeToken(attribute?.label || "")
  let best = null

  corpus.forEach((entry) => {
    const text = entry.text
    const lowered = normalizeToken(text)
    const hasLabelHit = terms.some((term) => lowered.includes(normalizeToken(term)))
    if (!hasLabelHit) return

    const value =
      extractDimensionLikeValue(text, terms)
      || extractLabeledValue(text, terms)

    if (!value) return

    let confidence = 0.45
    if (text.includes(":")) confidence += 0.14
    if (lowered.includes(normalizedLabel)) confidence += 0.16
    confidence += classifyEvidenceQuality(entry.source) * 0.18
    confidence *= 0.75 + (Math.max(0, Math.min(1, Number(context.productMatchScore) || 0)) * 0.25)
    confidence = Math.max(0, Math.min(0.99, confidence))

    const candidate = {
      attributeId: attribute.id,
      label: attribute.label,
      value,
      confidence: Number(confidence.toFixed(2)),
      status: confidence >= 0.6 ? "filled" : "ambiguous",
      sourceStage: "webpage",
      evidence: [
        {
          sourceType: "webpage",
          url: context.url || "",
          snippet: text,
          section: entry.source
        }
      ]
    }

    if (!best || candidate.confidence > best.confidence) {
      best = candidate
    }
  })

  return best || {
    attributeId: attribute.id,
    label: attribute.label,
    value: null,
    confidence: 0,
    status: "not_found",
    sourceStage: "webpage",
    evidence: []
  }
}

function classifyPdfDocument(item, productName, schema) {
  const label = normalizeWhitespace(item?.label || "")
  const href = normalizeWhitespace(item?.href || "")
  const combined = `${label} ${href}`.toLowerCase()
  let docType = "unknown"
  if (/price book|pricing|pricelist|price list/.test(combined)) docType = "price_book"
  else if (/spec|product sheet|cutsheet|cut sheet|guide/.test(combined)) docType = "spec_sheet"
  else if (/install/.test(combined)) docType = "install_guide"
  else if (/brochure|lookbook/.test(combined)) docType = "brochure"

  let relevanceScore = 0.3
  if (docType === "spec_sheet") relevanceScore += 0.35
  if (docType === "price_book") relevanceScore += 0.24
  if (docType === "install_guide") relevanceScore += 0.12
  if (docType === "brochure") relevanceScore += 0.08

  const productTokens = tokenize(productName)
  const schemaTokens = normalizeAttributeSchema(schema).flatMap((attribute) => tokenize(attribute.label))
  const lowerCombined = combined.toLowerCase()
  productTokens.forEach((token) => {
    if (token && lowerCombined.includes(token)) relevanceScore += 0.08
  })
  schemaTokens.slice(0, 10).forEach((token) => {
    if (token && lowerCombined.includes(token)) relevanceScore += 0.02
  })

  return {
    id: href || label.toLowerCase().replace(/[^a-z0-9]+/g, "-"),
    url: href,
    title: label || getResourceLabelFromHref(href) || "PDF",
    docType,
    relevanceScore: Number(Math.max(0, Math.min(0.99, relevanceScore)).toFixed(2))
  }
}

function isHighValueWebsitePdf(item) {
  const combined = `${normalizeWhitespace(item?.title || "")} ${normalizeWhitespace(item?.url || "")}`.toLowerCase()
  return /price|spec/.test(combined)
}

function buildWebsitePdfDiscoveryResult(summary, productName, schema) {
  const normalizedSchema = normalizeAttributeSchema(schema)
  const candidateDocs = (summary?.resources || [])
    .map((item) => classifyPdfDocument(item, productName, normalizedSchema))
    .filter((item) => item.url)
    .filter((item) => isHighValueWebsitePdf(item))
    .sort((left, right) => right.relevanceScore - left.relevanceScore)
    .slice(0, 6)

  return {
    productName: normalizeWhitespace(productName),
    sourceUrl: summary?.url || "",
    productMatch: {
      score: 1,
      status: "strong"
    },
    attributes: normalizedSchema.map((attribute) => ({
      attributeId: attribute.id,
      label: attribute.label,
      value: null,
      confidence: 0,
      status: "not_found",
      sourceStage: "webpage",
      evidence: []
    })),
    completionScore: 0,
    overallConfidence: 0,
    outcome: "needs_pdf_escalation",
    message: candidateDocs.length
      ? "Using the webpage to find high-value PDFs for this product. Starting PDF extraction from the strongest price books and spec sheets."
      : "The webpage is being used as document discovery only, but no high-value PDFs were found here.",
    ambiguity: null,
    incompleteWebEvidence: {
      reasons: [
        "Website mode is using the webpage as PDF discovery context rather than as the final extraction source."
      ]
    },
    escalation: candidateDocs.length
      ? {
          message: "Prioritizing price books and spec sheets first.",
          candidateDocs
        }
      : null
  }
}

function detectIncompleteWebEvidence(rawHtml, summary, totalCandidateCount, selectionTerms = []) {
  const html = String(rawHtml || "")
  const text = normalizeToken([
    summary?.title || "",
    ...(Array.isArray(summary?.headings) ? summary.headings : []),
    ...(Array.isArray(summary?.bullets) ? summary.bullets : []),
    ...(Array.isArray(summary?.paragraphs) ? summary.paragraphs : []),
    stripHtmlTags(html)
  ].join(" "))

  const hasShowAllControl = /\bshow all\b/i.test(html) || text.includes("show all")
  const hasNextControl = /\bnext\b/i.test(html) || text.includes("next")
  const hasModelsSection = /\bmodels\b/i.test(html) || text.includes("models")
  const hasScrollableModelArea = /overflow|swiper|carousel|slider|slick|scroll/i.test(html)
  const hasPartialEnumerationRisk = (hasShowAllControl && hasNextControl) || (hasModelsSection && (hasNextControl || hasScrollableModelArea))

  return {
    isIncomplete: hasPartialEnumerationRisk && Number(totalCandidateCount || 0) <= 8 && !(Array.isArray(selectionTerms) && selectionTerms.length),
    reasons: [
      hasShowAllControl ? "The page exposes a Show All control in the model area." : "",
      hasNextControl ? "The page exposes a Next/pagination control in the model area." : "",
      hasScrollableModelArea ? "The model set appears to be inside a scrollable/carousel component." : ""
    ].filter(Boolean)
  }
}

function buildAttributeExtractionResultFromModel(modelResult, summary, rawHtml, productName, schema, userInputs = {}) {
  const normalizedSchema = normalizeAttributeSchema(schema)
  const attributes = normalizeAiAttributes(modelResult?.attributes, normalizedSchema)
  const requiredAttributes = normalizedSchema.filter((attribute) => attribute.required !== false)
  const filledRequired = attributes.filter((attribute) => attribute.status === "filled" && Number(attribute.confidence || 0) >= 0.6)
  const completionScore = requiredAttributes.length ? filledRequired.length / requiredAttributes.length : 0
  const overallConfidence = attributes.length
    ? attributes.reduce((sum, attribute) => sum + Number(attribute.confidence || 0), 0) / attributes.length
    : 0
  const candidateDocs = (summary?.resources || [])
    .map((item) => classifyPdfDocument(item, productName, normalizedSchema))
    .filter((item) => item.url)
    .sort((left, right) => right.relevanceScore - left.relevanceScore)
    .slice(0, 6)
  const outcome = normalizeAiOutcome(modelResult?.outcome)
  const detectedCandidates = detectProductCandidates(summary, rawHtml).map((label) => ({ label, description: "" }))
  const selectionTerms = normalizeSelectionTerms(userInputs)
  const mergedProductCandidates = normalizeProductCandidates(modelResult?.ambiguity?.product_candidates, detectedCandidates)
    .filter((candidate) => candidateMatchesSelections(candidate, selectionTerms))
    .slice(0, AMBIGUITY_CANDIDATE_LIMIT)
  const modelCandidateCount = Number(modelResult?.ambiguity?.candidate_count || 0)
  const totalCandidateCount = Math.max(
    normalizeProductCandidates(modelResult?.ambiguity?.product_candidates, detectedCandidates).length,
    detectedCandidates.length,
    Number.isFinite(modelCandidateCount) ? modelCandidateCount : 0
  )
  const filteredCandidateCount = selectionTerms.length
    ? mergedProductCandidates.length
    : totalCandidateCount
  const incompleteWebEvidence = detectIncompleteWebEvidence(rawHtml, summary, totalCandidateCount, selectionTerms)
  const effectiveOutcome = incompleteWebEvidence.isIncomplete && outcome === "ambiguous"
    ? "needs_pdf_escalation"
    : outcome
  const ambiguity = outcome === "ambiguous"
    ? {
        reasons: Array.isArray(modelResult?.ambiguity?.reasons) ? modelResult.ambiguity.reasons.map((item) => normalizeWhitespace(item)).filter(Boolean).slice(0, 4) : [],
        productCandidates: mergedProductCandidates,
        candidateCount: filteredCandidateCount,
        totalCandidateCount,
        narrowingGroups: normalizeNarrowingGroups(modelResult?.ambiguity?.narrowing_groups, mergedProductCandidates, summary),
        requiredUserInputs: Array.isArray(modelResult?.ambiguity?.required_user_inputs) ? modelResult.ambiguity.required_user_inputs.map((item) => normalizeWhitespace(item)).filter(Boolean).slice(0, 4) : ["model_number", "variant_focus"],
        actionPrompts: Array.isArray(modelResult?.ambiguity?.action_prompts) ? modelResult.ambiguity.action_prompts.map((item) => normalizeWhitespace(item)).filter(Boolean).slice(0, 4) : [
          "Can you confirm the model number?",
          "Do you know which specific product or variation you're looking for?",
          "Select or describe the option you want me to focus on."
        ]
      }
    : null

  return {
    productName: normalizeWhitespace(modelResult?.product_name || productName),
    sourceUrl: summary?.url || "",
    productMatch: normalizeAiProductMatch(modelResult?.product_match),
    attributes,
    completionScore: Number(completionScore.toFixed(2)),
    overallConfidence: Number(overallConfidence.toFixed(2)),
    outcome: effectiveOutcome,
    message: normalizeWhitespace(modelResult?.message || "")
      || (effectiveOutcome === "ambiguous"
        ? "This page contains multiple products or configurations. I need a bit more direction to pull the right information."
        : effectiveOutcome === "needs_pdf_escalation"
          ? incompleteWebEvidence.isIncomplete
            ? "This webpage appears to show only part of the available models or configurations, so I may not be seeing the full set reliably here."
            : "The webpage doesn't appear to contain the information needed to complete these attributes."
          : ""),
    incompleteWebEvidence: incompleteWebEvidence.isIncomplete
      ? { reasons: incompleteWebEvidence.reasons }
      : null,
    ambiguity: effectiveOutcome === "ambiguous"
      ? ambiguity
      : null,
    escalation: effectiveOutcome === "needs_pdf_escalation" && candidateDocs.length
      ? {
        message: incompleteWebEvidence.isIncomplete
          ? "I found high-value PDFs and can check those instead."
          : "We found high-value documents. Would you like me to check those?",
        candidateDocs
      }
      : null
  }
}

async function discoverRelatedPageResources(baseHtml, baseUrl) {
  const candidatePages = extractLikelyResourcePageLinks(baseHtml, baseUrl)
  if (!candidatePages.length) return []

  const discoveredResources = []
  for (const page of candidatePages) {
    try {
      const response = await fetch(page.href, {
        method: "GET",
        redirect: "follow",
        signal: AbortSignal.timeout(RELATED_RESOURCE_FETCH_TIMEOUT_MS),
        headers: BROWSER_LIKE_FETCH_HEADERS
      })
      if (!response.ok) continue
      const html = await response.text()
      const pageSummary = summarizeHtmlDocument(html, response.url || page.href)
      discoveredResources.push(...pageSummary.resources)
    } catch {
      // Ignore one-hop fetch failures and keep primary summary intact.
    }
  }
  return discoveredResources
}

function normalizeHeaderValue(value) {
  return String(value || "").trim()
}

function parseFrameAncestors(cspValue) {
  const directives = String(cspValue || "")
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
  const frameAncestorsDirective = directives.find((directive) => /^frame-ancestors\b/i.test(directive))
  if (!frameAncestorsDirective) return []
  return frameAncestorsDirective
    .replace(/^frame-ancestors\b/i, "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
}

function evaluateEmbeddability(targetUrl, viewerOrigin, headers) {
  const xFrameOptions = normalizeHeaderValue(headers.get("x-frame-options"))
  const contentSecurityPolicy = normalizeHeaderValue(headers.get("content-security-policy"))
  const frameAncestors = parseFrameAncestors(contentSecurityPolicy)
  const reasons = []
  const targetOrigin = new URL(targetUrl).origin

  if (xFrameOptions) {
    const normalized = xFrameOptions.toUpperCase()
    if (normalized.includes("DENY")) {
      reasons.push("x-frame-options=DENY")
    } else if (normalized.includes("SAMEORIGIN")) {
      if (viewerOrigin !== targetOrigin) {
        reasons.push("x-frame-options=SAMEORIGIN")
      }
    }
  }

  if (frameAncestors.length) {
    const normalizedAncestors = frameAncestors.map((value) => value.replace(/^"+|"+$/g, "").replace(/^'+|'+$/g, ""))
    if (normalizedAncestors.includes("'none'") || normalizedAncestors.includes("none")) {
      reasons.push("csp frame-ancestors 'none'")
    } else {
      const allowsAll = normalizedAncestors.includes("*")
      const allowsSelfOnly = normalizedAncestors.includes("'self'") || normalizedAncestors.includes("self")
      const explicitlyAllowsViewerOrigin = normalizedAncestors.some((value) => value === viewerOrigin)
      if (!allowsAll && allowsSelfOnly && viewerOrigin !== targetOrigin) {
        reasons.push("csp frame-ancestors 'self'")
      } else if (!allowsAll && !allowsSelfOnly && !explicitlyAllowsViewerOrigin) {
        reasons.push("csp frame-ancestors does not allow viewer origin")
      }
    }
  }

  return {
    blocked: reasons.length > 0,
    reasons,
    xFrameOptions,
    contentSecurityPolicy
  }
}

async function fetchEmbeddability(url, viewerOrigin) {
  const attempts = ["HEAD", "GET"]
  let lastError = null

  for (const method of attempts) {
    try {
      const response = await fetch(url, {
        method,
        redirect: "follow",
        headers: BROWSER_LIKE_FETCH_HEADERS
      })
      return {
        ok: true,
        method,
        finalUrl: response.url || url,
        status: response.status,
        ...evaluateEmbeddability(response.url || url, viewerOrigin, response.headers)
      }
    } catch (error) {
      lastError = error
    }
  }

  return {
    ok: false,
    blocked: false,
    reasons: [],
    error: lastError instanceof Error ? lastError.message : "Unable to check embeddability."
  }
}

async function handleEmbedCheckRequest(request, response) {
  try {
    const requestUrl = new URL(request.url, `http://localhost:${PORT}`)
    const rawTargetUrl = requestUrl.searchParams.get("url") || ""
    let targetUrl = ""
    try {
      const parsedUrl = new URL(rawTargetUrl)
      if (!/^https?:$/i.test(parsedUrl.protocol)) {
        sendJson(response, 400, { error: "Only http:// and https:// URLs are supported." })
        return
      }
      targetUrl = parsedUrl.toString()
    } catch {
      sendJson(response, 400, { error: "Invalid URL." })
      return
    }

    const forwardedProto = normalizeHeaderValue(request.headers["x-forwarded-proto"])
    const host = normalizeHeaderValue(request.headers.host) || `localhost:${PORT}`
    const viewerOrigin = `${forwardedProto || "http"}://${host}`
    const result = await fetchEmbeddability(targetUrl, viewerOrigin)
    sendJson(response, 200, result)
  } catch (error) {
    sendJson(response, 500, { error: error instanceof Error ? error.message : "Unable to check embeddability." })
  }
}

async function handleWebsiteSummaryRequest(request, response) {
  try {
    const requestUrl = new URL(request.url, `http://localhost:${PORT}`)
    const rawTargetUrl = requestUrl.searchParams.get("url") || ""
    let targetUrl = ""
    try {
      const parsedUrl = new URL(rawTargetUrl)
      if (!/^https?:$/i.test(parsedUrl.protocol)) {
        sendJson(response, 400, { error: "Only http:// and https:// URLs are supported." })
        return
      }
      targetUrl = parsedUrl.toString()
    } catch {
      sendJson(response, 400, { error: "Invalid URL." })
      return
    }

    const upstreamResponse = await fetch(targetUrl, {
      method: "GET",
      redirect: "follow",
      signal: AbortSignal.timeout(UPSTREAM_FETCH_TIMEOUT_MS),
      headers: BROWSER_LIKE_FETCH_HEADERS
    })

    const html = await upstreamResponse.text()
    const summary = summarizeHtmlDocument(html, upstreamResponse.url || targetUrl)
    sendJson(response, 200, {
      ok: true,
      status: upstreamResponse.status,
      finalUrl: upstreamResponse.url || targetUrl,
      summary
    })
  } catch (error) {
    sendJson(response, 500, { ok: false, error: error instanceof Error ? error.message : "Unable to summarize website." })
  }
}

async function handleWebsitePdfDiscoveryRequest(request, response) {
  try {
    const requestUrl = new URL(request.url, `http://localhost:${PORT}`)
    const rawTargetUrl = requestUrl.searchParams.get("url") || ""
    let targetUrl = ""
    try {
      const parsedUrl = new URL(rawTargetUrl)
      if (!/^https?:$/i.test(parsedUrl.protocol)) {
        sendJson(response, 400, { error: "Only http:// and https:// URLs are supported." })
        return
      }
      targetUrl = parsedUrl.toString()
    } catch {
      sendJson(response, 400, { error: "Invalid URL." })
      return
    }

    const upstreamResponse = await fetch(targetUrl, {
      method: "GET",
      redirect: "follow",
      signal: AbortSignal.timeout(WEBSITE_PDF_DISCOVERY_TIMEOUT_MS),
      headers: BROWSER_LIKE_FETCH_HEADERS
    })

    const finalUrl = upstreamResponse.url || targetUrl
    const pageHtml = await upstreamResponse.text()
    let resources = extractDirectWebsitePdfResources(pageHtml)
    let scannedUrl = finalUrl

    if (!resources.length) {
      const fallbackPage = buildAdditionalResourcePageCandidates(finalUrl, pageHtml)[0] || null
      if (fallbackPage?.href) {
        try {
          const fallbackResponse = await fetch(fallbackPage.href, {
            method: "GET",
            redirect: "follow",
            signal: AbortSignal.timeout(RELATED_RESOURCE_FETCH_TIMEOUT_MS),
            headers: BROWSER_LIKE_FETCH_HEADERS
          })
          const fallbackHtml = await fallbackResponse.text()
          const fallbackResources = extractDirectWebsitePdfResources(fallbackHtml)
          if (fallbackResources.length) {
            resources = fallbackResources
            scannedUrl = fallbackResponse.url || fallbackPage.href
          }
        } catch {
          // Ignore fallback page fetch failures and return the direct-page result.
        }
      }
    }

    sendJson(response, 200, {
      ok: true,
      status: upstreamResponse.status,
      finalUrl: scannedUrl,
      resources
    })
  } catch (error) {
    sendJson(response, 500, { ok: false, error: error instanceof Error ? error.message : "Unable to scan webpage PDFs." })
  }
}

function readJsonBody(request) {
  return new Promise((resolve, reject) => {
    let rawBody = ""
    request.on("data", (chunk) => {
      rawBody += chunk
      if (rawBody.length > 1_000_000) {
        reject(new Error("Request body too large."))
      }
    })
    request.on("end", () => {
      try {
        resolve(JSON.parse(rawBody || "{}"))
      } catch {
        reject(new Error("Invalid JSON body."))
      }
    })
    request.on("error", (error) => reject(error))
  })
}

async function handleAttributeAnswerRequest(request, response) {
  try {
    const body = await readJsonBody(request)
    const productUrl = normalizeWhitespace(body?.product_url || body?.productUrl || "")
    const productName = normalizeWhitespace(body?.product_name || body?.productName || "")
    const attributeSchema = normalizeAttributeSchema(body?.attribute_schema || body?.attributeSchema)

    if (!productUrl) {
      sendJson(response, 400, { ok: false, error: "product_url is required." })
      return
    }

    if (!productName) {
      sendJson(response, 400, { ok: false, error: "product_name is required." })
      return
    }

    if (!attributeSchema.length) {
      sendJson(response, 400, { ok: false, error: "attribute_schema is required." })
      return
    }

    let targetUrl = ""
    try {
      const parsedUrl = new URL(productUrl)
      if (!/^https?:$/i.test(parsedUrl.protocol)) {
        sendJson(response, 400, { ok: false, error: "Only http:// and https:// URLs are supported." })
        return
      }
      targetUrl = parsedUrl.toString()
    } catch {
      sendJson(response, 400, { ok: false, error: "Invalid product_url." })
      return
    }

    const upstreamResponse = await fetch(targetUrl, {
      method: "GET",
      redirect: "follow",
      signal: AbortSignal.timeout(UPSTREAM_FETCH_TIMEOUT_MS),
      headers: BROWSER_LIKE_FETCH_HEADERS
    })

    const html = await upstreamResponse.text()
    const summary = summarizeHtmlDocument(html, upstreamResponse.url || targetUrl)

    const extraction = buildWebsitePdfDiscoveryResult(summary, productName, attributeSchema)
    sendJson(response, 200, {
      ok: true,
      status: upstreamResponse.status,
      finalUrl: upstreamResponse.url || targetUrl,
      extraction,
      source: {
        title: summary.title,
        description: summary.description,
        headings: summary.headings,
        resources: summary.resources
      }
    })
  } catch (error) {
    sendJson(response, 500, { ok: false, error: error instanceof Error ? error.message : "Unable to extract webpage attributes." })
  }
}

async function handlePdfProxyRequest(request, response) {
  try {
    const requestUrl = new URL(request.url, `http://localhost:${PORT}`)
    const rawTargetUrl = requestUrl.searchParams.get("url") || ""
    let targetUrl = ""
    try {
      const parsedUrl = new URL(rawTargetUrl)
      if (!/^https?:$/i.test(parsedUrl.protocol)) {
        sendJson(response, 400, { error: "Only http:// and https:// URLs are supported." })
        return
      }
      targetUrl = parsedUrl.toString()
    } catch {
      sendJson(response, 400, { error: "Invalid URL." })
      return
    }

    const upstreamResponse = await fetch(targetUrl, {
      method: "GET",
      redirect: "follow",
      headers: BROWSER_LIKE_FETCH_HEADERS
    })

    if (!upstreamResponse.ok) {
      sendJson(response, upstreamResponse.status, { error: `Unable to fetch remote PDF (${upstreamResponse.status}).` })
      return
    }

    const contentType = String(upstreamResponse.headers.get("content-type") || "")
    const finalUrl = upstreamResponse.url || targetUrl
    if (!/application\/pdf/i.test(contentType) && !/\.pdf(?:[\?#]|$)/i.test(finalUrl)) {
      sendJson(response, 400, { error: "The requested resource does not appear to be a PDF." })
      return
    }

    const fileBuffer = Buffer.from(await upstreamResponse.arrayBuffer())
    response.writeHead(200, {
      "Content-Type": "application/pdf",
      "Cache-Control": "no-store",
      "Content-Length": String(fileBuffer.length),
      "Content-Disposition": `inline; filename="${path.basename(new URL(finalUrl).pathname || "resource.pdf") || "resource.pdf"}"`
    })
    response.end(fileBuffer)
  } catch (error) {
    sendJson(response, 500, { error: error instanceof Error ? error.message : "Unable to fetch remote PDF." })
  }
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
  const cacheControl =
    extension === ".html" || extension === ".js" || extension === ".css"
      ? "no-store"
      : "public, max-age=3600"

  fs.readFile(filePath, (error, content) => {
    if (error) {
      sendJson(response, 404, { error: "Not found" })
      return
    }

    let responseBody = content
    if (extension === ".html") {
      const getVersion = (assetPath) => {
        try {
          const stats = fs.statSync(path.join(ROOT, assetPath))
          return String(Math.floor(stats.mtimeMs))
        } catch {
          return String(Date.now())
        }
      }
      const appVersion = getVersion("app.js")
      const styleVersion = getVersion("styles.css")
      responseBody = Buffer.from(
        String(content)
          .replace(/\.\/styles\.css(?:\?[^"']*)?/g, `./styles.css?v=${styleVersion}`)
          .replace(/\.\/app\.js(?:\?[^"']*)?/g, `./app.js?v=${appVersion}`),
        "utf8"
      )
    }

    response.writeHead(200, {
      "Content-Type": contentType,
      "Cache-Control": cacheControl
    })
    response.end(responseBody)
  })
}

function getLiveStyleVersion() {
  return LIVE_WATCH_FILES.reduce((latest, filePath) => {
    try {
      const stats = fs.statSync(filePath)
      return Math.max(latest, stats.mtimeMs)
    } catch {
      return latest
    }
  }, 0)
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

async function handleOpenAiResponsesProxyRequest(request, response) {
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
        body: JSON.stringify(body)
      })

      const payload = await upstreamResponse.text()
      response.writeHead(upstreamResponse.status, {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
      })
      response.end(payload)
    } catch (error) {
      sendJson(response, 500, { error: error instanceof Error ? error.message : "OpenAI proxy failed." })
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

  if (request.method === "GET" && request.url.startsWith("/api/live-css")) {
    sendJson(response, 200, { version: getLiveStyleVersion() })
    return
  }

  if (request.method === "GET" && request.url.startsWith("/api/embed-check")) {
    handleEmbedCheckRequest(request, response)
    return
  }

  if (request.method === "GET" && request.url.startsWith("/api/website-summary")) {
    handleWebsiteSummaryRequest(request, response)
    return
  }

  if (request.method === "GET" && request.url.startsWith("/api/website-pdfs")) {
    handleWebsitePdfDiscoveryRequest(request, response)
    return
  }

  if (request.method === "POST" && request.url.startsWith("/api/attribute-answer")) {
    handleAttributeAnswerRequest(request, response)
    return
  }

  if (request.method === "GET" && request.url.startsWith("/api/pdf-proxy")) {
    handlePdfProxyRequest(request, response)
    return
  }

  if (request.method === "POST" && request.url.startsWith("/api/vision")) {
    handleVisionRequest(request, response)
    return
  }

  if (request.method === "POST" && request.url.startsWith("/api/openai-responses")) {
    handleOpenAiResponsesProxyRequest(request, response)
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

server.listen(PORT, HOST, () => {
  process.stdout.write(`Assisted Spec Capture POC server running at http://${HOST}:${PORT}\n`)
  process.stdout.write(`OPENAI_API_KEY configured: ${process.env.OPENAI_API_KEY ? "yes" : "no"}\n`)
})
