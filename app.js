(function () {
  const PDFJS_VERSION = "4.8.69"
  const PERSISTENCE_DB = "assisted-spec-capture-poc"
  const PERSISTENCE_STORE = "defaults"
  const PERSISTENCE_KEY = "active-default-v2"
  const DEFAULT_BUNDLED_PDF_PATH = "./PBHCL%20(4).pdf"
  const DEFAULT_BUNDLED_PDF_NAME = "PBHCL (4).pdf"
  const DEFAULT_VISION_MODEL = "gpt-4.1"
  const PRODUCT_FIRST_CONFIDENCE_THRESHOLD = 0.78
  const HELP_SPEC_LOADING_INTERVAL_MS = 4200
  const HELP_SPEC_LOADING_LINES = [
    "Analyzing details... good specs take a beat.",
    "Reading the relevant pages, not just the one you clicked.",
    "Sorting dimensions, finishes, upholstery, and pricing into something human.",
    "Architects love clean lines. PDFs do not always cooperate.",
    "COM means Customer's Own Material. COL means Customer's Own Leather.",
    "Cross-checking reference pages so one finish table does not hijack the whole result.",
    "Looking for the exact product before it looks for the exact finish.",
    "Separating family-level pages from variant-level details.",
    "Comparing headers, pricing tables, and dimensions across candidate pages.",
    "Checking whether this page is one product, several products, or a family overview.",
    "Trying not to let marketing photography win the argument over the spec table.",
    "Matching finishes, materials, and dimensions to the right context.",
    "Pulling structure out of a layout that definitely did not volunteer it.",
    "Scanning for clues in titles, captions, tables, and tiny labels.",
    "Figuring out which pages are shared references versus product-specific pages.",
    "Looking for the point where options become actual spec decisions.",
    "Keeping the useful pages and ignoring the decorative detours.",
    "Turning price-list chaos into tabs you can actually use.",
    "Checking whether a nearby page is a continuation, a sibling variant, or neither.",
    "Separating configuration choices from finish choices.",
    "Making sure dimensions stay attached to the right variant.",
    "Trying to keep one powder coat legend from becoming the answer to everything.",
    "Looking for repeated family names, model cues, and variant language.",
    "Reducing the odds that one ambiguous SKU sends the whole result sideways.",
    "Searching for the pages that matter after the click, not just next to it.",
    "Comparing shared spec pages against variant-specific pages.",
    "Finding where the actual specification content begins.",
    "Checking if the page is broad enough to require a selector first.",
    "Grouping together pages that describe the same family context.",
    "Making the PDF admit what should have been structured data.",
    "Reading the page like a spec editor, not like a thumbnail strip.",
    "Trying to preserve exact units, fractions, and notation while sorting content.",
    "Looking for the details that change the spec, not just the details that fill the page."
  ]

  const sampleDocuments = buildSampleDocuments()
  const initialSpec = {
    originalSpecName: "Eames Molded Plywood Chairs",
    specDisplayName: "Eames Molded Plywood Chairs",
    originalBrand: "Herman Miller",
    brandDisplayName: "Herman Miller",
    category: "Lounge Seating",
    attributes: [
      { key: "width", label: "Width", type: "text", value: "" },
      { key: "length", label: "Length", type: "text", value: "" },
      { key: "overall-height", label: "Overall Height", type: "text", value: "" },
      { key: "seat-height", label: "Seat Height", type: "text", value: "" },
      { key: "arm-height", label: "Arm Height", type: "text", value: "" },
      { key: "body-frame", label: "Body/Frame", type: "text", value: "" },
      { key: "legs-base", label: "Legs/Base", type: "text", value: "" }
    ]
  }

  const appState = {
    spec: cloneSpec(initialSpec),
    documents: [],
    activeDocumentId: "",
    activePageNumber: 1,
    rankedPages: [],
    selectionText: "",
    selectionSource: "",
    selectionRect: null,
    copiedText: "",
    pickerPosition: null,
    confirmPosition: null,
    pendingFieldKey: "",
    pendingInsertText: "",
    highlightedFieldKey: "",
    highlightTimeoutId: null,
    toastMessage: "",
    toastTimeoutId: null,
    loadingMessage: "",
    errorMessage: "",
    assistantResult: null,
    assistantSelection: "",
    pagePreviews: {},
    pageRenderImages: {},
    pageRenderTextByKey: {},
    pageRenderMetricsByKey: {},
    pageRenderStatusByKey: {},
    pageTextVisibleByKey: {},
    pageReadableHtmlByKey: {},
    pageReadableStatusByKey: {},
    pageReadableErrorByKey: {},
    wordStatsPage: null,
    ocrErrorByPage: {},
    ocrStatusByPage: {},
    visionApiKey: "",
    visionModel: DEFAULT_VISION_MODEL,
    productImageDataUrl: "",
    productImageName: "",
    productImageUrl: "",
    aiRerankLoading: false,
    analyzeRequestLoading: false,
    aiRerankResult: null,
    aiRerankDocumentId: "",
    aiRerankCacheByDocumentId: {},
    structureRouting: null,
    productFirstSelection: null,
    aiRerankError: "",
    sourceSelectionScores: [],
    sourceSelectionChosenId: "",
    decisionAssistLoading: false,
    helpSpecLoadingLineIndex: 0,
    helpSpecLoadingLineOrder: [],
    decisionAssistResult: null,
    familyPageProductsByKey: {},
    familyPageProductsStatusByKey: {},
    familyPageProductsErrorByKey: {},
    activeDecisionChoiceIndex: -1,
    decisionTabsWindowStart: 0,
    decisionAssistError: "",
    summaryPanelOpen: false,
    pdfZoom: 1,
    pageStripScrollLeft: 0,
    pageStripCanScrollLeft: false,
    pageStripCanScrollRight: false,
    decisionTabsScrollLeft: 0,
    decisionTabsCanScrollLeft: false,
    decisionTabsCanScrollRight: false,
    uploadFiles: [],
    inputDraft: {
      productName: initialSpec.specDisplayName,
      brandName: initialSpec.brandDisplayName,
      category: initialSpec.category,
      attributes: initialSpec.attributes.map((attribute) => attribute.label).join("\n")
    }
  }

  let pdfjsLibPromise = null
  let persistenceDbPromise = null
  let selectionChangeTimeoutId = null
  let helpSpecLoadingIntervalId = null
  let floatingCopyButton = null
  const app = document.getElementById("app")

  function cloneSpec(spec) {
    return {
      ...spec,
      attributes: spec.attributes.map((attribute) => ({ ...attribute }))
    }
  }

  function getHelpSpecLoadingLine() {
    const lineOrder = Array.isArray(appState.helpSpecLoadingLineOrder) ? appState.helpSpecLoadingLineOrder : []
    const activeIndex = lineOrder[appState.helpSpecLoadingLineIndex]
    return HELP_SPEC_LOADING_LINES[activeIndex] || HELP_SPEC_LOADING_LINES[0]
  }

  function buildHelpSpecLoadingLineOrder() {
    const order = HELP_SPEC_LOADING_LINES.map((_, index) => index)
    for (let index = order.length - 1; index > 0; index -= 1) {
      const swapIndex = Math.floor(Math.random() * (index + 1))
      const current = order[index]
      order[index] = order[swapIndex]
      order[swapIndex] = current
    }
    return order
  }

  function startHelpSpecLoadingTicker() {
    if (helpSpecLoadingIntervalId) return
    if (!appState.helpSpecLoadingLineOrder.length) {
      appState.helpSpecLoadingLineOrder = buildHelpSpecLoadingLineOrder()
    }
    helpSpecLoadingIntervalId = window.setInterval(() => {
      if (!appState.decisionAssistLoading) {
        stopHelpSpecLoadingTicker()
        return
      }
      const nextIndex = appState.helpSpecLoadingLineIndex + 1
      if (nextIndex >= HELP_SPEC_LOADING_LINES.length) {
        appState.helpSpecLoadingLineOrder = buildHelpSpecLoadingLineOrder()
        appState.helpSpecLoadingLineIndex = 0
      } else {
        appState.helpSpecLoadingLineIndex = nextIndex
      }
      renderPreservingViewerScroll()
    }, HELP_SPEC_LOADING_INTERVAL_MS)
  }

  function stopHelpSpecLoadingTicker() {
    if (!helpSpecLoadingIntervalId) return
    window.clearInterval(helpSpecLoadingIntervalId)
    helpSpecLoadingIntervalId = null
  }

  function hydrateSyntheticDocument(documentRecord) {
    return {
      ...documentRecord,
      sourceType: "synthetic",
      pages: extractPagesFromPdfText(documentRecord.pdfText)
    }
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;")
  }

  function normalizeText(value) {
    return String(value || "").replace(/\s+/g, " ").trim()
  }

  function slugify(value, fallback) {
    const slug = normalizeText(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "")
    return slug || fallback
  }

  function getAttributeType(label) {
    return /notes|description|remarks|details/i.test(label) ? "textarea" : "text"
  }

  function buildConfigurationOptionsFromSections(characteristic) {
    const sections = Array.isArray(characteristic?.configuration_sections)
      ? characteristic.configuration_sections
      : Array.isArray(characteristic?.configurationSections)
        ? characteristic.configurationSections
        : []
    if (!sections.length) return []

    const matrix = characteristic?.pricing_matrix || characteristic?.pricingMatrix || null
    const matrixRows = Array.isArray(matrix?.rows) ? matrix.rows : []
    const rowMap = new Map(
      matrixRows
        .map((row) => ({
          name: normalizeText(row?.row_name || row?.rowName || row?.name || ""),
          cells: Array.isArray(row?.cells) ? row.cells : []
        }))
        .filter((row) => row.name)
        .map((row) => [row.name.toLowerCase(), row.cells])
    )

    return sections
      .map((section) => {
        const name = normalizeText(section?.title || section?.name || section?.label || "")
        const optionList = Array.isArray(section?.options) ? section.options : []
        const values = optionList.map((item) => normalizeText(item)).filter(Boolean).join("; ")
        const matchingCells = rowMap.get(name.toLowerCase()) || []
        const pricing = matchingCells
          .map((cell) => {
            const label = normalizeText(cell?.column || cell?.label || "")
            const price = normalizeText(cell?.price || "")
            const model = normalizeText(cell?.model || "")
            if (!label || !price) return ""
            return `${label}: ${price}${model ? ` [${model}]` : ""}`
          })
          .filter(Boolean)
          .join("; ")
        return {
          name,
          values,
          pricing,
          difference: "",
          evidence: ""
        }
      })
      .filter((item) => item.name)
  }

  function getConfigurationSections(characteristic) {
    const sections = Array.isArray(characteristic?.configuration_sections)
      ? characteristic.configuration_sections
      : Array.isArray(characteristic?.configurationSections)
        ? characteristic.configurationSections
        : []
    if (sections.length) {
      return sections
        .map((section) => ({
          title: normalizeText(section?.title || section?.name || section?.label || ""),
          options: Array.isArray(section?.options)
            ? section.options.map((item) => formatConfigurationDisplayLabel(item)).filter(Boolean)
            : []
        }))
        .filter((section) => section.title)
    }

    return getCharacteristicOptions(characteristic).map((option) => ({
      title: formatConfigurationDisplayLabel(option.name),
      options: parseConfigurationSectionItems(option.values).map((item) => formatConfigurationDisplayLabel(item))
    }))
  }

  function getConfigurationMatrix(characteristic) {
    const matrix = characteristic?.pricing_matrix || characteristic?.pricingMatrix || null
    if (!matrix) return null
    const rowLabel = formatConfigurationDisplayLabel(matrix?.row_label || matrix?.rowLabel || "")
    const columns = Array.isArray(matrix?.column_labels)
      ? matrix.column_labels.map((item) => formatConfigurationDisplayLabel(item)).filter(Boolean)
      : Array.isArray(matrix?.columnLabels)
        ? matrix.columnLabels.map((item) => formatConfigurationDisplayLabel(item)).filter(Boolean)
        : []
    const rows = Array.isArray(matrix?.rows)
      ? matrix.rows.map((row) => ({
        name: formatConfigurationDisplayLabel(row?.row_name || row?.rowName || row?.name || ""),
        cells: Array.isArray(row?.cells)
          ? row.cells.map((cell) => ({
              column: formatConfigurationDisplayLabel(cell?.column || cell?.label || ""),
              price: normalizeText(cell?.price || ""),
              model: normalizeText(cell?.model || ""),
              matchKey: getConfigurationMatchKey(cell?.column || cell?.label || "")
            }))
          : []
      })).filter((row) => row.name)
      : []
    if (!columns.length || !rows.length) return null
    return { rowLabel, columns, rows }
  }

  function isFinishOrShellDescriptor(value) {
    const normalized = normalizeText(value).toLowerCase()
    if (!normalized) return false
    return /finish|veneer|chrome|walnut|ash|oak|ebony|palisander|stain|black base|chrome base|shell|nonupholstered|upholstered/.test(normalized)
  }

  function shouldHideConfigurationCharacteristic(characteristic, selectedVariantId, characteristics) {
    if (!/configuration/i.test(normalizeText(characteristic?.label || ""))) return false

    const hasShellFinish = hasCharacteristicLabel(characteristics, /shell finish|wood finish/i)
    const hasBaseFrameFinish = hasCharacteristicLabel(characteristics, /base finish|frame finish|base \/ frame finish/i)
    if (!hasShellFinish && !hasBaseFrameFinish) return false

    const sections = getConfigurationSections(characteristic)
    const options = getCharacteristicOptions(characteristic)
    const matrix = getConfigurationMatrix(characteristic)

    const hasMeaningfulSections = sections.some((section) => {
      if (!section.options.length) return false
      if (isFinishOrShellDescriptor(section.title)) return false
      return section.options.some((option) => !isFinishOrShellDescriptor(option))
    })

    const hasMeaningfulOptions = options.some((option) => !isFinishOrShellDescriptor(option.name) && !isGenericConfigurationOptionName(option.name))

    const hasMeaningfulMatrix = Boolean(
      matrix
      && !isFinishOrShellDescriptor(matrix.rowLabel)
      && matrix.rows.some((row) => !isFinishOrShellDescriptor(row.name))
      && matrix.columns.some((column) => !isFinishOrShellDescriptor(column))
      && matrix.rows.some((row) => row.cells.some((cell) => normalizeText(cell.price)))
    )

    const onlyGenericShellStub = !hasMeaningfulSections && !hasMeaningfulOptions && !hasMeaningfulMatrix
    if (selectedVariantId && onlyGenericShellStub) return true
    return onlyGenericShellStub
  }

  function pruneResolvedConfigurationCharacteristics(characteristics, selectedVariantId) {
    const items = Array.isArray(characteristics) ? characteristics : []
    return items.filter((characteristic) => !shouldHideConfigurationCharacteristic(characteristic, selectedVariantId, items))
  }

  function getCharacteristicOptions(characteristic) {
    if (!Array.isArray(characteristic?.options)) {
      return buildConfigurationOptionsFromSections(characteristic)
    }
    return characteristic.options
      .map((item) => {
        if (typeof item === "string") {
          return { name: normalizeText(item), values: "", pricing: "", difference: "", evidence: "" }
        }
        return {
          name: normalizeText(item?.name || item?.label || ""),
          values: normalizeText(item?.values || item?.spec_values || ""),
          pricing: normalizeText(item?.pricing || item?.price || ""),
          difference: normalizeText(item?.difference || ""),
          evidence: normalizeText(item?.evidence || "")
        }
      })
      .filter((item) => item.name)
  }

  function isCompactCharacteristicLabel(label) {
    return /shell finish|frame finish|wood finish|base finish|finish|upholstery|material/i.test(normalizeText(label))
  }

  function parseOptionValueLines(rawValueText) {
    return normalizeText(rawValueText)
      .split(/\s*;\s*|\s*,\s*(?=[A-Z][^,]{0,60}:)/)
      .map((segment) => normalizeText(segment))
      .filter(Boolean)
      .map((segment) => {
        const colonIndex = segment.indexOf(":")
        if (colonIndex === -1) {
          const trailingValueMatch = segment.match(/^(.*?)(\b\d[\d\s./"]*(?:in|ft|sq ft|mm|cm|m)?\b.*)$/i)
          if (trailingValueMatch) {
            return {
              label: normalizeText(trailingValueMatch[1]),
              value: normalizeText(trailingValueMatch[2])
            }
          }
          return { label: "", value: segment }
        }
        return {
          label: normalizeText(segment.slice(0, colonIndex)),
          value: normalizeText(segment.slice(colonIndex + 1))
        }
      })
  }

  function parseConfigurationSectionItems(rawValueText) {
    return normalizeText(rawValueText)
      .split(/\s*;\s*|\n+/)
      .map((segment) => normalizeText(segment))
      .filter(Boolean)
      .map((segment) => {
        const colonIndex = segment.indexOf(":")
        return colonIndex === -1 ? normalizeText(segment) : normalizeText(segment.slice(colonIndex + 1))
      })
      .filter(Boolean)
  }

  function buildFallbackConciseComparison(characteristic) {
    const options = getCharacteristicOptions(characteristic)
    if (options.length < 2) return ""
    const optionNames = options.map((item) => item.name).join(" vs ")
    const firstDifference = options.find((item) => item.difference)?.difference || normalizeText(characteristic?.dependencyNote || characteristic?.blurb || "")
    return firstDifference ? `${optionNames}: ${firstDifference}` : ""
  }

  function buildOptionDifferenceSummary(characteristic) {
    const options = getCharacteristicOptions(characteristic)
      .filter((item) => normalizeText(item.difference))
      .map((item) => ({
        name: item.name,
        difference: normalizeText(item.difference).replace(/\.$/, "")
      }))
    if (!options.length) return ""
    if (options.length === 1) return options[0].difference
    if (options.length === 2) {
      return `${options[0].name} is ${options[0].difference.toLowerCase()}, while ${options[1].name} is ${options[1].difference.toLowerCase()}.`
    }
    return options.map((item) => `${item.name}: ${item.difference}`).join(" • ")
  }

  function formatPageRangeLabel(pageNumbers) {
    if (!pageNumbers?.length) return ""
    const sorted = [...new Set(pageNumbers)].sort((a, b) => a - b)
    if (sorted.length === 1) return String(sorted[0])
    const segments = []
    let rangeStart = sorted[0]
    let previous = sorted[0]

    for (let index = 1; index <= sorted.length; index += 1) {
      const current = sorted[index]
      const isBreak = current !== previous + 1
      if (isBreak) {
        segments.push(rangeStart === previous ? String(rangeStart) : `${rangeStart}-${previous}`)
        rangeStart = current
      }
      previous = current
    }

    return segments.join(", ")
  }

  function detectFamilyPdfArchetype(documentRecord = null) {
    const documentToCheck = documentRecord || getActiveDocument()
    if (!documentToCheck?.pages?.length) return false

    const samplePages = documentToCheck.pages.slice(0, 16)
    let familySignals = 0

    samplePages.forEach((page) => {
      const text = normalizeText(getPageCombinedText(page))
      if (!text) return
      const lowered = text.toLowerCase()
      const modelMatches = text.match(/\b[A-Z]{1,4}-\d{1,3}[A-Z0-9-]*\b/g) || []
      if (modelMatches.length >= 2) familySignals += 2
      if (/item\s+description/i.test(text) && /list price/i.test(text)) familySignals += 2
      if (/series\b/i.test(text) && /low back|high back|wire base|wood base|pedestal/i.test(lowered)) familySignals += 1
      if (/required to specify|metal finish|veneer finish|shell finish/i.test(lowered)) familySignals += 1
    })

    return familySignals >= 4
  }

  function getSpecParsingMode() {
    const brand = normalizeText(appState.spec.brandDisplayName || appState.spec.originalBrand || "").toLowerCase()
    if (/\bdavis\b/.test(brand)) return "family"
    if (detectFamilyPdfArchetype()) return "family"
    return "contiguous"
  }

  function getSpecParsingModeConfig(modeOverride = "") {
    const mode = normalizeText(modeOverride) || getSpecParsingMode()
    if (mode === "family") {
      return {
        mode,
        label: "family",
        scopeNoun: "family",
        pageLabel: "family-relevant",
        selectorProductTitle: "Select the product in this family",
        selectorProductCopy: "This Davis page appears to contain more than one distinct product within the same family. Choose the product first, then the system will gather the shared spec and reference pages tied to that family context.",
        selectorVariantTitle: "Select the variant to review",
        selectorVariantCopy: "This Davis page appears to contain multiple related variants in the same family. Pick the variant first, then the system will gather the family-level spec and reference pages that apply to it."
      }
    }

    return {
      mode,
      label: "contiguous product",
      scopeNoun: "product",
      pageLabel: "contiguous",
      selectorProductTitle: "Select the product on this page",
      selectorProductCopy: "This starting page appears to contain more than one product. Choose the one you want, then the system will scan forward until the next page is no longer the same product.",
      selectorVariantTitle: "Select the variant to review",
      selectorVariantCopy: "This starting page appears to contain more than one variant for the same product. Pick the one you want before the spec characteristics are shown."
    }
  }

  function getFamilyResultPageLimit() {
    return 8
  }

  function normalizeStructureType(value) {
    const normalized = normalizeText(value).toLowerCase().replace(/\s+/g, "_")
    if (normalized === "product_family") return "product_family"
    if (normalized === "single_product") return "single_product"
    return ""
  }

  function normalizeInteractionModel(value) {
    const normalized = normalizeText(value).toLowerCase().replace(/\s+/g, "_")
    if (normalized === "product_first") return "product_first"
    if (normalized === "page_first") return "page_first"
    return ""
  }

  function parseModelBoolean(value) {
    if (typeof value === "boolean") return value
    const normalized = normalizeText(value).toLowerCase()
    if (normalized === "true" || normalized === "yes") return true
    if (normalized === "false" || normalized === "no") return false
    return false
  }

  function getStructureTypeFromArchetype(mode) {
    return mode === "family" ? "product_family" : "single_product"
  }

  function isAbstractProductBucket(label) {
    const normalized = normalizeText(label).toLowerCase()
    if (!normalized) return true
    return /\b(base type|upholstery type|finish type|variant|configuration|material|wood base|wire base|pedestal|shell finish|frame finish|base finish)\b/.test(normalized)
  }

  function getConcreteProductCandidateKey(value) {
    return normalizeText(value).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim()
  }

  function isConcreteProductLabel(label) {
    const normalized = normalizeText(label)
    const lowered = normalized.toLowerCase()
    if (!normalized || isAbstractProductBucket(normalized)) return false
    if (/^(general|base|standard|variant\s+\d+|option\s+\d+|related|dimensions?(?:\s*&\s*sizes?)?|finish options?|material(?: specifications?)?)$/i.test(normalized)) {
      return false
    }
    if (/(^| )(finish|oak|ash|walnut|ebony|palisander|upholstery|material|dimensions?|height|width|depth)( |$)/i.test(normalized) && !/(chair|stool|table|bench|sofa|ottoman|lounge|settee|desk|credenza|cabinet)/i.test(normalized)) {
      return false
    }
    const tokenCount = lowered.split(/\s+/).filter(Boolean).length
    return tokenCount >= 2 || /\b[A-Z]{1,4}-\d{1,3}[A-Z0-9-]*\b/.test(normalized)
  }

  function normalizeConcreteProductCandidates(items) {
    const unique = new Map()
    normalizeDecisionCandidates(items).forEach((candidate) => {
      const mapKey = getConcreteProductCandidateKey(candidate.label || candidate.id)
      if (!mapKey || !isConcreteProductLabel(candidate.label)) return
      if (!unique.has(mapKey)) {
        unique.set(mapKey, candidate)
      }
    })
    return [...unique.values()]
  }

  function getTopRerankedPageNumbers(aiResult) {
    if (!aiResult) return []
    const kept = Array.isArray(aiResult.keptPages) ? aiResult.keptPages : []
    if (kept.length) return kept.slice(0, 3)
    const ordered = Array.isArray(aiResult.orderedPages) ? aiResult.orderedPages.map((item) => item.pageNumber).filter(Number.isFinite) : []
    return ordered.slice(0, 3)
  }

  function buildGroundedProductCandidatesFromTopPages(documentRecord, aiResult) {
    if (!documentRecord || !aiResult) return []

    const pageNumbers = getTopRerankedPageNumbers(aiResult)
    if (!pageNumbers.length) return []
    const titles = pageNumbers.map((pageNumber) => getPrimaryHeaderTitle(documentRecord, pageNumber)).filter(Boolean)
    const sharedPrefix = getSharedTitlePrefix(titles)
    const prefixPattern = sharedPrefix
      ? new RegExp(`^${sharedPrefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "i")
      : null

    const candidates = pageNumbers
      .map((pageNumber) => {
        const title = normalizeText(getPrimaryHeaderTitle(documentRecord, pageNumber))
        if (!title) return null
        const stripped = prefixPattern
          ? normalizeText(title.replace(prefixPattern, "")).replace(/^[,:\-–—]\s*/, "")
          : ""
        const label = isConcreteProductLabel(stripped) ? stripped : title
        if (!isConcreteProductLabel(label)) return null
        return {
          id: slugify(label, `product-${pageNumber}`),
          label,
          description: title !== label ? title : "",
          evidence: `Visible top family page title on page ${pageNumber}`
        }
      })
      .filter(Boolean)

    return normalizeConcreteProductCandidates(candidates)
  }

  function getAggregatedTopFamilyProductCandidates(documentRecord, aiResult) {
    const combined = [
      ...normalizeConcreteProductCandidates(aiResult?.concreteProductCandidates),
      ...buildGroundedProductCandidatesFromTopPages(documentRecord, aiResult)
    ]
    return normalizeConcreteProductCandidates(combined)
  }

  function isProductFirstSessionActive() {
    return Boolean(appState.productFirstSelection?.active && normalizeText(appState.productFirstSelection?.productId))
  }

  function suppressProductFirstSiblingVariants(variantCandidates, selectedProductId, familyProductCandidates) {
    const candidates = Array.isArray(variantCandidates) ? variantCandidates : []
    if (!candidates.length) return []
    const selectedKey = getConcreteProductCandidateKey(selectedProductId)
    const familyKeys = new Set(
      (familyProductCandidates || [])
        .flatMap((candidate) => [candidate.id, candidate.label])
        .map((value) => getConcreteProductCandidateKey(value))
        .filter(Boolean)
    )

    const filtered = candidates.filter((candidate) => {
      const labelKey = getConcreteProductCandidateKey(candidate.label)
      const idKey = getConcreteProductCandidateKey(candidate.id)
      if (selectedKey && (labelKey === selectedKey || idKey === selectedKey)) return false
      if (familyKeys.has(labelKey) || familyKeys.has(idKey)) return false
      if (/^model\s+[A-Z]{1,4}-\d{1,3}[A-Z0-9-]*$/i.test(normalizeText(candidate.label))) return false
      return true
    })

    return filtered
  }

  function getProductCardModelCode(candidate) {
    const description = normalizeText(candidate?.description || "")
    const evidence = normalizeText(candidate?.evidence || "")
    const sourceText = [description, evidence].filter(Boolean).join(" ")
    const modelMatch = sourceText.match(/\b[A-Z]{1,4}-\d{1,3}[A-Z0-9-]*\b/)
    return modelMatch ? modelMatch[0] : ""
  }

  function shouldShowSeparateProductCardModelCode(candidate) {
    const modelCode = getProductCardModelCode(candidate)
    const label = normalizeText(candidate?.label || "").toLowerCase()
    if (!modelCode) return false
    return !label.includes(modelCode.toLowerCase())
  }

  function getProductCardTitle(candidate) {
    const label = normalizeText(candidate?.label || "")
    const subtitle = getProductCardSubtitle(candidate)
    if (!subtitle) return label

    const normalizedLabel = label.toLowerCase()
    const normalizedSubtitle = subtitle.toLowerCase()
    if (normalizedSubtitle.includes(normalizedLabel) && normalizedSubtitle.length > normalizedLabel.length) {
      return subtitle
    }

    return label
  }

  function getProductCardSubtitle(candidate) {
    const description = normalizeText(candidate?.description || "")
    const label = normalizeText(candidate?.label || "")
    const modelCode = getProductCardModelCode(candidate)
    const cleanedDescription = normalizeText(
      description
        .replace(/\bpage\s+\d+\b/gi, "")
        .replace(/\bsection\b/gi, "")
        .replace(/\s{2,}/g, " ")
    )
    if (!cleanedDescription) return ""

    let simplified = cleanedDescription
    if (label) {
      simplified = normalizeText(
        simplified.replace(new RegExp(label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i"), "")
      )
    }
    if (modelCode) {
      simplified = normalizeText(
        simplified.replace(new RegExp(modelCode.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "ig"), "")
      )
    }
    simplified = normalizeText(simplified.replace(/^[,:\-–—]\s*/, ""))

    if (!simplified || simplified.toLowerCase() === label.toLowerCase()) return ""
    return simplified
  }

  function getDisplayProductCardSubtitle(candidate) {
    const subtitle = getProductCardSubtitle(candidate)
    const title = getProductCardTitle(candidate)
    if (!subtitle) return ""
    if (subtitle.toLowerCase() === title.toLowerCase()) return ""
    const normalizedTitle = title.toLowerCase()
    const normalizedSubtitle = subtitle.toLowerCase()
    if (
      normalizedSubtitle.length <= 20
      && (
        normalizedTitle.includes(normalizedSubtitle)
        || /^(base|swivel|tilt|mechanism|fixed|memory return|four prong base)$/i.test(subtitle)
      )
    ) {
      return ""
    }
    return subtitle
  }

  // Routing is intentionally conservative: product-first only activates when the
  // reranked top pages expose grounded product labels we can safely present.
  function buildStructureRoutingState(documentRecord = null) {
    const activeDocument = documentRecord || getActiveDocument()
    const archetype = getSpecParsingMode()
    const fallbackPageNumber =
      appState.aiRerankResult?.bestPage
      || appState.rankedPages[0]?.pageNumber
      || activeDocument?.pages?.[0]?.pageNumber
      || 1
    const fallback = {
      structureType: getStructureTypeFromArchetype(archetype),
      interactionModel: "page_first",
      suggestedInteractionModel: "",
      confidence: 0,
      hasConcreteProducts: false,
      productCandidates: [],
      productFirstPageNumber: fallbackPageNumber,
      source: "fallback"
    }

    if (!activeDocument || appState.aiRerankDocumentId !== activeDocument.id || !appState.aiRerankResult) {
      return fallback
    }

    const aiResult = appState.aiRerankResult
    const structureType = normalizeStructureType(aiResult.structureType) || getStructureTypeFromArchetype(archetype)
    const suggestedInteractionModel = normalizeInteractionModel(aiResult.interactionModel)
    const confidence = Number.isFinite(aiResult.structureConfidence) ? aiResult.structureConfidence : 0
    const productCandidates = getAggregatedTopFamilyProductCandidates(activeDocument, aiResult)
    const topPageNumbers = getTopRerankedPageNumbers(aiResult)
    const hasConcreteProducts = Boolean(aiResult.hasConcreteProducts) && productCandidates.length > 1 && topPageNumbers.length > 0
    const familySignalSupportsProductFirst =
      structureType === "product_family"
      || archetype === "family"

    // Concrete visible product candidates outrank model confidence. If we can
    // safely ground product choices on the top reranked family pages, keep the
    // user in product-first mode and use pages only as downstream context.
    const eligibleForProductFirst =
      hasConcreteProducts
      && familySignalSupportsProductFirst
      && (
        suggestedInteractionModel === "product_first"
        || confidence >= PRODUCT_FIRST_CONFIDENCE_THRESHOLD
        || archetype === "family"
      )

    return {
      structureType,
      suggestedInteractionModel,
      interactionModel: eligibleForProductFirst ? "product_first" : "page_first",
      confidence,
      hasConcreteProducts,
      productCandidates,
      productFirstPageNumber: aiResult.bestPage || topPageNumbers[0] || fallbackPageNumber,
      source: eligibleForProductFirst ? "grounded_products" : "fallback"
    }
  }

  function updateStructureRoutingState(documentRecord = null) {
    appState.structureRouting = buildStructureRoutingState(documentRecord)
  }

  function getAiRerankCandidateLimit() {
    return getSpecParsingMode() === "family" ? 12 : 5
  }

  function getAiKeptPageLimit() {
    return getSpecParsingMode() === "family" ? 8 : 3
  }

  function buildFamilySearchTokens(documentRecord, selectedPage) {
    const productName = normalizeText(appState.spec.specDisplayName || appState.spec.originalSpecName || "")
    const selectedTitle = normalizeText(getPrimaryHeaderTitle(documentRecord, selectedPage.pageNumber))
    const selectedText = normalizeText(getPageCombinedText(selectedPage)).slice(0, 500)
    const variantStopwords = new Set([
      "wire",
      "wood",
      "base",
      "pedestal",
      "prong",
      "low",
      "high",
      "back",
      "footrest",
      "exposed",
      "veneer",
      "shell",
      "fully",
      "upholstered",
      "contrasting",
      "material"
    ])
    const genericStopwords = new Set([
      "the",
      "and",
      "for",
      "with",
      "from",
      "this",
      "that",
      "price",
      "list",
      "pages",
      "page",
      "spec",
      "specs",
      "specification",
      "specifications",
      "reference",
      "references",
      "finish",
      "finishes",
      "powder",
      "coat",
      "davis"
    ])

    return [...new Set(
      `${productName} ${selectedTitle} ${selectedText}`
        .toLowerCase()
        .split(/[^a-z0-9]+/)
        .filter((token) => token.length >= 3)
        .filter((token) => !genericStopwords.has(token))
        .filter((token) => !variantStopwords.has(token))
    )].slice(0, 8)
  }

  function getFamilyDecisionCandidatePages(documentRecord, selectedPage, selectionHints = {}) {
    const selectedProductHint = normalizeText(selectionHints.selectedProductId || "")
    const selectedVariantHint = normalizeText(selectionHints.selectedVariantId || "")
    const allowFamilyExpansion = selectionHints.allowFamilyExpansion === true
    if (!selectedProductHint && !selectedVariantHint) {
      return selectedPage ? [selectedPage] : []
    }
    if (!allowFamilyExpansion) {
      return selectedPage ? [selectedPage] : []
    }

    const familyTokens = buildFamilySearchTokens(documentRecord, selectedPage)
    const selectedTitle = normalizeText(getPrimaryHeaderTitle(documentRecord, selectedPage.pageNumber)).toLowerCase()
    const exactProduct = normalizeText(appState.spec.specDisplayName || appState.spec.originalSpecName || "").toLowerCase()
    const rankedScoreByPage = new Map(appState.rankedPages.map((page) => [page.pageNumber, Number(page.score) || 0]))
    const referencePattern = /finish|finishes|powder\s*coat|specification|specifications|dimension|dimensions|material|upholstery|veneer|glide|pricing|price list|reference/i
    const variantPattern = /wire base|wood base|four prong pedestal|low back|high back|footrest|exposed veneer shell|fully upholstered|contrasting material/i

    const scoredPages = documentRecord.pages
      .map((page) => {
        const pageTitle = normalizeText(getPrimaryHeaderTitle(documentRecord, page.pageNumber))
        const pageText = normalizeText(getPageCombinedText(page))
        const haystack = `${pageTitle}\n${pageText}`.toLowerCase()
        const tokenMatches = familyTokens.filter((token) => haystack.includes(token))
        const titleTokenMatches = familyTokens.filter((token) => pageTitle.toLowerCase().includes(token))
        const hasReferenceSignals = referencePattern.test(pageTitle) || referencePattern.test(pageText)
        const hasVariantSignals = variantPattern.test(haystack)
        let score = page.pageNumber === selectedPage.pageNumber ? 200 : 0

        if (exactProduct && haystack.includes(exactProduct)) score += 40
        score += tokenMatches.length * 8
        score += titleTokenMatches.length * 6
        score += Math.min(24, Math.round((rankedScoreByPage.get(page.pageNumber) || 0) / 2))
        if (selectedTitle && pageTitle && pageTitle.toLowerCase().includes(selectedTitle)) score += 18
        if (tokenMatches.length && hasReferenceSignals) score += 14
        if (hasVariantSignals) score += 8

        return {
          page,
          score
        }
      })
      .filter((entry) => entry.score >= 14 || entry.page.pageNumber === selectedPage.pageNumber)
      .sort((a, b) => b.score - a.score || a.page.pageNumber - b.page.pageNumber)

    const topPages = scoredPages.slice(0, 10).map((entry) => entry.page)
    if (topPages.length > 1) {
      return [...topPages].sort((a, b) => a.pageNumber - b.pageNumber)
    }

    return documentRecord.pages
      .filter((page) => page.pageNumber >= selectedPage.pageNumber)
      .slice(0, 4)
  }

  function getDecisionCandidatePages(documentRecord, selectedPage, selectionHints = {}) {
    if (getSpecParsingMode() === "family") {
      return getFamilyDecisionCandidatePages(documentRecord, selectedPage, selectionHints)
    }

    return documentRecord.pages
      .filter((page) => page.pageNumber >= selectedPage.pageNumber)
      .slice(0, 6)
  }

  function normalizeDecisionCandidates(items) {
    if (!Array.isArray(items)) return []
    return items
      .map((item, index) => ({
        id: normalizeText(item?.id || item?.key || item?.name || item?.label || `candidate-${index + 1}`),
        label: normalizeText(item?.label || item?.name || `Option ${index + 1}`),
        description: normalizeText(item?.description || item?.summary || ""),
        evidence: normalizeText(item?.evidence || item?.reason || "")
      }))
      .filter((item) => item.id && item.label)
  }

  function extractReferencedPageNumbers(value) {
    const matches = String(value || "").match(/\bpage\s+(\d+)\b/gi) || []
    return matches
      .map((match) => Number(match.replace(/[^\d]/g, "")))
      .filter((pageNumber) => Number.isFinite(pageNumber))
  }

  function filterCandidatesToStartingPage(candidates, startPageNumber) {
    if (!Array.isArray(candidates) || candidates.length <= 1) return candidates || []
    const withPageMentions = candidates.filter((candidate) => extractReferencedPageNumbers(candidate.evidence).length)
    if (!withPageMentions.length) return candidates
    const startPageCandidates = candidates.filter((candidate) => {
      const referencedPages = extractReferencedPageNumbers(candidate.evidence)
      return !referencedPages.length || referencedPages.includes(startPageNumber)
    })
    return startPageCandidates.length ? startPageCandidates : candidates
  }

  function buildVariantEvidenceTokens(label) {
    const genericTokens = new Set([
      "ginkgo",
      "ply",
      "lounge",
      "chair",
      "chairs",
      "seating",
      "base",
      "model",
      "series",
      "davis"
    ])

    return tokenize(label)
      .filter((token) => token.length >= 4 || token === "wood" || token === "wire")
      .filter((token) => !genericTokens.has(token))
  }

  function extractConcreteVariantCandidatesFromPageText(pageText, subtypeHint = "") {
    const rawText = String(pageText || "")
    const modelPattern = /\b([A-Z]{1,4}-\d{1,3}[A-Z0-9-]*)\b/g
    const itemMatches = [...rawText.matchAll(modelPattern)]
    if (itemMatches.length < 2) return []

    const phraseMatchers = [
      { pattern: /exposed veneer shell/gi, label: "Exposed Veneer Shell" },
      { pattern: /interior upholstery/gi, label: "Interior Upholstery" },
      { pattern: /fully upholstered/gi, label: "Fully Upholstered" },
      { pattern: /all same material/gi, label: "All Same Material" },
      { pattern: /contrasting material/gi, label: "Contrasting Material" },
      { pattern: /wire base/gi, label: "Wire Base" },
      { pattern: /wood base/gi, label: "Wood Base" },
      { pattern: /four prong pedestal/gi, label: "Four Prong Pedestal" },
      { pattern: /memory return swivel/gi, label: "Memory Return Swivel" },
      { pattern: /swivel/gi, label: "Swivel" },
      { pattern: /fixed/gi, label: "Fixed" },
      { pattern: /low back/gi, label: "Low Back" },
      { pattern: /high back/gi, label: "High Back" }
    ]

    const items = itemMatches.map((match, index) => {
      const model = normalizeText(match[1] || "")
      const start = match.index || 0
      const end = itemMatches[index + 1]?.index || rawText.length
      const block = rawText.slice(start, end)
      const matchedLabels = phraseMatchers
        .filter((entry) => entry.pattern.test(block))
        .map((entry) => entry.label)

      return {
        model,
        block,
        matchedLabels: [...new Set(matchedLabels)]
      }
    })

    const labelFrequency = new Map()
    items.forEach((item) => {
      item.matchedLabels.forEach((label) => {
        labelFrequency.set(label, (labelFrequency.get(label) || 0) + 1)
      })
    })

    return items.map((item, index) => {
      const distinctiveLabels = item.matchedLabels.filter((label) => (labelFrequency.get(label) || 0) < items.length)
      const preferredLabels = distinctiveLabels.filter((label) => !/low back|high back|wire base|wood base|four prong pedestal|fixed|swivel/i.test(label))
      const labelParts = (preferredLabels.length ? preferredLabels : distinctiveLabels.length ? distinctiveLabels : item.matchedLabels).slice(0, 3)
      const label = labelParts.join(" / ")
      return {
        id: slugify(`${item.model}-${label || `variant-${index + 1}`}`, `variant-${index + 1}`),
        label: label || `Model ${item.model}`,
        description: item.model,
        evidence: `Page-local model ${item.model}`,
        matchedLabels: item.matchedLabels
      }
    })
      .filter((candidate) => normalizeText(candidate.label))
      .filter((candidate) => {
        const hint = normalizeText(subtypeHint).toLowerCase()
        if (!hint) return true
        const candidateText = [candidate.label, ...(candidate.matchedLabels || [])].map((value) => normalizeText(value).toLowerCase()).join(" ")
        if (/wire base/.test(hint)) return /wire base/.test(candidateText)
        if (/wood base/.test(hint)) return /wood base/.test(candidateText)
        if (/memory return/.test(hint)) return /memory return/.test(candidateText)
        if (/pedestal/.test(hint) && /fixed/.test(hint)) return /four prong pedestal/.test(candidateText) && /\bfixed\b/.test(candidateText)
        if (/pedestal/.test(hint) && /swivel/.test(hint)) return /four prong pedestal/.test(candidateText) && /\bswivel\b/.test(candidateText)
        if (/pedestal/.test(hint)) return /four prong pedestal/.test(candidateText)
        return true
      })
      .map(({ matchedLabels, ...candidate }) => candidate)
  }

  function preferConcretePageVariants(candidates, selectedPage, selectedVariantHint = "", documentRecord = null) {
    if (getSpecParsingMode() !== "family" || selectedVariantHint) return candidates || []
    const subtypeHint = getSelectedPageSubtypeHint(documentRecord, selectedPage)
    const concreteCandidates = extractConcreteVariantCandidatesFromPageText(getPageCombinedText(selectedPage), subtypeHint)
    if (concreteCandidates.length >= 2) return concreteCandidates
    if ((String(getPageCombinedText(selectedPage)).match(/\b[A-Z]{1,4}-\d{1,3}[A-Z0-9-]*\b/g) || []).length >= 2) {
      return []
    }

    const genericVariantLabels = new Set(["base type", "upholstery type", "material type", "finish type"])
    const filtered = (candidates || []).filter((candidate) => !genericVariantLabels.has(normalizeText(candidate.label).toLowerCase()))
    return filtered.length ? filtered : candidates || []
  }

  function filterVariantCandidatesToSelectedPage(candidates, selectedPage, documentRecord) {
    if (getSpecParsingMode() !== "family" || !selectedPage) return candidates || []
    if (!Array.isArray(candidates) || candidates.length <= 1) return candidates || []

    const pageTitle = normalizeText(getPrimaryHeaderTitle(documentRecord, selectedPage.pageNumber)).toLowerCase()
    const pageText = normalizeText(getPageCombinedText(selectedPage)).toLowerCase()
    const haystack = `${pageTitle}\n${pageText}`

    const filtered = candidates.filter((candidate) => {
      const label = normalizeText(candidate.label).toLowerCase()
      const simplifiedLabel = label.replace(/[(),]/g, " ").replace(/\s+/g, " ").trim()
      if (simplifiedLabel && haystack.includes(simplifiedLabel)) return true

      const tokens = buildVariantEvidenceTokens(candidate.label)
      if (!tokens.length) return false
      return tokens.every((token) => haystack.includes(token))
    })

    return filtered.length ? filtered : candidates
  }

  function filterCandidatesToResolvedProduct(candidates, productName) {
    if (!Array.isArray(candidates) || candidates.length <= 1) return candidates || []
    const normalizedProductName = normalizeText(productName).toLowerCase()
    if (!normalizedProductName) return candidates
    const matchingCandidates = candidates.filter((candidate) => {
      const normalizedLabel = normalizeText(candidate.label).toLowerCase()
      return normalizedLabel && (normalizedProductName.includes(normalizedLabel) || normalizedLabel.includes(normalizedProductName))
    })
    return matchingCandidates.length ? matchingCandidates : candidates
  }

  function shouldTreatVariantCandidatesAsConfiguration(variantCandidates, sectionPages) {
    const candidates = Array.isArray(variantCandidates) ? variantCandidates : []
    if (candidates.length < 2) return false
    const looksLikeModelSplit = candidates.every((candidate) => /\bEA\d+\b/i.test(candidate.label) || /(no arms|arms with arm pads|arms)/i.test(candidate.label))
    if (!looksLikeModelSplit) return false
    const combinedText = (sectionPages || []).map((page) => getPageCombinedText(page)).join("\n")
    const configSignals = extractConfigurationComboPrices(combinedText)
    return Boolean(configSignals.prefix && configSignals.priceMap.size && configSignals.armCodeMap.size)
  }

  function getVariantMatchTokens(variantCandidates, selectedVariantId) {
    const selectedVariant = (variantCandidates || []).find((candidate) => normalizeText(candidate.id) === normalizeText(selectedVariantId))
    const source = selectedVariant || { id: selectedVariantId, label: selectedVariantId }
    return [source.id, source.label]
      .map((value) => normalizeText(value).toLowerCase())
      .filter(Boolean)
  }

  function getSelectedCandidateMatchTokens(candidates, selectedId) {
    const selectedCandidate = (candidates || []).find((candidate) => normalizeText(candidate.id) === normalizeText(selectedId))
    const source = selectedCandidate || { id: selectedId, label: selectedId }
    const modelCode = getProductCardModelCode(source)
    return [source.id, source.label, modelCode]
      .map((value) => normalizeText(value).toLowerCase())
      .filter(Boolean)
      .filter((value) => value.length >= 5 || /\b[a-z]{1,4}-\d{1,4}\b/i.test(value))
  }

  function extractModelCodes(value) {
    return [...new Set((normalizeText(value).match(/\b[A-Z]{1,4}-\d{1,4}[A-Z0-9-]*\b/g) || []).map((item) => normalizeText(item)))]
  }

  function getSelectedCandidate(candidates, selectedId) {
    return (candidates || []).find((candidate) => normalizeText(candidate.id) === normalizeText(selectedId)) || null
  }

  function optionContainsSelectedAndSiblingModels(optionName, selectedModelCodes, siblingModelCodes) {
    const optionModels = extractModelCodes(optionName)
    if (!optionModels.length) return false
    const hasSelected = optionModels.some((code) => selectedModelCodes.has(code))
    const hasSibling = optionModels.some((code) => siblingModelCodes.has(code))
    return hasSelected && hasSibling
  }

  function relabelCombinedOptionForSelectedProduct(option, selectedCandidate, selectedModelCodes, siblingModelCodes) {
    const optionName = normalizeText(option?.name || "")
    if (!optionContainsSelectedAndSiblingModels(optionName, selectedModelCodes, siblingModelCodes)) {
      return option
    }

    const selectedLabel = normalizeText(selectedCandidate?.label || "")
    const selectedModelCode = extractModelCodes(selectedCandidate?.description || selectedCandidate?.label || selectedCandidate?.id || "")
      .find((code) => selectedModelCodes.has(code)) || [...selectedModelCodes][0] || ""

    return {
      ...option,
      name: selectedLabel || selectedModelCode || optionName
    }
  }

  function optionContainsSiblingModels(option, siblingModelCodes) {
    const optionModels = extractModelCodes(option?.name || "")
    if (!optionModels.length) return false
    return optionModels.some((code) => siblingModelCodes.has(code))
  }

  function optionMatchesVariant(option, tokens) {
    if (!tokens.length) return false
    const haystack = [
      option?.name,
      option?.values,
      option?.difference,
      option?.evidence
    ]
      .map((value) => normalizeText(value).toLowerCase())
      .join(" ")
    return tokens.some((token) => haystack.includes(token))
  }

  function filterCharacteristicsForSelectedRecord(characteristics, candidates, selectedId) {
    const selectedCandidate = getSelectedCandidate(candidates, selectedId)
    const selectedTokens = getSelectedCandidateMatchTokens(candidates, selectedId)
    const selectedModelCodes = new Set(extractModelCodes([
      selectedCandidate?.id,
      selectedCandidate?.label,
      selectedCandidate?.description
    ].filter(Boolean).join(" ")))
    const siblingTokens = (candidates || [])
      .filter((candidate) => normalizeText(candidate.id) !== normalizeText(selectedId))
      .flatMap((candidate) => [candidate.id, candidate.label, getProductCardModelCode(candidate)])
      .map((value) => normalizeText(value).toLowerCase())
      .filter(Boolean)
      .filter((value) => value.length >= 5 || /\b[a-z]{1,4}-\d{1,4}\b/i.test(value))
    const siblingModelCodes = new Set(
      (candidates || [])
        .filter((candidate) => normalizeText(candidate.id) !== normalizeText(selectedId))
        .flatMap((candidate) => extractModelCodes([candidate.id, candidate.label, candidate.description].filter(Boolean).join(" ")))
    )

    if (!selectedTokens.length) return characteristics || []

    return (characteristics || []).map((characteristic) => {
      const options = getCharacteristicOptions(characteristic)
      const relabeledOptions = options.map((option) =>
        relabelCombinedOptionForSelectedProduct(option, selectedCandidate, selectedModelCodes, siblingModelCodes)
      )
      const selectedMatches = relabeledOptions.filter((option) => optionMatchesVariant(option, selectedTokens))
      const siblingMatches = options.filter((option) => optionMatchesVariant(option, siblingTokens))
      if (!selectedMatches.length) return characteristic
      if (!siblingMatches.length) {
        return {
          ...characteristic,
          options: relabeledOptions
        }
      }

      const containsCombinedSharedOptions = relabeledOptions.some((option) =>
        optionContainsSelectedAndSiblingModels(option.name, selectedModelCodes, siblingModelCodes)
      )
      if (containsCombinedSharedOptions) {
        const sharedOptions = relabeledOptions.map((option) => ({
          ...option,
          name: normalizeText(selectedCandidate?.label || option.name || "Selected Product")
        }))
        return {
          ...characteristic,
          options: sharedOptions,
          blurb: normalizeText(characteristic?.blurb || "") || "Shared details that apply to the selected product."
        }
      }

      const filteredSelectedOnly = selectedMatches.filter((option) => !optionContainsSiblingModels(option, siblingModelCodes))
      if (filteredSelectedOnly.length) {
        return {
          ...characteristic,
          options: filteredSelectedOnly
        }
      }

      return {
        ...characteristic,
        options: selectedMatches
      }
    })
  }

  function filterCharacteristicsForSelectedVariant(characteristics, variantCandidates, selectedVariantId) {
    const selectedTokens = getVariantMatchTokens(variantCandidates, selectedVariantId)
    const siblingTokens = (variantCandidates || [])
      .filter((candidate) => normalizeText(candidate.id) !== normalizeText(selectedVariantId))
      .flatMap((candidate) => [candidate.id, candidate.label])
      .map((value) => normalizeText(value).toLowerCase())
      .filter(Boolean)

    if (!selectedTokens.length) return characteristics || []

    return (characteristics || []).map((characteristic) => {
      const options = getCharacteristicOptions(characteristic)
      if (options.length < 2) return characteristic
      const selectedMatches = options.filter((option) => optionMatchesVariant(option, selectedTokens))
      const siblingMatches = options.filter((option) => optionMatchesVariant(option, siblingTokens))
      if (!selectedMatches.length || !siblingMatches.length) return characteristic
      return {
        ...characteristic,
        options: selectedMatches
      }
    })
  }

  function getSelectedProductDisplayLabel(decisionResult, structureRouting) {
    const selectedId = normalizeText(decisionResult?.selectedProductId || "")
    if (!selectedId) return ""
    const candidates = [
      ...(decisionResult?.productCandidates || []),
      ...(structureRouting?.productCandidates || [])
    ]
    const match = candidates.find((candidate) => normalizeText(candidate.id) === selectedId)
    return normalizeText(match?.label || selectedId)
  }

  function mergeComColIntoDimensions(characteristics) {
    const items = Array.isArray(characteristics) ? characteristics.map((item) => ({
      ...item,
      options: getCharacteristicOptions(item)
    })) : []
    const dimensionsIndex = items.findIndex((item) => /size|dimension/i.test(normalizeText(item.label)))
    if (dimensionsIndex === -1) return items

    const dimensions = { ...items[dimensionsIndex], options: [...items[dimensionsIndex].options] }
    const mergedIndexes = new Set()

    items.forEach((item, index) => {
      if (index === dimensionsIndex) return
      const normalizedLabel = normalizeText(item.label).toLowerCase()
      const isTextileRequirement = /\bcom\b|\bcol\b|customer'?s own leather|requirements/i.test(normalizedLabel)
      if (!isTextileRequirement) return
      mergedIndexes.add(index)

      const textileOptions = getCharacteristicOptions(item)
      textileOptions.forEach((textileOption) => {
        const textileName = normalizeText(textileOption.name).toLowerCase()
        const matchingDimension = dimensions.options.find((dimensionOption) => {
          const optionName = normalizeText(dimensionOption.name).toLowerCase()
          return textileName && optionName && (textileName.includes(optionName) || optionName.includes(textileName))
        })

        if (matchingDimension) {
          if (!normalizeText(matchingDimension.pricing) && normalizeText(textileOption.pricing)) {
            matchingDimension.pricing = textileOption.pricing
          }
          return
        }

        if (dimensions.options.length === 1 && textileOptions.length === 1 && !normalizeText(dimensions.options[0].pricing) && normalizeText(textileOption.pricing)) {
          dimensions.options[0].pricing = textileOption.pricing
        }
      })
    })

    return items
      .map((item, index) => (index === dimensionsIndex ? dimensions : item))
      .filter((_, index) => !mergedIndexes.has(index))
  }

  function removeTextileRequirementOptionsFromUpholstery(characteristics) {
    const items = Array.isArray(characteristics) ? characteristics.map((item) => ({
      ...item,
      options: getCharacteristicOptions(item)
    })) : []
    const hasDimensions = items.some((item) => /size|dimension/i.test(normalizeText(item.label)))
    if (!hasDimensions) return items

    return items
      .map((item) => {
        if (!/upholstery|material/i.test(normalizeText(item.label))) return item
        const filteredOptions = item.options.filter((option) => !/^(COM|COL)$/i.test(normalizeText(option.name)))
        return {
          ...item,
          options: filteredOptions
        }
      })
      .filter((item) => item.options.length >= 2 || /size|dimension|configuration/i.test(item.label))
  }

  function normalizeDecisionCharacteristics(items) {
    if (!Array.isArray(items)) return []
    return items
      .map((item, index) => {
        const label = normalizeText(item?.label || item?.name || `Characteristic ${index + 1}`)
        const compactLabel = isCompactCharacteristicLabel(label)
        return {
          id: normalizeText(item?.id || item?.key || item?.label || `characteristic-${index + 1}`),
          label,
          blurb: normalizeText(item?.blurb || item?.summary || ""),
          dependencyNote: normalizeText(item?.dependency_note || item?.dependencyNote || item?.note || ""),
          pricingNote: normalizeText(item?.pricing_note || item?.pricingNote || ""),
          configuration_sections: Array.isArray(item?.configuration_sections)
            ? item.configuration_sections
            : Array.isArray(item?.configurationSections)
              ? item.configurationSections
              : [],
          pricing_matrix: item?.pricing_matrix || item?.pricingMatrix || null,
          options: getCharacteristicOptions(item).map((option) => compactLabel
            ? { ...option, values: "", difference: "" }
            : option
          )
        }
      })
      .filter((item) => item.id && item.label)
      .filter((item) => {
        const normalizedLabel = item.label.toLowerCase()
        const looksPricingOnly = /base price|pricing|price list|order information/i.test(normalizedLabel)
        return !looksPricingOnly
      })
      .filter((item) => item.options.length >= 2 || /size|dimension|configuration/i.test(item.label))
  }

  function mapConfigurationSectionToCharacteristicLabel(sectionTitle) {
    const normalized = normalizeText(sectionTitle).toLowerCase()
    if (!normalized) return ""
    if (/shell finish|wood veneer|veneer/i.test(normalized)) return "Shell Finish"
    if (/base finish|frame finish|base\/frame finish|legs\/base/i.test(normalized)) return "Base / Frame Finish"
    if (/upholstery|fabric|material/i.test(normalized)) return "Upholstery / Fabric"
    return ""
  }

  function isDerivedNonStructuralConfigurationSection(sectionTitle) {
    return Boolean(mapConfigurationSectionToCharacteristicLabel(sectionTitle))
  }

  function isGenericConfigurationOptionName(name) {
    const normalized = normalizeText(name).toLowerCase()
    if (!normalized) return false
    return /^(shell|nonupholstered|upholstered|fully upholstered|exposed veneer shell|shell only|side chair shell)$/.test(normalized)
  }

  function splitOvergroupedConfiguration(characteristics) {
    const items = Array.isArray(characteristics) ? [...characteristics] : []
    const derivedCharacteristics = []

    const normalizedExistingLabels = new Set(items.map((item) => normalizeText(item.label).toLowerCase()))

    const nextItems = items
      .map((item) => {
        if (!/configuration/i.test(normalizeText(item.label)) || item.pricing_matrix || item.pricingMatrix) {
          return item
        }

        const rawSections = Array.isArray(item.configuration_sections)
          ? item.configuration_sections
          : Array.isArray(item.configurationSections)
            ? item.configurationSections
            : []

        if (!rawSections.length) return item

        const keptSections = []
        let derivedSectionCount = 0
        rawSections.forEach((section) => {
          const title = normalizeText(section?.title || section?.name || section?.label || "")
          const mappedLabel = mapConfigurationSectionToCharacteristicLabel(title)
          const sectionOptions = Array.isArray(section?.options)
            ? section.options.map((option) => normalizeText(option)).filter(Boolean)
            : []

          if (!mappedLabel) {
            keptSections.push(section)
            return
          }

          derivedSectionCount += 1

          if (!normalizedExistingLabels.has(mappedLabel.toLowerCase())) {
            derivedCharacteristics.push({
              id: slugify(mappedLabel, mappedLabel.toLowerCase()),
              label: mappedLabel,
              blurb: "",
              dependencyNote: "",
              pricingNote: "",
              options: sectionOptions.map((option) => ({
                name: option,
                values: "",
                pricing: "",
                difference: "",
                evidence: ""
              }))
            })
            normalizedExistingLabels.add(mappedLabel.toLowerCase())
          }
        })

        const nextItem = {
          ...item,
          configuration_sections: keptSections
        }

        const optionNames = getCharacteristicOptions(nextItem).map((option) => normalizeText(option.name)).filter(Boolean)
        const hasStructuralSections = keptSections.some((section) => !isDerivedNonStructuralConfigurationSection(section?.title || section?.name || section?.label || ""))
        const hasStructuralOptions = optionNames.some((name) => !/finish|veneer|upholstery|fabric|material/i.test(name) && !isGenericConfigurationOptionName(name))
        const onlyDerivedSectionsRemain = rawSections.length > 0 && derivedSectionCount === rawSections.length && !hasStructuralSections
        if (!nextItem.pricing_matrix && !nextItem.pricingMatrix && onlyDerivedSectionsRemain && !hasStructuralOptions) {
          return null
        }
        if (!nextItem.pricing_matrix && !nextItem.pricingMatrix && !keptSections.length && optionNames.length > 0 && optionNames.every((name) => isGenericConfigurationOptionName(name))) {
          return null
        }

        return nextItem
      })
      .filter(Boolean)
      .filter((item) => {
        if (!/configuration/i.test(normalizeText(item.label))) return true
        const sections = Array.isArray(item.configuration_sections) ? item.configuration_sections : []
        const options = getCharacteristicOptions(item)
        return Boolean(item.pricing_matrix || item.pricingMatrix || sections.length || options.length)
      })

    return [...nextItems, ...derivedCharacteristics]
  }

  function hasCharacteristicLabel(characteristics, pattern) {
    return (characteristics || []).some((item) => pattern.test(normalizeText(item?.label || "")))
  }

  function formatCompactSourceLabel(value, fallbackPageNumber = "") {
    const normalized = normalizeText(value)
    if (!normalized) return fallbackPageNumber ? `p. ${fallbackPageNumber}` : ""
    const pageMatch = normalized.match(/page\s+(\d+)/i)
    if (pageMatch) return `p. ${pageMatch[1]}`
    return fallbackPageNumber ? `p. ${fallbackPageNumber}` : normalized
  }

  function formatOptionPricingLabel(value) {
    const normalized = normalizeText(value)
    if (!normalized) return ""
    const quantityMatch = normalized.match(/\b(COM|COL)\b[^0-9]*?(\d+(?:\.\d+)?)\s*(sq ft|sqft|yd|yards?)\b/i)
    if (quantityMatch) {
      const unit = /sq\s*ft|sqft/i.test(quantityMatch[3]) ? "SqFt" : quantityMatch[3]
      return `${quantityMatch[1].toUpperCase()}: ${quantityMatch[2]} ${unit}`
    }
    return normalized
  }

  function getCompactOptionPricing(option) {
    const directPricing = formatOptionPricingLabel(option?.pricing || "")
    if (directPricing) return directPricing

    const valueText = normalizeText(option?.values || "")
    if (!valueText) return ""

    const textileMatch = valueText.match(/\b(COM|COL)\b[^0-9]*?(\d+(?:\.\d+)?)\s*(sq ft|sqft|yd|yards?)\b/i)
    if (textileMatch) {
      return formatOptionPricingLabel(textileMatch[0])
    }

    const priceMatch = valueText.match(/[$+][\d,]+(?:\.\d+)?/)
    if (priceMatch) return normalizeText(priceMatch[0])

    return ""
  }

  function extractTextileRequirementRows(value) {
    const normalized = normalizeText(value)
    if (!normalized) return []
    return [...normalized.matchAll(/\b(COM|COL)\b[^0-9]*?(\d+(?:\.\d+)?)\s*(sq ft|sqft|yd|yards?)\b/gi)]
      .map((match) => {
        const unit = /sq\s*ft|sqft/i.test(match[3]) ? "SqFt" : match[3]
        return `${match[1].toUpperCase()}: ${match[2]} ${unit}`
      })
  }

  function parseConfigurationDisplayRows(rawValueText) {
    return normalizeText(rawValueText)
      .split(/\s*;\s*|\n+/)
      .map((segment) => normalizeText(segment))
      .filter(Boolean)
      .map((segment) => {
        let match = segment.match(/^(.+?):\s*([$+][\d,]+(?:\.\d+)?)\s*\[(EA\d+)\]$/i)
        if (match) {
          return {
            label: formatConfigurationDisplayLabel(match[1]),
            price: normalizeText(match[2]),
            model: normalizeText(match[3]).toUpperCase()
          }
        }
        match = segment.match(/^(.+?):\s*([$+][\d,]+(?:\.\d+)?)$/i)
        if (match) {
          return {
            label: formatConfigurationDisplayLabel(match[1]),
            price: normalizeText(match[2]),
            model: ""
          }
        }
        match = segment.match(/^([SF])\s+([$+][\d,]+(?:\.\d+)?)$/i)
        if (match) {
          return {
            label: formatConfigurationDisplayLabel(match[1]),
            price: normalizeText(match[2]),
            model: ""
          }
        }
        return {
          label: formatConfigurationDisplayLabel(segment),
          price: "",
          model: ""
        }
      })
      .filter((row) => row.label)
  }

  function parseConfigurationInlinePricing(rawPricingText) {
    return normalizeText(rawPricingText)
      .split(/\s*;\s*|\n+/)
      .map((segment) => normalizeText(segment))
      .filter(Boolean)
      .map((segment) => {
        const parts = segment.split(/:\s+/)
        if (parts.length < 2) return null
        return {
          label: normalizeText(parts[0]),
          price: normalizeText(parts.slice(1).join(": "))
        }
      })
      .filter(Boolean)
  }

  function formatConfigurationDisplayLabel(value) {
    const normalized = normalizeText(value)
    if (!normalized) return ""
    if (/^(s\/f|swivel\/fixed)$/i.test(normalized)) return "Swivel / Fixed"
    if (/^s$/i.test(normalized)) return "Swivel"
    if (/^f$/i.test(normalized)) return "Fixed"
    return normalized
      .replace(/\bS\/F\b/gi, "Swivel / Fixed")
      .replace(/\bSwivel\b/gi, "Swivel")
      .replace(/\bFixed\b/gi, "Fixed")
      .replace(/\bEA\d+\b/gi, "")
      .replace(/\s{2,}/g, " ")
      .trim()
  }

  function getConfigurationMatchKey(value) {
    const formatted = formatConfigurationDisplayLabel(value)
    if (!formatted) return ""
    return formatted
      .toLowerCase()
      .replace(/\((?:s|f)\)/gi, "")
      .replace(/\bno swivel\b/gi, "fixed")
      .replace(/\bfixed base\b/gi, "fixed")
      .replace(/\bswivel base\b/gi, "swivel")
      .replace(/[^a-z0-9]+/g, " ")
      .trim()
  }

  function parseConfigurationMatrix(rawPricingText) {
    const normalized = String(rawPricingText || "")
    if (!/rows?\s*:|columns?\s*:|cells?\s*:/i.test(normalized)) return null

    const rowsMatch = normalized.match(/rows?\s*:\s*([\s\S]*?)(?=columns?\s*:|cells?\s*:|$)/i)
    const columnsMatch = normalized.match(/columns?\s*:\s*([\s\S]*?)(?=cells?\s*:|$)/i)
    const cellsMatch = normalized.match(/cells?\s*:\s*([\s\S]*)$/i)

    const rows = normalizeText(rowsMatch?.[1] || "")
      .split(/\s*\|\s*|\s*;\s*/)
      .map((item) => normalizeText(item))
      .filter(Boolean)
    const columns = normalizeText(columnsMatch?.[1] || "")
      .split(/\s*\|\s*|\s*;\s*/)
      .map((item) => normalizeText(item))
      .filter(Boolean)
    const cells = new Map()

    normalizeText(cellsMatch?.[1] || "")
      .split(/\s*;\s*/)
      .map((item) => normalizeText(item))
      .filter(Boolean)
      .forEach((entry) => {
        const cellMatch = entry.match(/(.+?)\s*[x/]\s*(.+?)\s*(?:=|:)\s*(.+)$/i)
        if (!cellMatch) return
        const rowKey = normalizeText(cellMatch[1]).toLowerCase()
        const columnKey = normalizeText(cellMatch[2]).toLowerCase()
        const price = normalizeText(cellMatch[3])
        cells.set(`${rowKey}:::${columnKey}`, price)
      })

    if (!rows.length || !columns.length || !cells.size) return null
    return { rows, columns, cells }
  }

  function sortCharacteristicsForDisplay(characteristics) {
    const items = Array.isArray(characteristics) ? [...characteristics] : []
    const sizeItems = items.filter((item) => /size|dimension/i.test(normalizeText(item.label)))
    const remaining = items.filter((item) => !/size|dimension/i.test(normalizeText(item.label)))
    return [...sizeItems, ...remaining]
  }

  function buildVariantTokenSet(variantCandidates, selectedVariantId) {
    const selectedTokens = getVariantMatchTokens(variantCandidates, selectedVariantId)
    return new Set(selectedTokens)
  }

  function measurementMatchesSelectedVariant(measurement, variantTokenSet) {
    if (!variantTokenSet?.size) return true
    const variantValue = normalizeText(measurement?.variant || "").toLowerCase()
    if (!variantValue) return true
    for (const token of variantTokenSet) {
      if (variantValue.includes(token) || token.includes(variantValue)) return true
    }
    return false
  }

  async function fillMissingDimensionOptions(documentRecord, includedPages, characteristics, variantCandidates, selectedVariantId) {
    const items = Array.isArray(characteristics) ? characteristics.map((item) => ({
      ...item,
      options: getCharacteristicOptions(item)
    })) : []
    const dimensionIndex = items.findIndex((item) => /size|dimension/i.test(normalizeText(item.label)))
    if (dimensionIndex === -1) return items
    if (items[dimensionIndex].options.length) return items

    const variantTokenSet = buildVariantTokenSet(variantCandidates, selectedVariantId)
    const groupedMeasurements = new Map()
    let fallbackSummary = ""

    for (const page of includedPages || []) {
      const canvas = await renderPdfPageToCanvas(documentRecord, page.pageNumber, 1.25)
      const result = await analyzeDimensionsWithVision(page.pageNumber, canvas)
      const measurements = Array.isArray(result?.analysis?.measurements) ? result.analysis.measurements : []
      if (!measurements.length) continue
      fallbackSummary = fallbackSummary || normalizeText(result.analysis?.summary || "")
      measurements
        .filter((measurement) => measurementMatchesSelectedVariant(measurement, variantTokenSet))
        .forEach((measurement) => {
          const variantLabel = normalizeText(measurement.variant || "") || "Dimensions"
          if (!groupedMeasurements.has(variantLabel)) {
            groupedMeasurements.set(variantLabel, {
              name: variantLabel === "Dimensions" ? "Dimensions" : titleCaseWords(variantLabel),
              values: [],
              pricing: "",
              difference: "",
              evidence: `Detected on page ${page.pageNumber}`
            })
          }
          const entry = groupedMeasurements.get(variantLabel)
          const label = normalizeText(measurement.label || "")
          const value = normalizeText(measurement.value || "")
          if (label && value) {
            entry.values.push(`${label}: ${value}`)
          }
        })
      if (groupedMeasurements.size) break
    }

    if (!groupedMeasurements.size) return items

    items[dimensionIndex] = {
      ...items[dimensionIndex],
      blurb: items[dimensionIndex].blurb || fallbackSummary,
      options: [...groupedMeasurements.values()].map((option) => ({
        ...option,
        values: option.values.join("; ")
      }))
    }
    return items
  }

  function mergeReferenceTextileIntoDimensions(characteristics, referenceInfo) {
    const items = Array.isArray(characteristics) ? characteristics.map((item) => ({
      ...item,
      options: getCharacteristicOptions(item)
    })) : []
    const dimensionsIndex = items.findIndex((item) => /size|dimension/i.test(normalizeText(item.label)))
    if (dimensionsIndex === -1) return items

    const textileRows = (referenceInfo || [])
      .flatMap((entry) => extractTextileRequirementRows(entry))
      .filter(Boolean)
    if (!textileRows.length) return items

    const dimensions = { ...items[dimensionsIndex], options: [...items[dimensionsIndex].options] }
    if (!dimensions.options.length) return items

    const existingPricingRows = dimensions.options[0].pricing
      ? extractTextileRequirementRows(dimensions.options[0].pricing)
      : []
    const mergedRows = [...new Set([...existingPricingRows, ...textileRows])]
    if (!mergedRows.length) return items

    dimensions.options[0] = {
      ...dimensions.options[0],
      pricing: mergedRows.join("; ")
    }

    return items.map((item, index) => (index === dimensionsIndex ? dimensions : item))
  }

  function extractOptionPricingFromText(text, optionName) {
    const normalizedOption = normalizeText(optionName)
    if (!normalizedOption) return ""

    const lines = String(text || "")
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean)

    const normalizedSearchTerms = [...new Set([
      normalizedOption,
      normalizedOption.replace(/^[A-Z0-9]+\s+/, ""),
      normalizedOption.replace(/\([^)]*\)/g, ""),
      normalizedOption.replace(/^[A-Z0-9]+\s+/, "").replace(/\([^)]*\)/g, "")
    ].map((value) => normalizeText(value)).filter((value) => value.length >= 3))]
    const codeMatch = normalizedOption.match(/\(([A-Z0-9]{1,4})\)|\b([A-Z0-9]{1,4})$/i)
    const optionCode = normalizeText(codeMatch?.[1] || codeMatch?.[2] || "").toUpperCase()
    const plainLabel = normalizedSearchTerms[normalizedSearchTerms.length - 1] || normalizedOption
    const reversedSearchTerms = optionCode && plainLabel
      ? [`${optionCode} ${plainLabel}`, `${plainLabel} ${optionCode}`]
      : []
    const allSearchTerms = [...new Set([...normalizedSearchTerms, ...reversedSearchTerms].map((value) => normalizeText(value)).filter((value) => value.length >= 2))]

    const extractPriceFromCandidate = (candidateLine) => {
      const dollarMatch = candidateLine.match(/[+\-]?\$[\d,]+(?:\.\d+)?(?:\s*(?:ea|each|list|net))?/i)
      if (dollarMatch) return normalizeText(dollarMatch[0])

      const upchargeMatch = candidateLine.match(/(?:upcharge|add|added|premium|list price|price)\s*[:\-]?\s*([+\-]?\$[\d,]+(?:\.\d+)?(?:\s*(?:ea|each|sq ft))?)/i)
      if (upchargeMatch) return `${candidateLine.match(/upcharge|add|added|premium|list price|price/i)?.[0] || "Price"} ${normalizeText(upchargeMatch[1])}`

      const textileMatch = candidateLine.match(/\b(?:COM|COL)\b[^.;]*?\b\d+(?:\.\d+)?\s*(?:sq ft|yd|yards?)\b/i)
      if (textileMatch) return textileMatch[0]

      return ""
    }

    const fullText = normalizeText(text)
    if (fullText) {
      for (const term of allSearchTerms) {
        const escapedTerm = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
        const forwardPattern = new RegExp(`${escapedTerm}[^$+\\-]{0,80}([+\\-]?\\$[\\d,]+(?:\\.\\d+)?)`, "i")
        const backwardPattern = new RegExp(`([+\\-]?\\$[\\d,]+(?:\\.\\d+)?)[^$+\\-]{0,80}${escapedTerm}`, "i")
        const forwardMatch = fullText.match(forwardPattern)
        if (forwardMatch?.[1]) return normalizeText(forwardMatch[1])
        const backwardMatch = fullText.match(backwardPattern)
        if (backwardMatch?.[1]) return normalizeText(backwardMatch[1])
      }

      if (optionCode) {
        const compactCodePattern = new RegExp(`\\b${optionCode}\\b[^$+\\-]{0,40}([+\\-]?\\$[\\d,]+(?:\\.\\d+)?)`, "i")
        const compactCodeMatch = fullText.match(compactCodePattern)
        if (compactCodeMatch?.[1]) return normalizeText(compactCodeMatch[1])
      }
    }

    for (let index = 0; index < lines.length; index += 1) {
      const line = lines[index]
      const loweredLine = line.toLowerCase()
      const matchesOption = allSearchTerms.some((term) => loweredLine.includes(term.toLowerCase()))
      if (!matchesOption) continue

      const candidateLines = [
        line,
        lines[index + 1] || "",
        lines[index + 2] || "",
        lines[index - 1] || ""
      ].filter(Boolean)

      for (const candidateLine of candidateLines) {
        const price = extractPriceFromCandidate(candidateLine)
        if (price) return price
      }
    }

    return ""
  }

  function getCharacteristicSearchTerms(characteristicLabel, optionName) {
    const normalizedLabel = normalizeText(characteristicLabel)
    const normalizedOption = normalizeText(optionName)
    const strippedOption = normalizeText(normalizedOption.replace(/^[A-Z0-9]+\s+/, "").replace(/\([^)]*\)/g, ""))
    const codeMatch = normalizedOption.match(/\(([A-Z0-9]{1,4})\)|\b([A-Z0-9]{1,4})$/i)
    const optionCode = normalizeText(codeMatch?.[1] || codeMatch?.[2] || "").toUpperCase()
    const terms = [normalizedOption, strippedOption]

    if (optionCode && strippedOption) {
      terms.push(`${optionCode} ${strippedOption}`, `${strippedOption} ${optionCode}`, optionCode)
    }

    if (/shell finish/i.test(normalizedLabel)) {
      terms.push(`${normalizedOption} shell finish`, `${strippedOption} shell finish`, `${normalizedOption} veneer`, `${strippedOption} veneer`)
    }
    if (/base|frame finish/i.test(normalizedLabel)) {
      terms.push(`${normalizedOption} base finish`, `${strippedOption} base finish`, `${normalizedOption} frame finish`, `${strippedOption} frame finish`)
    }

    return [...new Set(terms.map((term) => normalizeText(term)).filter((term) => term.length >= 3))]
  }

  function extractDetailedUpholsteryOptionsFromText(text, pageNumber) {
    const options = new Map()
    String(text || "")
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean)
      .forEach((line) => {
        const match = line.match(/^Price Category\s+(.+?)\s+([$+][\d,]+(?:\.\d+)?)$/i)
        if (!match) return
        const rawLabel = normalizeText(match[1])
        const price = normalizeText(match[2])
        const optionName = /^(\d+|[A-Z])$/.test(rawLabel)
          ? `Category ${rawLabel}`
          : /^COM|COL$/i.test(rawLabel)
            ? rawLabel.toUpperCase()
            : `Category ${rawLabel}`
        options.set(optionName, {
          name: optionName,
          values: "",
          pricing: price,
          difference: "",
          evidence: `Detected on page ${pageNumber}`
        })
      })
    return options
  }

  function shouldReplaceWithDetailedUpholstery(options) {
    const names = (options || []).map((option) => normalizeText(option.name).toLowerCase())
    if (!names.length) return true
    return names.every((name) => /standard categories|material options|fabric price category|price categories|^com$|^col$/.test(name))
  }

  function extractSupplementalFinishOptionsFromText(text, pageNumber) {
    const shellOptions = new Map()
    const baseFrameOptions = new Map()
    const lines = String(text || "")
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean)

    const addOption = (map, name, pricing = "") => {
      const key = normalizeText(name)
      if (!key) return
      map.set(key, {
        name: key,
        values: "",
        pricing: normalizeText(pricing),
        difference: "",
        evidence: `Detected on page ${pageNumber}`
      })
    }

    lines.forEach((line) => {
      let match = line.match(/^(?:-)?([A-Z0-9]{1,4})\s+(.+?)\s+([+\-]?\$[\d,]+(?:\.\d+)?)$/i)
      if (match) {
        const code = normalizeText(match[1]).toUpperCase()
        const label = titleCaseWords(match[2])
        const price = normalizeText(match[3])
        if (/black|chrome|powder coat|metal/i.test(label)) {
          addOption(baseFrameOptions, `${label} (${code})`, price)
          return
        }
        addOption(shellOptions, `${label} (${code})`, price)
        return
      }

      match = line.match(/^(?:-)?([A-Z0-9]{1,4})\s+(.+)$/i)
      if (!match) return
      const code = normalizeText(match[1]).toUpperCase()
      const label = normalizeText(match[2])
      if (!label || /^(list price|grades|com\/col|item|description)$/i.test(label)) return

      const pricing = extractOptionPricingFromText(text, `${label} (${code})`) || extractOptionPricingFromText(text, `${code} ${label}`)
      if (/black|chrome|powder coat|metal/i.test(label)) {
        addOption(baseFrameOptions, `${titleCaseWords(label)} (${code})`, pricing)
        return
      }
      if (/oak|walnut|ash|ebony|palisander|stain|paint|veneer/i.test(label)) {
        addOption(shellOptions, `${titleCaseWords(label)} (${code})`, pricing)
      }
    })

    return {
      shellOptions,
      baseFrameOptions
    }
  }

  function extractConfigurationComboPrices(text) {
    const normalizedLines = String(text || "")
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean)
    const armCodeMap = new Map()
    const baseCodeMap = new Map()
    let prefix = ""

    normalizedLines.forEach((line) => {
      const prefixMatch = line.match(/\b(EA\d+)\b/i)
      if (prefixMatch && !prefix) {
        const raw = prefixMatch[1].toUpperCase()
        prefix = raw.length > 3 ? raw.slice(0, -1) : raw
      }
      const armMatch = line.match(/^(\d+)\s+(no arms|arms with arm pads|arms)$/i)
      if (armMatch) {
        armCodeMap.set(normalizeText(armMatch[2]).toLowerCase(), armMatch[1])
      }
      const baseMatch = line.match(/^([SF])\s+(swivel base|fixed base(?:,\s*no swivel)?)$/i)
      if (baseMatch) {
        baseCodeMap.set(normalizeText(baseMatch[2]).toLowerCase(), baseMatch[1].toUpperCase())
        if (/fixed/i.test(baseMatch[2])) baseCodeMap.set("fixed base", baseMatch[1].toUpperCase())
        if (/swivel/i.test(baseMatch[2])) baseCodeMap.set("swivel base", baseMatch[1].toUpperCase())
      }
    })

    const priceMap = new Map()
    normalizedLines.forEach((line) => {
      const priceMatch = line.match(/\b(EA\d+)\b\s+([SF])\s+([$+][\d,]+(?:\.\d+)?)$/i)
      if (!priceMatch) return
      priceMap.set(`${priceMatch[1].toUpperCase()}:::${priceMatch[2].toUpperCase()}`, normalizeText(priceMatch[3]))
    })

    return { prefix, armCodeMap, baseCodeMap, priceMap }
  }

  function buildConfigurationMatrixFromSignals(configSignals, fallbackOptions = []) {
    const armEntries = [...configSignals.armCodeMap.entries()]
      .map(([label, code]) => ({ label: formatConfigurationDisplayLabel(label), code }))
    const baseEntries = [
      { label: "Swivel", code: configSignals.baseCodeMap.get("swivel base") || "" },
      { label: "Fixed", code: configSignals.baseCodeMap.get("fixed base") || "" }
    ].filter((entry) => entry.code)

    if (!armEntries.length || !baseEntries.length || !configSignals.prefix || !configSignals.priceMap.size) {
      return null
    }

    const rows = armEntries
      .map((arm) => {
        const model = `${configSignals.prefix}${arm.code}`
        const cells = baseEntries
          .map((base) => ({
            column: base.label,
            price: configSignals.priceMap.get(`${model}:::${base.code}`) || "",
            model
          }))
          .filter((cell) => cell.price)
        return {
          row_name: arm.label,
          cells
        }
      })
      .filter((row) => row.cells.length)

    if (!rows.length) return null

    const configurationSections = armEntries.length
      ? [{ title: "Arms", options: armEntries.map((entry) => entry.label) }, { title: "Base Type", options: baseEntries.map((entry) => entry.label) }]
      : fallbackOptions

    return {
      sections: configurationSections,
      matrix: {
        row_label: "Arm Type",
        column_labels: baseEntries.map((entry) => entry.label),
        rows
      }
    }
  }

  function enrichCharacteristicsFromSectionPages(characteristics, sectionPages) {
    const items = Array.isArray(characteristics) ? characteristics.map((item) => ({
      ...item,
      options: getCharacteristicOptions(item)
    })) : []
    const combinedText = (sectionPages || []).map((page) => getPageCombinedText(page)).join("\n")
    const configSignals = extractConfigurationComboPrices(combinedText)

    return items.map((item) => {
      const normalizedLabel = normalizeText(item.label).toLowerCase()

      if (/upholstery|material/i.test(normalizedLabel)) {
        const detailedOptions = new Map()
        ;(sectionPages || []).forEach((page) => {
          const pageOptions = extractDetailedUpholsteryOptionsFromText(getPageCombinedText(page), page.pageNumber)
          pageOptions.forEach((value, key) => detailedOptions.set(key, value))
        })
        if (detailedOptions.size >= 4 && shouldReplaceWithDetailedUpholstery(item.options)) {
          return {
            ...item,
            options: [...detailedOptions.values()]
          }
        }
      }

      if (/configuration/i.test(normalizedLabel)) {
        const fallbackSections = item.options.length
          ? [
              {
                title: "Arms",
                options: item.options.map((option) => formatConfigurationDisplayLabel(option.name))
              }
            ]
          : []
        const builtMatrix = buildConfigurationMatrixFromSignals(configSignals, fallbackSections)
        const enrichedOptions = item.options.map((option) => {
          const optionName = normalizeText(option.name).toLowerCase()
          const armCode = configSignals.armCodeMap.get(optionName)
          if (armCode && configSignals.baseCodeMap.size && configSignals.prefix) {
            const comboRows = []
            ;[
              ["swivel base", "Swivel"],
              ["fixed base", "Fixed"]
            ].forEach(([lookupLabel, displayLabel]) => {
              const baseCode = configSignals.baseCodeMap.get(lookupLabel)
              if (!baseCode) return
              const model = `${configSignals.prefix}${armCode}`
              const price = configSignals.priceMap.get(`${model}:::${baseCode}`)
              if (!price) return
              comboRows.push(`${displayLabel}: ${price} [${model}]`)
            })
            if (comboRows.length) {
              return {
                ...option,
                values: comboRows.join("; "),
                pricing: ""
              }
            }
          }
          if (normalizeText(option.pricing)) return option
          const parts = normalizeText(option.name).split(/\s*\/\s*/)
          if (parts.length !== 2) return option
          const armLabel = parts[0].toLowerCase()
          const baseLabel = parts[1].toLowerCase()
          const armComboCode = configSignals.armCodeMap.get(armLabel)
          const baseCode = configSignals.baseCodeMap.get(baseLabel)
          if (!armComboCode || !baseCode || !configSignals.prefix) return option
          const model = `${configSignals.prefix}${armComboCode}`
          const price = configSignals.priceMap.get(`${model}:::${baseCode}`)
          return price ? { ...option, pricing: price, values: `${formatConfigurationDisplayLabel(baseLabel)}: ${price} [${model}]` } : option
        })
        return {
          ...item,
          options: enrichedOptions,
          configuration_sections: builtMatrix?.sections || item.configuration_sections || item.configurationSections || [],
          pricing_matrix: builtMatrix?.matrix || item.pricing_matrix || item.pricingMatrix || null
        }
      }

      if (/shell finish|frame finish|base finish|wood finish|finish/i.test(normalizedLabel)) {
        const enrichedOptions = item.options.map((option) => {
          if (normalizeText(option.pricing)) return option
          const pricing = (sectionPages || [])
            .map((page) => {
              const pageText = getPageCombinedText(page)
              const searchTerms = getCharacteristicSearchTerms(item.label, option.name)
              return searchTerms
                .map((term) => extractOptionPricingFromText(pageText, term))
                .find(Boolean) || ""
            })
            .find(Boolean)
          return pricing ? { ...option, pricing } : option
        })

        return {
          ...item,
          options: enrichedOptions
        }
      }

      return item
    })
  }

  function buildSupplementalCharacteristics(documentRecord, sectionPages, characteristics) {
    const supplemental = []
    const existing = characteristics || []
    const shellFinishOptions = new Map()
    const baseFrameFinishOptions = new Map()
    const upholsteryOptions = new Map()

    ;(sectionPages || []).forEach((page) => {
      const text = getPageCombinedText(page)
      const signals = extractPageThemeSignals(text, documentRecord, page.pageNumber)
      const extractedFinishOptions = extractSupplementalFinishOptionsFromText(text, page.pageNumber)

      signals.finishLabels.forEach((label) => {
        const key = normalizeText(label)
        if (!key) return
        shellFinishOptions.set(key, {
          name: key.replace(/\s+Finish$/i, ""),
          values: "",
          pricing: extractOptionPricingFromText(text, key) || extractOptionPricingFromText(text, key.replace(/\s+Finish$/i, "")),
          difference: "",
          evidence: `Detected on page ${page.pageNumber}`
        })
      })
      extractedFinishOptions.shellOptions.forEach((value, key) => shellFinishOptions.set(key, value))
      extractedFinishOptions.baseFrameOptions.forEach((value, key) => baseFrameFinishOptions.set(key, value))

      const categoryMatches = text.match(/\b(?:COM|COL|Category\s+\d+|Category\s+[A-Z])\b/gi) || []
      categoryMatches.forEach((match) => {
        const key = normalizeText(match).replace(/\s+/g, " ")
        if (!key) return
        upholsteryOptions.set(key, {
          name: key.toUpperCase().startsWith("CATEGORY") ? titleCaseWords(key) : key.toUpperCase(),
          values: "",
          pricing: extractOptionPricingFromText(text, key),
          difference: "",
          evidence: `Detected on page ${page.pageNumber}`
        })
      })

      if (signals.hasUpholstery && !upholsteryOptions.size) {
        upholsteryOptions.set("Material Options", {
          name: "Material Options",
          values: "",
          pricing: "",
          difference: "",
          evidence: `Upholstery signals detected on page ${page.pageNumber}`
        })
      }
    })

    if (!hasCharacteristicLabel(existing, /shell finish|wood finish/i) && shellFinishOptions.size >= 2) {
      supplemental.push({
        id: "shell-finish",
        label: "Shell Finish",
        blurb: "Finish options detected across the pages gathered for this product or family.",
        dependencyNote: "",
        pricingNote: "",
        options: [...shellFinishOptions.values()]
      })
    }

    if (!hasCharacteristicLabel(existing, /base finish|frame finish|base \/ frame finish/i) && baseFrameFinishOptions.size >= 2) {
      supplemental.push({
        id: "base-frame-finish",
        label: "Base / Frame Finish",
        blurb: "Base or frame finish options detected across the pages gathered for this product or family.",
        dependencyNote: "",
        pricingNote: "",
        options: [...baseFrameFinishOptions.values()]
      })
    }

    if (!hasCharacteristicLabel(existing, /upholstery|material/i) && upholsteryOptions.size >= 2) {
      supplemental.push({
        id: "upholstery-material",
        label: "Upholstery / Material",
        blurb: "Upholstery or material-related options detected across the pages gathered for this product or family.",
        dependencyNote: "",
        pricingNote: "",
        options: [...upholsteryOptions.values()]
      })
    }

    return supplemental
  }

  function buildAttributesFromText(rawAttributes) {
    const labels = rawAttributes
      .split(/\n|,/)
      .map((item) => normalizeText(item))
      .filter(Boolean)

    return labels.map((label, index) => ({
      key: slugify(label, `field-${index + 1}`),
      label,
      type: getAttributeType(label),
      value: ""
    }))
  }

  function getActiveDocument() {
    return appState.documents.find((document) => document.id === appState.activeDocumentId) || appState.documents[0]
  }

  function getPageCombinedText(page) {
    return [page?.text || "", page?.ocrText || ""].filter(Boolean).join("\n")
  }

  function getPageRenderKey(documentRecord, pageNumber) {
    return `${documentRecord?.id || "unknown"}:${pageNumber}`
  }

  function clampPdfZoom(value) {
    return Math.min(2.25, Math.max(0.75, Math.round(value * 100) / 100))
  }

  function getPdfZoomLabel() {
    return `${Math.round(appState.pdfZoom * 100)}%`
  }

  function getViewerBaseRenderWidth() {
    const viewerScroll = document.getElementById("viewer-scroll")
    const viewerWidth = viewerScroll?.clientWidth || 0
    if (!viewerWidth) return 420
    return Math.max(320, viewerWidth - 72)
  }

  function escapeHtmlAttribute(value) {
    return escapeHtml(value).replaceAll("`", "&#96;")
  }

  function sanitizeReadableHtml(html) {
    const parser = new DOMParser()
    const documentFragment = parser.parseFromString(`<div>${html || ""}</div>`, "text/html")
    const allowedTags = new Set(["DIV", "H3", "H4", "P", "UL", "OL", "LI", "STRONG", "EM", "BR"])

    function sanitizeNode(node) {
      if (node.nodeType === Node.TEXT_NODE) {
        return document.createTextNode(node.textContent || "")
      }

      if (node.nodeType !== Node.ELEMENT_NODE) {
        return document.createDocumentFragment()
      }

      if (!allowedTags.has(node.tagName)) {
        const fragment = document.createDocumentFragment()
        ;[...node.childNodes].forEach((child) => fragment.appendChild(sanitizeNode(child)))
        return fragment
      }

      const safeNode = document.createElement(node.tagName.toLowerCase())
      ;[...node.childNodes].forEach((child) => safeNode.appendChild(sanitizeNode(child)))
      return safeNode
    }

    const wrapper = document.createElement("div")
    ;[...documentFragment.body.firstChild.childNodes].forEach((child) => {
      wrapper.appendChild(sanitizeNode(child))
    })
    return wrapper.innerHTML.trim()
  }

  function getAiOrderedPages(topPages) {
    if (!appState.aiRerankResult?.orderedPages?.length) return topPages

    const pageMap = new Map(topPages.map((page) => [page.pageNumber, page]))
    const ordered = []
    const keptPageSet = new Set(appState.aiRerankResult.keptPages || [])
    const scoredPageSet = new Set(appState.aiRerankResult.orderedPages.map((item) => item.pageNumber))

    appState.aiRerankResult.orderedPages.forEach((item) => {
      const page = pageMap.get(item.pageNumber)
      if (!page) return
      const variantMatch = appState.aiRerankResult.variantComparison?.find((variant) => variant.pageNumber === item.pageNumber)
      ordered.push({
        ...page,
        aiRole: item.role || "",
        aiReason: item.reason || "",
        aiConfidence: item.confidence || "",
        aiScore: typeof item.aiScore === "number" ? item.aiScore : "",
        aiRank: ordered.length + 1,
        aiSelected: item.pageNumber === appState.aiRerankResult.bestPage,
        aiVariant: variantMatch?.label || "",
        aiDifference: variantMatch?.difference || "",
        aiStatus: keptPageSet.has(item.pageNumber) ? "kept" : "filtered_out"
      })
      pageMap.delete(item.pageNumber)
    })

    pageMap.forEach((page, pageNumber) =>
      ordered.push({
        ...page,
        aiStatus: scoredPageSet.has(pageNumber) ? "filtered_out" : "not_scored"
      })
    )
    return ordered.slice(0, getAiKeptPageLimit())
  }

  function extractResponseText(payload) {
    if (typeof payload?.output_text === "string" && payload.output_text.trim()) {
      return payload.output_text.trim()
    }

    const outputBlocks = payload?.output || []
    const textParts = []
    outputBlocks.forEach((block) => {
      ;(block?.content || []).forEach((item) => {
        if (item?.type === "output_text" && item.text) {
          textParts.push(item.text)
        }
      })
    })
    return textParts.join("\n").trim()
  }

  function parseJsonPayload(rawText) {
    if (!rawText) return null

    try {
      return JSON.parse(rawText)
    } catch (error) {
      const fencedMatch = rawText.match(/```(?:json)?\s*([\s\S]*?)```/i)
      if (fencedMatch) {
        try {
          return JSON.parse(fencedMatch[1].trim())
        } catch (fencedError) {
          return null
        }
      }
    }

    return null
  }

  function formatVisionAnalysisText(analysis) {
    if (!analysis) return ""

    const lines = []
    if (analysis.summary) lines.push(`Summary: ${analysis.summary}`)

    if (Array.isArray(analysis.measurements) && analysis.measurements.length) {
      lines.push("Measurements:")
      analysis.measurements.forEach((item) => {
        const parts = [item.variant, item.label, item.value].filter(Boolean)
        if (parts.length) {
          lines.push(`- ${parts.join(": ")}`)
        }
      })
    }

    if (Array.isArray(analysis.raw_visible_dimensions) && analysis.raw_visible_dimensions.length) {
      lines.push("Visible dimensions:")
      analysis.raw_visible_dimensions.forEach((item) => {
        lines.push(`- ${item}`)
      })
    }

    if (Array.isArray(analysis.notes) && analysis.notes.length) {
      lines.push("Notes:")
      analysis.notes.forEach((item) => {
        lines.push(`- ${item}`)
      })
    }

    return lines.join("\n")
  }

  function extractPagesFromPdfText(pdfText) {
    const pageMatches = [...pdfText.matchAll(/%%PAGE:(\d+)\n([\s\S]*?)(?=%%PAGE:|%%EOF)/g)]
    return pageMatches.map((match) => ({
      pageNumber: Number(match[1]),
      text: decodePdfStream(match[2])
    }))
  }

  function decodePdfStream(pageSource) {
    const lines = []
    const textMatches = [...pageSource.matchAll(/\((.*?)\)\s*Tj/g)]
    textMatches.forEach((match) => {
      const line = match[1]
        .replaceAll("\\(", "(")
        .replaceAll("\\)", ")")
        .replaceAll("\\\\", "\\")
      lines.push(line)
    })
    return lines.join("\n")
  }

  function buildRetrievalSignals() {
    const { spec } = appState
    const signals = [
      { value: spec.originalSpecName, weight: 16, type: "product" },
      { value: spec.specDisplayName, weight: 18, type: "product" }
    ]

    if (appState.assistantSelection) {
      signals.push({ value: appState.assistantSelection, weight: 22, type: "clarifier" })
    }

    return signals.filter((signal) => normalizeText(signal.value))
  }

  function tokenize(value) {
    return normalizeText(value.toLowerCase())
      .split(/[^a-z0-9]+/)
      .filter(Boolean)
  }

  function buildProductMatchLadder() {
    const productName = normalizeText(appState.spec.specDisplayName || appState.spec.originalSpecName)
    const tokens = tokenize(productName)
    const seen = new Set()
    const ladder = []
    const genericTokens = new Set(["chair", "chairs", "seating", "collection", "standard", "series"])

    if (productName) {
      ladder.push({ label: "Full product phrase", phrase: productName.toLowerCase(), weight: 18 })
      seen.add(productName.toLowerCase())
    }

    for (let size = Math.min(tokens.length - 1, 3); size >= 2; size -= 1) {
      for (let index = 0; index <= tokens.length - size; index += 1) {
        const phrase = tokens.slice(index, index + size).join(" ")
        if (seen.has(phrase)) continue
        ladder.push({ label: "Product subphrase", phrase, weight: 12 })
        seen.add(phrase)
      }
    }

    tokens.forEach((token) => {
      if (seen.has(token)) return
      ladder.push({ label: genericTokens.has(token) ? "Downweighted token" : "Product token", phrase: token, weight: genericTokens.has(token) ? 1 : 4 })
      seen.add(token)
    })

    return ladder
  }

  function getQueryInterpretation() {
    const ladder = buildProductMatchLadder()
    return {
      primary: ladder.find((entry) => entry.label === "Full product phrase")?.phrase || "",
      secondary: ladder.filter((entry) => entry.label === "Product subphrase").map((entry) => entry.phrase).slice(0, 4),
      weakTokens: ladder.filter((entry) => entry.label === "Downweighted token").map((entry) => entry.phrase),
      strongTokens: ladder.filter((entry) => entry.label === "Product token").map((entry) => entry.phrase)
    }
  }

  function scorePage(pageText) {
    const haystack = ` ${normalizeText(pageText.toLowerCase())} `
    const signals = buildRetrievalSignals()
    const ladder = buildProductMatchLadder()
    let score = 0
    let productHit = false

    signals.forEach((signal) => {
      const lowered = signal.value.toLowerCase()
      if (!lowered) return

      if (haystack.includes(` ${lowered} `) || haystack.includes(lowered)) {
        score += signal.weight
        if (signal.type === "product") productHit = true
        return
      }

      const tokens = tokenize(lowered)
      const matchedTokens = tokens.filter((token) => haystack.includes(` ${token} `))
      if (!tokens.length || !matchedTokens.length) return

      const tokenCoverage = matchedTokens.length / tokens.length
      if (tokenCoverage >= 0.5) {
        score += Math.round(signal.weight * tokenCoverage * 0.72)
        if (signal.type === "product" && tokenCoverage >= 0.75) productHit = true
      }
    })

    ladder.forEach((entry) => {
      const count = countOccurrences(haystack, entry.phrase)
      if (!count) return
      score += entry.weight * Math.min(count, entry.label === "Product token" ? 2 : 1)
      if (entry.label !== "Product token") productHit = true
    })

    if (productHit) score += 8
    if (/product information|description|notes/i.test(pageText)) score += 4

    return score
  }

  function rankPages(documentRecord) {
    return documentRecord.pages
      .map((page) => ({
        pageNumber: page.pageNumber,
        score: scorePage(getPageCombinedText(page)),
        snippet: normalizeText(getPageCombinedText(page)).slice(0, 180)
      }))
      .sort((a, b) => b.score - a.score)
  }

  function getDocumentKeywordDensityScores(documents) {
    return documents
      .map((document) => {
        const ranked = rankPages(document)
        const topScore = ranked[0]?.score || 0
        const topThreeScore = ranked.slice(0, 3).reduce((total, page) => total + page.score, 0)
        const pageCount = Math.max(1, document.pages?.length || 1)
        // Larger PDFs have more chances to produce extreme top-page matches.
        // Apply a mild log penalty so relevance dominates over sheer size.
        const sizePenalty = 1 + 0.22 * Math.log2(pageCount)
        const selectionScore = Math.round(((topScore * 0.65) + ((topThreeScore / 3) * 0.35)) / sizePenalty)
        return {
          documentId: document.id,
          title: document.title,
          pageCount,
          topScore,
          topThreeScore,
          sizePenalty,
          selectionScore
        }
      })
      .sort((a, b) => b.selectionScore - a.selectionScore || b.topScore - a.topScore || b.topThreeScore - a.topThreeScore)
  }

  function chooseBestDocumentByKeywordDensity(documents) {
    const scoredDocuments = getDocumentKeywordDensityScores(documents)
    const bestId = scoredDocuments[0]?.documentId
    return documents.find((document) => document.id === bestId) || null
  }

  function runRanking(jumpToTopPage) {
    const documentRecord = getActiveDocument()
    if (!documentRecord) {
      appState.rankedPages = []
      appState.activePageNumber = 1
      updateStructureRoutingState(null)
      return
    }

    const ranked = rankPages(documentRecord)
    appState.rankedPages = ranked
    appState.assistantResult = analyzeRankedPages(documentRecord, ranked)
    const topMatch = ranked[0]
    if (jumpToTopPage) {
      appState.activePageNumber = topMatch?.pageNumber || 1
    }
    updateStructureRoutingState(documentRecord)
  }

  function cloneAiRerankResult(result) {
    if (!result) return null
    return {
      ...result,
      keptPages: Array.isArray(result.keptPages) ? [...result.keptPages] : [],
      orderedPages: Array.isArray(result.orderedPages) ? result.orderedPages.map((item) => ({ ...item })) : [],
      variantComparison: Array.isArray(result.variantComparison) ? result.variantComparison.map((item) => ({ ...item })) : [],
      concreteProductCandidates: Array.isArray(result.concreteProductCandidates) ? result.concreteProductCandidates.map((item) => ({ ...item })) : []
    }
  }

  function setPage(pageNumber) {
    if (appState.decisionAssistResult?.pageNumber && appState.decisionAssistResult.pageNumber !== pageNumber) {
      appState.decisionAssistResult = null
      appState.activeDecisionChoiceIndex = -1
      appState.decisionTabsWindowStart = 0
      appState.decisionAssistError = ""
    }
    appState.activePageNumber = pageNumber
    render()
    ensureFamilyPageProducts(pageNumber).catch(() => {})
    requestAnimationFrame(() => {
      const viewerScroll = document.getElementById("viewer-scroll")
      const targetPage = document.querySelector(`.pdf-page[data-page-number="${pageNumber}"]`)
      if (viewerScroll) {
        viewerScroll.scrollTo({ top: 0, behavior: "smooth" })
      }
      if (targetPage) {
        targetPage.scrollIntoView({ behavior: "smooth", block: "start" })
      }
    })
  }

  function renderPreservingViewerScroll() {
    const viewerScroll = document.getElementById("viewer-scroll")
    const previousScrollTop = viewerScroll?.scrollTop || 0
    const previousScrollLeft = viewerScroll?.scrollLeft || 0
    const previousWindowScrollY = window.scrollY || 0
    const previousWindowScrollX = window.scrollX || 0
    render()
    requestAnimationFrame(() => {
      const nextViewerScroll = document.getElementById("viewer-scroll")
      if (nextViewerScroll) {
        nextViewerScroll.scrollTop = previousScrollTop
        nextViewerScroll.scrollLeft = previousScrollLeft
      }
      window.scrollTo(previousWindowScrollX, previousWindowScrollY)
    })
  }

  function syncPageStripState() {
    const track = document.getElementById("page-strip-track")
    if (!track) {
      appState.pageStripCanScrollLeft = false
      appState.pageStripCanScrollRight = false
      appState.pageStripScrollLeft = 0
      return
    }

    if (Math.abs(track.scrollLeft - appState.pageStripScrollLeft) > 1) {
      track.scrollLeft = appState.pageStripScrollLeft
    }

    const maxScrollLeft = Math.max(0, track.scrollWidth - track.clientWidth)
    appState.pageStripScrollLeft = track.scrollLeft
    appState.pageStripCanScrollLeft = track.scrollLeft > 4
    appState.pageStripCanScrollRight = track.scrollLeft < maxScrollLeft - 4
  }

  function queuePageStripSync() {
    requestAnimationFrame(() => {
      const beforeLeft = appState.pageStripCanScrollLeft
      const beforeRight = appState.pageStripCanScrollRight
      syncPageStripState()
      if (beforeLeft !== appState.pageStripCanScrollLeft || beforeRight !== appState.pageStripCanScrollRight) {
        renderPreservingViewerScroll()
      }
    })
  }

  function syncDecisionTabsState() {
    const track = document.getElementById("decision-choice-tabs-track")
    if (!track) {
      appState.decisionTabsCanScrollLeft = false
      appState.decisionTabsCanScrollRight = false
      appState.decisionTabsScrollLeft = 0
      return
    }

    const maxScrollLeft = Math.max(0, track.scrollWidth - track.clientWidth)
    if (appState.decisionTabsScrollLeft > maxScrollLeft) {
      appState.decisionTabsScrollLeft = maxScrollLeft
    }
    if (Math.abs(track.scrollLeft - appState.decisionTabsScrollLeft) > 1) {
      track.scrollLeft = appState.decisionTabsScrollLeft
    }
    appState.decisionTabsScrollLeft = track.scrollLeft
    appState.decisionTabsCanScrollLeft = track.scrollLeft > 4
    appState.decisionTabsCanScrollRight = track.scrollLeft < maxScrollLeft - 4
  }

  function queueDecisionTabsSync() {
    requestAnimationFrame(() => {
      syncDecisionTabsState()
    })
  }

  function setActiveDocument(documentId) {
    appState.activeDocumentId = documentId
    const cachedAiResult = appState.aiRerankCacheByDocumentId[documentId] || null
    appState.aiRerankResult = cloneAiRerankResult(cachedAiResult)
    appState.aiRerankDocumentId = cachedAiResult ? documentId : ""
    appState.aiRerankError = ""
    appState.summaryPanelOpen = false
    appState.decisionAssistResult = null
    appState.decisionAssistError = ""
    appState.productFirstSelection = null
    clearSelectionState()
    runRanking(true)
    if (cachedAiResult?.bestPage) {
      appState.activePageNumber = cachedAiResult.bestPage
    } else if (appState.rankedPages[0]?.pageNumber) {
      appState.activePageNumber = appState.rankedPages[0].pageNumber
    }
    updateStructureRoutingState(getActiveDocument())
    render()
    requestAnimationFrame(() => setPage(appState.activePageNumber))
  }

  function selectionSummary() {
    return normalizeText(appState.selectionText).slice(0, 140)
  }

  function getVisiblePages(documentRecord) {
    if (!documentRecord) return []
    const activeIndex = Math.max(0, documentRecord.pages.findIndex((page) => page.pageNumber === appState.activePageNumber))
    const start = activeIndex
    const end = Math.min(documentRecord.pages.length, activeIndex + 2)
    return documentRecord.pages.slice(start, end)
  }

  function buildPageClusters(rankedPages) {
    const sortedByPage = [...rankedPages]
      .filter((page) => page.score > 0)
      .sort((a, b) => a.pageNumber - b.pageNumber)

    const clusters = []
    sortedByPage.forEach((page) => {
      const previous = clusters[clusters.length - 1]
      if (!previous || page.pageNumber - previous.endPage > 2) {
        clusters.push({
          startPage: page.pageNumber,
          endPage: page.pageNumber,
          pages: [page],
          totalScore: page.score
        })
        return
      }

      previous.endPage = page.pageNumber
      previous.pages.push(page)
      previous.totalScore += page.score
    })

    return clusters
      .map((cluster) => ({
        ...cluster,
        averageScore: Math.round(cluster.totalScore / cluster.pages.length)
      }))
      .sort((a, b) => b.totalScore - a.totalScore)
  }

  function formatPageRange(startPage, endPage) {
    return startPage === endPage ? `page ${startPage}` : `pages ${startPage}-${endPage}`
  }

  function getPageTextMetrics(documentRecord, pageNumber) {
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    const text = getPageCombinedText(page)
    const normalized = normalizeText(text)
    const words = normalized ? normalized.split(/\s+/).length : 0
    const lines = text
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean).length

    return {
      pageNumber,
      characters: normalized.length,
      words,
      lines
    }
  }

  function titleCaseWords(value) {
    return normalizeText(value)
      .split(/\s+/)
      .map((word) => {
        if (!word) return ""
        if (/^[A-Z0-9]+$/.test(word)) return word
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
      })
      .join(" ")
  }

  function getCompactDescriptorLabel(value) {
    const normalized = normalizeText(value)
    if (!normalized) return ""

    const product = normalizeText(appState.spec.specDisplayName || appState.spec.originalSpecName || "")
    let shortened = normalized

    if (product) {
      shortened = normalizeText(
        shortened.replace(new RegExp(product.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i"), "")
      ).replace(/^[,:\-–—]\s*/, "")
    }

    if (!shortened) return "General"

    const directTerms = [
      "general",
      "ebony",
      "white ash",
      "white oak",
      "walnut",
      "palisander",
      "classic",
      "tall",
      "standard"
    ]
    const directMatch = directTerms.find((term) => shortened.toLowerCase() === term)
    if (directMatch) return titleCaseWords(directMatch)

    const commaTail = normalized.split(",").slice(1).join(",").trim()
    if (commaTail && commaTail.length <= 24) {
      return titleCaseWords(commaTail)
    }

    const specificMatch = normalized.match(/\b(Ebony|White Ash|White Oak|Walnut|Palisander|Classic|Tall|Standard|General)\b/i)
    if (specificMatch) return titleCaseWords(specificMatch[1])

    if (shortened.length <= 24) return titleCaseWords(shortened)

    return titleCaseWords(shortened.split(/\s+/).slice(0, 2).join(" "))
  }

  function isLikelyFileName(value) {
    const normalized = normalizeText(value)
    return /\.pdf$/i.test(normalized) || normalized.includes("_")
  }

  function getPageTitleCandidates(documentRecord, pageNumber) {
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    const text = getPageCombinedText(page)
    const brand = normalizeText(appState.spec.brandDisplayName || "")
    const product = normalizeText(appState.spec.specDisplayName || appState.spec.originalSpecName || "")
    const productTokens = product
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter((token) => token.length >= 3)

    return text
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean)
      .filter((line) => {
        const lowered = line.toLowerCase()
        if (brand && lowered === brand.toLowerCase()) return false
        if (/^page\s+\d+/i.test(line)) return false
        if (/^[\d\s."'/,-]+$/.test(line)) return false
        if (line.length < 8 || line.length > 120) return false
        if (/^[A-Z]{1,4}\d{2,}[A-Z0-9\s!/-]*$/.test(line)) return false
        if (/^[A-Z0-9!/-]{4,16}$/.test(line)) return false
        const tokenMatches = productTokens.filter((token) => lowered.includes(token)).length
        if (tokenMatches === 0 && line.length < 18) return false
        return true
      })
      .sort((a, b) => {
        const aHasProduct = product && a.toLowerCase().includes(product.toLowerCase())
        const bHasProduct = product && b.toLowerCase().includes(product.toLowerCase())
        const aTokenMatches = productTokens.filter((token) => a.toLowerCase().includes(token)).length
        const bTokenMatches = productTokens.filter((token) => b.toLowerCase().includes(token)).length
        if (aHasProduct && !bHasProduct) return -1
        if (!aHasProduct && bHasProduct) return 1
        if (aTokenMatches !== bTokenMatches) return bTokenMatches - aTokenMatches
        return b.length - a.length
      })
      .slice(0, 6)
  }

  function getPrimaryHeaderTitle(documentRecord, pageNumber) {
    const candidates = getPageTitleCandidates(documentRecord, pageNumber)
    const product = normalizeText(appState.spec.specDisplayName || appState.spec.originalSpecName || "")
    const productTokens = product
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter((token) => token.length >= 4)

    const strongCandidate = candidates.find((line) => {
      const lowered = line.toLowerCase()
      const tokenMatches = productTokens.filter((token) => lowered.includes(token)).length
      return tokenMatches >= Math.min(2, productTokens.length || 2)
    })

    if (strongCandidate) return strongCandidate

    const commaCandidate = candidates.find((line) => /,/.test(line) && line.length >= 16)
    if (commaCandidate) return commaCandidate

    const longCandidate = candidates.find((line) => line.length >= 18)
    return longCandidate || candidates[0] || ""
  }

  function getSharedTitlePrefix(titles) {
    const normalizedTitles = titles.map((title) => normalizeText(title)).filter(Boolean)
    if (!normalizedTitles.length) return ""

    const tokenLists = normalizedTitles.map((title) => title.split(/\s+/))
    const shortest = tokenLists.reduce((best, current) => current.length < best.length ? current : best, tokenLists[0])
    const sharedTokens = []

    for (let index = 0; index < shortest.length; index += 1) {
      const token = shortest[index]
      const normalizedToken = token.toLowerCase().replace(/[^a-z0-9]+/g, "")
      const everyMatches = tokenLists.every((list) => {
        const candidate = list[index] || ""
        return candidate.toLowerCase().replace(/[^a-z0-9]+/g, "") === normalizedToken
      })
      if (!everyMatches) break
      sharedTokens.push(token)
    }

    return normalizeText(sharedTokens.join(" "))
  }

  function getTitleBasedDescriptor(documentRecord, pageNumber, peerPages) {
    const currentTitle = getPrimaryHeaderTitle(documentRecord, pageNumber)
    if (!currentTitle) return ""

    const titles = [...new Set((peerPages || []).map((candidate) => getPrimaryHeaderTitle(documentRecord, candidate.pageNumber)).filter(Boolean))]
    if (titles.length <= 1) return ""

    const sharedPrefix = getSharedTitlePrefix(titles)
    if (!sharedPrefix) return ""

    const prefixPattern = new RegExp(`^${sharedPrefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "i")
    const strippedCurrent = normalizeText(currentTitle.replace(prefixPattern, "")).replace(/^[,:\-–—]\s*/, "")

    const strippedPeers = titles
      .map((title) => normalizeText(title.replace(prefixPattern, "")))
      .map((value) => value.replace(/^[,:\-–—]\s*/, ""))

    const nonEmptyPeerSuffixes = strippedPeers.filter(Boolean)
    if (!strippedCurrent) {
      if (nonEmptyPeerSuffixes.length) return "General"
      return ""
    }

    if (nonEmptyPeerSuffixes.includes(strippedCurrent)) {
      return getCompactDescriptorLabel(strippedCurrent)
    }

    return getCompactDescriptorLabel(strippedCurrent)
  }

  function extractPageThemeSignals(text, documentRecord, pageNumber) {
    const lowered = text.toLowerCase()
    const pageTitle = normalizeText(documentRecord?.pageTitles?.[pageNumber - 1] || "")
    const titleLowered = pageTitle.toLowerCase()
    const pageTitleUsable = pageTitle && !isLikelyFileName(pageTitle) && !/^\s*page\s+\d+/i.test(pageTitle)

    const finishMatches = [
      { pattern: /white oak|whn white oak/i, label: "White Oak Finish" },
      { pattern: /ebony|ebonized ash/i, label: "Ebony Finish" },
      { pattern: /walnut|ou walnut|oiled walnut/i, label: "Walnut Finish" },
      { pattern: /ash\b/i, label: "Ash Finish" },
      { pattern: /santos palisander|palisander/i, label: "Palisander Finish" }
    ]

    const matchedFinishLabels = [...new Set(
      finishMatches
        .filter((item) => item.pattern.test(text))
        .map((item) => item.label)
    )]
    return {
      multiFinish: matchedFinishLabels.length > 1,
      finishLabels: matchedFinishLabels,
      hasUpholstery: /price category|upholstery|fabric|leather|col\b|com\b/i.test(lowered),
      hasFinish: /finish options?|finishes|veneer|wood finish|shell finish|base finish/i.test(lowered),
      hasDimensions: /dimension|overall width|overall depth|overall height|seat height|arm height|sizes?/i.test(lowered),
      hasSpecs: /specifications?|spec sheet|materials?/i.test(lowered),
      titleLabel:
        pageTitleUsable && /dimension|size/i.test(titleLowered) ? "Dimensions & Sizes" :
        pageTitleUsable && /finish|veneer|wood|oak|walnut|ash|ebony/i.test(titleLowered) ? "Finish Options" :
        pageTitleUsable && /material|upholstery|fabric|leather/i.test(titleLowered) ? "Material (Upholstery)" :
        pageTitleUsable ? titleCaseWords(pageTitle) : ""
    }
  }

  function getRelativePageDescriptor(documentRecord, pageNumber, pageData, peerPages) {
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    const text = getPageCombinedText(page)
    const signals = extractPageThemeSignals(text, documentRecord, pageNumber)
    const peerSignals = (peerPages || [])
      .filter((candidate) => candidate.pageNumber !== pageNumber)
      .map((candidate) => {
        const peerPage = documentRecord?.pages.find((item) => item.pageNumber === candidate.pageNumber)
        return extractPageThemeSignals(getPageCombinedText(peerPage), documentRecord, candidate.pageNumber)
      })

    const titleDescriptor = getTitleBasedDescriptor(documentRecord, pageNumber, peerPages)

    if (pageData?.aiVariant) return getCompactDescriptorLabel(pageData.aiVariant)

    const aiDifference = normalizeText(pageData?.aiDifference || "")
    if (aiDifference) {
      if (signals.multiFinish && /ebony|ebonized|white oak|white ash|ash|walnut|palisander/i.test(aiDifference)) {
        return "General Finish Options"
      }
      if (/ebony|ebonized/i.test(aiDifference)) return "Ebony Finish"
      if (/white oak|white ash|ash/i.test(aiDifference)) return "White Ash Finish"
      if (/walnut/i.test(aiDifference)) return "Walnut Finish"
      if (/palisander/i.test(aiDifference) && !signals.multiFinish) return "Palisander Finish"
      if (/dimension|height|width|depth|size/i.test(aiDifference)) return "Dimensions & Sizes"
    }

    const aiReason = normalizeText(pageData?.aiReason || "")
    if (aiReason) {
      if (/dimension|drawing|size|measurement/i.test(aiReason)) return "Dimensions & Sizes"
      if (/finish|veneer|oak|ash|walnut|ebony/i.test(aiReason)) {
        return signals.multiFinish ? "General Finish Options" : "Finish Option"
      }
      if (/material|upholstery|fabric|leather/i.test(aiReason)) return "Material (Upholstery)"
      if (/spec/i.test(aiReason)) return "Material Specifications"
    }

    if (pageData?.aiRole) {
      if (/dimension|size/i.test(pageData.aiRole)) return "Dimensions & Sizes"
      if (/finish|veneer/i.test(pageData.aiRole)) return signals.multiFinish ? "General Finish Options" : "Finish Option"
      if (/material|upholstery/i.test(pageData.aiRole)) return "Material (Upholstery)"
      return titleCaseWords(pageData.aiRole)
    }

    if (titleDescriptor) return titleDescriptor
    if (signals.hasDimensions) return "Dimensions & Sizes"
    if (signals.hasUpholstery) return "Material (Upholstery)"

    const peerHasSpecificFinish = peerSignals.some((peer) => peer.finishLabels.length === 1 && !peer.multiFinish)
    if (signals.multiFinish) return peerHasSpecificFinish ? "General Finish Options" : "Finish Options"
    if (signals.finishLabels.length === 1 || signals.hasFinish) {
      return peerHasSpecificFinish ? "General Finish Options" : "Finish Options"
    }
    if (signals.hasSpecs) return "Material Specifications"
    if (signals.titleLabel) return signals.titleLabel

    return analyzePageForVisualCues(text).primaryLabel
  }

  function getDistinctTopPageLabels(documentRecord, pages) {
    const sharedTitles = pages.map((page) => getPrimaryHeaderTitle(documentRecord, page.pageNumber))
    const sharedPrefix = getSharedTitlePrefix(sharedTitles)
    const prefixPattern = sharedPrefix
      ? new RegExp(`^${sharedPrefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "i")
      : null

    const firstPass = pages.map((page) => {
      const aiLabel = getCompactDescriptorLabel(page.aiVariant || "")
      if (aiLabel) return aiLabel

      const title = getPrimaryHeaderTitle(documentRecord, page.pageNumber)
      if (title && prefixPattern) {
        const suffix = normalizeText(title.replace(prefixPattern, "")).replace(/^[,:\-–—]\s*/, "")
        const compactSuffix = getCompactDescriptorLabel(suffix)
        if (compactSuffix) return compactSuffix
        if (!suffix) return "General"
      }

      return getRelativePageDescriptor(documentRecord, page.pageNumber, page, pages)
    })

    const counts = new Map()
    firstPass.forEach((label) => counts.set(label, (counts.get(label) || 0) + 1))

    return pages.map((page, index) => {
      const initialLabel = firstPass[index]
      if ((counts.get(initialLabel) || 0) === 1) return initialLabel

      const candidates = getPageTitleCandidates(documentRecord, page.pageNumber)
      const title = getPrimaryHeaderTitle(documentRecord, page.pageNumber)

      const extraCandidates = [
        ...candidates.map((line) => getCompactDescriptorLabel(line)),
        getCompactDescriptorLabel(title),
        getCompactDescriptorLabel(page.aiDifference || ""),
        getCompactDescriptorLabel(page.aiReason || "")
      ]
        .map((label) => normalizeText(label))
        .filter(Boolean)
        .filter((label) => label.toLowerCase() !== initialLabel.toLowerCase())

      const uniqueAlternative = extraCandidates.find((candidate) =>
        !firstPass.some((label, labelIndex) => labelIndex !== index && label.toLowerCase() === candidate.toLowerCase())
      )
      if (uniqueAlternative) return uniqueAlternative

      if (initialLabel.toLowerCase() === "general") {
        return index === 0 ? "General" : `Variant ${index + 1}`
      }

      return `${initialLabel} ${index + 1}`
    })
  }

  function getAiOnlyTopPageLabel(page) {
    const rawCandidates = [
      page?.aiVariant || "",
      page?.aiDifference || "",
      page?.aiReason || "",
      page?.aiRole || ""
    ]

    for (const rawValue of rawCandidates) {
      const compact = getCompactDescriptorLabel(rawValue)
      if (compact) return compact

      const normalized = normalizeText(rawValue)
      if (!normalized) continue

      const product = normalizeText(appState.spec.specDisplayName || appState.spec.originalSpecName || "")
      let shortened = normalized
      if (product) {
        shortened = normalizeText(
          shortened.replace(new RegExp(product.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i"), "")
        ).replace(/^[,:\-–—]\s*/, "")
      }

      const fallback = normalizeText(shortened || normalized)
      if (!fallback) continue

      const commaTail = fallback.split(",").slice(1).join(",").trim()
      if (commaTail) return titleCaseWords(commaTail.split(/\s+/).slice(0, 3).join(" "))

      return titleCaseWords(fallback.split(/\s+/).slice(0, 3).join(" "))
    }

    return ""
  }

  function getSelectedPageSubtypeHint(documentRecord, selectedPage) {
    if (!selectedPage) return ""

    const aiOrderedPages = getAiOrderedPages(appState.rankedPages.slice(0, getAiRerankCandidateLimit()).map((page) => ({
      ...page,
      metrics: getPageTextMetrics(documentRecord, page.pageNumber)
    })))
    const aiPage = aiOrderedPages.find((page) => page.pageNumber === selectedPage.pageNumber)
    const aiLabel = aiPage ? getAiOnlyTopPageLabel(aiPage) : ""
    if (aiLabel) return aiLabel

    const rankedPage = appState.rankedPages.find((page) => page.pageNumber === selectedPage.pageNumber)
    if (rankedPage) {
      const fallbackLabel = getRelativePageDescriptor(documentRecord, selectedPage.pageNumber, rankedPage, appState.rankedPages.slice(0, 5))
      if (fallbackLabel) return fallbackLabel
    }

    return ""
  }

  function getFamilyPageSelectionCandidates(documentRecord, pageNumber) {
    if (getSpecParsingMode() !== "family") return []
    const selectedPage = documentRecord?.pages?.find((page) => page.pageNumber === pageNumber)
    if (!selectedPage) return []
    const subtypeHint = getSelectedPageSubtypeHint(documentRecord, selectedPage)
    return extractConcreteVariantCandidatesFromPageText(getPageCombinedText(selectedPage), subtypeHint)
  }

  function getFamilyPageSelectionKey(documentRecord, pageNumber) {
    return `${documentRecord?.id || "unknown"}:${pageNumber}`
  }

  function getPageRetrievalExplanation(documentRecord, pageNumber) {
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    const text = getPageCombinedText(page)
    const lowered = text.toLowerCase()
    const signals = [
      { label: "Product phrase", value: appState.spec.specDisplayName, weight: 18 },
      { label: "Product phrase", value: appState.spec.originalSpecName, weight: 16 }
    ]
    const ladder = buildProductMatchLadder()
    const matches = []

    signals.forEach((signal) => {
      const phrase = normalizeText(signal.value).toLowerCase()
      if (!phrase) return
      const count = countOccurrences(lowered, phrase)
      if (!count) return
      matches.push({
        label: signal.label,
        phrase: signal.value,
        count,
        effect: `+${signal.weight}${count > 1 ? ` x${count}` : ""}`
      })
    })

    ladder.forEach((entry) => {
      const count = countOccurrences(lowered, entry.phrase)
      if (!count) return
      matches.push({
        label: entry.label,
        phrase: entry.phrase,
        count,
        effect: `+${entry.weight}${count > 1 ? ` x${count}` : ""}`
      })
    })

    const metrics = getPageTextMetrics(documentRecord, pageNumber)
    return {
      metrics,
      matches
    }
  }

  function shouldOfferDimensionAnalysis(documentRecord, pageNumber) {
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    if (!page) return false

    const cues = analyzePageForVisualCues(getPageCombinedText(page))
    return cues.labels.includes("Likely dimensions drawing") || cues.labels.includes("Spec content detected")
  }

  function countOccurrences(haystack, phrase) {
    if (!phrase) return 0
    const escaped = phrase.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    const regex = new RegExp(`\\b${escaped}\\b`, "gi")
    return [...haystack.matchAll(regex)].length
  }

  function getActiveSourceSummary() {
    if (appState.uploadFiles.length) {
      const names = appState.uploadFiles
        .slice(0, 2)
        .map((file) => file.name)
        .join(", ")
      const moreCount = Math.max(0, appState.uploadFiles.length - 2)
      const namesSuffix = moreCount ? `, +${moreCount} more` : ""
      return `Pending uploads: ${appState.uploadFiles.length} PDF${appState.uploadFiles.length === 1 ? "" : "s"} selected (${names}${namesSuffix}). Click "Ask AI to Analyze PDFs".`
    }

    const activeDocument = getActiveDocument()
    if (!activeDocument) {
      return "No active source loaded."
    }

    if (activeDocument.title === DEFAULT_BUNDLED_PDF_NAME) {
      return `Active source: ${activeDocument.title} (bundled default)`
    }

    if (activeDocument.sourceType === "uploaded") {
      return `Active source: ${activeDocument.title} (uploaded or saved default)`
    }

    return `Active source: ${activeDocument.title} (sample)`
  }

  function getDisplayScore(page) {
    return Number.isFinite(page.aiScore) ? page.aiScore : page.score
  }

  function formatUiScore(value) {
    if (!Number.isFinite(value)) return "Pending"
    const normalized = Math.max(0, Math.min(100, Math.round(value)))
    return `${normalized}/100`
  }

  function formatPageChipScore(page, topPages = []) {
    if (!topPages.length) {
      return formatUiScore(getDisplayScore(page))
    }
    const topScore = getDisplayScore(topPages[0])
    const currentScore = getDisplayScore(page)
    if (!Number.isFinite(topScore) || !Number.isFinite(currentScore)) {
      return formatUiScore(currentScore)
    }
    if (page?.pageNumber === topPages[0]?.pageNumber) return "Best"
    const diff = topScore - currentScore
    if (diff <= 2) return "Close"
    if (diff <= 8) return "Related"
    return "Weak"
  }

  function buildTopPageTiers(topPages) {
    if (!topPages.length) {
      return []
    }

    const topScore = getDisplayScore(topPages[0])
    const primary = topPages.filter((page) => getDisplayScore(page) >= topScore - 2)
    const remaining = topPages.filter((page) => !primary.includes(page))
    const secondary = remaining.filter((page) => getDisplayScore(page) >= 40)
    const lowConfidence = remaining.filter((page) => !secondary.includes(page))

    return [
      { key: "primary", label: "Best matches", pages: primary },
      { key: "secondary", label: "Related", pages: secondary },
      { key: "low", label: "Low confidence", pages: lowConfidence }
    ].filter((group) => group.pages.length)
  }

  function render() {
    const documentRecord = getActiveDocument()
    if (documentRecord && !appState.rankedPages.length) {
      runRanking(true)
    }

    scheduleVisiblePageRendering()

    const topMatch = appState.rankedPages[0]
    const hasAiForActiveDocument = Boolean(appState.aiRerankResult && appState.aiRerankDocumentId === documentRecord?.id)
    const topPages = getAiOrderedPages(appState.rankedPages.slice(0, getAiRerankCandidateLimit()).map((page) => ({
      ...page,
      metrics: getPageTextMetrics(documentRecord, page.pageNumber)
    })))
    const sourceScoreById = new Map(
      appState.sourceSelectionScores.map((item) => [item.documentId, item.selectionScore])
    )
    const activeDocumentSourceScore = Number.isFinite(sourceScoreById.get(documentRecord?.id))
      ? sourceScoreById.get(documentRecord?.id)
      : null
    const decisionResult = appState.decisionAssistResult
    const parsingModeConfig = getSpecParsingModeConfig(decisionResult?.parsingMode || "")
    const isFamilyMode = getSpecParsingMode() === "family"
    const selectedProductId = normalizeText(decisionResult?.selectedProductId || "")
    const selectedVariantId = normalizeText(decisionResult?.selectedVariantId || "")
    const showProductSelector = (decisionResult?.productCandidates?.length || 0) > 1 && !selectedProductId
    const showVariantSelector = !showProductSelector && (decisionResult?.variantCandidates?.length || 0) > 1 && !selectedVariantId
    const decisionCharacteristics = !showProductSelector && !showVariantSelector ? (decisionResult?.characteristics || []) : []
    const canShowDecisionTabNav = decisionCharacteristics.length > 3
    const maxDecisionTabsWindowStart = Math.max(0, decisionCharacteristics.length - 3)
    const decisionTabsWindowStart = Math.min(appState.decisionTabsWindowStart, maxDecisionTabsWindowStart)
    const visibleDecisionChoices = canShowDecisionTabNav
      ? decisionCharacteristics.slice(decisionTabsWindowStart, decisionTabsWindowStart + 3).map((choice, offset) => ({
          choice,
          index: decisionTabsWindowStart + offset
        }))
      : decisionCharacteristics.map((choice, index) => ({ choice, index }))
    const showDecisionContext = false
    const activeDecisionChoice =
      appState.activeDecisionChoiceIndex >= 0 && appState.activeDecisionChoiceIndex < decisionCharacteristics.length
        ? decisionCharacteristics[appState.activeDecisionChoiceIndex]
        : null
    const pageExplanation = appState.wordStatsPage ? getPageRetrievalExplanation(documentRecord, appState.wordStatsPage) : null
    const viewerPageLimit = getSpecParsingMode() === "family" ? getFamilyResultPageLimit() : 3
    const viewerTopPages = (hasAiForActiveDocument ? topPages.slice(0, viewerPageLimit) : appState.rankedPages.slice(0, viewerPageLimit).map((page) => ({
      ...page,
      metrics: getPageTextMetrics(documentRecord, page.pageNumber)
    })))
      .filter(Boolean)
    const topPageLabels = documentRecord
      ? hasAiForActiveDocument
        ? viewerTopPages.map((page) => getAiOnlyTopPageLabel(page))
        : []
      : []
    const waitingOnAiLabels = Boolean(
      documentRecord &&
      appState.visionApiKey &&
      (appState.analyzeRequestLoading || appState.aiRerankLoading) &&
      !hasAiForActiveDocument
    )
    const gateViewerUntilAiReady = waitingOnAiLabels
    const productFirstSessionActive = isProductFirstSessionActive()
    const canShowPageStripNav = viewerTopPages.length > 3
    const showHelpSpecPanel = Boolean(
      documentRecord &&
      (appState.decisionAssistLoading || decisionResult || appState.errorMessage || appState.aiRerankError || appState.decisionAssistError)
    )
    const structureRouting = appState.structureRouting || buildStructureRoutingState(documentRecord)
    const selectedProductDisplayLabel = getSelectedProductDisplayLabel(decisionResult, structureRouting)
    const showProductFirstEntry = Boolean(
      documentRecord
      && !gateViewerUntilAiReady
      && isFamilyMode
      && structureRouting.interactionModel === "product_first"
      && structureRouting.productCandidates.length > 1
      && !productFirstSessionActive
      && !decisionResult
      && !appState.decisionAssistLoading
    )
    const familyPageSelectionKey = documentRecord ? getFamilyPageSelectionKey(documentRecord, appState.activePageNumber) : ""
    const activeFamilyPageSelections = familyPageSelectionKey ? (appState.familyPageProductsByKey[familyPageSelectionKey] || []) : []
    const activeFamilyPageSelectionStatus = familyPageSelectionKey ? (appState.familyPageProductsStatusByKey[familyPageSelectionKey] || "") : ""
    const activeFamilyPageSelectionError = familyPageSelectionKey ? (appState.familyPageProductsErrorByKey[familyPageSelectionKey] || "") : ""
    const statusScore = Number.isFinite(activeDocumentSourceScore)
      ? activeDocumentSourceScore
      : topMatch
        ? Math.min(99, Math.max(12, Math.round(getDisplayScore(topMatch))))
        : null
    const baseRenderWidth = getViewerBaseRenderWidth()

    app.innerHTML = `
      <div class="app-shell">
        <section class="session-bar">
          <div class="session-bar-title">
            <span class="session-bar-kicker">DEV:</span>
            <span class="session-bar-name">Session Configuration</span>
          </div>
          <div class="session-bar-controls">
            <div class="session-control session-control-file">
              <label for="pdf-upload">PDF File(s)</label>
              <input id="pdf-upload" type="file" accept="application/pdf,.pdf" multiple />
              <span class="session-control-meta">${escapeHtml(getActiveSourceSummary())}</span>
            </div>
            <div class="session-control">
              <label for="product-name">Product Name</label>
              <input id="product-name" data-draft-input="productName" value="${escapeHtml(appState.inputDraft.productName)}" />
            </div>
            <div class="session-control">
              <label for="vision-api-key">OpenAI API Key</label>
              <input id="vision-api-key" type="password" value="${escapeHtml(appState.visionApiKey)}" placeholder="sk-..." />
            </div>
            <div class="session-submit">
              <button class="session-submit-btn" id="session-submit-btn" type="button" ${appState.analyzeRequestLoading ? "disabled" : ""}>
                ${appState.analyzeRequestLoading || appState.loadingMessage ? "Working..." : "Submit"}
              </button>
            </div>
          </div>
          ${appState.errorMessage ? `<p class="session-status session-status-error">${escapeHtml(appState.errorMessage)}</p>` : ""}
        </section>

        <section class="workspace-layout">
          <div class="viewer-column">
            ${documentRecord
              ? `
                <div class="workspace-header">
                  <div class="workspace-header-main">
                    <div class="workspace-doc-pill">
                      <select class="workspace-doc-select" id="document-select" ${appState.documents.length ? "" : "disabled"}>
                        ${appState.documents
                          .map((document) => {
                            const sourceScore = sourceScoreById.get(document.id)
                            const prefix = Number.isFinite(sourceScore) ? `${sourceScore} · ` : ""
                            return `
                              <option value="${document.id}" ${document.id === appState.activeDocumentId ? "selected" : ""}>
                                ${escapeHtml(`${prefix}${document.title}`)}
                              </option>
                            `
                          })
                          .join("")}
                      </select>
                    </div>
                    <div class="workspace-status">
                      <span class="confidence-dot confidence-dot-single" aria-hidden="true"></span>
                      <span class="workspace-score">${formatUiScore(statusScore)}</span>
                    </div>
                    <div class="viewer-zoom-controls" aria-label="PDF zoom controls">
                      <button class="viewer-zoom-btn" id="zoom-out-btn" type="button" aria-label="Zoom out">-</button>
                      <button class="viewer-zoom-btn viewer-zoom-reset" id="zoom-reset-btn" type="button" aria-label="Reset zoom">${getPdfZoomLabel()}</button>
                      <button class="viewer-zoom-btn" id="zoom-in-btn" type="button" aria-label="Zoom in">+</button>
                    </div>
                  </div>
                </div>
              `
              : ""}

            ${documentRecord && !gateViewerUntilAiReady && !showProductFirstEntry && !productFirstSessionActive
              ? `
                <div class="page-strip ${viewerTopPages.length <= 3 ? "page-strip-centered" : ""}">
                  ${canShowPageStripNav && appState.pageStripCanScrollLeft ? '<button class="page-strip-nav page-strip-nav-left" id="page-strip-nav-left" type="button" aria-label="Previous page results">‹</button>' : ""}
                  <div class="page-strip-track" id="page-strip-track">
                    ${
                      viewerTopPages.length
                        ? viewerTopPages
                            .map((page) => {
                              const isActive = page.pageNumber === appState.activePageNumber
                              const label =
                                topPageLabels[viewerTopPages.findIndex((candidate) => candidate.pageNumber === page.pageNumber)]
                                || (waitingOnAiLabels ? "Finding distinction..." : "")
                              return `
                                <div class="page-chip-shell ${isActive ? "is-active" : ""}">
                                  <button class="page-chip ${isActive ? "is-active" : ""}" data-jump-page="${page.pageNumber}" type="button">
                                    <div class="page-chip-topline">
                                      <span class="page-chip-number">Page ${page.pageNumber}</span>
                                      <span class="page-chip-score">${escapeHtml(formatPageChipScore(page, viewerTopPages))}</span>
                                    </div>
                                    ${label ? `<span class="page-chip-label">${escapeHtml(label)}</span>` : ""}
                                  </button>
                                  ${
                                    isActive
                                      ? (isFamilyMode ? "" : `<button class="page-chip-ai-btn" data-help-spec-page="${page.pageNumber}" type="button" aria-label="Help me spec this page">✧</button>`)
                                      : ""
                                  }
                                </div>
                              `
                            })
                            .join("")
                        : `<div class="page-strip-empty">No ranked pages available yet.</div>`
                    }
                  </div>
                  ${canShowPageStripNav && appState.pageStripCanScrollRight ? '<button class="page-strip-nav page-strip-nav-right" id="page-strip-nav-right" type="button" aria-label="More page results">›</button>' : ""}
                </div>
              `
              : ""}

            ${
              showProductFirstEntry
                ? `
                  <div class="help-spec-panel">
                    <div class="help-spec-card">
                      <div class="help-spec-selector-block">
                        <div class="help-spec-selector-head">
                          <p class="help-spec-selector-title">Select the product</p>
                          <p class="help-spec-selector-copy">These product options are grounded in concrete visible labels or groups on the top strong or related family page${getTopRerankedPageNumbers(appState.aiRerankResult).length > 1 ? "s" : ""}. Choose one to start Help Me Spec without clicking a page first.</p>
                        </div>
                        <div class="help-spec-selector-grid">
                          ${structureRouting.productCandidates
                            .map((candidate) => `
                              <button class="help-spec-selector-card" data-product-first-product="${escapeHtmlAttribute(candidate.id)}" type="button">
                                <strong>${escapeHtml(getProductCardTitle(candidate))}</strong>
                                ${shouldShowSeparateProductCardModelCode(candidate) ? `<span>${escapeHtml(getProductCardModelCode(candidate))}</span>` : ""}
                                ${getDisplayProductCardSubtitle(candidate) ? `<em>${escapeHtml(getDisplayProductCardSubtitle(candidate))}</em>` : ""}
                              </button>
                            `)
                            .join("")}
                        </div>
                      </div>
                    </div>
                  </div>
                `
                : ""
            }

            ${documentRecord && gateViewerUntilAiReady ? `<div class="page-strip"></div>` : ""}

            ${
              documentRecord && isFamilyMode && !showProductFirstEntry && !productFirstSessionActive && !gateViewerUntilAiReady && (Boolean(activeFamilyPageSelections.length) || activeFamilyPageSelectionStatus === "loading" || activeFamilyPageSelectionStatus === "error")
                ? `
                  <div class="help-spec-panel">
                    <div class="help-spec-card">
                      <div class="help-spec-selector-block">
                        <div class="help-spec-selector-head">
                          <p class="help-spec-selector-title">Select the product on page ${appState.activePageNumber}</p>
                          <p class="help-spec-selector-copy">Choose the exact product row shown on this page. Help Me Spec will start only after this selection, so the system does not need to infer the product from the whole family.</p>
                        </div>
                        ${
                          activeFamilyPageSelectionStatus === "loading"
                            ? `<p class="help-spec-loading-inline">${escapeHtml("Reading visible product rows on this page...")}</p>`
                            : activeFamilyPageSelectionStatus === "error"
                              ? `<p class="inline-ai-error">${escapeHtml(activeFamilyPageSelectionError || "Unable to detect the visible products on this page.")}</p>`
                              : `
                                <div class="help-spec-selector-grid">
                                  ${activeFamilyPageSelections
                                    .map((candidate) => `
                                      <button class="help-spec-selector-card" data-help-spec-page-product="${escapeHtmlAttribute(candidate.id)}" type="button">
                                        <strong>${escapeHtml(candidate.label)}</strong>
                                        ${candidate.description ? `<span>${escapeHtml(candidate.description)}</span>` : ""}
                                      </button>
                                    `)
                                    .join("")}
                                </div>
                              `
                        }
                      </div>
                    </div>
                  </div>
                `
                : ""
            }

            ${showHelpSpecPanel
              ? `
                <div class="help-spec-panel">
                  ${
                    appState.decisionAssistLoading
                      ? `
                        <div class="help-spec-card help-spec-card-loading">
                          <div class="loading-orb" aria-hidden="true"></div>
                          <p class="help-spec-loading-inline">${escapeHtml(getHelpSpecLoadingLine())}</p>
                        </div>
                      `
                      : decisionResult
                      ? `
                        <div class="help-spec-card">
                          ${
                            decisionResult.fromProductFirst && selectedProductDisplayLabel
                              ? `
                                <div class="help-spec-selection-lock">
                                  <div>
                                    <p class="help-spec-selection-label">Selected product</p>
                                    <strong>${escapeHtml(selectedProductDisplayLabel)}</strong>
                                  </div>
                                  <button class="ghost-btn" data-change-product-selection type="button">Change selection</button>
                                </div>
                              `
                              : ""
                          }
                          ${
                            showDecisionContext
                              ? `
                                <div class="help-spec-context">
                                  <div class="help-spec-context-head">
                                    <div>
                                      <p class="help-spec-eyebrow">Help Me Spec</p>
                                      <h3 class="help-spec-title">${escapeHtml(decisionResult.productName || "Product under review")}</h3>
                                      ${
                                        decisionResult.modelFamily
                                          ? `<p class="help-spec-subtitle">${escapeHtml(decisionResult.modelFamily)}</p>`
                                          : ""
                                      }
                                    </div>
                                    ${
                                      decisionResult.pageRangeLabel
                                        ? `<span class="help-spec-page-range">Pages ${escapeHtml(decisionResult.pageRangeLabel)}</span>`
                                        : ""
                                    }
                                  </div>
                                  ${
                                    decisionResult.summary
                                      ? `<p class="help-spec-summary">${escapeHtml(decisionResult.summary)}</p>`
                                      : ""
                                  }
                                  <div class="help-spec-trust-grid">
                                    ${
                                      decisionResult.includedPages?.length
                                        ? `
                                          <div class="help-spec-trust-item">
                                            <span>Pages used</span>
                                            <strong>${escapeHtml(formatPageRangeLabel(decisionResult.includedPages))}</strong>
                                          </div>
                                        `
                                        : ""
                                    }
                                    ${
                                      decisionResult.inclusionSummary
                                        ? `
                                          <div class="help-spec-trust-item">
                                            <span>Why included</span>
                                            <strong>${escapeHtml(decisionResult.inclusionSummary)}</strong>
                                          </div>
                                        `
                                        : ""
                                    }
                                    ${
                                      decisionResult.stopReason
                                        ? `
                                          <div class="help-spec-trust-item">
                                            <span>Where it stopped</span>
                                            <strong>${escapeHtml(decisionResult.stopReason)}</strong>
                                          </div>
                                        `
                                        : ""
                                    }
                                  </div>
                                </div>
                              `
                              : ""
                          }

                          ${
                            showProductSelector
                              ? `
                                <div class="help-spec-selector-block">
                                  <div class="help-spec-selector-head">
                                    <p class="help-spec-selector-title">${escapeHtml(parsingModeConfig.selectorProductTitle)}</p>
                                    <p class="help-spec-selector-copy">${escapeHtml(parsingModeConfig.selectorProductCopy)}</p>
                                  </div>
                                  <div class="help-spec-selector-grid">
                                    ${decisionResult.productCandidates
                                      .map((candidate) => `
                                        <button class="help-spec-selector-card" data-help-spec-product="${escapeHtmlAttribute(candidate.id)}" type="button">
                                          <strong>${escapeHtml(candidate.label)}</strong>
                                        </button>
                                      `)
                                      .join("")}
                                  </div>
                                </div>
                              `
                              : ""
                          }

                          ${
                            showVariantSelector
                              ? `
                                <div class="help-spec-selector-block">
                                  <div class="help-spec-selector-head">
                                    <p class="help-spec-selector-title">${escapeHtml(parsingModeConfig.selectorVariantTitle)}</p>
                                    <p class="help-spec-selector-copy">${escapeHtml(parsingModeConfig.selectorVariantCopy)}</p>
                                  </div>
                                  <div class="help-spec-selector-grid">
                                    ${decisionResult.variantCandidates
                                      .map((candidate) => `
                                        <button class="help-spec-selector-card" data-help-spec-variant="${escapeHtmlAttribute(candidate.id)}" type="button">
                                          <strong>${escapeHtml(candidate.label)}</strong>
                                        </button>
                                      `)
                                      .join("")}
                                  </div>
                                </div>
                              `
                              : ""
                          }

                          ${
                            decisionCharacteristics.length
                              ? `
                                <div class="decision-choice-strip ${canShowDecisionTabNav && decisionTabsWindowStart > 0 ? "has-left-nav" : ""} ${canShowDecisionTabNav && decisionTabsWindowStart < maxDecisionTabsWindowStart ? "has-right-nav" : ""}">
                                  ${canShowDecisionTabNav && decisionTabsWindowStart > 0 ? '<button class="decision-choice-nav decision-choice-nav-left" id="decision-choice-nav-left" type="button" aria-label="Previous spec tabs">‹</button>' : ""}
                                  <div class="decision-choice-tabs" id="decision-choice-tabs-track">
                                    ${visibleDecisionChoices
                                      .map(({ choice, index }) => `
                                        <button class="decision-choice-tab ${index === appState.activeDecisionChoiceIndex ? "is-active" : ""}" data-choice-tab="${index}" type="button">
                                          ${escapeHtml(choice.label || `Characteristic ${index + 1}`)}
                                        </button>
                                      `)
                                      .join("")}
                                  </div>
                                  ${canShowDecisionTabNav && decisionTabsWindowStart < maxDecisionTabsWindowStart ? '<button class="decision-choice-nav decision-choice-nav-right" id="decision-choice-nav-right" type="button" aria-label="More spec tabs">›</button>' : ""}
                                </div>
                                ${
                                  activeDecisionChoice
                                    ? `
                                      <div class="help-spec-grid">
                                        ${(() => {
                                          const parsedOptions = getCharacteristicOptions(activeDecisionChoice)
                                          const conciseComparison = buildFallbackConciseComparison(activeDecisionChoice)
                                          const isDimensionsCharacteristic = /size|dimension/i.test(activeDecisionChoice.label || "")
                                          const isConfigurationCharacteristic = /configuration/i.test(activeDecisionChoice.label || "")
                                          const isUpholsteryCharacteristic = /upholstery|material/i.test(activeDecisionChoice.label || "")
                                          const isShellFinishCharacteristic = /shell finish|frame finish|wood finish|finish/i.test(activeDecisionChoice.label || "")
                                          const configurationSections = isConfigurationCharacteristic ? getConfigurationSections(activeDecisionChoice) : []
                                          const configurationMatrix = isConfigurationCharacteristic ? getConfigurationMatrix(activeDecisionChoice) : null
                                          const useCompactOptionCards = isUpholsteryCharacteristic || isShellFinishCharacteristic
                                          const useCompactSourceRow = useCompactOptionCards || isDimensionsCharacteristic
                                          const useCompactPriceStyle = useCompactOptionCards
                                          const sourceFallbackPage = decisionResult?.includedPages?.[0] || decisionResult?.startPage || appState.activePageNumber
                                          return `
                                            <div class="option-card option-card-expanded">
                                              <div class="option-card-header">
                                                <span class="option-card-title">${escapeHtml(activeDecisionChoice.label || "Characteristic")}</span>
                                                ${
                                                  activeDecisionChoice.blurb
                                                    ? `<span class="option-card-summary">${escapeHtml(activeDecisionChoice.blurb)}</span>`
                                                    : conciseComparison
                                                      ? `<span class="option-card-summary">${escapeHtml(conciseComparison)}</span>`
                                                      : ""
                                                }
                                              </div>
                                              ${parsedOptions.length
                                                ? `
                                                  ${isConfigurationCharacteristic
                                                    ? `
                                                      <div class="configuration-section-list">
                                                        ${!configurationMatrix
                                                          ? configurationSections
                                                              .map((section, index) => {
                                                                return `
                                                                  <section class="configuration-section ${index > 0 ? "configuration-section-with-divider" : ""}">
                                                                    <h4 class="configuration-section-title">${escapeHtml(section.title)}</h4>
                                                                    ${
                                                                      section.options.length
                                                                        ? `
                                                                          <div class="configuration-section-items">
                                                                            ${section.options
                                                                              .map((item) => `
                                                                                <div class="configuration-section-item">
                                                                                  <span>${escapeHtml(item)}</span>
                                                                                </div>
                                                                              `)
                                                                              .join("")}
                                                                          </div>
                                                                        `
                                                                          : ""
                                                                    }
                                                                  </section>
                                                                `
                                                              })
                                                              .join("")
                                                          : ""}
                                                        ${
                                                          configurationMatrix
                                                            ? `
                                                              <div class="configuration-matrix">
                                                                <div class="configuration-matrix-row configuration-matrix-head">
                                                                  <span>${escapeHtml(configurationMatrix.rowLabel || "Configuration")}</span>
                                                                  ${configurationMatrix.columns.map((column) => `<span>${escapeHtml(column)}</span>`).join("")}
                                                                </div>
                                                                ${configurationMatrix.rows
                                                                  .map((row) => `
                                                                    <div class="configuration-matrix-row">
                                                                      <span class="configuration-matrix-label">${escapeHtml(row.name)}</span>
                                                                      ${configurationMatrix.columns
                                                                        .map((column) => {
                                                                          const columnMatchKey = getConfigurationMatchKey(column)
                                                                          const matchingCell = row.cells.find((cell) => {
                                                                            return (cell.matchKey || getConfigurationMatchKey(cell.column)) === columnMatchKey
                                                                          })
                                                                          const price = matchingCell?.price || "-"
                                                                          return `
                                                                            <span class="configuration-matrix-cell">
                                                                              ${price !== "-" ? `<button class="option-copy-trigger configuration-section-price" data-copy-value="${escapeHtml(price)}" type="button">${escapeHtml(price)}</button>` : `<span>${escapeHtml(price)}</span>`}
                                                                              ${matchingCell?.model ? `<button class="option-copy-trigger configuration-section-model" data-copy-value="${escapeHtml(matchingCell.model)}" type="button">${escapeHtml(matchingCell.model)}</button>` : ""}
                                                                            </span>
                                                                          `
                                                                        })
                                                                        .join("")}
                                                                    </div>
                                                                  `)
                                                                  .join("")}
                                                              </div>
                                                            `
                                                            : ""
                                                        }
                                                      </div>
                                                    `
                                                    : `
                                                      <div class="option-detail-grid ${useCompactOptionCards ? "option-detail-grid-compact" : ""}">
                                                        ${parsedOptions
                                                          .map((option) => {
                                                            const valueLines = useCompactOptionCards ? [] : parseOptionValueLines(option.values)
                                                            const formattedPricing = useCompactOptionCards
                                                              ? getCompactOptionPricing(option)
                                                              : formatOptionPricingLabel(option.pricing)
                                                            const pricingRows = extractTextileRequirementRows(option.pricing)
                                                            const pricingParts = formattedPricing.split(/:\s+/)
                                                            return `
                                                              <div class="option-detail-card ${useCompactOptionCards ? "option-detail-card-compact" : ""}">
                                                                <button class="option-copy-trigger option-copy-name" data-copy-value="${escapeHtml(option.name)}" type="button">${escapeHtml(option.name)}</button>
                                                                ${
                                                                  valueLines.length
                                                                    ? `
                                                                      <div class="option-value-list">
                                                                        ${valueLines
                                                                          .map((line) => `
                                                                            <div class="option-value-row">
                                                                              ${
                                                                                line.label
                                                                                  ? `<span class="option-value-label">${escapeHtml(line.label)}:</span>`
                                                                                  : ""
                                                                              }
                                                                              <button class="option-copy-trigger option-copy-value option-copy-value-inline" data-copy-value="${escapeHtml(line.value)}" type="button">${escapeHtml(line.value)}</button>
                                                                            </div>
                                                                          `)
                                                                          .join("")}
                                                                      </div>
                                                                    `
                                                                    : ""
                                                                }
                                                                ${
                                                                  option.pricing
                                                                    ? isDimensionsCharacteristic && pricingRows.length
                                                                      ? `
                                                                        ${pricingRows
                                                                          .map((row) => {
                                                                            const rowParts = row.split(/:\s+/)
                                                                            if (rowParts.length !== 2) return ""
                                                                            return `
                                                                              <div class="option-value-row">
                                                                                <span class="option-value-label">${escapeHtml(rowParts[0])}:</span>
                                                                                <button class="option-copy-trigger option-copy-value option-copy-value-inline" data-copy-value="${escapeHtml(rowParts[1])}" type="button">${escapeHtml(rowParts[1])}</button>
                                                                              </div>
                                                                            `
                                                                          })
                                                                          .join("")}
                                                                      `
                                                                      : isDimensionsCharacteristic && pricingParts.length === 2
                                                                      ? `
                                                                        <div class="option-value-row">
                                                                          <span class="option-value-label">${escapeHtml(pricingParts[0])}:</span>
                                                                          <button class="option-copy-trigger option-copy-value option-copy-value-inline" data-copy-value="${escapeHtml(pricingParts[1])}" type="button">${escapeHtml(pricingParts[1])}</button>
                                                                        </div>
                                                                      `
                                                                      : `<span class="option-copy-meta ${useCompactPriceStyle ? "option-copy-price-emphasis" : ""}">${useCompactPriceStyle ? escapeHtml(formattedPricing) : `Pricing: ${escapeHtml(formattedPricing)}`}</span>`
                                                                    : ""
                                                                }
                                                                ${option.difference && !isDimensionsCharacteristic && !useCompactOptionCards
                                                                  ? `<span class="option-copy-meta">${escapeHtml(option.difference)}</span>`
                                                                  : ""}
                                                                ${
                                                                  useCompactSourceRow && option.evidence
                                                                    ? `<span class="option-source-chip option-source-chip-subtle">${escapeHtml(formatCompactSourceLabel(option.evidence, sourceFallbackPage))}</span>`
                                                                    : ""
                                                                }
                                                                ${
                                                                  useCompactSourceRow && !option.evidence
                                                                    ? `<span class="option-source-chip option-source-chip-subtle">${escapeHtml(formatCompactSourceLabel("", sourceFallbackPage))}</span>`
                                                                    : ""
                                                                }
                                                                ${
                                                                  !useCompactSourceRow && option.evidence
                                                                    ? `<span class="option-copy-meta">${escapeHtml(option.evidence)}</span>`
                                                                    : ""
                                                                }
                                                              </div>
                                                            `
                                                          })
                                                          .join("")}
                                                      </div>
                                                    `}
                                                `
                                                : `<div class="option-row"><span class="option-copy-meta">${escapeHtml(activeDecisionChoice.blurb || "No explicit options found for this characteristic.")}</span></div>`
                                              }
                                            </div>
                                          `
                                        })()}
                                      </div>
                                    `
                                    : ""
                                }
                              `
                              : !showProductSelector && !showVariantSelector
                                ? `
                                  <div class="help-spec-empty">
                                    <p>No characteristic modules were identified yet for this product section.</p>
                                  </div>
                                `
                                : ""
                          }

                          ${
                            decisionResult.referenceInfo?.length
                              ? `
                                <div class="help-spec-reference-block">
                                  <p class="help-spec-selector-title">Reference details</p>
                                  <ul class="help-spec-reference-list">
                                    ${decisionResult.referenceInfo.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
                                  </ul>
                                </div>
                              `
                              : ""
                          }
                        </div>
                      `
                      : ""
                  }

                  ${appState.errorMessage ? `<p class="error-banner">${escapeHtml(appState.errorMessage)}</p>` : ""}
                  ${appState.aiRerankError ? `<p class="error-banner">${escapeHtml(appState.aiRerankError)}</p>` : ""}
                  ${appState.decisionAssistError ? `<p class="error-banner">${escapeHtml(appState.decisionAssistError)}</p>` : ""}
                </div>
              `
              : ""}

            <div class="pdf-panel">
              <div class="viewer-scroll" id="viewer-scroll">
                ${
                  gateViewerUntilAiReady
                    ? `
                      <div class="viewer-loading-state">
                        <div class="loading-orb loading-orb-large" aria-hidden="true"></div>
                        <p class="viewer-loading-title">Analyzing pages...</p>
                        <p class="viewer-loading-copy">The viewer will appear once AI finishes distinguishing the best matching pages.</p>
                      </div>
                    `
                    : documentRecord
                    ? getVisiblePages(documentRecord)
                        .map((page) => {
                          const isTarget = page.pageNumber === appState.activePageNumber
                          return `
                            <article class="pdf-page ${isTarget ? "is-target" : ""}" data-page-number="${page.pageNumber}">
                              ${
                                (() => {
                                  const renderKey = getPageRenderKey(documentRecord, page.pageNumber)
                                  const renderImage = appState.pageRenderImages[renderKey]
                                  const renderMetrics = appState.pageRenderMetricsByKey[renderKey]
                                  const renderStatus = appState.pageRenderStatusByKey[renderKey]
                                  const textVisible = Boolean(appState.pageTextVisibleByKey[renderKey])
                                  const readableHtml = appState.pageReadableHtmlByKey[renderKey]
                                  const readableStatus = appState.pageReadableStatusByKey[renderKey]
                                  const readableError = appState.pageReadableErrorByKey[renderKey]
                                  const renderWidthStyle = `style="width:${Math.round(baseRenderWidth * appState.pdfZoom)}px; max-width:none;"`
                                  return renderImage
                                    ? `
                                        <div class="page-render-shell" ${renderWidthStyle}>
                                          <img class="page-render-image" src="${renderImage}" alt="Rendered PDF page ${page.pageNumber}" />
                                          ${
                                            renderMetrics && appState.pageRenderTextByKey[renderKey]
                                              ? `
                                                <div class="page-text-layer-shell" aria-hidden="true">
                                                  <div
                                                    class="page-text-layer"
                                                    data-page-text-layer-key="${escapeHtmlAttribute(renderKey)}"
                                                    data-page-width="${renderMetrics.width}"
                                                    data-page-height="${renderMetrics.height}"
                                                  ></div>
                                                </div>
                                              `
                                              : ""
                                          }
                                        </div>
                                        ${
                                          textVisible
                                            ? `
                                              <div class="page-section page-inline-text-block" style="max-width:${Math.round(baseRenderWidth)}px;">
                                                <div class="page-inline-text-head">
                                                  <p class="page-section-label">${readableHtml ? "AI-formatted text" : "Selectable PDF text"}</p>
                                                </div>
                                                ${
                                                  readableStatus === "loading"
                                                    ? `<p class="inline-ai-loading">Formatting text for easier reading...</p>`
                                                    : ""
                                                }
                                                ${
                                                  readableError
                                                    ? `<p class="inline-ai-error">${escapeHtml(readableError)}</p>`
                                                    : ""
                                                }
                                                ${
                                                  readableStatus === "loading"
                                                    ? ""
                                                    : readableHtml
                                                      ? `<div class="page-text page-text-inline-select page-text-rich">${readableHtml}</div>`
                                                      : `<div class="page-text page-text-inline-select">${escapeHtml(page.text || "[No extractable text found on this page]")}</div>`
                                                }
                                              </div>
                                            `
                                            : `
                                              <div class="page-inline-action-row" style="max-width:${Math.round(baseRenderWidth)}px;">
                                                <button class="inline-ai-btn inline-ai-reveal-btn" type="button" data-reveal-page-text="${page.pageNumber}">
                                                  ${appState.visionApiKey ? "View Selectable Text" : "View Raw Selectable Text"}
                                                </button>
                                              </div>
                                            `
                                        }
                                        ${
                                          page.ocrText
                                            ? `
                                              <div class="page-section page-ocr-block">
                                                <p class="page-section-label">Image-added text</p>
                                                <div class="page-text page-text-ocr">${escapeHtml(page.ocrText)}</div>
                                              </div>
                                            `
                                            : ""
                                        }
                                      `
                                    : `
                                        <div class="page-render-shell page-render-loading" ${renderWidthStyle}>
                                          <p class="subtle">${renderStatus === "loading" ? "Rendering PDF page..." : "Rendered preview unavailable."}</p>
                                        </div>
                                        ${
                                          textVisible
                                            ? `
                                              <div class="page-section page-inline-text-block" style="max-width:${Math.round(baseRenderWidth)}px;">
                                                <div class="page-inline-text-head">
                                                  <p class="page-section-label">${readableHtml ? "AI-formatted text" : "Selectable PDF text"}</p>
                                                </div>
                                                ${
                                                  readableStatus === "loading"
                                                    ? `<p class="inline-ai-loading">Formatting text for easier reading...</p>`
                                                    : ""
                                                }
                                                ${
                                                  readableError
                                                    ? `<p class="inline-ai-error">${escapeHtml(readableError)}</p>`
                                                    : ""
                                                }
                                                ${
                                                  readableStatus === "loading"
                                                    ? ""
                                                    : readableHtml
                                                      ? `<div class="page-text page-text-inline-select page-text-rich">${readableHtml}</div>`
                                                      : `<div class="page-text page-text-inline-select">${escapeHtml(page.text || "[No extractable text found on this page]")}</div>`
                                                }
                                              </div>
                                            `
                                            : `
                                              <div class="page-inline-action-row" style="max-width:${Math.round(baseRenderWidth)}px;">
                                                <button class="inline-ai-btn inline-ai-reveal-btn" type="button" data-reveal-page-text="${page.pageNumber}">
                                                  ${appState.visionApiKey ? "View Selectable Text" : "View Raw Selectable Text"}
                                                </button>
                                              </div>
                                            `
                                        }
                                        ${
                                          page.ocrText
                                            ? `
                                              <div class="page-section page-ocr-block">
                                                <p class="page-section-label">Image-added text</p>
                                                <div class="page-text page-text-ocr">${escapeHtml(page.ocrText)}</div>
                                              </div>
                                            `
                                            : ""
                                        }
                                      `
                                })()
                              }
                            </article>
                          `
                        })
                        .join("")
                    : '<div class="empty-state">No document loaded yet. Upload PDFs and run analysis, or use the bundled sample.</div>'
                }
              </div>

              <div class="viewer-controls">
                <div class="viewer-footer-note">${escapeHtml(documentRecord ? documentRecord.title : "")}</div>
              </div>
            </div>
          </div>

          <div class="field-panel">
            <div class="field-panel-header">
              <h2>Attributes</h2>
              <p>Capture the details from the ranked pages and spec comparisons.</p>
            </div>

            <div class="field-body">
              ${
                pageExplanation
                  ? `
                    <section class="setup-note-panel">
                      <div class="setup-note">
                        <span class="setup-note-label">Rank context</span>
                        <span class="setup-note-value">Page ${pageExplanation.pageNumber} matched ${pageExplanation.matches.length} retrieval cue${pageExplanation.matches.length === 1 ? "" : "s"}</span>
                      </div>
                    </section>
                  `
                  : ""
              }

              ${appState.spec.attributes
                .map((attribute) => {
                  const stateClass = attribute.key === appState.highlightedFieldKey
                    ? "is-highlighted"
                    : attribute.value.trim()
                      ? "is-filled"
                      : "is-empty"
                  return `
                    <div class="field-card ${stateClass}" data-field-card="${attribute.key}">
                      <div class="field-card-head">
                        <label for="field-${attribute.key}">${escapeHtml(attribute.label)}</label>
                        ${
                          attribute.key === appState.highlightedFieldKey
                            ? '<span class="field-indicator field-indicator-pulse"></span>'
                            : attribute.value.trim()
                              ? '<span class="field-indicator field-indicator-check">✓</span>'
                              : ""
                        }
                        ${
                          appState.copiedText
                            ? `<button class="field-paste-btn" data-paste-field="${attribute.key}" type="button">Paste</button>`
                            : ""
                        }
                      </div>
                      ${
                        attribute.type === "textarea"
                          ? `<textarea id="field-${attribute.key}" data-field-input="${attribute.key}">${escapeHtml(attribute.value)}</textarea>`
                          : `<input id="field-${attribute.key}" data-field-input="${attribute.key}" value="${escapeHtml(attribute.value)}" />`
                      }
                    </div>
                  `
                })
                .join("")}
            </div>

            <div class="field-panel-footer">
              <span>${escapeHtml(appState.copiedText ? `Ready to paste: ${appState.copiedText.slice(0, 48)}${appState.copiedText.length > 48 ? "..." : ""}` : "Select text or click an option value to copy it into a field.")}</span>
            </div>
          </div>
        </section>

        ${
          appState.wordStatsPage
            ? `
              <div class="lightbox-backdrop" id="word-stats-close"></div>
              <div class="lightbox-panel">
                <div class="lightbox-head">
                  <div>
                    <p class="eyebrow">Why This Page Ranked</p>
                    <h3>Page ${appState.wordStatsPage}</h3>
                  </div>
                  <button class="ghost-btn" id="word-stats-close-btn" type="button">Close</button>
                </div>
                <p class="subtle">Matches against the active retrieval inputs and ranking penalties.</p>
                ${
                  pageExplanation
                    ? `
                      <div class="doc-metrics compact-metrics">
                        <div class="doc-metric"><span>Words</span><strong>${pageExplanation.metrics.words}</strong></div>
                        <div class="doc-metric"><span>Lines</span><strong>${pageExplanation.metrics.lines}</strong></div>
                        <div class="doc-metric"><span>Characters</span><strong>${pageExplanation.metrics.characters}</strong></div>
                      </div>
                    `
                    : ""
                }
                <div class="word-stats-list">
                  ${
                    pageExplanation?.matches.length
                      ? pageExplanation.matches
                          .map(
                            (item) => `
                              <div class="word-stat-row">
                                <div>
                                  <strong>${escapeHtml(item.phrase)}</strong>
                                  <div class="metric-subline">${escapeHtml(item.label)} • ${item.count} hit${item.count === 1 ? "" : "s"}</div>
                                </div>
                                <span>${escapeHtml(item.effect)}</span>
                              </div>
                            `
                          )
                          .join("")
                      : '<span class="subtle">No product-derived phrase or token matches were found for this page.</span>'
                  }
                </div>
              </div>
            `
            : ""
        }
        ${
          appState.pickerPosition
            ? `
              <div
                class="field-picker"
                style="top:${appState.pickerPosition.top}px; left:${appState.pickerPosition.left}px;"
              >
                <h3>Choose destination field</h3>
                <p>Only visible attributes from the form are eligible targets.</p>
                <div class="field-list">
                  ${appState.spec.attributes
                    .map(
                      (attribute) => `
                        <button class="field-option" data-pick-field="${attribute.key}" type="button">
                          ${escapeHtml(attribute.label)}
                        </button>
                      `
                    )
                    .join("")}
                  <button class="field-option" id="cancel-picker-btn" type="button">Cancel</button>
                </div>
              </div>
            `
            : ""
        }

        ${
          appState.confirmPosition
            ? `
              <div
                class="confirm-dialog"
                style="top:${appState.confirmPosition.top}px; left:${appState.confirmPosition.left}px;"
              >
                <h3>Field already has a value</h3>
                <p>${escapeHtml(appState.pendingInsertText.slice(0, 120))}</p>
                <div class="field-list">
                  <button class="confirm-option" data-confirm-insert="replace" type="button">Replace existing value</button>
                  <button class="confirm-option" data-confirm-insert="append" type="button">Append to existing value</button>
                  <button class="confirm-option" data-confirm-insert="cancel" type="button">Cancel</button>
                </div>
              </div>
            `
            : ""
        }

        ${appState.toastMessage ? `<div class="toast">${escapeHtml(appState.toastMessage)}</div>` : ""}
      </div>
    `

    bindEvents()
    schedulePageTextLayerHydration()
  }

  function bindEvents() {
    document.querySelectorAll("[data-draft-input]").forEach((input) => {
      input.addEventListener("input", (event) => {
        const key = event.target.getAttribute("data-draft-input")
        appState.inputDraft[key] = event.target.value
      })
    })

    document.getElementById("pdf-upload")?.addEventListener("change", (event) => {
      appState.uploadFiles = Array.from(event.target.files || [])
      render()
    })

    document.getElementById("vision-api-key")?.addEventListener("input", (event) => {
      appState.visionApiKey = event.target.value.trim()
      render()
    })

    document.getElementById("session-submit-btn")?.addEventListener("click", async () => {
      appState.analyzeRequestLoading = true
      render()
      try {
        await handleSubmit()
        if (appState.errorMessage || !getActiveDocument()) return
        if (appState.visionApiKey) {
          await aiRerankTopPages()
        }
      } finally {
        appState.analyzeRequestLoading = false
        render()
      }
    })

    document.getElementById("ai-rerank-btn")?.addEventListener("click", async () => {
      appState.analyzeRequestLoading = true
      render()
      try {
        await handleSubmit()
        if (appState.errorMessage || !getActiveDocument()) return
        await aiRerankTopPages()
      } finally {
        appState.analyzeRequestLoading = false
        render()
      }
    })

    document.getElementById("document-select")?.addEventListener("change", async (event) => {
      setActiveDocument(event.target.value)
      if (appState.visionApiKey && !appState.aiRerankResult) {
        await aiRerankTopPages(true)
      }
    })

    document.getElementById("zoom-out-btn")?.addEventListener("click", () => {
      appState.pdfZoom = clampPdfZoom(appState.pdfZoom - 0.2)
      renderPreservingViewerScroll()
    })

    document.getElementById("zoom-reset-btn")?.addEventListener("click", () => {
      appState.pdfZoom = 1
      renderPreservingViewerScroll()
    })

    document.getElementById("zoom-in-btn")?.addEventListener("click", () => {
      appState.pdfZoom = clampPdfZoom(appState.pdfZoom + 0.2)
      renderPreservingViewerScroll()
    })

    document.querySelectorAll("[data-assistant-option]").forEach((button) => {
      button.addEventListener("click", () => {
        appState.assistantSelection = button.getAttribute("data-assistant-option") || ""
        runRanking(true)
        render()
        setPage(appState.activePageNumber)
        showToast(`Refined to ${appState.assistantSelection}`)
      })
    })

    document.querySelectorAll("[data-open-document]").forEach((button) => {
      button.addEventListener("click", () => setActiveDocument(button.getAttribute("data-open-document")))
    })

    document.querySelectorAll("[data-download-document]").forEach((button) => {
      button.addEventListener("click", () => downloadDocument(button.getAttribute("data-download-document")))
    })

    document.querySelectorAll("[data-jump-page]").forEach((button) => {
      button.addEventListener("click", () => setPage(Number(button.getAttribute("data-jump-page"))))
    })

    document.querySelectorAll("[data-help-spec-page]").forEach((button) => {
      button.addEventListener("click", async (event) => {
        event.stopPropagation()
        const pageNumber = Number(button.getAttribute("data-help-spec-page"))
        if (Number.isFinite(pageNumber)) {
          setPage(pageNumber)
          await analyzeSpecDecisions(pageNumber)
        }
      })
    })

    document.querySelectorAll("[data-help-spec-page-product]").forEach((button) => {
      button.addEventListener("click", async () => {
        const productId = button.getAttribute("data-help-spec-page-product")
        const pageNumber = appState.activePageNumber
        if (!productId || !pageNumber) return
        await analyzeSpecDecisions(pageNumber, { selectedProductId: productId, selectedVariantId: productId })
      })
    })

    document.querySelectorAll("[data-product-first-product]").forEach((button) => {
      button.addEventListener("click", async () => {
        const productId = button.getAttribute("data-product-first-product")
        const pageNumber = appState.structureRouting?.productFirstPageNumber || appState.aiRerankResult?.bestPage || appState.activePageNumber
        if (!productId || !pageNumber) return
        appState.productFirstSelection = {
          active: true,
          productId
        }
        setPage(pageNumber)
        await analyzeSpecDecisions(pageNumber, { selectedProductId: productId, fromProductFirst: true })
      })
    })

    document.querySelectorAll("[data-change-product-selection]").forEach((button) => {
      button.addEventListener("click", () => {
        appState.productFirstSelection = null
        appState.decisionAssistResult = null
        appState.activeDecisionChoiceIndex = -1
        appState.decisionTabsWindowStart = 0
        appState.decisionAssistError = ""
        renderPreservingViewerScroll()
      })
    })

    document.getElementById("page-strip-track")?.addEventListener("scroll", () => {
      const track = document.getElementById("page-strip-track")
      if (!track) return
      const maxScrollLeft = Math.max(0, track.scrollWidth - track.clientWidth)
      const nextCanScrollLeft = track.scrollLeft > 4
      const nextCanScrollRight = track.scrollLeft < maxScrollLeft - 4
      appState.pageStripScrollLeft = track.scrollLeft
      if (nextCanScrollLeft !== appState.pageStripCanScrollLeft || nextCanScrollRight !== appState.pageStripCanScrollRight) {
        appState.pageStripCanScrollLeft = nextCanScrollLeft
        appState.pageStripCanScrollRight = nextCanScrollRight
        renderPreservingViewerScroll()
      }
    })

    document.getElementById("page-strip-nav-left")?.addEventListener("click", () => {
      const track = document.getElementById("page-strip-track")
      if (!track) return
      const distance = Math.max(180, Math.floor(track.clientWidth * 0.72))
      appState.pageStripScrollLeft = Math.max(0, track.scrollLeft - distance)
      track.scrollTo({ left: appState.pageStripScrollLeft, behavior: "smooth" })
      queuePageStripSync()
    })

    document.getElementById("page-strip-nav-right")?.addEventListener("click", () => {
      const track = document.getElementById("page-strip-track")
      if (!track) return
      const distance = Math.max(180, Math.floor(track.clientWidth * 0.72))
      const maxScrollLeft = Math.max(0, track.scrollWidth - track.clientWidth)
      appState.pageStripScrollLeft = Math.min(maxScrollLeft, track.scrollLeft + distance)
      track.scrollTo({ left: appState.pageStripScrollLeft, behavior: "smooth" })
      queuePageStripSync()
    })

    document.getElementById("decision-choice-tabs-track")?.addEventListener("scroll", () => {
      const track = document.getElementById("decision-choice-tabs-track")
      if (!track) return
      appState.decisionTabsScrollLeft = track.scrollLeft
    })

    document.getElementById("decision-choice-nav-left")?.addEventListener("click", () => {
      appState.decisionTabsWindowStart = Math.max(0, appState.decisionTabsWindowStart - 1)
      renderPreservingViewerScroll()
    })

    document.getElementById("decision-choice-nav-right")?.addEventListener("click", () => {
      appState.decisionTabsWindowStart += 1
      renderPreservingViewerScroll()
    })

    window.addEventListener("resize", queuePageStripSync, { once: true })
    window.addEventListener("resize", queueDecisionTabsSync, { once: true })
    queuePageStripSync()
    queueDecisionTabsSync()

    document.querySelectorAll("[data-choice-tab]").forEach((button) => {
      button.addEventListener("click", () => {
        const nextIndex = Number(button.getAttribute("data-choice-tab"))
        appState.activeDecisionChoiceIndex =
          appState.activeDecisionChoiceIndex === nextIndex ? -1 : nextIndex
        if (appState.activeDecisionChoiceIndex >= 0) {
          const minVisibleIndex = appState.decisionTabsWindowStart
          const maxVisibleIndex = appState.decisionTabsWindowStart + 2
          if (nextIndex < minVisibleIndex) {
            appState.decisionTabsWindowStart = nextIndex
          } else if (nextIndex > maxVisibleIndex) {
            appState.decisionTabsWindowStart = Math.max(0, nextIndex - 2)
          }
        }
        renderPreservingViewerScroll()
      })
    })

    document.querySelectorAll("[data-help-spec-product]").forEach((button) => {
      button.addEventListener("click", async () => {
        const productId = button.getAttribute("data-help-spec-product")
        const pageNumber = appState.decisionAssistResult?.pageNumber
        if (!productId || !pageNumber) return
        await analyzeSpecDecisions(pageNumber, { selectedProductId: productId })
      })
    })

    document.querySelectorAll("[data-help-spec-variant]").forEach((button) => {
      button.addEventListener("click", async () => {
        const variantId = button.getAttribute("data-help-spec-variant")
        const pageNumber = appState.decisionAssistResult?.pageNumber
        const productId = appState.decisionAssistResult?.selectedProductId || ""
        if (!variantId || !pageNumber) return
        await analyzeSpecDecisions(pageNumber, { selectedProductId: productId, selectedVariantId: variantId })
      })
    })

    document.querySelectorAll("[data-run-ocr]").forEach((button) => {
      button.addEventListener("click", async (event) => {
        event.stopPropagation()
        await analyzePageImage(Number(button.getAttribute("data-run-ocr")))
      })
    })

    document.querySelectorAll("[data-open-word-stats]").forEach((button) => {
      button.addEventListener("click", (event) => {
        event.stopPropagation()
        appState.wordStatsPage = Number(button.getAttribute("data-open-word-stats"))
        render()
      })
    })

    document.getElementById("word-stats-close")?.addEventListener("click", () => {
      appState.wordStatsPage = null
      render()
    })

    document.getElementById("word-stats-close-btn")?.addEventListener("click", () => {
      appState.wordStatsPage = null
      render()
    })

    document.querySelectorAll("[data-field-input]").forEach((input) => {
      input.addEventListener("input", (event) => {
        const key = event.target.getAttribute("data-field-input")
        const attribute = appState.spec.attributes.find((item) => item.key === key)
        if (attribute) attribute.value = event.target.value
      })
    })

    document.querySelectorAll("[data-copy-value]").forEach((button) => {
      button.addEventListener("click", (event) => {
        const value = button.getAttribute("data-copy-value") || ""
        if (!value) return
        const rect = event.currentTarget.getBoundingClientRect()
        setSelectionTarget(value, "assistant", rect)
      })
    })

    document.querySelectorAll("[data-reveal-page-text]").forEach((button) => {
      button.addEventListener("click", async () => {
        const pageNumber = Number(button.getAttribute("data-reveal-page-text"))
        if (Number.isFinite(pageNumber)) {
          await revealPageText(pageNumber)
        }
      })
    })

    document.querySelectorAll("[data-paste-field]").forEach((button) => {
      button.addEventListener("click", () => {
        const fieldKey = button.getAttribute("data-paste-field")
        if (!fieldKey || !appState.copiedText) return
        insertIntoField(fieldKey, appState.copiedText, "replace")
      })
    })

    document.querySelectorAll("[data-pick-field]").forEach((button) => {
      button.addEventListener("click", () => handleFieldPick(button.getAttribute("data-pick-field")))
    })

    document.querySelectorAll("[data-confirm-insert]").forEach((button) => {
      button.addEventListener("click", () => resolveInsert(button.getAttribute("data-confirm-insert")))
    })

  }

  async function handleSubmit() {
    const draft = appState.inputDraft
    const attributes = buildAttributesFromText(draft.attributes)
    if (!attributes.length) {
      appState.errorMessage = "Add at least one attribute before submitting."
      showToast(appState.errorMessage)
      render()
      return
    }

    appState.errorMessage = ""
    appState.spec = {
      originalSpecName: normalizeText(draft.productName),
      specDisplayName: normalizeText(draft.productName),
      originalBrand: normalizeText(draft.brandName),
      brandDisplayName: normalizeText(draft.brandName),
      category: normalizeText(draft.category),
      attributes
    }
    clearSelectionState()
    appState.assistantSelection = ""
    appState.aiRerankResult = null
    appState.aiRerankDocumentId = ""
    appState.aiRerankCacheByDocumentId = {}
    appState.structureRouting = null
    appState.productFirstSelection = null
    appState.aiRerankError = ""
    appState.sourceSelectionScores = []
    appState.sourceSelectionChosenId = ""
    appState.decisionAssistResult = null
    appState.familyPageProductsByKey = {}
    appState.familyPageProductsStatusByKey = {}
    appState.familyPageProductsErrorByKey = {}
    appState.activeDecisionChoiceIndex = -1
    appState.decisionTabsWindowStart = 0
    appState.decisionAssistError = ""
    appState.summaryPanelOpen = false

    if (!appState.uploadFiles.length && !appState.documents.length) {
      appState.errorMessage = "No PDF is loaded yet. Upload a PDF or wait for the default document to finish loading."
      appState.loadingMessage = ""
      showToast(appState.errorMessage)
      render()
      saveCurrentAsDefault().catch(() => {})
      return
    }

    if (!appState.uploadFiles.length && appState.documents.length) {
      const sourceScores = getDocumentKeywordDensityScores(appState.documents)
      appState.sourceSelectionScores = sourceScores
      appState.sourceSelectionChosenId = sourceScores[0]?.documentId || appState.activeDocumentId
      const bestDocument = chooseBestDocumentByKeywordDensity(appState.documents)
      if (bestDocument) {
        appState.activeDocumentId = bestDocument.id
      }
      runRanking(true)
      render()
      saveCurrentAsDefault().catch(() => {})
      requestAnimationFrame(() => setPage(appState.activePageNumber))
      return
    }

    try {
      appState.loadingMessage = "Parsing uploaded PDFs..."
      render()
      const uploadedDocuments = await parseUploadedFiles(appState.uploadFiles)
      appState.documents = uploadedDocuments
      const sourceScores = getDocumentKeywordDensityScores(uploadedDocuments)
      appState.sourceSelectionScores = sourceScores
      appState.sourceSelectionChosenId = sourceScores[0]?.documentId || ""
      const bestDocument = chooseBestDocumentByKeywordDensity(uploadedDocuments)
      appState.activeDocumentId = bestDocument?.id || uploadedDocuments[0]?.id || ""
      appState.uploadFiles = []
      appState.aiRerankCacheByDocumentId = {}
      appState.structureRouting = null
      appState.productFirstSelection = null
      appState.rankedPages = []
      appState.pagePreviews = {}
      appState.pageRenderImages = {}
      appState.pageRenderTextByKey = {}
      appState.pageRenderMetricsByKey = {}
      appState.pageRenderStatusByKey = {}
      appState.pageTextVisibleByKey = {}
      appState.pageReadableHtmlByKey = {}
      appState.pageReadableStatusByKey = {}
      appState.pageReadableErrorByKey = {}
      appState.loadingMessage = ""
      runRanking(true)
      render()
      schedulePreviewRendering()
      saveCurrentAsDefault().catch(() => {})
      requestAnimationFrame(() => setPage(appState.activePageNumber))
      showToast(`Loaded ${uploadedDocuments.length} PDF${uploadedDocuments.length === 1 ? "" : "s"}`)
    } catch (error) {
      appState.loadingMessage = ""
      appState.errorMessage = error instanceof Error ? error.message : "Unable to parse the uploaded PDFs."
      render()
    }
  }

  async function resetToBundledDefault() {
    appState.spec = cloneSpec(initialSpec)
    appState.documents = []
    appState.activeDocumentId = ""
    appState.rankedPages = []
    appState.assistantResult = null
    appState.assistantSelection = ""
    appState.aiRerankResult = null
    appState.aiRerankDocumentId = ""
    appState.aiRerankCacheByDocumentId = {}
    appState.structureRouting = null
    appState.productFirstSelection = null
    appState.aiRerankError = ""
    appState.sourceSelectionScores = []
    appState.sourceSelectionChosenId = ""
    appState.decisionAssistResult = null
    appState.familyPageProductsByKey = {}
    appState.familyPageProductsStatusByKey = {}
    appState.familyPageProductsErrorByKey = {}
    appState.activeDecisionChoiceIndex = -1
    appState.decisionTabsWindowStart = 0
    appState.decisionAssistError = ""
    appState.summaryPanelOpen = false
    appState.pagePreviews = {}
    appState.pageRenderImages = {}
    appState.pageRenderTextByKey = {}
    appState.pageRenderMetricsByKey = {}
    appState.pageRenderStatusByKey = {}
    appState.pageTextVisibleByKey = {}
    appState.pageReadableHtmlByKey = {}
    appState.pageReadableStatusByKey = {}
    appState.pageReadableErrorByKey = {}
    appState.uploadFiles = []
    appState.errorMessage = ""
    appState.loadingMessage = "Loading bundled default PDF..."
    appState.inputDraft = {
      productName: initialSpec.specDisplayName,
      brandName: initialSpec.brandDisplayName,
      category: initialSpec.category,
      attributes: initialSpec.attributes.map((attribute) => attribute.label).join("\n")
    }
    clearSelectionState()
    render()
    await loadBundledDefaultCase()
  }

  function openPersistenceDb() {
    if (!("indexedDB" in window)) {
      return Promise.reject(new Error("IndexedDB is not available in this browser."))
    }

    if (!persistenceDbPromise) {
      persistenceDbPromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(PERSISTENCE_DB, 1)
        request.onupgradeneeded = () => {
          const database = request.result
          if (!database.objectStoreNames.contains(PERSISTENCE_STORE)) {
            database.createObjectStore(PERSISTENCE_STORE)
          }
        }
        request.onsuccess = () => resolve(request.result)
        request.onerror = () => reject(request.error || new Error("Unable to open IndexedDB."))
      })
    }

    return persistenceDbPromise
  }

  async function saveCurrentAsDefault() {
    const database = await openPersistenceDb()
    const payload = {
      savedAt: Date.now(),
      spec: appState.spec,
      inputDraft: appState.inputDraft,
      visionApiKey: appState.visionApiKey,
      activeDocumentId: appState.activeDocumentId,
      documents: appState.documents.map((document) => ({
        id: document.id,
        title: document.title,
        description: document.description,
        useCase: document.useCase,
        sourceType: document.sourceType,
        pageTitles: document.pageTitles,
        pdfText: document.pdfText || null,
        pdfBytes: document.pdfBytes ? [...document.pdfBytes] : null,
        pages: document.pages
      }))
    }

    await new Promise((resolve, reject) => {
      const transaction = database.transaction(PERSISTENCE_STORE, "readwrite")
      transaction.objectStore(PERSISTENCE_STORE).put(payload, PERSISTENCE_KEY)
      transaction.oncomplete = () => resolve()
      transaction.onerror = () => reject(transaction.error || new Error("Unable to save default session."))
    })
  }

  async function loadSavedDefault() {
    const database = await openPersistenceDb()
    return new Promise((resolve, reject) => {
      const transaction = database.transaction(PERSISTENCE_STORE, "readonly")
      const request = transaction.objectStore(PERSISTENCE_STORE).get(PERSISTENCE_KEY)
      request.onsuccess = () => resolve(request.result || null)
      request.onerror = () => reject(request.error || new Error("Unable to load saved default session."))
    })
  }

  async function clearSavedDefault() {
    const database = await openPersistenceDb()
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(PERSISTENCE_STORE, "readwrite")
      transaction.objectStore(PERSISTENCE_STORE).delete(PERSISTENCE_KEY)
      transaction.oncomplete = () => resolve()
      transaction.onerror = () => reject(transaction.error || new Error("Unable to clear saved default session."))
    })
  }

  async function hydrateSavedDraftState() {
    try {
      const saved = await loadSavedDefault()
      if (!saved) return false

      appState.inputDraft = {
        ...appState.inputDraft,
        ...(saved.inputDraft || {})
      }
      appState.visionApiKey = saved.visionApiKey || ""
      appState.spec = {
        ...cloneSpec(initialSpec),
        originalSpecName: normalizeText(appState.inputDraft.productName) || initialSpec.originalSpecName,
        specDisplayName: normalizeText(appState.inputDraft.productName) || initialSpec.specDisplayName,
        originalBrand: normalizeText(appState.inputDraft.brandName) || initialSpec.originalBrand,
        brandDisplayName: normalizeText(appState.inputDraft.brandName) || initialSpec.brandDisplayName,
        category: normalizeText(appState.inputDraft.category) || initialSpec.category,
        attributes: buildAttributesFromText(appState.inputDraft.attributes || "").length
          ? buildAttributesFromText(appState.inputDraft.attributes)
          : cloneSpec(initialSpec).attributes
      }
      appState.documents = []
      appState.activeDocumentId = ""
      appState.uploadFiles = []
      appState.rankedPages = []
      appState.assistantResult = null
      appState.assistantSelection = ""
      appState.aiRerankResult = null
      appState.aiRerankDocumentId = ""
      appState.aiRerankCacheByDocumentId = {}
      appState.structureRouting = null
      appState.productFirstSelection = null
      appState.aiRerankError = ""
      appState.sourceSelectionScores = []
      appState.sourceSelectionChosenId = ""
      appState.decisionAssistResult = null
      appState.activeDecisionChoiceIndex = -1
      appState.decisionTabsWindowStart = 0
      appState.decisionAssistError = ""
      appState.pagePreviews = {}
      appState.pageRenderImages = {}
      appState.pageRenderTextByKey = {}
      appState.pageRenderMetricsByKey = {}
      appState.pageRenderStatusByKey = {}
      appState.pageTextVisibleByKey = {}
      appState.pageReadableHtmlByKey = {}
      appState.pageReadableStatusByKey = {}
      appState.pageReadableErrorByKey = {}
      appState.loadingMessage = ""
      render()
      return true
    } catch (error) {
      // If persistence fails, keep the in-memory defaults.
      return false
    }
  }

  async function loadPdfJsLibrary() {
    if (!pdfjsLibPromise) {
      pdfjsLibPromise = import(`https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.min.mjs`)
        .then((module) => {
          module.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.mjs`
          return module
        })
        .catch(() => {
          throw new Error("PDF.js could not be loaded from the CDN. Serve the prototype over http:// and make sure you are online.")
        })
    }

    return pdfjsLibPromise
  }

  async function parseUploadedFiles(files) {
    const results = []
    const nameCounts = new Map()

    for (const file of files) {
      const seenCount = (nameCounts.get(file.name) || 0) + 1
      nameCounts.set(file.name, seenCount)
      const displayName = seenCount > 1 ? `${file.name} (${seenCount})` : file.name
      const bytes = new Uint8Array(await file.arrayBuffer())
      results.push(await buildUploadedDocument(file.name, bytes, results.length + 1, displayName))
    }

    return results
  }

  async function buildUploadedDocument(fileName, bytes, ordinal, displayTitle) {
    const pdfjsLib = await loadPdfJsLibrary()
    const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(bytes) })
    const pdfDocument = await loadingTask.promise
    const pages = []

    for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
      const page = await pdfDocument.getPage(pageNumber)
      const textContent = await page.getTextContent()
      const pageText = extractTextFromContentItems(textContent.items)
      pages.push({
        pageNumber,
        text: pageText,
        ocrText: ""
      })
    }

    return {
      id: `${slugify(fileName.replace(/\.pdf$/i, ""), "upload")}-${ordinal}`,
      title: displayTitle || fileName,
      description: "Uploaded by user for ranking validation.",
      useCase: "Uploaded PDF",
      sourceType: "uploaded",
      pageTitles: pages.map(() => displayTitle || fileName),
      pdfBytes: Array.from(bytes),
      pages
    }
  }

  async function loadBundledDefaultCase() {
    try {
      const response = await fetch(DEFAULT_BUNDLED_PDF_PATH)
      if (!response.ok) {
        throw new Error(`Bundled default PDF could not be loaded (${response.status}).`)
      }

      const bytes = new Uint8Array(await response.arrayBuffer())
      const bundledDocument = await buildUploadedDocument(DEFAULT_BUNDLED_PDF_NAME, bytes, 1)
      appState.documents = [bundledDocument]
      appState.activeDocumentId = bundledDocument.id
      appState.rankedPages = []
      appState.errorMessage = ""
      appState.loadingMessage = ""
      runRanking(true)
      render()
      schedulePreviewRendering()
      saveCurrentAsDefault().catch(() => {})
      requestAnimationFrame(() => setPage(appState.activePageNumber))
    } catch (error) {
      appState.documents = sampleDocuments.map(hydrateSyntheticDocument)
      appState.activeDocumentId = sampleDocuments[1].id
      appState.loadingMessage = ""
      appState.errorMessage = error instanceof Error
        ? `${error.message} Falling back to built-in samples.`
        : "Bundled default PDF could not be loaded. Falling back to built-in samples."
      runRanking(true)
      render()
      requestAnimationFrame(() => setPage(appState.activePageNumber))
    }
  }

  async function renderPdfPageToCanvas(documentRecord, pageNumber, scale) {
    if (!documentRecord?.pdfBytes) {
      throw new Error("Image analysis is only available for real PDF sources.")
    }

    const pdfjsLib = await loadPdfJsLibrary()
    // PDF.js can detach the underlying buffer during worker handoff, so OCR analysis
    // always works against a fresh copy of the stored bytes.
    const loadingTask = pdfjsLib.getDocument({ data: Uint8Array.from(documentRecord.pdfBytes) })
    const pdfDocument = await loadingTask.promise
    const page = await pdfDocument.getPage(pageNumber)
    const viewport = page.getViewport({ scale })
    const canvas = document.createElement("canvas")
    const context = canvas.getContext("2d")
    canvas.width = Math.ceil(viewport.width)
    canvas.height = Math.ceil(viewport.height)
    await page.render({ canvasContext: context, viewport }).promise
    return canvas
  }

  async function getPdfPageRenderData(documentRecord, pageNumber, scale) {
    if (!documentRecord?.pdfBytes) {
      throw new Error("Image analysis is only available for real PDF sources.")
    }

    const pdfjsLib = await loadPdfJsLibrary()
    const loadingTask = pdfjsLib.getDocument({ data: Uint8Array.from(documentRecord.pdfBytes) })
    const pdfDocument = await loadingTask.promise
    const page = await pdfDocument.getPage(pageNumber)
    const viewport = page.getViewport({ scale })
    const canvas = document.createElement("canvas")
    const context = canvas.getContext("2d")
    canvas.width = Math.ceil(viewport.width)
    canvas.height = Math.ceil(viewport.height)
    await page.render({ canvasContext: context, viewport }).promise
    const textContent = await page.getTextContent()

    return {
      canvas,
      textContent,
      viewport
    }
  }

  async function analyzeDimensionsWithVision(pageNumber, canvas) {
    if (!appState.visionApiKey) {
      throw new Error("Add an OpenAI API key above before using Analyze dimensions.")
    }

    const imageUrl = canvas.toDataURL("image/png")
    const prompt = [
      "You are analyzing a furniture specification PDF page image.",
      `Product name: ${appState.spec.specDisplayName || appState.spec.originalSpecName || "Unknown"}`,
      "Determine whether the page contains dimension drawings for the product and extract the visible measurements.",
      "Pay attention to variants such as upholstered vs nonupholstered if shown.",
      "Return JSON only with no markdown fences.",
      "",
      "Required JSON shape:",
      '{',
      '  "has_dimension_diagram": true,',
      '  "summary": "short sentence",',
      '  "measurements": [',
      '    {"variant": "nonupholstered", "label": "overall height", "value": "29 3/4\\"", "evidence": "right elevation"}',
      "  ],",
      '  "raw_visible_dimensions": ["20 3/4\\"", "16 1/8\\"", "29 3/4\\""],',
      '  "notes": ["short note"]',
      '}',
      "If no reliable dimensions are visible, return an empty measurements array and explain why in notes."
    ].join("\n")

    const response = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${appState.visionApiKey}`
      },
      body: JSON.stringify({
        model: appState.visionModel || DEFAULT_VISION_MODEL,
        input: [
          {
            role: "user",
            content: [
              { type: "input_text", text: prompt },
              { type: "input_image", image_url: imageUrl }
            ]
          }
        ]
      })
    })

    if (!response.ok) {
      const errorBody = await response.text()
      throw new Error(`Vision analysis failed (${response.status}): ${errorBody.slice(0, 180)}`)
    }

    const payload = await response.json()
    const rawText = extractResponseText(payload)
    const analysis = parseJsonPayload(rawText)
    if (!analysis) {
      throw new Error("The vision model returned a response that could not be parsed as JSON.")
    }

    return {
      rawText,
      analysis,
      formattedText: formatVisionAnalysisText(analysis)
    }
  }

  async function analyzeFamilyPageProductsWithVision(documentRecord, pageNumber) {
    if (!appState.visionApiKey) {
      throw new Error("Add an OpenAI API key before extracting page products.")
    }

    const canvas = await renderPdfPageToCanvas(documentRecord, pageNumber, 1.6)
    const imageUrl = canvas.toDataURL("image/png")
    const selectedPage = documentRecord?.pages?.find((page) => page.pageNumber === pageNumber)
    const subtypeHint = getSelectedPageSubtypeHint(documentRecord, selectedPage)
    const prompt = [
      "You are reading a furniture manufacturer PDF page image.",
      "Goal: detect the concrete product rows shown on this single page so the UI can let the user pick the exact product before any spec analysis runs.",
      "Return only products visibly present on this page image.",
      "Do not infer additional family products from other pages.",
      "Do not return abstract buckets such as Base Type, Upholstery Type, Finish Type, or Variant.",
      "Prefer one card per visible item/model row.",
      "Each product should include the model code if visible and a short distinguishing label based on the visible description.",
      subtypeHint ? `The page itself appears to be scoped to: ${subtypeHint}. Only return rows that match that scope.` : "Subtype hint: none",
      "If the page shows exactly three products, return three products.",
      "Return JSON only with no markdown fences.",
      "",
      "Required JSON shape:",
      "{",
      '  "products": [',
      '    {"id": "gl-10", "label": "Exposed Veneer Shell / Interior Upholstery", "description": "GL-10", "evidence": "visible item row"}',
      "  ]",
      "}"
    ].join("\n")

    const response = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${appState.visionApiKey}`
      },
      body: JSON.stringify({
        model: appState.visionModel || DEFAULT_VISION_MODEL,
        input: [
          {
            role: "user",
            content: [
              { type: "input_text", text: prompt },
              { type: "input_image", image_url: imageUrl }
            ]
          }
        ]
      })
    })

    if (!response.ok) {
      const errorBody = await response.text()
      throw new Error(`Family page extraction failed (${response.status}): ${errorBody.slice(0, 180)}`)
    }

    const payload = await response.json()
    const rawText = extractResponseText(payload)
    const parsed = parseJsonPayload(rawText)
    if (!parsed) {
      throw new Error("Family page extraction returned unreadable JSON.")
    }

    return normalizeDecisionCandidates(parsed.products)
  }

  async function ensureFamilyPageProducts(pageNumber) {
    const documentRecord = getActiveDocument()
    if (!documentRecord || getSpecParsingMode() !== "family" || !Number.isFinite(Number(pageNumber))) return
    const key = getFamilyPageSelectionKey(documentRecord, pageNumber)
    if (appState.familyPageProductsByKey[key]?.length) return
    if (appState.familyPageProductsStatusByKey[key] === "loading") return
    if (!appState.visionApiKey) return

    try {
      appState.familyPageProductsStatusByKey[key] = "loading"
      appState.familyPageProductsErrorByKey[key] = ""
      renderPreservingViewerScroll()

      const products = await analyzeFamilyPageProductsWithVision(documentRecord, pageNumber)
      appState.familyPageProductsByKey[key] = products
      appState.familyPageProductsStatusByKey[key] = "done"
      appState.familyPageProductsErrorByKey[key] = ""
    } catch (error) {
      appState.familyPageProductsByKey[key] = []
      appState.familyPageProductsStatusByKey[key] = "error"
      appState.familyPageProductsErrorByKey[key] = error instanceof Error ? error.message : "Unable to extract products from this page."
    } finally {
      renderPreservingViewerScroll()
    }
  }

  async function aiRerankTopPages(useActiveDocumentOnly) {
    let documentRecord = getActiveDocument()
    if (!useActiveDocumentOnly && appState.documents.length > 1) {
      const sourceScores = getDocumentKeywordDensityScores(appState.documents)
      appState.sourceSelectionScores = sourceScores
      const bestDocument = chooseBestDocumentByKeywordDensity(appState.documents)
      appState.sourceSelectionChosenId = bestDocument?.id || appState.activeDocumentId
      if (bestDocument && bestDocument.id !== appState.activeDocumentId) {
        appState.activeDocumentId = bestDocument.id
        appState.pagePreviews = {}
        appState.pageRenderImages = {}
        appState.pageRenderTextByKey = {}
        appState.pageRenderMetricsByKey = {}
        appState.pageRenderStatusByKey = {}
        appState.pageTextVisibleByKey = {}
        appState.pageReadableHtmlByKey = {}
        appState.pageReadableStatusByKey = {}
        appState.pageReadableErrorByKey = {}
        runRanking(true)
        render()
        showToast(`Best source selected: ${bestDocument.title}`)
      }
      documentRecord = getActiveDocument()
    } else {
      const sourceScores = getDocumentKeywordDensityScores(appState.documents)
      appState.sourceSelectionScores = sourceScores
      appState.sourceSelectionChosenId = sourceScores[0]?.documentId || appState.activeDocumentId
    }

    if (!documentRecord) {
      appState.aiRerankError = "No ranked pages are available yet."
      render()
      return
    }

    const candidatePages = rankPages(documentRecord).slice(0, getAiRerankCandidateLimit())
    if (!candidatePages.length) {
      appState.aiRerankError = "No ranked pages are available yet."
      render()
      return
    }
    if (!appState.visionApiKey) {
      appState.aiRerankError = "Add an OpenAI API key above before using AI rerank."
      showToast("Add an OpenAI API key first")
      render()
      return
    }

    try {
      appState.aiRerankLoading = true
      appState.aiRerankError = ""
      showToast("AI reranking top pages...")
      render()

      const pageImages = []
      for (const candidate of candidatePages) {
        const canvas = await renderPdfPageToCanvas(documentRecord, candidate.pageNumber, 1.2)
        pageImages.push({
          pageNumber: candidate.pageNumber,
          imageUrl: canvas.toDataURL("image/png"),
          score: candidate.score
        })
      }

      const prompt = [
        "You are reranking candidate PDF pages for a furniture product.",
        `Product name: ${appState.spec.specDisplayName || appState.spec.originalSpecName || "Unknown"}`,
        "Goal: determine whether the candidate pages show the same product, the same product family, or different products.",
        "Treat this as hierarchical classification, not page summarization.",
        "Your job is to identify the highest-level meaningful visible difference between pages so a user can quickly pick the right one.",
        "When comparing pages, evaluate differences in this order:",
        "1. Product category",
        "2. Product subtype or configuration",
        "3. Base type",
        "4. Size / geometry / height",
        "5. Material or surface",
        "6. Finish or color",
        "Always return the highest-level meaningful difference you can see.",
        "Do not assume the pages show the same product.",
        "First determine whether the pages represent the same product, the same product family, or different products.",
        "Prefer visual signals over repeated text, especially product imagery, silhouette, section headers, dimension diagrams, and layout structure.",
        "For any top pages with nearly the same scores, compare those pages directly and assign each page a short distinct label describing what differentiates it from the others.",
        "If one page is the base or generic version and nearby pages are variants, label it as Base or General only when that is visually supported by comparison.",
        "Do not invent a finish, material, or size difference unless the page actually shows it.",
        "Normalize synonymous outputs for UI consistency. Examples: wood legs -> Wood base, barstool -> Bar stool, cafe table -> Cafe table, rect table -> Rectangular table.",
        "Keep labels extremely short, ideally 1-4 words.",
        "Examples of good labels: Base, Metal base, Wood base, Side chair, Lounge chair, Mid-back lounge, Bar height, Ebony finish, White Ash finish.",
        "Score each candidate page from 0 to 100 for this task: how likely is it to be the best product/spec page for this product.",
        "Prefer focused product detail pages, dimensions pages, and specification content.",
        "Deprioritize appendix charts, textile application matrices, broad tables, and pages that list many unrelated products.",
        "If a product reference image is provided, use it to compare visually against the candidate pages.",
        "Additionally, determine the structural pattern of the candidate pages.",
        "Does this appear to represent a single product with variants, or a family of multiple distinct products?",
        "Identify whether multiple concrete products are visibly present.",
        "Prefer decisions based on visible structure such as titles, section blocks, repeated layouts, and distinct spec tables, not assumptions.",
        "Indicate confidence in this structure determination as a number from 0.0 to 1.0.",
        "Indicate whether concrete visible product candidates are safely extractable from the top reranked page set.",
        "If concrete visible product labels or groups are present on the top reranked page(s), return only those grounded options.",
        "Do not return abstract buckets such as Base Type, Upholstery Type, Finish Type, Variant, or Configuration as concrete product candidates.",
        "Return JSON only with no markdown fences.",
        "",
        "Required JSON shape:",
        '{',
        '  "status": "single_match | similar_variants | unclear",',
        '  "relationship": "same_product | same_family | different_products",',
        '  "best_page": 5,',
        '  "confidence": "high",',
        '  "summary": "short summary",',
        '  "question": "optional short question for the user",',
        '  "structure_type": "single_product | product_family",',
        '  "interaction_model": "page_first | product_first",',
        '  "structure_confidence": 0.0,',
        '  "has_concrete_products": true,',
        '  "concrete_product_candidates": [',
        '    {"id": "mid-back-lounge", "label": "Mid Back Lounge Chair", "description": "visible section title", "evidence": "top reranked page title"}',
        "  ],",
        '  "variant_comparison": [',
        '    {"page_number": 86, "label": "Base", "difference": "base product page compared with nearby finish variants", "reason": "same chair and ottoman without a variant qualifier in the page title"},',
        '    {"page_number": 88, "label": "Ebony finish", "difference": "same product with Ebony called out", "reason": "page title or visible finish callout identifies Ebony"},',
        '    {"page_number": 89, "label": "White Ash finish", "difference": "same product with White Ash called out", "reason": "page title or visible finish callout identifies White Ash"}',
        "  ],",
        '  "ordered_pages": [',
        '    {"page_number": 5, "ai_score": 94, "role": "dimensions/spec page", "reason": "focused product drawings and dimensions", "confidence": "high"}',
        "  ]",
        '}'
      ].join("\n")

      const content = [{ type: "input_text", text: prompt }]

      if (appState.productImageUrl) {
        content.push({ type: "input_text", text: "Reference product image:" })
        content.push({ type: "input_image", image_url: appState.productImageUrl })
      }

      pageImages.forEach((page) => {
        content.push({
          type: "input_text",
          text: `Candidate page ${page.pageNumber}. Deterministic score: ${page.score}.`
        })
        content.push({ type: "input_image", image_url: page.imageUrl })
      })

      const response = await fetch("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${appState.visionApiKey}`
        },
        body: JSON.stringify({
          model: appState.visionModel || DEFAULT_VISION_MODEL,
          input: [{ role: "user", content }]
        })
      })

      if (!response.ok) {
        const errorBody = await response.text()
        throw new Error(`AI rerank failed (${response.status}): ${errorBody.slice(0, 180)}`)
      }

      const payload = await response.json()
      const rawText = extractResponseText(payload)
      const parsed = parseJsonPayload(rawText)
      if (!parsed?.ordered_pages?.length) {
        throw new Error("AI rerank returned an unreadable response.")
      }

      const orderedPages = parsed.ordered_pages
        .map((item) => ({
          pageNumber: Number(item.page_number),
          aiScore: Number(item.ai_score),
          role: item.role || "",
          reason: item.reason || "",
          confidence: item.confidence || ""
        }))
        .filter((item) => Number.isFinite(item.pageNumber))
      const strongPages = orderedPages.filter((item) => Number.isFinite(item.aiScore) && item.aiScore >= 60)
      const keptPages = (strongPages.length ? strongPages : orderedPages.slice(0, getAiKeptPageLimit())).map((item) => item.pageNumber)
      const bestPage = keptPages.includes(Number(parsed.best_page)) ? Number(parsed.best_page) : keptPages[0]
      const variantComparison = Array.isArray(parsed.variant_comparison)
        ? parsed.variant_comparison
            .map((item) => ({
              pageNumber: Number(item.page_number),
              label: item.label || "",
              difference: item.difference || "",
              reason: item.reason || ""
            }))
            .filter((item) => keptPages.includes(item.pageNumber))
        : []

      appState.aiRerankResult = {
        status: parsed.status || "single_match",
        relationship: parsed.relationship || "",
        bestPage,
        confidence: parsed.confidence || "",
        summary: parsed.summary || "",
        question: parsed.question || "",
        structureType: normalizeStructureType(parsed.structure_type),
        interactionModel: normalizeInteractionModel(parsed.interaction_model),
        structureConfidence: Number.isFinite(Number(parsed.structure_confidence)) ? Number(parsed.structure_confidence) : 0,
        hasConcreteProducts: parseModelBoolean(parsed.has_concrete_products),
        concreteProductCandidates: normalizeConcreteProductCandidates(parsed.concrete_product_candidates),
        keptPages,
        variantComparison,
        orderedPages
      }
      appState.aiRerankDocumentId = documentRecord.id
      appState.aiRerankCacheByDocumentId[documentRecord.id] = cloneAiRerankResult(appState.aiRerankResult)
      updateStructureRoutingState(documentRecord)
      appState.summaryPanelOpen = false

      if (appState.aiRerankResult.bestPage) {
        setPage(appState.aiRerankResult.bestPage)
      } else {
        render()
      }
      showToast("AI rerank complete")
    } catch (error) {
      appState.aiRerankDocumentId = ""
      updateStructureRoutingState(documentRecord)
      appState.aiRerankError = error instanceof Error ? error.message : "AI rerank failed."
      showToast("AI rerank failed")
      render()
    } finally {
      appState.aiRerankLoading = false
      render()
    }
  }

  async function analyzeSpecDecisions(targetPageNumber, selectionHints = {}) {
    const documentRecord = getActiveDocument()
    if (!documentRecord) return
    if (!appState.visionApiKey) {
      appState.decisionAssistError = "Add an OpenAI API key above before using spec decision analysis."
      showToast("Add an OpenAI API key first")
      render()
      return
    }

    try {
      appState.decisionAssistLoading = true
      appState.helpSpecLoadingLineIndex = 0
      appState.helpSpecLoadingLineOrder = buildHelpSpecLoadingLineOrder()
      appState.decisionAssistError = ""
      appState.decisionAssistResult = null
      appState.activeDecisionChoiceIndex = -1
      appState.decisionTabsWindowStart = 0
      startHelpSpecLoadingTicker()
      showToast("Analyzing spec decisions...")
      renderPreservingViewerScroll()

      const requestedPageNumber = Number.isFinite(Number(targetPageNumber)) ? Number(targetPageNumber) : appState.activePageNumber
      const selectedPage = documentRecord.pages.find((page) => page.pageNumber === requestedPageNumber)
      const parsingModeConfig = getSpecParsingModeConfig()
      if (!selectedPage) {
        throw new Error("Select a page first before running spec decision analysis.")
      }

      const selectedProductHint = normalizeText(selectionHints.selectedProductId || "")
      const selectedVariantHint = normalizeText(selectionHints.selectedVariantId || "")
      const fromProductFirst = selectionHints.fromProductFirst === true
      const candidatePages = getDecisionCandidatePages(documentRecord, selectedPage, selectionHints)
      if (!candidatePages.length) {
        throw new Error("No candidate pages were found for spec analysis.")
      }
      const candidateRangeLabel = formatPageRangeLabel(candidatePages.map((page) => page.pageNumber))

      const content = [
        {
          type: "input_text",
          text: [
            "You are helping a furniture spec editor understand how to spec a product from a manufacturer PDF.",
            `Brand: ${appState.spec.brandDisplayName || appState.spec.originalBrand || "Unknown"}`,
            `Product name: ${appState.spec.specDisplayName || appState.spec.originalSpecName || "Unknown"}`,
            `The user selected page ${selectedPage.pageNumber}. Candidate ${parsingModeConfig.pageLabel} pages are ${candidateRangeLabel}.`,
            parsingModeConfig.mode === "family"
              ? "Treat the clicked page as the entry point into a product family, not necessarily a single product."
              : "Treat the clicked page as the entry point into a product section, not necessarily the whole answer.",
            parsingModeConfig.mode === "family" && !selectedProductHint && !selectedVariantHint
              ? "At this stage, analyze only the clicked page. Do not scan forward and do not use later family pages yet."
              : "",
            "First determine whether the starting page contains one exact product, multiple related variants of one family, or multiple distinct products.",
            parsingModeConfig.mode === "family"
              ? "If the clicked page contains multiple distinct products or multiple variant rows for the same family, assume that page is self-contained for Help Me Spec and do not expand to other pages."
              : "",
            "If more than one distinct product appears on the starting page and no selected product hint is provided, return product candidates and do not generate characteristics yet.",
            "Product selection happens only on the starting page.",
            parsingModeConfig.mode === "family"
              ? "If a selected product hint is provided, use that product as the active context, but stay on the clicked page unless expansion is explicitly required by referenced shared specs."
              : "If a selected product hint is provided, use that product as the active context for the forward scan.",
            parsingModeConfig.mode === "family"
              ? "For this Davis-style family document, do not assume one page equals one product and do not use adjacent-page proximity as the main rule."
              : "Then determine the contiguous related page range by scanning forward only.",
            parsingModeConfig.mode === "family"
              ? "Default to the clicked page only. Expand included_pages beyond the clicked page only when the clicked page explicitly depends on shared specification pages, finish references, or powder coat references that are required to understand the selected result."
              : "Keep including the next page only while it still belongs to the same product selected from the starting page.",
            parsingModeConfig.mode === "family"
              ? "If the starting page contains multiple related variants and no selected variant hint is provided, return variant candidates before characteristics."
              : "Stop at the first page that is clearly not the same contiguous product section.",
            parsingModeConfig.mode === "family" && !selectedProductHint && !selectedVariantHint
              ? "Variant candidates must come only from what is actually shown on the clicked page."
              : "",
            parsingModeConfig.mode === "family" && !selectedVariantHint
              ? "When the clicked page lists concrete model rows or item codes, return those exact on-page variants as the variant candidates. Do not return abstract buckets such as Base Type, Upholstery Type, or Finish Type."
              : "",
            parsingModeConfig.mode === "family"
              ? "For vague family-level matches, assume the user wants the family first, then progressively narrow by product selector, variant selector, and finally spec tabs."
              : "If later pages are clearly separate variant pages with their own page-selector entries in the main UI, do not include them in this Help Me Spec result.",
            "Only return variant candidates when the starting page itself contains multiple variants of the same product or family that would materially change the downstream characteristic modules.",
            "Once the product and variant context is clear, extract the specification characteristics from the included product pages.",
            "If a selected variant hint is provided, scope all downstream characteristics to that chosen variant only.",
            "Do not include sibling variant options, dimensions, finishes, or pricing once a specific variant has been selected.",
            parsingModeConfig.mode === "family"
              ? "Important Davis example: if the family has one page with multiple variants and another page with shared finish or powder coat references, include both when they apply to the selected product or variant even if they are not adjacent."
              : "Important example: if page 86 is the selected page and page 87 is its continuation, include page 87. If pages 88 and 89 are clearly separate variant pages with their own page-selector entries, do not include them in the page 86 Help Me Spec result.",
            "Map content into meaningful characteristic types such as Size / Dimensions, Configuration, Shell / Frame Finish, Upholstery / Material, Surface / Top Material, Compliance / Certifications, or Options / Add-ons when those fit.",
            "Use more specific labels like Shell Finish or Frame Finish when the source supports it.",
            "Configuration tab rendering rules: use a single Configuration tab only when pricing depends on a combination of multiple structural variables.",
            "Examples of structural variables: Arms, Base Type, Swivel / Fixed, Orientation, Leg Type.",
            "If structural options have independent additive pricing, do not combine them into one Configuration tab. Treat them as separate characteristics.",
            "Configuration tab layout: title should be Configuration. Include a brief description of which structural variables affect pricing or form.",
            "Show each structural variable as its own labeled section. Example: Arms: No arms, Arms, Arms with arm pads. Base Type: Swivel base, Fixed base.",
            "If pricing depends on multiple structural variables, show a pricing matrix.",
            "Required pricing matrix format: rows = primary structural variable, columns = secondary structural variable, cells = final price.",
            "Do not show grouped prose like 'Base: Swivel or Fixed'.",
            "Do not show derived SKU codes as the main content, but they may be included as secondary metadata in pricing cells.",
            "Do not repeat the same options in multiple blocks.",
            "Do not use Step 1 / Step 2 language.",
            "If pricing depends on only one structural variable, show a simple labeled list instead of a matrix. Example: No arms: $0; Arms: +$120.",
            "Do not behave like a configurator. This is an organized reading layer.",
            "Each characteristic should include a short blurb, all detected options, raw pricing details if present, and dependency notes where applicable.",
            "For Size / Dimensions, return one labeled dimension per row inside each option's values field using semicolon-separated label:value pairs.",
            'Example: "A (Overall Depth): 35 in; B (Seat Height): 16 in; C (Overall Width): 33 1/2 in".',
            "Pricing-only sections such as Base Price, Price List, or Order Information are not characteristics and must go into reference_info instead.",
            "Reference-only information should go into reference_info.",
            "Do not round, estimate, or convert units. Preserve exact source numbers and notation (including decimals and fractions such as 16.5 in or 16 1/2 in).",
            "Do not invent unsupported information.",
            selectedProductHint ? `Selected product hint: ${selectedProductHint}` : "Selected product hint: none",
            selectedVariantHint ? `Selected variant hint: ${selectedVariantHint}` : "Selected variant hint: none",
            "Return JSON only with no markdown fences.",
            "",
            "Required JSON shape:",
            '{',
            '  "included_pages": [86, 87],',
            '  "product_context": {',
            '    "product_name": "Eames Lounge Chair and Ottoman",',
            '    "model_family": "Lounge Seating",',
            '    "summary": "short product-level summary",',
            '    "inclusion_summary": "why those pages were included",',
            '    "stop_reason": "why the section stopped"', 
            "  },",
            '  "selected_product_id": "",',
            '  "product_candidates": [',
            '    {"id": "es670", "label": "Eames Lounge Chair and Ottoman", "description": "chair and ottoman spec section", "evidence": "page title and hero image"}',
            "  ],",
            '  "selected_variant_id": "",',
            '  "variant_candidates": [',
            '    {"id": "upholstered", "label": "Upholstered", "description": "variant with upholstery pricing", "evidence": "variant callout in spec pages"}',
            "  ],",
            '  "characteristics": [',
            '    {',
            '      "id": "size-dimensions",',
            '      "label": "Size / Dimensions",',
            '      "blurb": "Dimensions and required leather square footage vary by size.",',
            '      "dependency_note": "",',
            '      "pricing_note": "",',
            '      "options": [',
            '        {"name": "Classic", "values": "A (Overall Depth): 35 in; B (Seat Height): 16 in; C (Overall Width): 33 1/2 in; D (Overall Height): 31 1/2 in; E (Seat Depth): 21 1/4 in", "pricing": "COL 55 sq ft", "difference": "Smaller footprint and lower back height", "evidence": "Dimension callouts near chair profile"},',
            '        {"name": "Tall", "values": "A (Overall Depth): 37 3/4 in; B (Seat Height): 16 1/2 in; C (Overall Width): 33 1/2 in; D (Overall Height): 33 1/4 in; E (Seat Depth): 23 1/4 in", "pricing": "COL 70 sq ft", "difference": "0.5 in higher seat height and 1.75 in taller overall height", "evidence": "Dimension callouts and COM table"}',
            "      ]",
            "    },",
            '    {',
            '      "id": "configuration",',
            '      "label": "Configuration",',
            '      "blurb": "Core structural options for the product.",',
            '      "configuration_sections": [',
            '        {"title": "No arms", "options": ["Swivel", "Fixed"]},',
            '        {"title": "Arms", "options": ["Swivel", "Fixed"]},',
            '        {"title": "Arms with arm pads", "options": ["Swivel", "Fixed"]}',
            '      ],',
            '      "pricing_matrix": {',
            '        "row_label": "Arms",',
            '        "column_labels": ["Swivel", "Fixed"],',
            '        "rows": [',
            '          {"row_name": "No arms", "cells": [{"column": "Swivel", "price": "$1909", "model": "EA306"}, {"column": "Fixed", "price": "$2027", "model": "EA306"}]},',
            '          {"row_name": "Arms", "cells": [{"column": "Swivel", "price": "$2490", "model": "EA308"}, {"column": "Fixed", "price": "$2605", "model": "EA308"}]}',
            '        ]',
            '      }',
            '    }',
            "  ],",
            '  "reference_info": ["Dimensions table is on page 86", "Chair and ottoman included"]',
            '}'
          ].join("\n")
        }
      ]

      const selectedPageText = candidatePages
        .map((page) => {
          const pageText = normalizeText(getPageCombinedText(page)).slice(0, 2200)
          return pageText ? `Page ${page.pageNumber} text:\n${pageText}` : ""
        })
        .filter(Boolean)
        .join("\n\n")

      if (selectedPageText) {
        content.push({
          type: "input_text",
          text: `Extracted text from candidate pages ${candidateRangeLabel}:\n${selectedPageText}`
        })
      }

      for (const page of candidatePages) {
        const canvas = await renderPdfPageToCanvas(documentRecord, page.pageNumber, 1.35)
        content.push({
          type: "input_text",
          text: `Candidate product page ${page.pageNumber}.`
        })
        content.push({
          type: "input_image",
          image_url: canvas.toDataURL("image/png")
        })
      }

      const response = await fetch("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${appState.visionApiKey}`
        },
        body: JSON.stringify({
          model: appState.visionModel || DEFAULT_VISION_MODEL,
          input: [{ role: "user", content }]
        })
      })

      if (!response.ok) {
        const errorBody = await response.text()
        throw new Error(`Decision analysis failed (${response.status}): ${errorBody.slice(0, 180)}`)
      }

      const payload = await response.json()
      const rawText = extractResponseText(payload)
      const parsed = parseJsonPayload(rawText)
      if (!parsed) {
        throw new Error("Decision analysis returned an unreadable response.")
      }

      const candidatePageSet = new Set(candidatePages.map((page) => page.pageNumber))
      const includedPagesRaw = Array.isArray(parsed.included_pages)
        ? parsed.included_pages.map((value) => Number(value)).filter((value) => Number.isFinite(value) && candidatePageSet.has(value))
        : []
      const sortedIncluded = [...new Set([selectedPage.pageNumber, ...includedPagesRaw])].sort((a, b) => a - b)
      const includedPages = parsingModeConfig.mode === "family"
        ? sortedIncluded
        : (() => {
            const pages = [selectedPage.pageNumber]
            let expectedPage = selectedPage.pageNumber + 1
            sortedIncluded.forEach((pageNumber) => {
              if (pageNumber === expectedPage) {
                pages.push(pageNumber)
                expectedPage += 1
              }
            })
            return pages
          })()
      const pageRangeLabel = formatPageRangeLabel(includedPages)
      const resolvedProductName = normalizeText(parsed.product_context?.product_name || "")
      const productCandidates = filterCandidatesToResolvedProduct(
        filterCandidatesToStartingPage(
          normalizeDecisionCandidates(parsed.product_candidates),
          selectedPage.pageNumber
        ),
        resolvedProductName
      )
      const rawVariantCandidates = preferConcretePageVariants(
        filterVariantCandidatesToSelectedPage(
          normalizeDecisionCandidates(parsed.variant_candidates),
          selectedPage,
          documentRecord
        ),
        selectedPage,
        selectedVariantHint || parsed.selected_variant_id || "",
        documentRecord
      )
      const characteristics = splitOvergroupedConfiguration(normalizeDecisionCharacteristics(parsed.characteristics))
      const includedPageRecords = includedPages
        .map((pageNumber) => documentRecord.pages.find((page) => page.pageNumber === pageNumber))
        .filter(Boolean)
      const variantCandidates = shouldTreatVariantCandidatesAsConfiguration(rawVariantCandidates, includedPageRecords)
        ? []
        : (
          fromProductFirst && parsingModeConfig.mode === "family"
            ? suppressProductFirstSiblingVariants(rawVariantCandidates, selectedProductHint, appState.structureRouting?.productCandidates || [])
            : rawVariantCandidates
        )
      const supplementalCharacteristics = buildSupplementalCharacteristics(documentRecord, includedPageRecords, characteristics)
      const referenceInfo = Array.isArray(parsed.reference_info) ? parsed.reference_info : []
      const productScopedCharacteristics = fromProductFirst && selectedProductHint
        ? filterCharacteristicsForSelectedRecord(
            enrichCharacteristicsFromSectionPages(
              mergeReferenceTextileIntoDimensions(
                removeTextileRequirementOptionsFromUpholstery(
                  mergeComColIntoDimensions([...characteristics, ...supplementalCharacteristics])
                ),
                referenceInfo
              ),
              includedPageRecords
            ),
            appState.structureRouting?.productCandidates || [],
            selectedProductHint
          )
        : enrichCharacteristicsFromSectionPages(
            mergeReferenceTextileIntoDimensions(
              removeTextileRequirementOptionsFromUpholstery(
                mergeComColIntoDimensions([...characteristics, ...supplementalCharacteristics])
              ),
              referenceInfo
            ),
            includedPageRecords
          )
      const mergedCharacteristics = sortCharacteristicsForDisplay(
        filterCharacteristicsForSelectedVariant(
          productScopedCharacteristics,
          variantCandidates,
          selectedVariantHint || parsed.selected_variant_id || ""
        )
      )
      const hydratedCharacteristics = await fillMissingDimensionOptions(
        documentRecord,
        includedPageRecords,
        mergedCharacteristics,
        variantCandidates,
        selectedVariantHint || parsed.selected_variant_id || ""
      )
      const prunedCharacteristics = pruneResolvedConfigurationCharacteristics(
        hydratedCharacteristics,
        selectedVariantHint || parsed.selected_variant_id || ""
      )
      const resolvedSelectedProductId = normalizeText(parsed.selected_product_id || selectedProductHint || (productCandidates.length === 1 ? productCandidates[0].id : ""))
      const resolvedSelectedVariantId = normalizeText(parsed.selected_variant_id || selectedVariantHint || (variantCandidates.length === 1 ? variantCandidates[0].id : ""))

      appState.decisionAssistResult = {
        pageNumber: selectedPage.pageNumber,
        parsingMode: parsingModeConfig.mode,
        fromProductFirst,
        pageRangeLabel,
        includedPages,
        productName: resolvedProductName,
        modelFamily: normalizeText(parsed.product_context?.model_family || ""),
        summary: normalizeText(parsed.product_context?.summary || ""),
        inclusionSummary: normalizeText(parsed.product_context?.inclusion_summary || ""),
        stopReason: normalizeText(parsed.product_context?.stop_reason || ""),
        productCandidates,
        selectedProductId: resolvedSelectedProductId,
        variantCandidates,
        selectedVariantId: resolvedSelectedVariantId,
        characteristics: prunedCharacteristics,
        referenceInfo
      }
      appState.activeDecisionChoiceIndex = prunedCharacteristics.length ? 0 : -1
      appState.decisionTabsWindowStart = 0

      renderPreservingViewerScroll()
      showToast(
        productCandidates.length > 1 && !resolvedSelectedProductId
          ? "Select the product to continue"
          : variantCandidates.length > 1 && !resolvedSelectedVariantId
            ? "Select the variant to continue"
            : `Spec details ready (pages ${pageRangeLabel})`
      )
    } catch (error) {
      appState.decisionAssistError = error instanceof Error ? error.message : "Decision analysis failed."
      showToast("Spec decision analysis failed")
      renderPreservingViewerScroll()
    } finally {
      appState.decisionAssistLoading = false
      stopHelpSpecLoadingTicker()
      renderPreservingViewerScroll()
    }
  }

  async function analyzePageImage(pageNumber) {
    const documentRecord = getActiveDocument()
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    if (!documentRecord || !page) return

    try {
      appState.ocrStatusByPage[pageNumber] = "running"
      appState.ocrErrorByPage[pageNumber] = ""
      render()
      const canvas = await renderPdfPageToCanvas(documentRecord, pageNumber, 2)
      const result = await analyzeDimensionsWithVision(pageNumber, canvas)
      page.ocrText = result.formattedText || normalizeText(result.rawText || "")
      appState.ocrStatusByPage[pageNumber] = "done"
      appState.ocrErrorByPage[pageNumber] = ""
      runRanking(false)
      appState.activePageNumber = pageNumber
      render()
      saveCurrentAsDefault().catch(() => {})
      showToast(page.ocrText ? `Vision analysis added to page ${pageNumber}` : `No image text found on page ${pageNumber}`)
    } catch (error) {
      appState.ocrStatusByPage[pageNumber] = "error"
      appState.ocrErrorByPage[pageNumber] = error instanceof Error ? error.message : "Image analysis failed."
      render()
      showToast(`Image analysis failed for page ${pageNumber}`)
    }
  }

  async function revealPageText(pageNumber) {
    const documentRecord = getActiveDocument()
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    if (!documentRecord || !page) return

    const renderKey = getPageRenderKey(documentRecord, pageNumber)
    appState.pageTextVisibleByKey[renderKey] = true
    renderPreservingViewerScroll()

    if (!appState.visionApiKey || appState.pageReadableHtmlByKey[renderKey] || appState.pageReadableStatusByKey[renderKey] === "loading") {
      return
    }

    await formatPageTextWithAi(pageNumber)
  }

  async function formatPageTextWithAi(pageNumber) {
    const documentRecord = getActiveDocument()
    const page = documentRecord?.pages.find((item) => item.pageNumber === pageNumber)
    if (!documentRecord || !page) return
    if (!appState.visionApiKey) {
      showToast("Add an OpenAI API key first")
      renderPreservingViewerScroll()
      return
    }

    const renderKey = getPageRenderKey(documentRecord, pageNumber)
    const rawText = getPageCombinedText(page)
    if (!normalizeText(rawText)) {
      appState.pageReadableErrorByKey[renderKey] = "No extractable page text was found."
      renderPreservingViewerScroll()
      return
    }

    try {
      appState.pageReadableStatusByKey[renderKey] = "loading"
      appState.pageReadableErrorByKey[renderKey] = ""
      renderPreservingViewerScroll()

      const prompt = [
        "You are cleaning up extracted PDF text for a furniture spec editor.",
        "Reformat the text so it is easier to scan and copy from.",
        "Preserve the exact wording, numbers, dimensions, and factual content from the source text.",
        "Do not summarize. Do not add facts. Do not delete facts.",
        "You may only improve structure: restore headings, combine broken lines into paragraphs, and convert obvious lists into bullet lists.",
        "If the source text is ambiguous, keep it literal rather than guessing.",
        "Return JSON only with no markdown fences.",
        'Use only these HTML tags inside the "html" field: h3, h4, p, ul, ol, li, strong, em, br.',
        "",
        "Required JSON shape:",
        '{',
        '  "html": "<h3>Section</h3><p>Readable paragraph...</p>"',
        '}',
        "",
        `Page ${pageNumber} extracted text:`,
        rawText
      ].join("\n")

      const response = await fetch("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${appState.visionApiKey}`
        },
        body: JSON.stringify({
          model: appState.visionModel || DEFAULT_VISION_MODEL,
          input: [{ role: "user", content: [{ type: "input_text", text: prompt }] }]
        })
      })

      if (!response.ok) {
        const errorBody = await response.text()
        throw new Error(`AI formatting failed (${response.status}): ${errorBody.slice(0, 180)}`)
      }

      const payload = await response.json()
      const rawResponseText = extractResponseText(payload)
      const parsed = parseJsonPayload(rawResponseText)
      const sanitizedHtml = sanitizeReadableHtml(parsed?.html || "")
      if (!sanitizedHtml) {
        throw new Error("AI formatting returned an unreadable result.")
      }

      appState.pageReadableHtmlByKey[renderKey] = sanitizedHtml
      appState.pageReadableStatusByKey[renderKey] = "done"
      appState.pageReadableErrorByKey[renderKey] = ""
      renderPreservingViewerScroll()
      showToast(`Formatted page ${pageNumber}`)
    } catch (error) {
      appState.pageReadableStatusByKey[renderKey] = "error"
      appState.pageReadableErrorByKey[renderKey] = error instanceof Error ? error.message : "AI formatting failed."
      renderPreservingViewerScroll()
      showToast(`Formatting failed for page ${pageNumber}`)
    }
  }

  function extractTextFromContentItems(items) {
    const rows = new Map()

    items.forEach((item) => {
      if (!item.str || !item.str.trim()) return
      const y = Math.round((item.transform?.[5] || 0) / 4) * 4
      const x = item.transform?.[4] || 0
      if (!rows.has(y)) rows.set(y, [])
      rows.get(y).push({ x, text: item.str })
    })

    return [...rows.entries()]
      .sort((a, b) => b[0] - a[0])
      .map(([, rowItems]) =>
        rowItems
          .sort((a, b) => a.x - b.x)
          .map((item) => item.text)
          .join(" ")
      )
      .join("\n")
  }

  function clearSelectionState() {
    appState.selectionText = ""
    appState.selectionSource = ""
    appState.selectionRect = null
    appState.pickerPosition = null
    appState.confirmPosition = null
    appState.pendingFieldKey = ""
    appState.pendingInsertText = ""
    updateFloatingCopyButton()
  }

  function analyzeRankedPages(documentRecord, rankedPages) {
    const top = rankedPages[0]
    if (!documentRecord || !top) {
      return null
    }

    if (getSpecParsingMode() === "family") {
      const topPages = rankedPages.slice(0, 4).map((page) => ({
        ...page,
        text: getPageCombinedText(documentRecord.pages.find((candidate) => candidate.pageNumber === page.pageNumber))
      }))
      const visualSignals = summarizeVisualSignals(topPages)
      const options = extractAmbiguityOptions(topPages)
      const second = rankedPages[1]
      const weakTopMatch = top.score < 24
      const closeCompetition = second && top.score - second.score <= 8

      if (weakTopMatch) {
        return {
          status: "no_match",
          title: "No strong family match yet",
          confidenceLabel: "Needs more context",
          recommendedPage: 1,
          reason: "The current product-name signals are not distinctive enough to isolate the right family pages with confidence.",
          question: "Try a more specific family, product, or variant name.",
          options: [],
          visualSignals,
          topPages: []
        }
      }

      if (!appState.assistantSelection && options.length >= 2 && closeCompetition) {
        return {
          status: "ambiguous",
          title: "Multiple family matches detected",
          confidenceLabel: "Ambiguous family match",
          recommendedPage: top.pageNumber,
          reason: "Several non-adjacent pages look relevant to the same search. Pick the product or variant that matches, then the viewer will retarget that family context.",
          question: "Select the model, product, or variant that best matches the family you want.",
          options,
          visualSignals,
          topPages: rankedPages.slice(0, 3).map((page) => page.pageNumber)
        }
      }

      return {
        status: "match",
        title: "Likely family pages found",
        confidenceLabel: top.score >= 90 ? "High confidence" : "Medium confidence",
        recommendedPage: top.pageNumber,
        reason: "The strongest matches are based on family relevance and shared reference signals, not adjacent-page clustering.",
        question: "",
        options: [],
        visualSignals,
        topPages: rankedPages.slice(0, 3).map((page) => page.pageNumber)
      }
    }

    const clusters = buildPageClusters(rankedPages)
    const primaryCluster = clusters[0]
    const secondCluster = clusters[1]
    const clusterPages = primaryCluster.pages.map((page) => ({
      ...page,
          text: getPageCombinedText(documentRecord.pages.find((candidate) => candidate.pageNumber === page.pageNumber))
        }))
    const visualSignals = summarizeVisualSignals(clusterPages)
    const options = extractAmbiguityOptions(clusterPages)
    const weakTopMatch = top.score < 24
    const closeCompetition = secondCluster && primaryCluster.totalScore - secondCluster.totalScore <= 20

    if (weakTopMatch || !primaryCluster) {
      return {
        status: "no_match",
        title: "No strong product section yet",
        confidenceLabel: "Needs more context",
        recommendedPage: 1,
        reason: "The current product-name signals are not distinctive enough to isolate a likely product section with confidence.",
        question: "Try a more specific product name, family name, or variant.",
        options: [],
        visualSignals,
        topPages: []
      }
    }

    if (!appState.assistantSelection && options.length >= 2 && closeCompetition) {
      return {
        status: "ambiguous",
        title: "Multiple likely product sections detected",
        confidenceLabel: "Ambiguous section match",
        recommendedPage: primaryCluster.startPage,
        reason: `The strongest candidates form more than one nearby page cluster. The current best section is ${formatPageRange(primaryCluster.startPage, primaryCluster.endPage)}, but another section is competitive.`,
        question: "Select the model, variant, or descriptor that best matches the product, then the viewer will retarget that section.",
        options,
        visualSignals,
        topPages: primaryCluster.pages.map((page) => page.pageNumber).slice(0, 3)
      }
    }

    if (appState.assistantSelection) {
      return {
        status: "match",
        title: "Product section retargeted",
        confidenceLabel: "User-confirmed",
        recommendedPage: primaryCluster.startPage,
        reason: `The selected clarifier "${appState.assistantSelection}" pushed ${formatPageRange(primaryCluster.startPage, primaryCluster.endPage)} to the top of the retrieval results.`,
        question: "",
        options: [],
        visualSignals,
        topPages: primaryCluster.pages.map((page) => page.pageNumber).slice(0, 3)
      }
    }

    return {
      status: "match",
      title: "Likely product section found",
      confidenceLabel: primaryCluster.totalScore >= 90 ? "High confidence" : "Medium confidence",
      recommendedPage: primaryCluster.startPage,
      reason: `The best match is the section spanning ${formatPageRange(primaryCluster.startPage, primaryCluster.endPage)}. Retrieval is based on product-name signals, then the page analyst labels the section contents.`,
      question: "",
      options: [],
      visualSignals,
      topPages: primaryCluster.pages.map((page) => page.pageNumber).slice(0, 3)
    }
  }

  function analyzePageForVisualCues(pageText) {
    const text = pageText || ""
    const measurementMatches = text.match(/\b\d+(?:\s+\d\/\d)?(?:["”]| in\b| inches\b)?\b/g) || []
    const hasDimensionKeywords = /dimensions?|overall width|overall depth|overall height|seat height|dia\b/i.test(text)
    const hasSpecKeywords = /notes|description|finish|material|frame|upholstery|glides|veneer/i.test(text)
    const shortLineCount = text
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .filter((line) => line.length <= 24).length

    const likelyDrawingPage = measurementMatches.length >= 4 && (hasDimensionKeywords || shortLineCount >= 6)
    const likelySpecPage = hasSpecKeywords && (measurementMatches.length >= 2 || hasDimensionKeywords)
    const likelyIndexPage = /introduction|index|contents|price book|seating \d+/i.test(text) && measurementMatches.length < 3

    let primaryLabel = "Text-heavy page"
    if (likelyIndexPage) primaryLabel = "Index or navigation page"
    else if (likelyDrawingPage) primaryLabel = "Likely dimensions drawing"
    else if (likelySpecPage) primaryLabel = "Likely product spec page"

    const labels = []
    if (likelyDrawingPage) labels.push("Likely dimensions drawing")
    if (likelySpecPage) labels.push("Spec content detected")
    if (likelyIndexPage) labels.push("Index-like structure")
    if (!labels.length) labels.push("General content page")

    return {
      primaryLabel,
      labels
    }
  }

  function summarizeVisualSignals(candidatePages) {
    const signalSet = new Set()
    candidatePages.forEach((page) => {
      analyzePageForVisualCues(page.text).labels.forEach((label) => signalSet.add(label))
    })
    return [...signalSet].slice(0, 4)
  }

  async function schedulePreviewRendering() {
    const documentRecord = getActiveDocument()
    if (!documentRecord || documentRecord.sourceType !== "uploaded" || !documentRecord.pdfBytes) return

      const pagesToRender = appState.rankedPages.slice(0, 3).map((page) => page.pageNumber)
    if (!pagesToRender.length) return

    try {
      const pdfjsLib = await loadPdfJsLibrary()
      const loadingTask = pdfjsLib.getDocument({ data: Uint8Array.from(documentRecord.pdfBytes) })
      const pdfDocument = await loadingTask.promise

      for (const pageNumber of pagesToRender) {
        const previewKey = `${documentRecord.id}:${pageNumber}`
        if (appState.pagePreviews[previewKey]) continue
        const page = await pdfDocument.getPage(pageNumber)
        const viewport = page.getViewport({ scale: 0.32 })
        const canvas = document.createElement("canvas")
        const context = canvas.getContext("2d")
        canvas.width = Math.ceil(viewport.width)
        canvas.height = Math.ceil(viewport.height)
        await page.render({ canvasContext: context, viewport }).promise
        appState.pagePreviews[previewKey] = canvas.toDataURL("image/png")
      }

      render()
    } catch (error) {
      // Preview rendering should fail quietly; ranking should still work without thumbnails.
    }
  }

  async function scheduleVisiblePageRendering() {
    const documentRecord = getActiveDocument()
    if (!documentRecord || !documentRecord.pdfBytes) return

    const visiblePageNumbers = getVisiblePages(documentRecord).map((page) => page.pageNumber)
    if (!visiblePageNumbers.length) return

    for (const pageNumber of visiblePageNumbers) {
      const renderKey = getPageRenderKey(documentRecord, pageNumber)
      if (appState.pageRenderImages[renderKey] || appState.pageRenderStatusByKey[renderKey] === "loading") continue

      appState.pageRenderStatusByKey[renderKey] = "loading"
      render()

      try {
        const pageRender = await getPdfPageRenderData(documentRecord, pageNumber, 1.25)
        appState.pageRenderImages[renderKey] = pageRender.canvas.toDataURL("image/png")
        appState.pageRenderTextByKey[renderKey] = pageRender.textContent
        appState.pageRenderMetricsByKey[renderKey] = {
          width: pageRender.viewport.width,
          height: pageRender.viewport.height,
          viewport: pageRender.viewport
        }
        appState.pageRenderStatusByKey[renderKey] = "done"
      } catch (error) {
        appState.pageRenderStatusByKey[renderKey] = "error"
      }

      render()
    }
  }

  function syncPageTextLayerScale(layer) {
    if (!layer) return
    const inner = layer.firstElementChild
    const pageWidth = Number(layer.getAttribute("data-page-width"))
    const pageHeight = Number(layer.getAttribute("data-page-height"))
    if (!inner || !pageWidth || !pageHeight) return

    const scale = layer.clientWidth ? layer.clientWidth / pageWidth : 1
    inner.style.width = `${pageWidth}px`
    inner.style.height = `${pageHeight}px`
    inner.style.transform = `scale(${scale})`
    layer.style.height = `${pageHeight * scale}px`
  }

  async function hydratePageTextLayer(layer) {
    if (!layer) return

    const renderKey = layer.getAttribute("data-page-text-layer-key")
    const textContent = appState.pageRenderTextByKey[renderKey]
    const renderMetrics = appState.pageRenderMetricsByKey[renderKey]
    if (!renderKey || !textContent || !renderMetrics?.viewport) return

    if (!layer.firstElementChild) {
      const pdfjsLib = await loadPdfJsLibrary()
      if (typeof pdfjsLib.renderTextLayer !== "function") return

      const inner = document.createElement("div")
      inner.className = "page-text-layer-inner"
      layer.replaceChildren(inner)

      let renderTask
      try {
        renderTask = pdfjsLib.renderTextLayer({
          container: inner,
          textContentSource: textContent,
          viewport: renderMetrics.viewport,
          textDivs: []
        })
      } catch (error) {
        renderTask = pdfjsLib.renderTextLayer({
          container: inner,
          textContent,
          viewport: renderMetrics.viewport,
          textDivs: []
        })
      }

      if (renderTask?.promise) {
        await renderTask.promise
      } else if (typeof renderTask?.then === "function") {
        await renderTask
      }
    }

    syncPageTextLayerScale(layer)
  }

  function schedulePageTextLayerHydration() {
    requestAnimationFrame(() => {
      document.querySelectorAll(".page-text-layer").forEach((layer) => {
        hydratePageTextLayer(layer).catch(() => {})
      })
    })
  }

  function extractAmbiguityOptions(candidatePages) {
    const optionSet = new Set()

    candidatePages.forEach((page) => {
      const text = page.text || ""
      const modelMatches = text.match(/\b(?:[A-Z]{1,4}-?)?\d{2,}[A-Z0-9-]*\b/g) || []
      modelMatches.forEach((match) => {
        if (match.length >= 3 && !/^\d+$/.test(match)) optionSet.add(match)
      })

      ;[
        "high-back",
        "low-back",
        "mid-back",
        "lounge chair",
        "lounge seating",
        "settee",
        "ottoman",
        "armchair",
        "side chair"
      ].forEach((phrase) => {
        if (text.toLowerCase().includes(phrase)) {
          optionSet.add(phrase.replace(/\b\w/g, (letter) => letter.toUpperCase()))
        }
      })
    })

    return [...optionSet].slice(0, 6)
  }

  function handleSelectionChange() {
    const selection = window.getSelection()
    if (!selection || selection.rangeCount === 0) {
      clearSelectionState()
      return
    }

    const selectedText = normalizeText(selection.toString())
    const anchorContainer = selection.anchorNode?.parentElement?.closest(".pdf-page, .help-spec-card")
    const focusContainer = selection.focusNode?.parentElement?.closest(".pdf-page, .help-spec-card")
    if (!selectedText || !anchorContainer || anchorContainer !== focusContainer) {
      clearSelectionState()
      return
    }

    const rect = selection.getRangeAt(0).getBoundingClientRect()
    setSelectionTarget(selectedText, anchorContainer.classList.contains("pdf-page") ? "pdf" : "assistant", rect)
  }

  function setSelectionTarget(text, source, rect) {
    appState.selectionText = normalizeText(text)
    appState.selectionSource = source
    appState.selectionRect = rect
      ? {
          top: rect.top,
          left: rect.left,
          width: rect.width || 0,
          height: rect.height || 0
        }
      : null
    appState.pickerPosition = null
    appState.confirmPosition = null
    updateFloatingCopyButton()
  }

  async function copyActiveSelection() {
    const copiedText = normalizeText(appState.selectionText)
    if (!copiedText) return
    appState.copiedText = copiedText
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(copiedText)
      }
    } catch (error) {
      // Keep the in-app copied value even if clipboard permissions fail.
    }
    clearSelectionState()
    showToast("Copied for paste")
    render()
  }

  function getFloatingCopyButton() {
    if (floatingCopyButton?.isConnected) return floatingCopyButton

    floatingCopyButton = document.createElement("button")
    floatingCopyButton.type = "button"
    floatingCopyButton.className = "floating-copy-btn"
    floatingCopyButton.setAttribute("aria-label", "Copy selected text")
    floatingCopyButton.textContent = "Copy"
    floatingCopyButton.hidden = true
    floatingCopyButton.addEventListener("mousedown", (event) => {
      event.preventDefault()
    })
    floatingCopyButton.addEventListener("click", () => {
      copyActiveSelection().catch(() => {})
    })
    document.body.appendChild(floatingCopyButton)
    return floatingCopyButton
  }

  function updateFloatingCopyButton() {
    const button = getFloatingCopyButton()
    if (!appState.selectionRect || !appState.selectionText) {
      button.hidden = true
      return
    }

    const top = Math.min(window.innerHeight - 56, Math.max(16, appState.selectionRect.top + appState.selectionRect.height + 12))
    const left = Math.min(window.innerWidth - 56, Math.max(28, appState.selectionRect.left + ((appState.selectionRect.width || 0) / 2)))
    button.style.top = `${top}px`
    button.style.left = `${left}px`
    button.hidden = false
  }

  function handleFieldPick(fieldKey) {
    const attribute = appState.spec.attributes.find((item) => item.key === fieldKey)
    const insertText = normalizeText(appState.selectionText)
    if (!attribute || !insertText) {
      clearSelectionState()
      render()
      return
    }

    if (attribute.value.trim()) {
      appState.pendingFieldKey = fieldKey
      appState.pendingInsertText = insertText
      appState.pickerPosition = null
      appState.confirmPosition = {
        top: Math.min(window.innerHeight - 240, (appState.selectionRect?.top || 120) + 36),
        left: Math.min(window.innerWidth - 360, (appState.selectionRect?.left || 32) + 12)
      }
      render()
      return
    }

    insertIntoField(fieldKey, insertText, "replace")
  }

  function resolveInsert(mode) {
    if (mode === "cancel") {
      clearSelectionState()
      render()
      return
    }

    insertIntoField(appState.pendingFieldKey, appState.pendingInsertText, mode)
  }

  function insertIntoField(fieldKey, insertText, mode) {
    const attribute = appState.spec.attributes.find((item) => item.key === fieldKey)
    if (!attribute) return

    attribute.value =
      mode === "append" && attribute.value.trim()
        ? `${attribute.value.trim()} ${insertText}`.trim()
        : insertText

    clearSelectionState()
    highlightField(fieldKey)
    showToast(`Added to ${attribute.label}`)
    render()

    requestAnimationFrame(() => {
      const input = document.querySelector(`[data-field-input="${fieldKey}"]`)
      input?.focus()
      input?.scrollIntoView({ behavior: "smooth", block: "center" })
    })
  }

  function highlightField(fieldKey) {
    appState.highlightedFieldKey = fieldKey
    window.clearTimeout(appState.highlightTimeoutId)
    appState.highlightTimeoutId = window.setTimeout(() => {
      appState.highlightedFieldKey = ""
      render()
    }, 700)
  }

  function showToast(message) {
    appState.toastMessage = message
    window.clearTimeout(appState.toastTimeoutId)
    appState.toastTimeoutId = window.setTimeout(() => {
      appState.toastMessage = ""
      render()
    }, 2200)
  }

  function downloadDocument(documentId) {
    const documentRecord = appState.documents.find((document) => document.id === documentId)
    if (!documentRecord || !documentRecord.pdfText) return

    const blob = new Blob([documentRecord.pdfText], { type: "application/pdf" })
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement("a")
    anchor.href = url
    anchor.download = `${documentRecord.id}.pdf`
    anchor.click()
    URL.revokeObjectURL(url)
  }

  function buildAsciiPdf(pages) {
    return [
      "%PDF-1.4",
      ...pages.map((pageText, index) => {
        const pdfLines = pageText
          .split("\n")
          .map((line) => line.replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)"))
        return `%%PAGE:${index + 1}
BT
${pdfLines.map((line) => `(${line}) Tj`).join("\n")}
ET`
      }),
      "%%EOF"
    ].join("\n")
  }

  function buildSampleDocuments() {
    return [
      {
        id: "simple-product-sheet",
        title: "Geiger Reframe One-Page Sheet",
        description: "A simple one-page sheet to validate direct entry with minimal navigation.",
        useCase: "One-page PDF",
        pageTitles: ["Reframe Product Sheet"],
        pdfText: buildAsciiPdf([
          [
            "GEIGER",
            "Reframe High-Back Lounge Chair",
            "Overall Width: 33 in",
            "Overall Depth: 31.5 in",
            "Overall Height: 42 in",
            "Finish Options: Walnut veneer, espresso oak",
            "Frame: Formed steel frame with upholstered shell",
            "Material: Hardwood shell, steel base, layered foam",
            "Notes: Companion ottoman available"
          ].join("\n")
        ])
      },
      {
        id: "reframe-catalog",
        title: "Geiger Lounge Collection Catalog",
        description: "A longer manufacturer-style catalog where the target product appears after introductory pages.",
        useCase: "Long PDF, target after page 1",
        pageTitles: [
          "Brand Introduction",
          "Collection Index",
          "Reframe High-Back Lounge Chair",
          "Reframe Finishes",
          "Warranty"
        ],
        pdfText: buildAsciiPdf([
          [
            "GEIGER",
            "Crafted lounge and workplace products.",
            "This catalog introduces seating, tables, and private office collections.",
            "Materials, finishes, and dimensions vary by family."
          ].join("\n"),
          [
            "GEIGER COLLECTION INDEX",
            "Crosshatch Lounge",
            "Filo Sofa",
            "Reframe Lounge Seating",
            "C-Side Tables",
            "Occasional Chairs"
          ].join("\n"),
          [
            "GEIGER",
            "Reframe High-Back Lounge Chair",
            "Reframe Lounge Seating family",
            "Overall Width: 33 in",
            "Overall Depth: 31.5 in",
            "Overall Height: 42 in",
            "Seat Height: 17 in",
            "Frame: Sculpted steel base with wood reveal detail",
            "Material: Molded foam, hardwood shell, steel frame"
          ].join("\n"),
          [
            "GEIGER",
            "Reframe Finish Options",
            "Wood finish options: White oak, walnut, ebonized ash",
            "Base finish options: Polished aluminum, blackened steel",
            "Upholstery: Fabric or leather by graded-in program"
          ].join("\n"),
          [
            "GEIGER",
            "Warranty and testing details",
            "BIFMA testing references",
            "Regional fabric approvals"
          ].join("\n")
        ])
      },
      {
        id: "naming-variant-catalog",
        title: "Geiger Seating Reference Book",
        description: "Tests fuzzy matching when the display name differs from the wording in the document.",
        useCase: "Display name differs",
        pageTitles: ["Front Matter", "Reframe High Back Chair", "Dimension Table", "Materials"],
        pdfText: buildAsciiPdf([
          [
            "GEIGER SEATING REFERENCE",
            "Overview of lounge and guest seating products."
          ].join("\n"),
          [
            "GEIGER",
            "Reframe High Back Chair",
            "Part of the Reframe Lounge Seating series.",
            "High-back silhouette with optional headrest wrap."
          ].join("\n"),
          [
            "GEIGER",
            "Reframe dimensional data",
            "Overall width 33 in",
            "Overall depth 31.5 in",
            "Overall height 42 in"
          ].join("\n"),
          [
            "GEIGER",
            "Material and finish reference",
            "Walnut veneer shell, blackened steel base, contract textile upholstery"
          ].join("\n")
        ])
      }
    ]
  }

  document.addEventListener("selectionchange", () => {
    const active = document.activeElement
    if (active && /INPUT|TEXTAREA/.test(active.tagName)) return
    window.clearTimeout(selectionChangeTimeoutId)
    selectionChangeTimeoutId = window.setTimeout(() => {
      handleSelectionChange()
    }, 120)
  })

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      clearSelectionState()
      render()
    }
  })

  window.addEventListener("resize", () => {
    renderPreservingViewerScroll()
    document.querySelectorAll(".page-text-layer").forEach((layer) => {
      syncPageTextLayerScale(layer)
    })
    updateFloatingCopyButton()
  })

  render()
  hydrateSavedDraftState()
})()
