import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import * as pdfjsLib from 'pdfjs-dist'
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import { createWorker, PSM } from 'tesseract.js'
import {
  FileSpreadsheet, UploadCloud, Download, Settings2, Plus, Trash2,
  CheckCircle2, AlertCircle, ChevronRight, Database, RotateCcw, LogOut, Mail, KeyRound,
  PlayCircle, ShieldCheck
} from 'lucide-react'
import { supabase } from './supabase'
import './styles.css'

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker

const BASE_MODELS = [
  { id: 'piraquara', name: 'Rede Piraquara', description: 'Quantidade direta; ignora a seção Trocas Pendentes', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'ouro-branco', name: 'Ouro Branco', description: 'Quantidade final calculada por Embalagem × Qtde', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'multiply', fixed: true },
  { id: 'adega-brasil', name: 'WG Adega Brasil', description: 'Quantidade direta; embalagem é apenas a apresentação do produto', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'dalpar', name: 'Dalpar', description: 'Pedido recebido como imagem; separação por marca e produto', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'flex-cbn', name: 'Orçamento Flex CBN', description: 'Prefixos RB, LO e SB identificam a indústria', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'personalizado', name: 'Modelo variável', description: 'Leitura completa para Excel, PDF, foto ou print, com revisão antes da exportação', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: false },
  { id: 'confronto', name: 'Confronto de arquivos', description: 'Compare o orçamento enviado com o pedido recebido', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
]

const PIRAQUARA_EXECUTION = {
  id: 'executar-piraquara',
  name: 'Executar pedidos Piraquara',
  description: 'Registre a execução, deixe uma observação opcional e acompanhe o andamento.',
  fixed: true,
}

const PIRAQUARA_DEFAULT_INSTRUCTION = 'execute todos os pedidos novos da Rede Piraquara em ordem cronológica, começando pelo mais antigo.'
const buildPiraquaraCommand = (instruction, requestId) => `Execução pedidos Piraquara: ${instruction?.trim() || PIRAQUARA_DEFAULT_INSTRUCTION}${requestId ? ` Solicitação no painel: ${requestId}.` : ''}`

const EXECUTION_STATUS = {
  SOLICITADO: { title: 'Solicitação registrada', text: 'Cole o comando curto no ChatGPT Work para iniciar.', tone: 'waiting' },
  EM_EXECUCAO: { title: 'Pedidos sendo executados', text: 'Acompanhe esta tela. Ela será atualizada automaticamente.', tone: 'running' },
  AGUARDANDO_REVISAO: { title: 'Pedidos prontos para conferência', text: 'Os carrinhos estão prontos. Confira antes de concluir qualquer pedido.', tone: 'ready' },
  CONCLUIDO: { title: 'Execução concluída', text: 'O processamento desta solicitação foi concluído.', tone: 'ready' },
  ERRO: { title: 'Execução com pendência', text: 'Confira a observação abaixo e peça a continuação pelo chat.', tone: 'error' },
}

const normalize = value => String(value ?? '').trim().toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')

const inferIndustryFromItem = itemName => {
  const item = normalize(itemName)
  if (/scotch|ponjita|nexcare|post.?it|command|3m\b/.test(item)) return '3M'
  if (/elseve|loreal|l'oreal|niely|garnier|maybelline/.test(item)) return 'LOREAL'
  if (/veja|vanish|finish|harpic|sbp\b|lysoform|destac|repelex/.test(item)) return 'RECKITT'
  return 'NAO_IDENTIFICADA'
}

const normalizeCatalogCode = value => String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '')

const isMasterCatalogRecord = record => normalize(record?.origem).includes('codigos que trabalho')

const buildTrustedCatalogMaps = records => {
  const industryMap = {}
  const codeMap = {}
  for (const item of records.filter(isMasterCatalogRecord)) {
    const catalogEan = cleanEan(item.ean)
    const catalogCode = normalizeCatalogCode(item.codigo_cbn)
    if (catalogEan) {
      industryMap[catalogEan] = item.industria
      // Alguns PDFs completam o UPC de 12 dígitos com um zero à esquerda.
      // O alias serve apenas para classificar; o EAN exportado não é alterado.
      if (catalogEan.length === 12) industryMap[`0${catalogEan}`] = item.industria
    }
    if (catalogCode && catalogEan) codeMap[catalogCode] = { EAN: catalogEan, Industria: item.industria }
  }
  return { industryMap, codeMap }
}

const resolveTrustedIndustry = (item, industryMap, codeMap = {}) => {
  const ean = cleanEan(item?.EAN)
  if (ean && industryMap[ean]) return industryMap[ean]
  if (!ean) {
    const codeMatch = codeMap[normalizeCatalogCode(item?.CodigoInterno)]
    if (codeMatch?.Industria) return codeMatch.Industria
  }
  const declared = String(item?.Industria || '').toUpperCase()
  if (declared === 'LOREAL' || declared === '3M') return declared
  const inferred = inferIndustryFromItem(item?.Item)
  return inferred === 'RECKITT' ? 'NAO_IDENTIFICADA' : inferred
}

const enrichRowsFromTrustedCatalog = (rows, industryMap, codeMap) => rows.map(item => {
  const EAN = cleanEan(item.EAN)
  if (EAN && industryMap[EAN]) return { ...item, EAN, Industria: industryMap[EAN] }
  if (!EAN && item.CodigoInterno) {
    const match = codeMap[normalizeCatalogCode(item.CodigoInterno)]
    if (match) return { ...item, EAN: match.EAN, Industria: match.Industria }
  }
  return { ...item, EAN, Industria: resolveTrustedIndustry(item, industryMap, codeMap) }
})

const findSuggestedColumn = (headers, terms) => {
  const normalizedTerms = terms.map(normalize)
  return headers.find(header => normalizedTerms.some(term => normalize(header).includes(term))) ?? ''
}

const cleanEan = value => {
  if (value === null || value === undefined || value === '') return ''
  if (typeof value === 'number') return Number.isInteger(value) ? String(value) : String(Math.trunc(value))
  const text = String(value).trim()
  const exponential = text.match(/^([\d.,]+)e\+(\d+)$/i)
  if (exponential) return Number(text.replace(',', '.')).toLocaleString('fullwide', { useGrouping: false, maximumFractionDigits: 0 })
  return text.replace(/\.0+$/, '').replace(/[^0-9]/g, '')
}

const cleanQuantity = value => {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'number') return Math.trunc(value)
  const text = String(value).trim().replace(/\s/g, '')
  const normalized = text.includes(',')
    ? text.replace(/\./g, '').replace(',', '.')
    : text
  const number = Number(normalized)
  return Number.isFinite(number) ? Math.trunc(number) : null
}

const parsePdfNumber = value => {
  const normalized = String(value).trim().replace(/\./g, '').replace(',', '.')
  const number = Number(normalized)
  return Number.isFinite(number) ? Math.trunc(number) : null
}

const parsePdfMoney = value => {
  const text = String(value || '').trim().replace(/R\$\s*/gi, '')
  if (!text) return 0
  const comma = text.lastIndexOf(',')
  const dot = text.lastIndexOf('.')
  let normalized = text
  if (comma > dot) normalized = text.replace(/\./g, '').replace(',', '.')
  else if (dot > comma && comma >= 0) normalized = text.replace(/,/g, '')
  else if (comma >= 0) normalized = text.replace(',', '.')
  const number = Number(normalized)
  return Number.isFinite(number) ? number : 0
}

const extractPrintedDocumentTotal = text => {
  const values = []
  for (const rawLine of String(text || '').split(/\r?\n/)) {
    const line = normalize(rawLine)
    if (!line.includes('total')) continue
    if (!/(liquid|pedido|valor|vl\s*total)/.test(line)) continue
    for (const match of rawLine.matchAll(/\d{1,3}(?:[.\s]\d{3})*,\d{2}|\d+[.,]\d{2}/g)) {
      const value = parsePdfMoney(match[0].replace(/\s/g, '.'))
      if (value > 0) values.push(value)
    }
  }
  return values.length ? Math.max(...values) : 0
}

const extractPiraquaraStoreInfo = text => {
  const destinationMatch = String(text || '').match(/\bDestino:\s*(\d{1,3})\b/i)
  const cnpjMatch = String(text || '').match(/CNPJ\/CPF:\s*([\d./-]+)/i)
  const cnpjDigits = (cnpjMatch?.[1] || '').replace(/\D/g, '')
  const destinationNumber = Number(destinationMatch?.[1])
  if (cnpjDigits.length !== 14 || !Number.isFinite(destinationNumber) || destinationNumber < 1) return null
  return {
    cnpjFinal: cnpjDigits.slice(-4),
    loja: String(destinationNumber).padStart(2, '0'),
    destinoOriginal: destinationMatch[1].padStart(3, '0'),
  }
}

const groupPdfLines = items => {
  const lines = []
  for (const item of items.filter(entry => entry.str?.trim())) {
    const x = item.transform[4]
    const y = item.transform[5]
    let line = lines.find(entry => Math.abs(entry.y - y) <= 2.5)
    if (!line) {
      line = { y, items: [] }
      lines.push(line)
    }
    line.items.push({ text: item.str.trim(), x })
  }
  return lines
    .sort((a, b) => b.y - a.y)
    .map(line => ({ ...line, items: line.items.sort((a, b) => a.x - b.x) }))
}

const parseJhlPurchaseOrder = pages => {
  const rows = []

  for (const lines of pages) {
    const allItems = lines.flatMap(line => line.items)
    const quantityHeader = allItems.find(item => normalize(item.text) === 'qtde')
    const barcodeHeader = allItems.find(item => /\bcod\.?\s+de\s+barras\b/.test(normalize(item.text)))
      || allItems.find(item => normalize(item.text) === 'cod.')
    const productHeader = allItems.find(item => /^nome(?:\s+do\s+produto)?$/.test(normalize(item.text)))
    if (!quantityHeader || !barcodeHeader) continue

    for (const line of lines) {
      const referenceItem = line.items.find(item => item.x < (productHeader?.x ?? barcodeHeader.x) && /\b(?:RB|LO|SB)[A-Z0-9-]+\b/i.test(item.text))
      const reference = referenceItem?.text.match(/\b(?:RB|LO|SB)[A-Z0-9-]+\b/i)?.[0]
      if (!reference) continue

      const eanItem = line.items.find(item =>
        /\b\d{12,14}\b/.test(item.text) &&
        item.x >= barcodeHeader.x - 10 &&
        item.x < quantityHeader.x - 25
      )
      const numericItems = line.items.filter(item => /^\d[\d.,]*$/.test(item.text))
      const quantityItem = numericItems.length
        ? numericItems.reduce((nearest, item) =>
          Math.abs(item.x - quantityHeader.x) < Math.abs(nearest.x - quantityHeader.x) ? item : nearest
        )
        : null
      const Quantidade = parsePdfNumber(quantityItem?.text)
      if (!Quantidade) continue

      const priceItems = numericItems
        .filter(item => item.x > quantityHeader.x + 15)
        .sort((a, b) => a.x - b.x)
      const ValorUnitario = parsePdfMoney(priceItems[0]?.text)
      const ValorTotal = parsePdfMoney(priceItems.at(-1)?.text)
      const itemStart = productHeader ? productHeader.x - 3 : referenceItem.x + 35
      const Item = line.items
        .filter(item => item.x >= itemStart && item.x < barcodeHeader.x - 3 && normalize(item.text) !== 'un')
        .map(item => item.text)
        .join(' ')
        .trim() || 'Item para revisar'

      const CodigoInterno = reference.toUpperCase()
      // RB não confirma mais Reckitt: a separação final vem da planilha-mãe,
      // priorizando EAN exato e usando o código CBN somente quando o PDF não traz EAN.
      const Industria = CodigoInterno.startsWith('LO')
        ? 'LOREAL'
        : CodigoInterno.startsWith('SB')
          ? '3M'
          : 'NAO_IDENTIFICADA'

      rows.push({
        EAN: eanItem?.text.match(/\b\d{12,14}\b/)?.[0] || '',
        CodigoInterno,
        Item,
        Quantidade,
        ValorUnitario,
        ValorTotal,
        Industria,
      })
    }
  }

  return rows
}

const readPdfOrder = async file => {
  const data = new Uint8Array(await file.arrayBuffer())
  const document = await pdfjsLib.getDocument({ data }).promise
  const pages = []
  let fullText = ''

  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber)
    const content = await page.getTextContent()
    const lines = groupPdfLines(content.items)
    pages.push(lines)
    fullText += ' ' + lines.map(line => line.items.map(item => item.text).join(' ')).join(' ')
  }

  const normalizedText = normalize(fullText)
  let modelId = ''
  if (normalizedText.includes('hiper erp') || normalizedText.includes('supermercado piraquara')) modelId = 'piraquara'
  else if (normalizedText.includes('hermes') || normalizedText.includes('sugestao de pedido')) modelId = 'ouro-branco'
  else if (normalizedText.includes('cbn distribuidora') && normalizedText.includes('orcamento')) modelId = 'flex-cbn'
  else if (normalizedText.includes('jhl produtos de higiene') && normalizedText.includes('pedido de compra')) modelId = 'personalizado'

  if (!modelId) {
    throw new Error(
      normalizedText.trim()
        ? 'Este modelo de PDF ainda não foi reconhecido automaticamente.'
        : 'Este PDF é uma imagem escaneada e precisa do módulo de OCR, que será a próxima etapa.'
    )
  }

  if (modelId === 'personalizado' && normalizedText.includes('jhl produtos de higiene')) {
    const rows = parseJhlPurchaseOrder(pages)
    if (!rows.length) throw new Error('Não encontrei itens válidos nesta Ordem de Compra.')
    const documentTotal = Math.round(rows.reduce((sum, item) => sum + (item.ValorTotal || 0), 0) * 100) / 100
    return { modelId, rows, documentTotal, storeInfo: null }
  }

  const results = []
  const seen = new Map()
  let stopped = false

  for (const lines of pages) {
    const allItems = lines.flatMap(line => line.items)
    const brandHeader = allItems.find(item => normalize(item.text) === 'marca')
    const packageHeader = allItems.find(item => ['emb.', 'emb'].includes(normalize(item.text)))
    const quantityHeader = allItems.find(item => /^(qtde|qtd|quantidade|quant)$/.test(normalize(item.text)))

    for (const line of lines) {
      const joined = normalize(line.items.map(item => item.text).join(' '))
      if (joined.includes('trocas pendentes')) {
        stopped = true
        break
      }

      const eanIndex = line.items.findIndex(item => /^\d{8,14}$/.test(item.text))
      if (eanIndex < 0) continue
      const eanItem = line.items[eanIndex]

      const numericAfterItems = line.items
        .filter(item => item.x > eanItem.x + 8 && /^\d[\d.,]*$/.test(item.text))
      const numericAfter = numericAfterItems
        .map(item => parsePdfNumber(item.text))
        .filter(value => value !== null)

      let quantity = null
      let quantityItem = null
      if (modelId === 'piraquara' && quantityHeader && numericAfterItems.length) {
        quantityItem = numericAfterItems.reduce((nearest, item) =>
          Math.abs(item.x - quantityHeader.x) < Math.abs(nearest.x - quantityHeader.x) ? item : nearest
        )
        quantity = parsePdfNumber(quantityItem.text)
      } else if (modelId === 'ouro-branco' && numericAfter.length >= 2) {
        quantity = numericAfter[0] * numericAfter[1]
      } else if (numericAfter.length) {
        quantity = numericAfter[0]
      }
      if (!quantity) continue

      let itemName = ''
      if (modelId === 'flex-cbn') {
        const beforeEan = line.items.slice(0, eanIndex).map(item => item.text).join(' ')
        itemName = beforeEan.replace(/^\s*\d+\s+(?:RB|LO|SB)\w+\s+/i, '').trim()
      } else {
        const boundary = modelId === 'piraquara'
          ? (brandHeader?.x || numericAfterItems[0]?.x)
          : (packageHeader?.x || numericAfterItems[0]?.x)
        itemName = line.items
          .filter(item => item.x > eanItem.x + 4 && (!boundary || item.x < boundary - 2))
          .map(item => item.text)
          .join(' ')
          .trim()
      }

      const ean = cleanEan(eanItem.text)
      const valueItems = modelId === 'piraquara' && quantityHeader
        ? numericAfterItems.filter(item => item.x > quantityHeader.x + 20)
        : numericAfterItems
      const lineTotal = parsePdfMoney(valueItems.at(-1)?.text)
      const unitCandidates = valueItems.slice(0, -1)
        .filter(item => item !== quantityItem)
        .map(item => parsePdfMoney(item.text))
        .filter(value => value > 0)
      const ValorUnitario = unitCandidates.length && lineTotal
        ? unitCandidates.reduce((best, value) =>
          Math.abs(value * quantity - lineTotal) < Math.abs(best * quantity - lineTotal) ? value : best
        )
        : 0
      const beforeEan = line.items.slice(0, eanIndex).map(item => item.text).join(' ')
      const internalMatch = beforeEan.match(/\b(?:RB|LO|SB)[A-Z0-9-]+\b/i)
      const CodigoInterno = internalMatch?.[0] || ''
      const existing = seen.get(ean)
      if (existing) {
        existing.Quantidade += quantity
        existing.ValorTotal += lineTotal
        if (!existing.Item && itemName) existing.Item = itemName
        if (!existing.ValorUnitario && ValorUnitario) existing.ValorUnitario = ValorUnitario
        if (!existing.CodigoInterno && CodigoInterno) existing.CodigoInterno = CodigoInterno
      } else {
        seen.set(ean, { EAN: ean, CodigoInterno, Item: itemName || 'Item sem descrição', Quantidade: quantity, ValorUnitario, ValorTotal: lineTotal })
      }
    }
    if (stopped) break
  }

  for (const item of seen.values()) results.push(item)
  if (!results.length) throw new Error('Não encontrei itens válidos neste PDF. O arquivo precisa de revisão.')
  const printedTotalMatch = fullText.match(/Vl\s*Total\s*\$?\s*([\d.,]+)/i)
  const mainItemsTotal = Math.round(
    results.reduce((sum, item) => sum + (item.ValorTotal || 0), 0) * 100
  ) / 100
  const documentTotal = modelId === 'piraquara'
    ? mainItemsTotal
    : extractPrintedDocumentTotal(fullText) ||
      (printedTotalMatch ? parsePdfMoney(printedTotalMatch[1]) : 0)
  const storeInfo = modelId === 'piraquara' ? extractPiraquaraStoreInfo(fullText) : null
  return { modelId, rows: results, documentTotal, storeInfo }
}

const isValidEan13 = ean => {
  if (!/^\d{13}$/.test(ean)) return false
  const sum = ean.slice(0, 12).split('').reduce((total, digit, index) =>
    total + Number(digit) * (index % 2 === 0 ? 1 : 3), 0)
  return (10 - (sum % 10)) % 10 === Number(ean[12])
}

const normalizeOcrDigitToken = value => String(value || '')
  .toUpperCase()
  .replace(/[OQD]/g, '0')
  .replace(/[IL|]/g, '1')
  .replace(/Z/g, '2')
  .replace(/S/g, '5')
  .replace(/G/g, '6')
  .replace(/B/g, '8')
  .replace(/[^0-9]/g, '')

const editDistance = (left, right) => {
  const previous = Array.from({ length: right.length + 1 }, (_, index) => index)
  for (let row = 1; row <= left.length; row += 1) {
    const current = [row]
    for (let column = 1; column <= right.length; column += 1) {
      current[column] = Math.min(
        current[column - 1] + 1,
        previous[column] + 1,
        previous[column - 1] + (left[row - 1] === right[column - 1] ? 0 : 1),
      )
    }
    previous.splice(0, previous.length, ...current)
  }
  return previous[right.length]
}

const resolveOcrEan = (value, catalog = {}) => {
  const digits = normalizeOcrDigitToken(value)
  if (catalog[digits] || isValidEan13(digits)) return digits

  const candidates = [...new Set(Object.keys(catalog))]
    .filter(ean => Math.abs(ean.length - digits.length) <= 1 && /^\d{8,14}$/.test(ean))
  let best = null
  let bestDistance = Infinity
  let tied = false
  for (const candidate of candidates) {
    const distance = editDistance(digits, candidate)
    if (distance < bestDistance) {
      best = candidate
      bestDistance = distance
      tied = false
    } else if (distance === bestDistance) {
      tied = true
    }
  }
  const allowedDistance = digits.length >= 12 ? 3 : 2
  if (best && !tied && bestDistance <= allowedDistance) return best
  return digits
}

const cleanOcrEan = (value, catalog = {}) => resolveOcrEan(value, catalog)

const mergeAdegaOcrRows = rows => {
  const merged = new Map()
  for (const item of rows) {
    if (!item?.EAN) continue
    const current = merged.get(item.EAN)
    if (!current) {
      merged.set(item.EAN, { ...item })
      continue
    }
    const currentScore = (current.CodigoInterno ? 2 : 0) + (current.Item?.length || 0) + (current.ValorTotal > 0 ? 4 : 0)
    const itemScore = (item.CodigoInterno ? 2 : 0) + (item.Item?.length || 0) + (item.ValorTotal > 0 ? 4 : 0)
    const preferred = itemScore > currentScore ? item : current
    merged.set(item.EAN, {
      ...current,
      ...preferred,
      CodigoInterno: preferred.CodigoInterno || current.CodigoInterno || '',
      Item: preferred.Item || current.Item,
      Quantidade: Math.max(current.Quantidade || 0, item.Quantidade || 0),
      ValorUnitario: preferred.ValorUnitario || current.ValorUnitario || 0,
      ValorTotal: Math.max(current.ValorTotal || 0, item.ValorTotal || 0),
    })
  }
  return [...merged.values()]
}

const parseAdegaText = (text, catalog = {}) => {
  const rows = []
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.replace(/\s+/g, ' ').trim()
    if (!line) continue
    const tokens = [...line.matchAll(/[A-Za-z0-9]{11,16}/g)]
    const eanToken = tokens.find(match => (match[0].match(/\d/g) || []).length >= 8)
    if (!eanToken) continue
    const EAN = resolveOcrEan(eanToken[0], catalog)
    if (!/^\d{12,14}$/.test(EAN)) continue

    const tail = line.slice((eanToken.index || 0) + eanToken[0].length)
    const codePrefix = tail.match(/^\s*((?:[A-Z]{0,3}\d[A-Z0-9.,/-]*\s+){1,5})/i)?.[1] || ''
    const internalCandidates = [...codePrefix.matchAll(/\b\d{3,7}\b/g)].map(match => match[0])
    const CodigoInterno = internalCandidates.at(-1) || ''
    const dataTail = codePrefix ? tail.slice(tail.indexOf(codePrefix) + codePrefix.length) : tail
    const packageMatch = dataTail.match(/[_-]?0?\d{1,2}X\d+[A-Z0-9]*/i)
    const numericTail = packageMatch ? dataTail.slice((packageMatch.index || 0) + packageMatch[0].length) : dataTail
    const numericTokens = [...numericTail.matchAll(/\d[\d.,]*/g)].map(match => parsePdfMoney(match[0])).filter(Number.isFinite)
    if (numericTokens.length < 2) continue

    const ValorTotal = numericTokens.at(-1)
    let ValorUnitario = 0
    for (let index = numericTokens.length - 2; index >= 0; index -= 1) {
      const candidate = numericTokens[index]
      if (candidate > 0 && candidate < ValorTotal) {
        ValorUnitario = candidate
        break
      }
    }

    let Quantidade = ValorUnitario ? Math.round(ValorTotal / ValorUnitario) : 0
    if ((!Quantidade || Quantidade < 2) && packageMatch) {
      const afterPackage = numericTail.match(/\d{1,4}/)
      Quantidade = afterPackage ? Number(afterPackage[0]) : 0
    }
    if (!Quantidade || Quantidade < 1 || Quantidade > 5000) continue

    let Item = dataTail.slice(0, packageMatch?.index ?? Math.max(0, dataTail.search(/\s0[.,]00/)))
      .replace(/^\s*[A-Z0-9-]{3,12}\s+(?:\d{3,8}\s+)?/i, '')
      .replace(/\b(?:UN|PC|PK|CX|LT)\b\s*$/i, '')
      .replace(/[|_]+/g, ' ')
      .trim()
    if (!Item) Item = 'Item Adega Brasil para revisar'

    rows.push({ EAN, CodigoInterno, Item, Quantidade, ValorUnitario, ValorTotal })
  }
  return mergeAdegaOcrRows(rows)
}

const parseVariableText = (text, catalog = {}) => {
  const normalizedDocument = normalize(text)
  const hasAdegaLayout = normalizedDocument.includes('adega brasil') ||
    normalizedDocument.includes('adeg a brasil') ||
    normalizedDocument.includes('pedido emitido na unid') ||
    /[_-]?0?\d{1,2}X\d+[A-Z0-9]*/i.test(text)
  if (hasAdegaLayout) {
    const adegaRows = parseAdegaText(text, catalog)
    if (adegaRows.length) return adegaRows
  }

  const multiplyPackage = normalizedDocument.includes('sugestao de pedido') ||
    (normalizedDocument.includes('emb') && normalizedDocument.includes('qtde'))
  const results = new Map()

  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.replace(/\s+/g, ' ').trim()
    if (!line || normalize(line).includes('trocas pendentes')) break

    const cleanedLine = line.replace(/[|\[\]_]/g, ' ')
    const candidates = [...cleanedLine.matchAll(/\d{8,14}/g)]
    const eanMatch = candidates.find(match => match[0].length >= 12) || candidates[0]
    if (!eanMatch) continue
    const EAN = cleanOcrEan(eanMatch[0], catalog)
    const tail = cleanedLine.slice((eanMatch.index || 0) + eanMatch[0].length)
    const numbers = [...tail.matchAll(/\d+(?:[.,]\d+)?/g)]
    const required = multiplyPackage ? 4 : 3
    if (numbers.length < required) continue

    const quantityIndex = numbers.length - required
    const ordered = parsePdfNumber(numbers[quantityIndex][0])
    const pack = multiplyPackage ? parsePdfNumber(numbers[quantityIndex + 1][0]) : 1
    const Quantidade = ordered && pack ? ordered * pack : null
    if (!Quantidade) continue

    const Item = tail.slice(0, numbers[quantityIndex].index).trim().replace(/^[-:;\s]+/, '') || 'Item para revisar'
    const ValorTotal = parsePdfMoney(numbers.at(-1)[0])
    const existing = results.get(EAN)
    if (existing) {
      existing.Quantidade += Quantidade
      existing.ValorTotal += ValorTotal
    } else {
      results.set(EAN, { EAN, Item, Quantidade, ValorTotal })
    }
  }

  const compactText = text.replace(/\s+/g, ' ')
  const rowPattern = /(\d{12,14})[\s|\[\]_]+(.{3,220}?)[\s|\[\]_]+(\d{1,3}[.,]\d{3})[\s|\[\]_]+(\d{1,4}[.,]\d{3,4})[\s|\[\]_]+(\d{1,3}(?:\.\d{3})*[.,]\d{2})/g
  for (const match of compactText.matchAll(rowPattern)) {
    const EAN = cleanOcrEan(match[1], catalog)
    const Quantidade = parsePdfNumber(match[3])
    if (!EAN || !Quantidade) continue
    results.set(EAN, {
      EAN,
      Item: match[2].replace(/[|\[\]_]+/g, ' ').replace(/\s+/g, ' ').trim(),
      Quantidade,
      ValorTotal: parsePdfMoney(match[5]),
    })
  }

  return [...results.values()]
}

const prepareImageForOcr = async (source, rotation = 0) => {
  if (!(source instanceof Blob) || !window.createImageBitmap) return source
  try {
    const bitmap = await window.createImageBitmap(source)
    const sideways = Math.abs(rotation) % 180 === 90
    const rotatedWidth = sideways ? bitmap.height : bitmap.width
    const rotatedHeight = sideways ? bitmap.width : bitmap.height
    const scale = Math.max(1.8, Math.min(2.5, 3600 / rotatedWidth))
    const canvas = window.document.createElement('canvas')
    canvas.width = Math.round(rotatedWidth * scale)
    canvas.height = Math.round(rotatedHeight * scale)
    const context = canvas.getContext('2d', { willReadFrequently: true })
    context.fillStyle = '#ffffff'
    context.fillRect(0, 0, canvas.width, canvas.height)
    context.filter = 'grayscale(1) contrast(1.45)'
    context.save()
    context.translate(canvas.width / 2, canvas.height / 2)
    context.rotate(rotation * Math.PI / 180)
    context.drawImage(bitmap, -bitmap.width * scale / 2, -bitmap.height * scale / 2, bitmap.width * scale, bitmap.height * scale)
    context.restore()
    bitmap.close()

    const image = context.getImageData(0, 0, canvas.width, canvas.height)
    const pixels = image.data
    const rowDark = new Uint32Array(canvas.height)
    const columnDark = new Uint32Array(canvas.width)

    for (let y = 0; y < canvas.height; y += 1) {
      for (let x = 0; x < canvas.width; x += 1) {
        const index = (y * canvas.width + x) * 4
        if (pixels[index] < 60) {
          rowDark[y] += 1
          columnDark[x] += 1
        }
      }
    }

    for (let y = 0; y < canvas.height; y += 1) {
      if (rowDark[y] / canvas.width <= 0.65) continue
      for (let offset = -2; offset <= 2; offset += 1) {
        const target = y + offset
        if (target < 0 || target >= canvas.height) continue
        for (let x = 0; x < canvas.width; x += 1) {
          const index = (target * canvas.width + x) * 4
          pixels[index] = pixels[index + 1] = pixels[index + 2] = 255
        }
      }
    }
    for (let x = 0; x < canvas.width; x += 1) {
      if (columnDark[x] / canvas.height <= 0.60) continue
      for (let offset = -2; offset <= 2; offset += 1) {
        const target = x + offset
        if (target < 0 || target >= canvas.width) continue
        for (let y = 0; y < canvas.height; y += 1) {
          const index = (y * canvas.width + target) * 4
          pixels[index] = pixels[index + 1] = pixels[index + 2] = 255
        }
      }
    }
    context.putImageData(image, 0, 0)
    return canvas
  } catch {
    return source
  }
}

const scoreOcrText = text => {
  const digitCodes = (text.match(/\b\d{12,14}\b/g) || []).length
  const tableWords = (normalize(text).match(/adega brasil|cod.?barras|quantidade|descricao|pedido/g) || []).length
  return digitCodes * 5 + tableWords
}

const recognizeImage = async source => {
  const uploadedImage = source instanceof Blob
  const worker = await createWorker('eng')
  try {
    const rotations = uploadedImage ? [270, 90, 0] : [0]
    let bestText = ''
    let bestScore = -1
    let bestSource = source

    await worker.setParameters({
      tessedit_pageseg_mode: uploadedImage ? PSM.SINGLE_COLUMN : PSM.AUTO,
      preserve_interword_spaces: '1',
    })

    for (const rotation of rotations) {
      const preparedSource = await prepareImageForOcr(source, rotation)
      const { data } = await worker.recognize(preparedSource)
      const candidateText = data.text || ''
      const candidateScore = scoreOcrText(candidateText)
      if (candidateScore > bestScore) {
        bestText = candidateText
        bestScore = candidateScore
        bestSource = preparedSource
      }
    }

    if (uploadedImage) {
      await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_BLOCK, preserve_interword_spaces: '1' })
      const { data } = await worker.recognize(bestSource)
      if (data.text) bestText += '\n' + data.text
    }
    return bestText
  } finally {
    await worker.terminate()
  }
}

const recognizeScannedPdf = async file => {
  const data = new Uint8Array(await file.arrayBuffer())
  const pdfDocument = await pdfjsLib.getDocument({ data }).promise
  let text = ''

  for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
    const page = await pdfDocument.getPage(pageNumber)
    const viewport = page.getViewport({ scale: 2 })
    const canvas = window.document.createElement('canvas')
    canvas.width = viewport.width
    canvas.height = viewport.height
    const context = canvas.getContext('2d')
    await page.render({ canvasContext: context, viewport }).promise
    text += '\n' + await recognizeImage(canvas)
  }
  return text
}

const mergeItemsByEan = rows => {
  const merged = new Map()
  for (const [index, item] of rows.entries()) {
    if (!item) continue
    const key = item.EAN
      ? `EAN:${item.EAN}`
      : item.CodigoInterno
        ? `COD:${item.CodigoInterno}`
        : `ROW:${index}`
    const current = merged.get(key)
    if (current) {
      current.Quantidade += item.Quantidade || 0
      current.ValorTotal += item.ValorTotal || 0
      if ((!current.Item || current.Item === 'Item sem descrição' || current.Item === 'Item para revisar') && item.Item) current.Item = item.Item
      if (!current.ValorUnitario && item.ValorUnitario) current.ValorUnitario = item.ValorUnitario
      if (!current.CodigoInterno && item.CodigoInterno) current.CodigoInterno = item.CodigoInterno
      if (!current.Industria && item.Industria) current.Industria = item.Industria
    } else {
      merged.set(key, { ...item })
    }
  }
  return [...merged.values()]
}

const parseSpreadsheetOrder = async file => {
  const data = await file.arrayBuffer()
  const workbook = XLSX.read(data, { type: 'array', cellDates: true })
  const firstSheet = workbook.SheetNames[0]
  const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1, defval: '', raw: false })

  let bestHeader = -1
  let bestScore = -1
  for (let rowIndex = 0; rowIndex < Math.min(rawRows.length, 30); rowIndex += 1) {
    const headers = rawRows[rowIndex].map(normalize)
    let score = 0
    if (headers.some(value => /ean|cod.*barra|cód.*barra|gtin/.test(value))) score += 5
    if (headers.some(value => /quant|qtd|qtde|pedido.*und/.test(value))) score += 4
    if (headers.some(value => /descr|item|produto/.test(value))) score += 2
    if (headers.some(value => /total|valor pedido|vl pedido/.test(value))) score += 2
    if (score > bestScore) {
      bestScore = score
      bestHeader = rowIndex
    }
  }

  if (bestHeader < 0 || bestScore < 7) throw new Error('Não encontrei as colunas de EAN e quantidade neste arquivo.')
  const headers = rawRows[bestHeader].map(value => String(value || '').trim())
  const eanIndex = headers.findIndex(value => /ean|cod.*barra|cód.*barra|gtin/.test(normalize(value)))
  const quantityIndex = headers.findIndex(value => /quant|qtd|qtde|pedido.*und/.test(normalize(value)))
  const itemIndex = headers.findIndex(value => /descr|item|produto/.test(normalize(value)))
  const packageIndex = headers.findIndex(value => /^emb|embalagem/.test(normalize(value)))
  const totalIndex = headers.findIndex(value => /(^|\s)(vl\.?\s*)?total|valor pedido/.test(normalize(value)))
  const unitIndex = headers.findIndex(value => /unit|vlr\.?$|valor$|preco|preço|^cbn$/.test(normalize(value)))
  const internalCodeIndex = headers.findIndex((value, index) => {
    const header = normalize(value).replace(/\./g, '').trim()
    return index !== eanIndex && /^(cod|codigo|codigo interno|cod interno|codigo produto|cod produto)$/.test(header)
  })
  if (eanIndex < 0 || quantityIndex < 0) throw new Error('Não encontrei as colunas de EAN e quantidade neste arquivo.')
  const shouldMultiply = packageIndex >= 0 && quantityIndex >= 0 &&
    headers.some(value => /sugestao|sugestão/.test(normalize(value))) ||
    (packageIndex >= 0 && normalize(headers[quantityIndex]).includes('qtde'))

  const rows = []
  for (const row of rawRows.slice(bestHeader + 1)) {
    const EAN = cleanEan(row[eanIndex])
    if (!/^\d{8,14}$/.test(EAN)) continue
    const ordered = cleanQuantity(row[quantityIndex])
    const pack = shouldMultiply ? cleanQuantity(row[packageIndex]) : 1
    const Quantidade = ordered !== null && pack !== null ? ordered * pack : null
    if (!Quantidade) continue
    const ValorUnitario = unitIndex >= 0 ? parsePdfMoney(row[unitIndex]) : 0
    const ValorTotal = totalIndex >= 0 ? parsePdfMoney(row[totalIndex]) : ValorUnitario * Quantidade
    rows.push({
      EAN,
      CodigoInterno: internalCodeIndex >= 0 ? String(row[internalCodeIndex] || '').trim() : '',
      Item: itemIndex >= 0 ? String(row[itemIndex] || '').trim() || 'Item sem descrição' : 'Item sem descrição',
      Quantidade,
      ValorUnitario,
      ValorTotal,
    })
  }

  if (!rows.length) throw new Error('O arquivo foi aberto, mas nenhum item válido foi encontrado.')
  return rows
}

function LoginScreen() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [sent, setSent] = useState(false)
  const [error, setError] = useState('')
  const [sending, setSending] = useState(false)

  const signInWithPassword = async event => {
    event.preventDefault()
    setSending(true)
    setError('')
    const normalizedEmail = email.trim().toLowerCase()
    const { error: authError } = await supabase.auth.signInWithPassword({ email: normalizedEmail, password })
    setSending(false)
    if (authError) {
      const invalid = /invalid login credentials|email not confirmed/i.test(authError.message || '')
      setError(invalid ? 'E-mail ou senha incorretos.' : 'Não foi possível entrar agora. Tente novamente.')
    }
  }

  const requestRecovery = async () => {
    const normalizedEmail = email.trim().toLowerCase()
    if (!normalizedEmail) {
      setError('Informe seu e-mail para receber o link de recuperação.')
      return
    }
    setSending(true)
    setError('')
    const { error: authError } = await supabase.auth.resetPasswordForEmail(normalizedEmail, {
      redirectTo: 'https://felipefraxino.github.io/Automa-o-Pedidos/',
    })
    setSending(false)
    if (authError) {
      const rateLimited = authError.status === 429 || /rate.?limit|too many requests|email rate limit/i.test(authError.message || '')
      setError(
        rateLimited
          ? 'Foram solicitados vários links em pouco tempo. Aguarde alguns minutos e tente novamente apenas uma vez.'
          : 'Não foi possível enviar o link de recuperação agora.'
      )
      return
    }
    setSent(true)
  }

  return (
    <main className="login-page">
      <section className="login-card">
        <div className="login-brand"><FileSpreadsheet size={28} /></div>
        <p>GESTÃO DE PEDIDOS - CBN DISTRIBUIDORA</p>
        <h1>Acesse o conversor</h1>
        {!sent ? (
          <>
            <span>Entre com seu e-mail e sua senha.</span>
            <form onSubmit={signInWithPassword}>
              <label>E-mail
                <div className="email-field"><Mail size={18} /><input type="email" autoComplete="email" required value={email} onChange={event => setEmail(event.target.value)} placeholder="seuemail@exemplo.com" /></div>
              </label>
              <label>Senha
                <div className="email-field"><KeyRound size={18} /><input type="password" autoComplete="current-password" required value={password} onChange={event => setPassword(event.target.value)} placeholder="Sua senha" /></div>
              </label>
              {error && <div className="login-error">{error}</div>}
              <button disabled={sending}>{sending ? 'Entrando…' : 'Entrar'}</button>
              <button className="login-link-button" type="button" disabled={sending} onClick={requestRecovery}>Esqueci minha senha</button>
            </form>
          </>
        ) : (
          <div className="login-sent">
            <CheckCircle2 size={32} />
            <strong>Confira seu e-mail</strong>
            <span>Enviamos um link seguro para você definir uma nova senha em {email.trim().toLowerCase()}.</span>
            <button onClick={() => setSent(false)}>Voltar para entrar com senha</button>
          </div>
        )}
      </section>
    </main>
  )
}

function PasswordRecoveryScreen({ onComplete }) {
  const [password, setPassword] = useState('')
  const [confirmation, setConfirmation] = useState('')
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')

  const updatePassword = async event => {
    event.preventDefault()
    setError('')
    if (password.length < 8) {
      setError('A nova senha precisa ter pelo menos 8 caracteres.')
      return
    }
    if (password !== confirmation) {
      setError('As duas senhas não são iguais.')
      return
    }
    setSaving(true)
    const { error: authError } = await supabase.auth.updateUser({ password })
    setSaving(false)
    if (authError) {
      setError(/same_password/i.test(authError.code || authError.message || '')
        ? 'Escolha uma senha diferente da senha anterior.'
        : 'Não foi possível salvar a nova senha. Solicite outro link e tente novamente.')
      return
    }
    window.history.replaceState({}, document.title, window.location.pathname)
    onComplete()
  }

  return (
    <main className="login-page">
      <section className="login-card">
        <div className="login-brand"><KeyRound size={28} /></div>
        <p>GESTÃO DE PEDIDOS - CBN DISTRIBUIDORA</p>
        <h1>Defina sua nova senha</h1>
        <span>Crie uma senha segura. Ela não será exibida nem enviada por e-mail.</span>
        <form onSubmit={updatePassword}>
          <label>Nova senha
            <div className="email-field"><KeyRound size={18} /><input type="password" autoComplete="new-password" minLength="8" required value={password} onChange={event => setPassword(event.target.value)} placeholder="Pelo menos 8 caracteres" /></div>
          </label>
          <label>Confirme a nova senha
            <div className="email-field"><KeyRound size={18} /><input type="password" autoComplete="new-password" minLength="8" required value={confirmation} onChange={event => setConfirmation(event.target.value)} placeholder="Digite novamente" /></div>
          </label>
          {error && <div className="login-error">{error}</div>}
          <button disabled={saving}>{saving ? 'Salvando…' : 'Salvar nova senha e entrar'}</button>
        </form>
      </section>
    </main>
  )
}

function AuthenticatedApp() {
  const [session, setSession] = useState(undefined)
  const [recoveringPassword, setRecoveringPassword] = useState(() =>
    `${window.location.search}${window.location.hash}`.includes('type=recovery')
  )

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session))
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event, nextSession) => {
      setSession(nextSession)
      if (event === 'PASSWORD_RECOVERY') setRecoveringPassword(true)
    })
    return () => subscription.unsubscribe()
  }, [])

  if (session === undefined) return <div className="auth-loading">Carregando…</div>
  if (!session) return <LoginScreen />
  if (recoveringPassword) return <PasswordRecoveryScreen onComplete={() => setRecoveringPassword(false)} />
  return <App session={session} />
}

function App({ session }) {
  const [models, setModels] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('cbn-models') || '[]')
      const map = new Map([...BASE_MODELS, ...saved].map(model => [model.id, model]))
      return [...map.values()]
    } catch {
      return BASE_MODELS
    }
  })
  const [selectedId, setSelectedId] = useState('piraquara')
  const [fileName, setFileName] = useState('')
  const [workbookRows, setWorkbookRows] = useState([])
  const [orderDetails, setOrderDetails] = useState([])
  const [sourceOrderTotal, setSourceOrderTotal] = useState(0)
  const [piraquaraStoreInfo, setPiraquaraStoreInfo] = useState(null)
  const [sheetName, setSheetName] = useState('')
  const [headerRow, setHeaderRow] = useState(1)
  const [eanColumn, setEanColumn] = useState('')
  const [quantityColumn, setQuantityColumn] = useState('')
  const [packageColumn, setPackageColumn] = useState('')
  const [quantityMode, setQuantityMode] = useState('direct')
  const [outputIndustry, setOutputIndustry] = useState('RECKITT')
  const [catalogIndustries, setCatalogIndustries] = useState({})
  const [catalogByCode, setCatalogByCode] = useState({})
  const [message, setMessage] = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const [pendingFiles, setPendingFiles] = useState([])
  const [processingFiles, setProcessingFiles] = useState(false)
  const fileRef = useRef(null)
  const [showPasswordSetup, setShowPasswordSetup] = useState(false)
  const [newPassword, setNewPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [passwordMessage, setPasswordMessage] = useState(null)
  const [savingPassword, setSavingPassword] = useState(false)
  const [budgetFileName, setBudgetFileName] = useState('')
  const [orderFileName, setOrderFileName] = useState('')
  const [budgetRows, setBudgetRows] = useState([])
  const [comparisonOrderRows, setComparisonOrderRows] = useState([])
  const [comparisonLoading, setComparisonLoading] = useState('')
  const budgetFileRef = useRef(null)
  const orderFileRef = useRef(null)
  const [piraquaraObservation, setPiraquaraObservation] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('cbn-piraquara-execution') || 'null')
      return saved?.notes || ''
    } catch {
      return ''
    }
  })
  const [piraquaraSubmitting, setPiraquaraSubmitting] = useState(false)
  const [piraquaraRequest, setPiraquaraRequest] = useState(null)
  const [piraquaraStatusError, setPiraquaraStatusError] = useState('')
  const [piraquaraAlertsEnabled, setPiraquaraAlertsEnabled] = useState(false)
  const audioContextRef = useRef(null)
  const previousPiraquaraStatusRef = useRef('')

  const selected = selectedId === PIRAQUARA_EXECUTION.id
    ? PIRAQUARA_EXECUTION
    : models.find(model => model.id === selectedId) || models[0]
  const headers = workbookRows[Math.max(0, headerRow - 1)]?.map((value, index) => String(value || `Coluna ${index + 1}`).trim()) || []

  useEffect(() => {
    localStorage.setItem('cbn-piraquara-execution', JSON.stringify({ notes: piraquaraObservation }))
  }, [piraquaraObservation])

  const playReadySignal = () => {
    if (!piraquaraAlertsEnabled || !audioContextRef.current) return
    const context = audioContextRef.current
    if (context.state === 'suspended') context.resume()
    const oscillator = context.createOscillator()
    const gain = context.createGain()
    oscillator.type = 'sine'
    oscillator.frequency.setValueAtTime(784, context.currentTime)
    oscillator.frequency.setValueAtTime(988, context.currentTime + 0.18)
    gain.gain.setValueAtTime(0.0001, context.currentTime)
    gain.gain.exponentialRampToValueAtTime(0.16, context.currentTime + 0.02)
    gain.gain.exponentialRampToValueAtTime(0.0001, context.currentTime + 0.45)
    oscillator.connect(gain)
    gain.connect(context.destination)
    oscillator.start()
    oscillator.stop(context.currentTime + 0.46)
  }

  const enableReadySignal = async () => {
    const AudioContextClass = window.AudioContext || window.webkitAudioContext
    if (!AudioContextClass) {
      setMessage({ type: 'error', text: 'Este navegador não permite o aviso sonoro. O aviso visual continuará funcionando.' })
      return
    }
    audioContextRef.current ||= new AudioContextClass()
    await audioContextRef.current.resume()
    setPiraquaraAlertsEnabled(true)
    localStorage.setItem('cbn-ready-sound', 'on')
    const context = audioContextRef.current
    const oscillator = context.createOscillator()
    const gain = context.createGain()
    oscillator.frequency.value = 880
    gain.gain.setValueAtTime(0.12, context.currentTime)
    gain.gain.exponentialRampToValueAtTime(0.0001, context.currentTime + 0.22)
    oscillator.connect(gain)
    gain.connect(context.destination)
    oscillator.start()
    oscillator.stop(context.currentTime + 0.23)
    setMessage({ type: 'success', text: 'Aviso sonoro ativado. Este foi o som que tocará quando os pedidos estiverem prontos.' })
  }

  useEffect(() => {
    if (!session?.user?.id || selectedId !== PIRAQUARA_EXECUTION.id) return undefined
    let active = true
    const loadLatestExecution = async () => {
      const { data, error } = await supabase
        .from('piraquara_execucoes')
        .select('id, status, observacao, total_pedidos, total_itens, mensagem_erro, criado_em, concluido_em')
        .eq('user_id', session.user.id)
        .order('criado_em', { ascending: false })
        .limit(1)
        .maybeSingle()
      if (!active) return
      if (error) {
        setPiraquaraStatusError('Não foi possível atualizar o andamento agora.')
        return
      }
      setPiraquaraStatusError('')
      if (!data) return
      const previousStatus = previousPiraquaraStatusRef.current
      setPiraquaraRequest(data)
      previousPiraquaraStatusRef.current = data.status
      if (previousStatus && previousStatus !== data.status && ['AGUARDANDO_REVISAO', 'CONCLUIDO'].includes(data.status)) {
        document.title = '✅ Pedidos prontos - CBN'
        playReadySignal()
      }
    }
    loadLatestExecution()
    const timer = window.setInterval(loadLatestExecution, 10000)
    return () => {
      active = false
      window.clearInterval(timer)
    }
  }, [session?.user?.id, selectedId, piraquaraAlertsEnabled])

  const converted = useMemo(() => {
    if (!workbookRows.length || !eanColumn || !quantityColumn) return []
    const eanIndex = headers.indexOf(eanColumn)
    const quantityIndex = headers.indexOf(quantityColumn)
    const packageIndex = headers.indexOf(packageColumn)
    const industryIndex = headers.indexOf('Industria')
    if (eanIndex < 0 || quantityIndex < 0) return []
    const validRows = workbookRows
      .slice(headerRow)
      .map(row => {
        const EAN = cleanEan(row[eanIndex])
        const ordered = cleanQuantity(row[quantityIndex])
        const pack = quantityMode === 'multiply' ? cleanQuantity(row[packageIndex]) : 1
        const quantity = ordered !== null && pack !== null ? ordered * pack : null
        const declaredIndustry = industryIndex >= 0 ? String(row[industryIndex] || '').toUpperCase() : ''
        const candidate = {
          EAN,
          Quantidade: quantity,
          Industria: declaredIndustry,
        }
        return { ...candidate, Industria: resolveTrustedIndustry(candidate, catalogIndustries, catalogByCode) }
      })
      .filter(item =>
        item.EAN &&
        item.Quantidade !== null &&
        item.Quantidade !== 0 &&
        (outputIndustry === 'TODAS' || item.Industria === outputIndustry)
      )
    return mergeItemsByEan(validRows)
  }, [workbookRows, headerRow, eanColumn, quantityColumn, packageColumn, quantityMode, outputIndustry, catalogIndustries, catalogByCode, headers.join('|')])

  const detailedPreview = useMemo(() => {
    const order = { RECKITT: 0, LOREAL: 1, '3M': 2, NAO_IDENTIFICADA: 3 }
    return orderDetails
      .map(item => ({ ...item, Industria: resolveTrustedIndustry(item, catalogIndustries, catalogByCode) }))
      .sort((a, b) => (order[a.Industria] ?? 3) - (order[b.Industria] ?? 3))
  }, [orderDetails, catalogIndustries, catalogByCode])

  const completeItems = useMemo(() => detailedPreview, [detailedPreview])

  const comparisonResult = useMemo(() => {
    const budgetByEan = new Map()
    for (const item of budgetRows) {
      const current = budgetByEan.get(item.EAN)
      if (current) {
        current.Quantidade += item.Quantidade
        current.ValorTotal += item.ValorTotal || 0
        if (!current.Item && item.Item) current.Item = item.Item
      } else {
        budgetByEan.set(item.EAN, { ...item })
      }
    }

    const orderByEan = new Map()
    for (const item of comparisonOrderRows) {
      const budget = budgetByEan.get(item.EAN)
      const budgetUnit = budget?.ValorUnitario || (budget?.Quantidade ? (budget.ValorTotal || 0) / budget.Quantidade : 0)
      const value = item.ValorTotal || budgetUnit * item.Quantidade
      const industry = resolveTrustedIndustry({ ...item, Item: item.Item || budget?.Item }, catalogIndustries, catalogByCode)
      const normalizedItem = {
        ...item,
        Item: item.Item && item.Item !== 'Item sem descrição' ? item.Item : budget?.Item || 'Item sem descrição',
        CodigoInterno: item.CodigoInterno || budget?.CodigoInterno || '',
        ValorUnitario: item.ValorUnitario || budgetUnit,
        ValorTotal: value,
        Industria: industry,
      }
      const current = orderByEan.get(item.EAN)
      if (current) {
        current.Quantidade += normalizedItem.Quantidade
        current.ValorTotal += normalizedItem.ValorTotal
      } else {
        orderByEan.set(item.EAN, normalizedItem)
      }
    }

    const industryOrder = { RECKITT: 0, LOREAL: 1, '3M': 2, NAO_IDENTIFICADA: 3 }
    const pedido = [...orderByEan.values()].sort((a, b) =>
      (industryOrder[a.Industria] ?? 3) - (industryOrder[b.Industria] ?? 3)
    )
    const excluidos = [...budgetByEan.values()]
      .filter(item => !orderByEan.has(item.EAN))
      .map(item => ({
        ...item,
        Industria: resolveTrustedIndustry(item, catalogIndustries, catalogByCode),
      }))
      .sort((a, b) => (industryOrder[a.Industria] ?? 3) - (industryOrder[b.Industria] ?? 3))

    return { pedido, excluidos }
  }, [budgetRows, comparisonOrderRows, catalogIndustries, catalogByCode])

  const comparisonLoss = useMemo(() => {
    const groups = ['RECKITT', 'LOREAL', '3M', 'NAO_IDENTIFICADA']
    const byIndustry = Object.fromEntries(groups.map(industry => {
      const rows = comparisonResult.excluidos.filter(item => item.Industria === industry)
      return [industry, {
        quantidade: rows.reduce((sum, item) => sum + item.Quantidade, 0),
        valor: rows.reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
        itens: rows.length,
      }]
    }))
    return {
      byIndustry,
      quantidade: comparisonResult.excluidos.reduce((sum, item) => sum + item.Quantidade, 0),
      valor: comparisonResult.excluidos.reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
    }
  }, [comparisonResult])

  useEffect(() => {
    let active = true

    const loadCatalog = async () => {
      const records = []
      for (let start = 0; ; start += 1000) {
        const { data, error } = await supabase
          .from('catalogo_produtos')
          .select('id,ean,industria,codigo_cbn,origem,tipo_registro')
          .eq('user_id', session.user.id)
          .order('id', { ascending: true })
          .range(start, start + 999)

        if (error) {
          if (active) setMessage({ type: 'error', text: 'Não foi possível carregar a separação por indústria.' })
          return
        }
        records.push(...(data || []))
        if (!data || data.length < 1000) break
      }

      if (active) {
        const { industryMap, codeMap } = buildTrustedCatalogMaps(records)
        setCatalogIndustries(industryMap)
        setCatalogByCode(codeMap)
      }
    }

    loadCatalog()
    return () => { active = false }
  }, [session.user.id])

  useEffect(() => {
    setHeaderRow(selected.headerRow || 1)
    setEanColumn(selected.eanColumn || '')
    setQuantityColumn(selected.quantityColumn || '')
    setPackageColumn(selected.packageColumn || '')
    setQuantityMode(selected.quantityMode || 'direct')
    setMessage(null)
  }, [selectedId])

  useEffect(() => {
    if (!headers.length) return
    if (!headers.includes(eanColumn)) setEanColumn(findSuggestedColumn(headers, ['ean', 'codigo de barras', 'código de barras', 'gtin', 'codigo produto']))
    if (!headers.includes(quantityColumn)) setQuantityColumn(findSuggestedColumn(headers, ['quantidade', 'qtd', 'qtde', 'pedido', 'volume']))
  }, [headerRow, workbookRows])

  const persistModels = next => {
    setModels(next)
    localStorage.setItem('cbn-models', JSON.stringify(next.filter(model => !model.fixed || model.eanColumn || model.quantityColumn)))
  }

  const registerNewCatalogItems = async rows => {
    const candidates = rows.filter(item => item.EAN && !catalogIndustries[item.EAN])
    if (!candidates.length) return

    const eansToCheck = [...new Set(candidates.flatMap(item => {
      const values = [item.EAN]
      if (item.EAN.length === 13 && item.EAN.startsWith('0')) values.push(item.EAN.slice(1))
      return values
    }))]
    const { data: existing } = await supabase
      .from('catalogo_produtos')
      .select('ean')
      .eq('user_id', session.user.id)
      .in('ean', eansToCheck)

    const existingEans = new Set((existing || []).flatMap(item => [item.ean, item.ean?.length === 12 ? `0${item.ean}` : item.ean]))
    const newItems = candidates
      .filter(item => !existingEans.has(item.EAN))
      .map(item => ({
        user_id: session.user.id,
        codigo_cbn: `AUTO-${item.EAN}`,
        produto: item.Item || 'Item novo para revisar',
        ean: item.EAN,
        industria: 'NAO_IDENTIFICADA',
        tipo_registro: 'A_REVISAR',
        validado: false,
        origem: 'Leitura automática — pendente de conferência',
      }))

    if (newItems.length) {
      await supabase.from('catalogo_produtos').upsert(newItems, {
        onConflict: 'user_id,codigo_cbn,ean',
        ignoreDuplicates: true,
      })
    }
  }

  const addFileName = (previous, name, append) => append && previous ? `${previous}, ${name}` : name

  const enrichRowsWithCatalog = rows => enrichRowsFromTrustedCatalog(rows, catalogIndustries, catalogByCode)

  const readFile = async (file, append = false) => {
    if (!file) return
    const extension = file.name.split('.').pop()?.toLowerCase()
    if (!['xlsx', 'xls', 'csv', 'pdf', 'png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
      setMessage({ type: 'error', text: 'Use Excel, CSV, PDF, PNG, JPG ou WEBP.' })
      return -1
    }
    try {
      if (['png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
        setMessage({ type: 'success', text: `Lendo ${file.name} por OCR. Isso pode levar alguns minutos…` })
        const recognizedText = await recognizeImage(file)
        const rows = enrichRowsWithCatalog(parseVariableText(recognizedText, catalogIndustries))
        const printedTotal = extractPrintedDocumentTotal(recognizedText)
        if (!rows.length) throw new Error('A imagem foi lida, mas o formato da tabela ainda não foi reconhecido. O problema está no leitor, não na qualidade da foto.')
        await registerNewCatalogItems(rows)
        const imageModelId = selectedId === 'dalpar' ? 'dalpar' : 'personalizado'
        if (!append) setSelectedId(imageModelId)
        setSourceOrderTotal(previous => append ? previous + printedTotal : printedTotal)
        setOrderDetails(previous => mergeItemsByEan(append ? [...previous, ...rows] : rows))
        setWorkbookRows(previous => {
          const added = rows.map(row => [row.EAN, row.Quantidade])
          return append && previous.length ? [previous[0], ...previous.slice(1), ...added] : [['EAN', 'Quantidade'], ...added]
        })
        setFileName(previous => addFileName(previous, file.name, append))
        setSheetName(imageModelId === 'dalpar' ? 'Pedido Dalpar lido por OCR — revisar' : 'Imagens lidas por OCR — revisar')
        setHeaderRow(1)
        setEanColumn('EAN')
        setQuantityColumn('Quantidade')
        setPackageColumn('')
        setQuantityMode('direct')
        return rows.length
      }

      if (extension === 'pdf') {
        setMessage({ type: 'success', text: `Lendo e identificando ${file.name}…` })
        let modelId
        let rows
        let documentTotal = 0
        let storeInfo = null
        try {
          const parsed = await readPdfOrder(file)
          modelId = parsed.modelId
          rows = parsed.rows
          documentTotal = parsed.documentTotal || 0
          storeInfo = parsed.storeInfo || null
        } catch (pdfError) {
          setMessage({ type: 'success', text: `PDF escaneado detectado em ${file.name}. Iniciando OCR…` })
          const recognizedText = await recognizeScannedPdf(file)
          rows = parseVariableText(recognizedText, catalogIndustries)
          documentTotal = extractPrintedDocumentTotal(recognizedText)
          modelId = 'personalizado'
          if (!rows.length) throw pdfError
        }
        rows = enrichRowsWithCatalog(rows)
        await registerNewCatalogItems(rows)
        if (!append) setSelectedId(modelId)
        setSourceOrderTotal(previous => append ? previous + documentTotal : documentTotal)
        setPiraquaraStoreInfo(previous => {
          if (modelId !== 'piraquara') return append ? previous : null
          if (!append) return storeInfo
          if (!previous || !storeInfo) return previous || storeInfo
          return previous.cnpjFinal === storeInfo.cnpjFinal && previous.loja === storeInfo.loja
            ? previous
            : { multiple: true }
        })
        setOrderDetails(previous => mergeItemsByEan(append ? [...previous, ...rows] : rows))
        setWorkbookRows(previous => {
          const added = rows.map(row => [row.EAN, row.Quantidade, row.Industria || ''])
          return append && previous.length ? [previous[0], ...previous.slice(1), ...added] : [['EAN', 'Quantidade', 'Industria'], ...added]
        })
        setFileName(previous => addFileName(previous, file.name, append))
        setSheetName(
          modelId === 'piraquara' && storeInfo
            ? `Piraquara · CNPJ final ${storeInfo.cnpjFinal} · Loja ${storeInfo.loja}`
            : 'Pedidos extraídos de PDF'
        )
        setHeaderRow(1)
        setEanColumn('EAN')
        setQuantityColumn('Quantidade')
        setPackageColumn('')
        setQuantityMode('direct')
        return rows.length
      }

      const data = await file.arrayBuffer()
      const workbook = XLSX.read(data, { type: 'array', cellDates: true })
      const firstSheet = workbook.SheetNames[0]
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1, defval: '', raw: false })
      if (rows.length < 2) throw new Error('A planilha não contém linhas de produtos para processar.')
      setOrderDetails([])
      setSourceOrderTotal(previous => append ? previous : 0)
      setWorkbookRows(previous => append && previous.length ? [previous[0], ...previous.slice(1), ...rows.slice(1)] : rows)
      setFileName(previous => addFileName(previous, file.name, append))
      setSheetName(append ? 'Arquivos combinados' : firstSheet)
      return Math.max(0, rows.length - 1)
    } catch (readError) {
      setMessage({ type: 'error', text: `${file.name}: ${readError.message || 'não foi possível ler o arquivo.'}` })
      return -1
    }
  }

  const stageFiles = (fileList, append = false) => {
    const files = [...(fileList || [])]
    if (!files.length) return
    setPendingFiles(previous => append ? [...previous, ...files] : files)
    setMessage(null)
    if (fileRef.current) fileRef.current.value = ''
  }

  const readFiles = async (files, appendExisting = false) => {
    if (!files.length) return { totalRead: 0, failures: 0 }
    let totalRead = 0
    let failures = 0
    for (let index = 0; index < files.length; index += 1) {
      const result = await readFile(files[index], appendExisting || index > 0)
      if (result < 0) failures += 1
      else totalRead += result
    }

    if (totalRead > 0) {
      setMessage({
        type: 'success',
        text: `${files.length - failures} arquivo(s) processado(s). ${totalRead} linhas ou itens adicionados ao pedido.`,
      })
    } else {
      setMessage({
        type: 'error',
        text: 'Os arquivos foram abertos, mas nenhum item válido foi encontrado. Confira o formato ou envie os arquivos para ajustarmos o leitor.',
      })
    }
    return { totalRead, failures }
  }

  const processPendingFiles = async () => {
    if (!pendingFiles.length || processingFiles) return
    setProcessingFiles(true)
    const result = await readFiles(pendingFiles, workbookRows.length > 0)
    setProcessingFiles(false)
    if (result.totalRead > 0) setPendingFiles([])
  }

  const saveCurrentModel = async () => {
    const updated = { ...selected, headerRow, eanColumn, quantityColumn, packageColumn, quantityMode }
    const next = models.map(model => model.id === selected.id ? updated : model)
    persistModels(next)

    if (supabase) {
      const { error } = await supabase.from('modelos_planilha').upsert({
        id: updated.id,
        nome: updated.name,
        descricao: updated.description,
        linha_cabecalho: updated.headerRow,
        coluna_ean: updated.eanColumn,
        coluna_quantidade: updated.quantityColumn,
        coluna_embalagem: updated.packageColumn,
        regra_quantidade: updated.quantityMode,
        atualizado_em: new Date().toISOString(),
      })
      if (error) {
        setMessage({ type: 'error', text: 'Modelo salvo neste computador, mas o Supabase ainda precisa de autenticação/configuração.' })
        return
      }
    }
    setMessage({ type: 'success', text: 'Configuração do modelo salva.' })
  }

  const addCustomModel = () => {
    const name = window.prompt('Nome do novo modelo:')
    if (!name?.trim()) return
    const model = {
      id: `personalizado-${Date.now()}`,
      name: name.trim(),
      description: 'Modelo variável editável para outros clientes',
      headerRow: 1,
      eanColumn: '',
      quantityColumn: '',
      packageColumn: '',
      quantityMode: 'direct',
      fixed: false,
    }
    const next = [...models, model]
    persistModels(next)
    setSelectedId(model.id)
  }

  const removeModel = id => {
    const model = models.find(item => item.id === id)
    if (!model || model.fixed) return
    if (!window.confirm(`Excluir o modelo "${model.name}"?`)) return
    const next = models.filter(item => item.id !== id)
    persistModels(next)
    setSelectedId('piraquara')
  }

  const exportFile = () => {
    if (!converted.length) {
      setMessage({ type: 'error', text: 'Escolha as colunas de EAN e Quantidade antes de exportar.' })
      return
    }
    const repposRows = converted.map(item => ({ EAN: item.EAN, Quantidade: item.Quantidade }))
    const sheet = XLSX.utils.json_to_sheet(repposRows, { header: ['EAN', 'Quantidade'] })
    sheet['!cols'] = [{ wch: 18 }, { wch: 14 }]
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, 'Pedido')
    const baseName = fileName.replace(/\.[^.]+$/, '') || 'pedido'
    const suffix = outputIndustry === 'RECKITT' ? 'reckitt_reppos' : outputIndustry.toLowerCase().replace(/[^a-z0-9]/g, '')
    const isNamedPiraquaraReppos = selectedId === 'piraquara' &&
      outputIndustry === 'RECKITT' &&
      piraquaraStoreInfo &&
      !piraquaraStoreInfo.multiple
    const downloadName = isNamedPiraquaraReppos
      ? `${piraquaraStoreInfo.cnpjFinal}-${piraquaraStoreInfo.loja}.xlsx`
      : `${baseName}_${suffix}.xlsx`
    XLSX.writeFile(workbook, downloadName)
    setMessage({
      type: 'success',
      text: `Planilha ${downloadName} gerada com ${converted.length} itens da indústria selecionada.`,
    })
  }

  const exportCompleteOrder = () => {
    if (!completeItems.length) {
      setMessage({ type: 'error', text: 'O pedido completo está disponível após carregar um arquivo reconhecido.' })
      return
    }

    const groups = [
      { code: 'RECKITT', label: 'RECKITT' },
      { code: 'LOREAL', label: "L'ORÉAL" },
      { code: '3M', label: '3M' },
      { code: 'NAO_IDENTIFICADA', label: 'ITENS NOVOS / A REVISAR' },
    ]
    const sheetRows = [['EAN', 'COD.', 'Item', 'Quantidade', 'Valor unitário', 'Valor total do item']]
    const sectionRows = []

    for (const group of groups) {
      if (sheetRows.length > 1) sheetRows.push(['', '', '', '', '', ''])
      sectionRows.push(sheetRows.length)
      sheetRows.push([group.label, '', '', '', '', ''])
      for (const item of completeItems.filter(product => product.Industria === group.code)) {
        const unitValue = item.ValorUnitario || (item.Quantidade ? (item.ValorTotal || 0) / item.Quantidade : 0)
        sheetRows.push([item.EAN, item.CodigoInterno || '', item.Item, item.Quantidade, unitValue, item.ValorTotal || 0])
      }
    }

    const sheet = XLSX.utils.aoa_to_sheet(sheetRows)
    sheet['!cols'] = [{ wch: 18 }, { wch: 14 }, { wch: 58 }, { wch: 14 }, { wch: 18 }, { wch: 20 }, { wch: 43 }, { wch: 18 }]
    sheet['!merges'] = sectionRows.map(row => ({ s: { r: row, c: 0 }, e: { r: row, c: 5 } }))
    for (let row = 1; row < sheetRows.length; row += 1) {
      for (const column of ['E', 'F']) {
        const cell = `${column}${row + 1}`
        if (typeof sheetRows[row][column === 'E' ? 4 : 5] === 'number' && sheet[cell]) sheet[cell].z = 'R$ #,##0.00'
      }
    }

    const totals = {
      RECKITT: completeItems.filter(item => item.Industria === 'RECKITT').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
      LOREAL: completeItems.filter(item => item.Industria === 'LOREAL').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
      '3M': completeItems.filter(item => item.Industria === '3M').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
      A_REVISAR: completeItems.filter(item => item.Industria === 'NAO_IDENTIFICADA').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
    }
    const itemsTotal = completeItems.reduce((sum, item) => sum + (item.ValorTotal || 0), 0)
    const grandTotal = sourceOrderTotal > 0 ? sourceOrderTotal : itemsTotal
    const documentAdjustment = sourceOrderTotal > 0 ? sourceOrderTotal - itemsTotal : 0
    const summaryRows = [
      ['RESUMO DO PEDIDO', 'Valor'],
      ['Reckitt', totals.RECKITT],
      ["L'Oréal", totals.LOREAL],
      ['3M', totals['3M']],
      ['Itens novos / a revisar', totals.A_REVISAR],
    ]
    if (Math.abs(documentAdjustment) >= 0.005) {
      summaryRows.push(['Diferença de itens/valores não reconhecidos no OCR', documentAdjustment])
    }
    summaryRows.push(['TOTAL DO PEDIDO (conferido com o documento)', grandTotal])
    XLSX.utils.sheet_add_aoa(sheet, summaryRows, { origin: 'G1' })

    for (let row = 2; row <= summaryRows.length; row += 1) {
      const cell = `H${row}`
      if (sheet[cell]) sheet[cell].z = 'R$ #,##0.00'
    }

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, 'Pedido completo')
    const baseName = fileName.replace(/\.[^.]+$/, '') || 'pedido'
    const completeBaseName = selectedId === 'piraquara' && piraquaraStoreInfo && !piraquaraStoreInfo.multiple
      ? `${piraquaraStoreInfo.cnpjFinal}-${piraquaraStoreInfo.loja}`
      : baseName
    XLSX.writeFile(workbook, `${completeBaseName}-pedido-completo.xlsx`)
    setMessage({ type: 'success', text: `Pedido completo gerado com ${completeItems.length} itens. Total conferido: ${grandTotal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}.` })
  }

  const readComparisonFile = async (file, kind, append = true) => {
    if (!file) return 0
    const extension = file.name.split('.').pop()?.toLowerCase()
    if (!['xlsx', 'xls', 'csv', 'pdf', 'png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
      setMessage({ type: 'error', text: 'Use Excel, CSV, PDF, PNG, JPG ou WEBP.' })
      return 0
    }

    setComparisonLoading(kind)
    setMessage({ type: 'success', text: `Lendo ${file.name}…` })
    try {
      let rows
      if (['png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
        rows = parseVariableText(await recognizeImage(file), catalogIndustries)
      } else if (extension === 'pdf') {
        try {
          rows = (await readPdfOrder(file)).rows
        } catch {
          rows = parseVariableText(await recognizeScannedPdf(file), catalogIndustries)
        }
      } else {
        rows = await parseSpreadsheetOrder(file)
      }

      if (!rows?.length) throw new Error('Nenhum item foi reconhecido neste arquivo.')
      rows = enrichRowsWithCatalog(rows)
      await registerNewCatalogItems(rows)
      if (kind === 'budget') {
        setBudgetRows(previous => mergeItemsByEan(append ? [...previous, ...rows] : rows))
        setBudgetFileName(previous => addFileName(previous, file.name, append))
      } else {
        setComparisonOrderRows(previous => mergeItemsByEan(append ? [...previous, ...rows] : rows))
        setOrderFileName(previous => addFileName(previous, file.name, append))
      }
      return rows.length
    } catch (error) {
      setMessage({ type: 'error', text: `${file.name}: ${error.message || 'não foi possível ler o arquivo.'}` })
      return 0
    } finally {
      setComparisonLoading('')
    }
  }

  const readComparisonFiles = async (fileList, kind) => {
    const files = [...(fileList || [])]
    if (!files.length) return
    let totalRead = 0
    for (const file of files) totalRead += await readComparisonFile(file, kind, true)
    const input = kind === 'budget' ? budgetFileRef.current : orderFileRef.current
    if (input) input.value = ''
    setMessage({ type: 'success', text: `${files.length} arquivo(s) adicionados ao ${kind === 'budget' ? 'orçamento' : 'pedido'}, com ${totalRead} itens lidos.` })
  }

  const resetComparison = kind => {
    if (kind === 'budget') {
      setBudgetRows([])
      setBudgetFileName('')
      if (budgetFileRef.current) budgetFileRef.current.value = ''
    } else {
      setComparisonOrderRows([])
      setOrderFileName('')
      if (orderFileRef.current) orderFileRef.current.value = ''
    }
    setMessage(null)
  }

  const exportComparison = () => {
    const { pedido, excluidos } = comparisonResult
    if (!pedido.length || !budgetRows.length) {
      setMessage({ type: 'error', text: 'Carregue o orçamento e o pedido antes de gerar o confronto.' })
      return
    }

    const groups = [
      { code: 'RECKITT', label: 'RECKITT' },
      { code: 'LOREAL', label: "L'ORÉAL" },
      { code: '3M', label: '3M' },
      { code: 'NAO_IDENTIFICADA', label: 'ITENS NOVOS / A REVISAR' },
    ]
    const sheetRows = [['EAN', 'COD.', 'Item', 'Quantidade', 'Valor unitário', 'Valor total do item', 'Indústria']]
    const sectionRows = []

    for (const group of groups) {
      if (sheetRows.length > 1) sheetRows.push(['', '', '', '', '', '', ''])
      sectionRows.push(sheetRows.length)
      sheetRows.push([group.label, '', '', '', '', '', ''])
      for (const item of pedido.filter(product => product.Industria === group.code)) {
        const unitValue = item.ValorUnitario || (item.Quantidade ? (item.ValorTotal || 0) / item.Quantidade : 0)
        sheetRows.push([item.EAN, item.CodigoInterno || '', item.Item, item.Quantidade, unitValue, item.ValorTotal || 0, group.label])
      }
    }

    sheetRows.push(['', '', '', '', '', '', ''])
    sectionRows.push(sheetRows.length)
    sheetRows.push(['ITENS EXCLUÍDOS DO PEDIDO', '', '', '', '', '', ''])
    sheetRows.push(['EAN', 'COD.', 'Item', 'Quantidade no orçamento', 'Valor unitário', 'Valor no orçamento', 'Indústria'])
    for (const item of excluidos) {
      const industryLabel = item.Industria === 'LOREAL' ? "L'ORÉAL" : item.Industria === 'NAO_IDENTIFICADA' ? 'A REVISAR' : item.Industria
      const unitValue = item.ValorUnitario || (item.Quantidade ? (item.ValorTotal || 0) / item.Quantidade : 0)
      sheetRows.push([item.EAN, item.CodigoInterno || '', item.Item, item.Quantidade, unitValue, item.ValorTotal || 0, industryLabel])
    }
    if (!excluidos.length) sheetRows.push(['Nenhum item excluído', '', '', '', '', '', ''])

    sheetRows.push(['', '', '', '', '', '', ''])
    sectionRows.push(sheetRows.length)
    sheetRows.push(['RESUMO DA PERDA DO PEDIDO', '', '', '', '', '', ''])
    sheetRows.push(['Indústria', '', 'Quantidade retirada', '', '', 'Valor perdido', ''])
    for (const group of groups) {
      const loss = comparisonLoss.byIndustry[group.code]
      sheetRows.push([group.label, '', loss.quantidade, '', '', loss.valor, ''])
    }
    sheetRows.push(['TOTAL DA PERDA', '', comparisonLoss.quantidade, '', '', comparisonLoss.valor, ''])

    const sheet = XLSX.utils.aoa_to_sheet(sheetRows)
    sheet['!cols'] = [{ wch: 18 }, { wch: 14 }, { wch: 58 }, { wch: 23 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 22 }, { wch: 18 }]
    sheet['!merges'] = sectionRows.map(row => ({ s: { r: row, c: 0 }, e: { r: row, c: 6 } }))
    for (let row = 1; row < sheetRows.length; row += 1) {
      for (const column of ['E', 'F']) {
        const cell = `${column}${row + 1}`
        if (sheet[cell]) sheet[cell].z = 'R$ #,##0.00'
      }
    }

    const totals = {
      RECKITT: pedido.filter(item => item.Industria === 'RECKITT').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
      LOREAL: pedido.filter(item => item.Industria === 'LOREAL').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
      '3M': pedido.filter(item => item.Industria === '3M').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
      A_REVISAR: pedido.filter(item => item.Industria === 'NAO_IDENTIFICADA').reduce((sum, item) => sum + (item.ValorTotal || 0), 0),
    }
    const grandTotal = pedido.reduce((sum, item) => sum + (item.ValorTotal || 0), 0)
    const summaryRows = [
      ['RESUMO DO PEDIDO', 'Valor'],
      ['Reckitt', totals.RECKITT],
      ["L'Oréal", totals.LOREAL],
      ['3M', totals['3M']],
      ['Itens novos / a revisar', totals.A_REVISAR],
      ['TOTAL DO PEDIDO', grandTotal],
    ]
    XLSX.utils.sheet_add_aoa(sheet, summaryRows, { origin: 'H1' })
    for (let row = 2; row <= summaryRows.length; row += 1) {
      const cell = `I${row}`
      if (sheet[cell]) sheet[cell].z = 'R$ #,##0.00'
    }

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, 'Confronto')
    const baseName = orderFileName.replace(/\.[^.]+$/, '') || 'pedido'
    XLSX.writeFile(workbook, `${baseName}_confronto.xlsx`)
    setMessage({ type: 'success', text: `Confronto gerado: ${pedido.length} itens no pedido e ${excluidos.length} itens excluídos.` })
  }

  const savePassword = async event => {
    event.preventDefault()
    setPasswordMessage(null)
    if (newPassword.length < 8) {
      setPasswordMessage({ type: 'error', text: 'Use uma senha com pelo menos 8 caracteres.' })
      return
    }
    if (newPassword !== confirmPassword) {
      setPasswordMessage({ type: 'error', text: 'As duas senhas precisam ser iguais.' })
      return
    }

    setSavingPassword(true)
    const { error: authError } = await supabase.auth.updateUser({ password: newPassword })
    setSavingPassword(false)
    if (authError) {
      setPasswordMessage({ type: 'error', text: 'Não foi possível salvar a senha. Tente novamente.' })
      return
    }

    setNewPassword('')
    setConfirmPassword('')
    setPasswordMessage({ type: 'success', text: 'Senha criada com sucesso. Agora você pode entrar com e-mail e senha.' })
  }

  const requestPiraquaraExecution = async () => {
    const note = piraquaraObservation.trim()
    setPiraquaraSubmitting(true)
    setMessage(null)
    const { data, error } = await supabase
      .from('piraquara_execucoes')
      .insert({
        user_id: session.user.id,
        observacao: note,
        escopo: 'TODOS_OS_PEDIDOS_NOVOS_DO_EMAIL',
        versao_regras: 'piraquara-v2',
        avisar_whatsapp: true,
      })
      .select('id, status, observacao, criado_em')
      .single()
    setPiraquaraSubmitting(false)

    if (error) {
      setMessage({ type: 'error', text: 'Não foi possível registrar a execução. Tente novamente.' })
      return
    }

    setPiraquaraRequest(data)
    setPiraquaraObservation('')
    previousPiraquaraStatusRef.current = data.status
    const command = buildPiraquaraCommand(note, data.id)
    try {
      await navigator.clipboard.writeText(command)
      setMessage({ type: 'success', text: 'Solicitação registrada. O comando curto foi copiado: agora cole no ChatGPT Work para iniciar.' })
    } catch {
      setMessage({ type: 'success', text: 'Solicitação registrada. Use o botão “Copiar comando” e cole no ChatGPT Work para iniciar.' })
    }
  }

  const copyPiraquaraCommand = async () => {
    if (!piraquaraRequest?.id) return
    const command = buildPiraquaraCommand(piraquaraRequest.observacao, piraquaraRequest.id)
    try {
      await navigator.clipboard.writeText(command)
      setMessage({ type: 'success', text: 'Comando copiado. Cole aqui no ChatGPT Work para eu executar os pedidos.' })
    } catch {
      setMessage({ type: 'error', text: `Copie e cole no chat: ${command}` })
    }
  }

  const resetFile = () => {
    setPendingFiles([])
    setWorkbookRows([])
    setOrderDetails([])
    setSourceOrderTotal(0)
    setPiraquaraStoreInfo(null)
    setFileName('')
    setSheetName('')
    setEanColumn(selected.eanColumn || '')
    setQuantityColumn(selected.quantityColumn || '')
    setPackageColumn(selected.packageColumn || '')
    setMessage(null)
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark"><FileSpreadsheet size={25} /></div>
          <div><strong>Gestão de Pedidos</strong><span>CBN Distribuidora</span></div>
        </div>

        <div className="side-label automation-label">AUTOMAÇÃO</div>
        <button
          className={`automation-shortcut ${selectedId === PIRAQUARA_EXECUTION.id ? 'active' : ''}`}
          onClick={() => { setSelectedId(PIRAQUARA_EXECUTION.id); setMessage(null) }}
        >
          <PlayCircle size={19} />
          <span><strong>Executar pedidos Piraquara</strong><small>Preparar solicitação</small></span>
          <ChevronRight size={16} />
        </button>

        <div className="side-label">MODELOS DE PLANILHA</div>
        <nav className="model-list">
          {models.map(model => (
            <div className={`model-row ${selectedId === model.id ? 'active' : ''}`} key={model.id}>
              <button onClick={() => setSelectedId(model.id)}>
                <FileSpreadsheet size={18} />
                <span><strong>{model.name}</strong><small>{model.fixed ? 'Modelo principal' : 'Editável'}</small></span>
                <ChevronRight size={16} />
              </button>
              {!model.fixed && model.id !== 'personalizado' && (
                <button className="remove-model" title="Excluir modelo" onClick={() => removeModel(model.id)}><Trash2 size={15} /></button>
              )}
            </div>
          ))}
        </nav>
        <button className="add-model" onClick={addCustomModel}><Plus size={17} /> Adicionar modelo</button>

        <div className="connection-card">
          <Database size={18} />
          <div><strong>Armazenamento</strong><span>{supabase ? 'Supabase configurado' : 'Local — Supabase pendente'}</span></div>
          <i className={supabase ? 'online' : ''}></i>
        </div>
      </aside>

      <main>
        <header className="topbar">
          <div>
            <p>CONVERSOR DE PEDIDOS</p>
            <h1>{selected.name}</h1>
            <span>{selected.description}</span>
          </div>
<div className="top-actions">
            <button className="ghost-button" onClick={() => { setPasswordMessage(null); setShowPasswordSetup(true) }}><KeyRound size={17} /> Criar/alterar senha</button>
            {selectedId !== PIRAQUARA_EXECUTION.id && <button className="ghost-button" onClick={saveCurrentModel}><Settings2 size={17} /> Salvar configuração</button>}
            <button className="ghost-button" onClick={() => supabase.auth.signOut()} title={session.user.email}><LogOut size={17} /> Sair</button>
          </div>
        </header>

        <section className="content">
          {message && (
            <div className={`message ${message.type}`}>
              {message.type === 'success' ? <CheckCircle2 size={19} /> : <AlertCircle size={19} />}
              {message.text}
            </div>
          )}

          {selectedId === PIRAQUARA_EXECUTION.id ? (
            <>
              <div className="execution-intro">
                <div className="execution-intro-icon"><PlayCircle size={28} /></div>
                <div>
                  <h2>Executar pedidos Piraquara</h2>
                  <p>Escreva o recorte desejado; sem observação, o comando considera todos os pedidos novos, do mais antigo ao mais recente.</p>
                </div>
              </div>

              <div className="execution-card">
                <label className="execution-notes">
                  Observação (opcional)
                  <textarea
                    value={piraquaraObservation}
                    onChange={event => setPiraquaraObservation(event.target.value)}
                    placeholder="Ex.: execute os 3 pedidos do dia 16/09/26 começando pelo mais antigo."
                    rows="5"
                  />
                </label>

                <div className="safety-note">
                  <ShieldCheck size={22} />
                  <div><strong>Revisão manual obrigatória</strong><span>O processo prepara o carrinho no Reppos, mas nunca finaliza, confirma ou envia a compra automaticamente.</span></div>
                </div>

                <div className="execution-actions">
                  <button className="primary-button execution-submit" disabled={piraquaraSubmitting} onClick={requestPiraquaraExecution}>
                    <PlayCircle size={19} /> {piraquaraSubmitting ? 'Registrando…' : 'Registrar execução'}
                  </button>
                </div>
                {piraquaraRequest && (
                  <div className={`execution-status ${EXECUTION_STATUS[piraquaraRequest.status]?.tone || 'waiting'}`}>
                    <CheckCircle2 size={21} />
                    <div className="execution-status-copy">
                      <strong>{EXECUTION_STATUS[piraquaraRequest.status]?.title || 'Andamento atualizado'}</strong>
                      <span>{piraquaraRequest.mensagem_erro || EXECUTION_STATUS[piraquaraRequest.status]?.text}</span>
                      {(piraquaraRequest.total_pedidos > 0 || piraquaraRequest.total_itens > 0) && (
                        <small>{piraquaraRequest.total_pedidos || 0} pedido(s) · {piraquaraRequest.total_itens || 0} item(ns)</small>
                      )}
                    </div>
                    {piraquaraRequest.status === 'SOLICITADO' && (
                      <button className="status-action" onClick={copyPiraquaraCommand}>Copiar comando</button>
                    )}
                  </div>
                )}
                <div className="ready-alert-row">
                  <div><strong>Aviso de conclusão</strong><span>A tela fica verde e o título da aba mostra ✅ quando os carrinhos estiverem prontos.</span></div>
                  <button className="ghost-button" onClick={enableReadySignal} disabled={piraquaraAlertsEnabled}>
                    {piraquaraAlertsEnabled ? 'Som ativado' : 'Ativar aviso sonoro'}
                  </button>
                </div>
                {piraquaraStatusError && <p className="execution-help error-text">{piraquaraStatusError}</p>}
                <p className="execution-help">Para iniciar, registre e cole no ChatGPT Work o comando “Execução pedidos Piraquara: …”. A observação define o recorte desta execução.</p>
              </div>
            </>
          ) : selectedId === 'confronto' ? (
            <>
              <div className="comparison-intro">
                <h2>Confronto de arquivos</h2>
                <p>Envie primeiro o orçamento original e depois o pedido devolvido pelo cliente. Aceita Excel, CSV, PDF, foto ou print.</p>
              </div>
              <div className="comparison-upload-grid">
                <div className={`comparison-upload ${budgetRows.length ? 'loaded' : ''}`}>
                  <span className="comparison-step">1</span>
                  <FileSpreadsheet size={28} />
                  <h3>Orçamento</h3>
                  <p className="uploaded-file-names">{budgetFileName || 'Arquivo(s) que foram enviados ao cliente'}</p>
                  <input ref={budgetFileRef} type="file" multiple accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp" hidden onChange={event => readComparisonFiles(event.target.files, 'budget')} />
                  <button className="ghost-button" onClick={() => budgetFileRef.current?.click()} disabled={comparisonLoading === 'budget'}>
                    <UploadCloud size={17} /> {comparisonLoading === 'budget' ? 'Lendo…' : budgetRows.length ? 'Adicionar mais arquivos' : 'Selecionar orçamento(s)'}
                  </button>
                  {budgetRows.length > 0 && <button className="comparison-remove" onClick={() => resetComparison('budget')}>Remover</button>}
                  {budgetRows.length > 0 && <strong>{budgetRows.length} itens lidos</strong>}
                </div>
                <div className={`comparison-upload ${comparisonOrderRows.length ? 'loaded' : ''}`}>
                  <span className="comparison-step">2</span>
                  <FileSpreadsheet size={28} />
                  <h3>Pedido</h3>
                  <p className="uploaded-file-names">{orderFileName || 'Arquivo(s) devolvidos pelo cliente'}</p>
                  <input ref={orderFileRef} type="file" multiple accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp" hidden onChange={event => readComparisonFiles(event.target.files, 'order')} />
                  <button className="ghost-button" onClick={() => orderFileRef.current?.click()} disabled={comparisonLoading === 'order'}>
                    <UploadCloud size={17} /> {comparisonLoading === 'order' ? 'Lendo…' : comparisonOrderRows.length ? 'Adicionar mais arquivos' : 'Selecionar pedido(s)'}
                  </button>
                  {comparisonOrderRows.length > 0 && <button className="comparison-remove" onClick={() => resetComparison('order')}>Remover</button>}
                  {comparisonOrderRows.length > 0 && <strong>{comparisonOrderRows.length} itens lidos</strong>}
                </div>
              </div>

              {budgetRows.length > 0 && comparisonOrderRows.length > 0 && (
                <>
                  <div className="preview-card">
                    <div className="section-heading">
                      <div><span>3</span><div><h2>Prévia do confronto</h2><p>Pedido organizado por indústria e itens retirados pelo cliente.</p></div></div>
                      <b>{comparisonResult.pedido.length} itens no pedido</b>
                    </div>
                    <div className="comparison-summary">
                      {['RECKITT', 'LOREAL', '3M', 'NAO_IDENTIFICADA'].map(industry => {
                        const rows = comparisonResult.pedido.filter(item => item.Industria === industry)
                        const value = rows.reduce((sum, item) => sum + item.ValorTotal, 0)
                        const label = industry === 'LOREAL' ? "L'Oréal" : industry === 'NAO_IDENTIFICADA' ? 'A revisar' : industry
                        return <div key={industry}><span>{label}</span><strong>{value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</strong><small>{rows.length} itens</small></div>
                      })}
                      <div className="comparison-total"><span>Total do pedido</span><strong>{comparisonResult.pedido.reduce((sum, item) => sum + item.ValorTotal, 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</strong><small>{comparisonResult.pedido.length} itens</small></div>
                    </div>
                    <div className="comparison-excluded">
                      <h3>Itens excluídos do pedido <span>{comparisonResult.excluidos.length}</span></h3>
                      <div className="loss-report">
                        <div className="loss-total">
                          <span>Perda total do pedido</span>
                          <strong>{comparisonLoss.valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</strong>
                          <small>{comparisonLoss.quantidade} unidades retiradas</small>
                        </div>
                        {['RECKITT', 'LOREAL', '3M', 'NAO_IDENTIFICADA'].map(industry => {
                          const loss = comparisonLoss.byIndustry[industry]
                          const label = industry === 'LOREAL' ? "L'Oréal" : industry === 'NAO_IDENTIFICADA' ? 'A revisar' : industry
                          return (
                            <div key={industry}>
                              <span>{label}</span>
                              <strong>{loss.valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</strong>
                              <small>{loss.quantidade} unidades · {loss.itens} itens</small>
                            </div>
                          )
                        })}
                      </div>
                      <div className="table-wrap">
                        <table>
                          <thead><tr><th>EAN</th><th>Item</th><th>Quantidade no orçamento</th><th>Valor</th><th>Indústria</th></tr></thead>
                          <tbody>
                            {comparisonResult.excluidos.slice(0, 12).map((row, index) => (
                              <tr key={index}><td>{row.EAN}</td><td>{row.Item}</td><td>{row.Quantidade}</td><td>{(row.ValorTotal || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td><td>{row.Industria === 'LOREAL' ? "L'Oréal" : row.Industria === 'NAO_IDENTIFICADA' ? 'A revisar' : row.Industria}</td></tr>
                            ))}
                            {!comparisonResult.excluidos.length && <tr><td className="empty" colSpan="5">Nenhum item foi excluído do pedido.</td></tr>}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                  <div className="action-bar">
                    <div><strong>Planilha do confronto pronta</strong><span>Pedido completo por indústria, totais e itens excluídos.</span></div>
                    <button className="primary-button" onClick={exportComparison}><Download size={19} /> Baixar confronto</button>
                  </div>
                </>
              )}
            </>
          ) : (
            <>
          {!workbookRows.length ? (
            <div
              className={`dropzone ${isDragging ? 'dragging' : ''} ${pendingFiles.length ? 'has-pending-files' : ''}`}
              onDragOver={event => { event.preventDefault(); setIsDragging(true) }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={event => { event.preventDefault(); setIsDragging(false); stageFiles(event.dataTransfer.files, false) }}
              onClick={() => !pendingFiles.length && fileRef.current?.click()}
            >
              <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp" onChange={event => stageFiles(event.target.files, false)} hidden />
              {pendingFiles.length ? (
                <>
                  <div className="upload-icon"><FileSpreadsheet size={34} /></div>
                  <h2>{pendingFiles.length} arquivo(s) selecionado(s)</h2>
                  <div className="pending-file-list">
                    {pendingFiles.map((file, index) => <span key={`${file.name}-${index}`}>{file.name}</span>)}
                  </div>
                  <div className="pending-file-actions">
                    <button className="secondary-upload-button" onClick={event => { event.stopPropagation(); fileRef.current?.click() }}>Alterar seleção</button>
                    <button className="process-files-button" disabled={processingFiles} onClick={event => { event.stopPropagation(); processPendingFiles() }}>
                      {processingFiles ? 'Processando…' : 'Processar arquivos'}
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <div className="upload-icon"><UploadCloud size={34} /></div>
                  <h2>Envie a planilha do cliente</h2>
                  <p>Arraste um ou vários arquivos para cá ou clique para escolher</p>
                  <span>Excel, CSV, PDF, fotos ou prints</span>
                  <button>Selecionar arquivo(s)</button>
                </>
              )}
            </div>
          ) : (
            <>
              <div className="file-card">
                <div className="file-icon"><FileSpreadsheet size={25} /></div>
                <div><strong>{fileName}</strong><span>{sheetName} · {workbookRows.length} linhas combinadas</span></div>
                <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp" onChange={event => stageFiles(event.target.files, true)} hidden />
                <div className="file-card-actions">
                  <button onClick={() => fileRef.current?.click()}><Plus size={16} /> Adicionar mais arquivos</button>
                  <button onClick={resetFile}><RotateCcw size={16} /> Limpar</button>
                </div>
              </div>
              {pendingFiles.length > 0 && (
                <div className="pending-processing-card">
                  <div>
                    <strong>{pendingFiles.length} novo(s) arquivo(s) aguardando</strong>
                    <span>{pendingFiles.map(file => file.name).join(', ')}</span>
                  </div>
                  <button className="primary-button" disabled={processingFiles} onClick={processPendingFiles}>
                    {processingFiles ? 'Processando…' : 'Processar arquivos'}
                  </button>
                </div>
              )}
              <div className="mapping-card">
                <div className="section-heading">
                  <div><span>2</span><div><h2>Configure as colunas</h2><p>Indique onde estão os dados no arquivo recebido.</p></div></div>
                  <small>Detecção automática inicial</small>
                </div>
                <div className="field-grid">
                  <label>Arquivo de saída
                    <select value={outputIndustry} onChange={event => setOutputIndustry(event.target.value)}>
                      <option value="RECKITT">Reckitt — sistema Reppos</option>
                      <option value="LOREAL">L'Oréal</option>
                      <option value="3M">3M</option>
                      <option value="TODAS">Todas — somente conferência</option>
                    </select>
                  </label>
                  <label>Linha do cabeçalho
                    <input type="number" min="1" max={Math.max(1, workbookRows.length)} value={headerRow}
                      onChange={event => setHeaderRow(Math.max(1, Number(event.target.value)))} />
                  </label>
                  <label>Coluna do código EAN
                    <select value={eanColumn} onChange={event => setEanColumn(event.target.value)}>
                      <option value="">Selecione uma coluna</option>
                      {headers.map((header, index) => <option key={index} value={header}>{header}</option>)}
                    </select>
                  </label>
                  <label>Regra da quantidade
                    <select value={quantityMode} onChange={event => setQuantityMode(event.target.value)}>
                      <option value="direct">Usar quantidade direta</option>
                      <option value="multiply">Multiplicar embalagem × quantidade</option>
                    </select>
                  </label>
                  <label>Coluna da quantidade
                    <select value={quantityColumn} onChange={event => setQuantityColumn(event.target.value)}>
                      <option value="">Selecione uma coluna</option>
                      {headers.map((header, index) => <option key={index} value={header}>{header}</option>)}
                    </select>
                  </label>
                  {quantityMode === 'multiply' && (
                    <label>Coluna da embalagem
                      <select value={packageColumn} onChange={event => setPackageColumn(event.target.value)}>
                        <option value="">Selecione uma coluna</option>
                        {headers.map((header, index) => <option key={index} value={header}>{header}</option>)}
                      </select>
                    </label>
                  )}
                </div>
              </div>

              <div className="preview-card">
                <div className="section-heading">
                  <div>
                    <span>3</span>
                    <div>
                      <h2>Prévia do resultado</h2>
                      <p>{orderDetails.length ? 'Confira os itens lidos antes de baixar.' : 'O arquivo Reppos terá somente EAN e Quantidade.'}</p>
                    </div>
                  </div>
                  <b>{orderDetails.length ? detailedPreview.length : converted.length} itens lidos</b>
                </div>
                <div className="table-wrap">
                  {orderDetails.length ? (
                    <table>
                      <thead><tr><th>EAN</th><th>COD.</th><th>Item</th><th>Quantidade</th><th>Valor unitário</th><th>Valor total</th><th>Indústria</th></tr></thead>
                      <tbody>
                        {detailedPreview.slice(0, 12).map((row, index) => (
                          <tr key={index}>
                            <td>{row.EAN}</td>
                            <td>{row.CodigoInterno || '—'}</td>
                            <td>{row.Item}</td>
                            <td>{row.Quantidade}</td>
                            <td>{(row.ValorUnitario || (row.Quantidade ? (row.ValorTotal || 0) / row.Quantidade : 0)).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
                            <td>{(row.ValorTotal || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
                            <td>
                              <span className={row.Industria === 'NAO_IDENTIFICADA' ? 'login-error' : 'valid'}>
                                {row.Industria === 'LOREAL' ? "L'Oréal" : row.Industria === 'NAO_IDENTIFICADA' ? 'Revisar EAN' : row.Industria}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <table>
                      <thead><tr><th>EAN</th><th>Quantidade</th><th>Status</th></tr></thead>
                      <tbody>
                        {converted.slice(0, 8).map((row, index) => (
                          <tr key={index}><td>{row.EAN}</td><td>{row.Quantidade}</td><td><span className="valid">Pronto</span></td></tr>
                        ))}
                        {!converted.length && <tr><td colSpan="3" className="empty">Selecione as colunas para gerar a prévia.</td></tr>}
                      </tbody>
                    </table>
                  )}
                </div>
                {orderDetails.length > 12 && <div className="more-rows">Mais {orderDetails.length - 12} itens foram lidos.</div>}
                {orderDetails.length > 0 && (
                  <div className="total-check">
                    <label>
                      Total líquido impresso no pedido (R$)
                      <input
                        type="number"
                        inputMode="decimal"
                        min="0"
                        step="0.01"
                        value={sourceOrderTotal || ''}
                        placeholder="Ex.: 8541,18"
                        onChange={event => setSourceOrderTotal(Number(event.target.value) || 0)}
                      />
                    </label>
                    <span>O sistema tenta preencher automaticamente. Se o OCR errar algum dígito, este campo garante o total exato do documento.</span>
                  </div>
                )}
                {!orderDetails.length && converted.length > 8 && <div className="more-rows">Mais {converted.length - 8} itens serão incluídos no arquivo.</div>}
              </div>

              <div className="action-bar">
                <div><strong>Dois arquivos disponíveis</strong><span>Reppos separado e pedido completo com valores por indústria.</span></div>
                <div className="top-actions">
                  <button className="ghost-button" onClick={exportCompleteOrder} disabled={!completeItems.length}>
                    <Download size={19} /> Baixar pedido completo
                  </button>
                  <button className="primary-button" onClick={exportFile} disabled={!converted.length}>
                    <Download size={19} /> Baixar Reckitt/Reppos
                  </button>
                </div>
              </div>
            </>
          )}
            </>
          )}
        </section>
      </main>
      {showPasswordSetup && (
        <div className="modal-backdrop">
          <section className="password-modal" role="dialog" aria-modal="true" aria-labelledby="password-title">
            <div className="password-modal-icon"><KeyRound size={24} /></div>
            <h2 id="password-title">Criar ou alterar senha</h2>
            <p>Defina uma senha com pelo menos 8 caracteres. Ela será usada junto com o e-mail <strong>{session.user.email}</strong>.</p>
            <form onSubmit={savePassword}>
              <label>Nova senha
                <input type="password" autoComplete="new-password" minLength="8" required value={newPassword} onChange={event => setNewPassword(event.target.value)} />
              </label>
              <label>Confirmar nova senha
                <input type="password" autoComplete="new-password" minLength="8" required value={confirmPassword} onChange={event => setConfirmPassword(event.target.value)} />
              </label>
              {passwordMessage && <div className={`password-feedback ${passwordMessage.type}`}>{passwordMessage.text}</div>}
              <div className="modal-actions">
                <button type="button" className="ghost-button" onClick={() => { setShowPasswordSetup(false); setNewPassword(''); setConfirmPassword(''); setPasswordMessage(null) }}>Fechar</button>
                <button type="submit" className="primary-button" disabled={savingPassword}>{savingPassword ? 'Salvando…' : 'Salvar senha'}</button>
              </div>
            </form>
          </section>
        </div>
      )}
    </div>
  )
}

createRoot(document.getElementById('root')).render(<React.StrictMode><AuthenticatedApp /></React.StrictMode>)
