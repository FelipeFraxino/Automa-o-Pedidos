import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import * as pdfjsLib from 'pdfjs-dist'
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import { createWorker, PSM } from 'tesseract.js'
import {
  FileSpreadsheet, UploadCloud, Download, Settings2, Plus, Trash2,
  CheckCircle2, AlertCircle, ChevronRight, Database, RotateCcw, LogOut, Mail, KeyRound
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

const normalize = value => String(value ?? '').trim().toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')

const inferIndustryFromItem = itemName => {
  const item = normalize(itemName)
  if (/scotch|ponjita|nexcare|post.?it|command|3m\b/.test(item)) return '3M'
  if (/elseve|loreal|l'oreal|niely|garnier|maybelline/.test(item)) return 'LOREAL'
  if (/veja|vanish|finish|harpic|sbp\b|lysoform|destac|repelex/.test(item)) return 'RECKITT'
  return 'NAO_IDENTIFICADA'
}

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
  const text = String(value || '').trim()
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

  if (!modelId) {
    throw new Error(
      normalizedText.trim()
        ? 'Este modelo de PDF ainda não foi reconhecido automaticamente.'
        : 'Este PDF é uma imagem escaneada e precisa do módulo de OCR, que será a próxima etapa.'
    )
  }

  const results = []
  const seen = new Map()
  let stopped = false

  for (const lines of pages) {
    const allItems = lines.flatMap(line => line.items)
    const brandHeader = allItems.find(item => normalize(item.text) === 'marca')
    const packageHeader = allItems.find(item => ['emb.', 'emb'].includes(normalize(item.text)))

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
      if (modelId === 'ouro-branco' && numericAfter.length >= 2) {
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
      const lineTotal = parsePdfMoney(numericAfterItems.at(-1)?.text)
      const existing = seen.get(ean)
      if (existing) {
        existing.Quantidade += quantity
        existing.ValorTotal += lineTotal
        if (!existing.Item && itemName) existing.Item = itemName
      } else {
        seen.set(ean, { EAN: ean, Item: itemName || 'Item sem descrição', Quantidade: quantity, ValorTotal: lineTotal })
      }
    }
    if (stopped) break
  }

  for (const item of seen.values()) results.push(item)
  if (!results.length) throw new Error('Não encontrei itens válidos neste PDF. O arquivo precisa de revisão.')
  const printedTotalMatch = fullText.match(/Vl\s*Total\s*\$?\s*([\d.,]+)/i)
  const documentTotal = printedTotalMatch ? parsePdfMoney(printedTotalMatch[1]) : 0
  return { modelId, rows: results, documentTotal }
}

const cleanOcrEan = value => {
  const digits = cleanEan(value)
  // Em algumas tabelas o OCR lê 789... como 79... e elimina o 8.
  if (digits.length === 12 && digits.startsWith('79')) return `78${digits.slice(1)}`
  return digits
}

const parseVariableText = text => {
  const normalizedDocument = normalize(text)
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
    const EAN = cleanOcrEan(eanMatch[0])
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

  // A leitura compacta sempre roda para recuperar linhas que o OCR separou pela grade.
  const compactText = text.replace(/\s+/g, ' ')
  const rowPattern = /(\d{12,14})[\s|\[\]_]+(.{3,220}?)[\s|\[\]_]+(\d{1,3}[.,]\d{3})[\s|\[\]_]+(\d{1,4}[.,]\d{3,4})[\s|\[\]_]+(\d{1,3}(?:\.\d{3})*[.,]\d{2})/g
  for (const match of compactText.matchAll(rowPattern)) {
    const EAN = cleanOcrEan(match[1])
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

const prepareImageForOcr = async source => {
  if (!(source instanceof Blob) || !window.createImageBitmap) return source
  try {
    const bitmap = await window.createImageBitmap(source)
    const scale = Math.max(1, Math.min(3, 2200 / bitmap.width))
    const canvas = window.document.createElement('canvas')
    canvas.width = Math.round(bitmap.width * scale)
    canvas.height = Math.round(bitmap.height * scale)
    const context = canvas.getContext('2d', { willReadFrequently: true })
    context.fillStyle = '#ffffff'
    context.fillRect(0, 0, canvas.width, canvas.height)
    context.filter = 'grayscale(1) contrast(1.35)'
    context.drawImage(bitmap, 0, 0, canvas.width, canvas.height)
    bitmap.close()

    // Remove as linhas longas da grade sem apagar os números e textos.
    const image = context.getImageData(0, 0, canvas.width, canvas.height)
    const pixels = image.data
    const rowDark = new Uint32Array(canvas.height)
    const columnDark = new Uint32Array(canvas.width)

    for (let y = 0; y < canvas.height; y += 1) {
      for (let x = 0; x < canvas.width; x += 1) {
        const index = (y * canvas.width + x) * 4
        if (pixels[index] < 180) {
          rowDark[y] += 1
          columnDark[x] += 1
        }
      }
    }

    const eraseRow = y => {
      for (let offset = -2; offset <= 2; offset += 1) {
        const target = y + offset
        if (target < 0 || target >= canvas.height) continue
        for (let x = 0; x < canvas.width; x += 1) {
          const index = (target * canvas.width + x) * 4
          pixels[index] = pixels[index + 1] = pixels[index + 2] = 255
        }
      }
    }
    const eraseColumn = x => {
      for (let offset = -2; offset <= 2; offset += 1) {
        const target = x + offset
        if (target < 0 || target >= canvas.width) continue
        for (let y = 0; y < canvas.height; y += 1) {
          const index = (y * canvas.width + target) * 4
          pixels[index] = pixels[index + 1] = pixels[index + 2] = 255
        }
      }
    }

    for (let y = 0; y < canvas.height; y += 1) {
      if (rowDark[y] / canvas.width > 0.45) eraseRow(y)
    }
    for (let x = 0; x < canvas.width; x += 1) {
      if (columnDark[x] / canvas.height > 0.40) eraseColumn(x)
    }
    context.putImageData(image, 0, 0)
    return canvas
  } catch {
    return source
  }
}

const recognizeImage = async source => {
  const uploadedImage = source instanceof Blob
  const worker = await createWorker('eng')
  try {
    await worker.setParameters({
      tessedit_pageseg_mode: uploadedImage ? PSM.SINGLE_BLOCK : PSM.AUTO,
      preserve_interword_spaces: '1',
    })
    const preparedSource = await prepareImageForOcr(source)
    const { data } = await worker.recognize(preparedSource)
    return data.text || ''
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
    const { error: authError } = await supabase.auth.signInWithPassword({ email, password })
    setSending(false)
    if (authError) {
      const invalid = /invalid login credentials|email not confirmed/i.test(authError.message || '')
      setError(invalid ? 'E-mail ou senha incorretos.' : 'Não foi possível entrar agora. Tente novamente.')
    }
  }

  const requestLink = async () => {
    if (!email) {
      setError('Informe seu e-mail para receber o link de recuperação.')
      return
    }
    setSending(true)
    setError('')
    const { error: authError } = await supabase.auth.signInWithOtp({
      email,
      options: { emailRedirectTo: 'https://felipefraxino.github.io/Automa-o-Pedidos/', shouldCreateUser: true },
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
        <p>AUTOMATIZAÇÃO CBN</p>
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
              <button className="login-link-button" type="button" disabled={sending} onClick={requestLink}>Esqueci a senha ou preciso de um link</button>
            </form>
          </>
        ) : (
          <div className="login-sent">
            <CheckCircle2 size={32} />
            <strong>Confira seu e-mail</strong>
            <span>Enviamos um link de recuperação para {email}.</span>
            <button onClick={() => setSent(false)}>Voltar para entrar com senha</button>
          </div>
        )}
      </section>
    </main>
  )
}

function AuthenticatedApp() {
  const [session, setSession] = useState(undefined)

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session))
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, nextSession) => setSession(nextSession))
    return () => subscription.unsubscribe()
  }, [])

  if (session === undefined) return <div className="auth-loading">Carregando…</div>
  if (!session) return <LoginScreen />
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
  const [sheetName, setSheetName] = useState('')
  const [headerRow, setHeaderRow] = useState(1)
  const [eanColumn, setEanColumn] = useState('')
  const [quantityColumn, setQuantityColumn] = useState('')
  const [packageColumn, setPackageColumn] = useState('')
  const [quantityMode, setQuantityMode] = useState('direct')
  const [outputIndustry, setOutputIndustry] = useState('RECKITT')
  const [catalogIndustries, setCatalogIndustries] = useState({})
  const [message, setMessage] = useState(null)
  const [isDragging, setIsDragging] = useState(false)
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

  const selected = models.find(model => model.id === selectedId) || models[0]
  const headers = workbookRows[Math.max(0, headerRow - 1)]?.map((value, index) => String(value || `Coluna ${index + 1}`).trim()) || []

  const converted = useMemo(() => {
    if (!workbookRows.length || !eanColumn || !quantityColumn) return []
    const eanIndex = headers.indexOf(eanColumn)
    const quantityIndex = headers.indexOf(quantityColumn)
    const packageIndex = headers.indexOf(packageColumn)
    if (eanIndex < 0 || quantityIndex < 0) return []
    return workbookRows
      .slice(headerRow)
      .map(row => {
        const EAN = cleanEan(row[eanIndex])
        const ordered = cleanQuantity(row[quantityIndex])
        const pack = quantityMode === 'multiply' ? cleanQuantity(row[packageIndex]) : 1
        const quantity = ordered !== null && pack !== null ? ordered * pack : null
        return { EAN, Quantidade: quantity, Industria: catalogIndustries[EAN] || 'NAO_IDENTIFICADA' }
      })
      .filter(item =>
        item.EAN &&
        item.Quantidade !== null &&
        item.Quantidade !== 0 &&
        (outputIndustry === 'TODAS' || item.Industria === outputIndustry)
      )
  }, [workbookRows, headerRow, eanColumn, quantityColumn, packageColumn, quantityMode, outputIndustry, catalogIndustries, headers.join('|')])

  const detailedPreview = useMemo(() => {
    const order = { RECKITT: 0, LOREAL: 1, '3M': 2, NAO_IDENTIFICADA: 3 }
    return orderDetails
      .map(item => ({ ...item, Industria: catalogIndustries[item.EAN] || inferIndustryFromItem(item.Item) }))
      .sort((a, b) => (order[a.Industria] ?? 3) - (order[b.Industria] ?? 3))
  }, [orderDetails, catalogIndustries])

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
      const industry = catalogIndustries[item.EAN] || inferIndustryFromItem(item.Item || budget?.Item)
      const normalizedItem = {
        ...item,
        Item: item.Item && item.Item !== 'Item sem descrição' ? item.Item : budget?.Item || 'Item sem descrição',
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
        Industria: catalogIndustries[item.EAN] || inferIndustryFromItem(item.Item),
      }))
      .sort((a, b) => (industryOrder[a.Industria] ?? 3) - (industryOrder[b.Industria] ?? 3))

    return { pedido, excluidos }
  }, [budgetRows, comparisonOrderRows, catalogIndustries])

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
      for (let start = 0; start < 2000; start += 1000) {
        const { data, error } = await supabase
          .from('catalogo_produtos')
          .select('ean,industria')
          .eq('user_id', session.user.id)
          .range(start, start + 999)

        if (error) {
          if (active) setMessage({ type: 'error', text: 'Não foi possível carregar a separação por indústria.' })
          return
        }
        records.push(...(data || []))
        if (!data || data.length < 1000) break
      }

      if (active) {
        const industryMap = {}
        for (const item of records.filter(record => record.ean)) {
          const catalogEan = cleanEan(item.ean)
          industryMap[catalogEan] = item.industria
          // Alguns PDFs completam o UPC de 12 dígitos com um zero à esquerda.
          // O alias serve apenas para classificar; o EAN exportado não é alterado.
          if (catalogEan.length === 12) industryMap[`0${catalogEan}`] = item.industria
        }
        setCatalogIndustries(industryMap)
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
        industria: inferIndustryFromItem(item.Item),
        tipo_registro: 'A_REVISAR',
        validado: false,
        origem: 'Leitura automática — pendente de conferência',
      }))

    if (newItems.length) {
      await supabase.from('catalogo_produtos').upsert(newItems, {
        onConflict: 'user_id,codigo_cbn,ean',
        ignoreDuplicates: true,
      })
      setCatalogIndustries(previous => ({
        ...previous,
        ...Object.fromEntries(newItems.map(item => [item.ean, item.industria])),
      }))
    }
  }

  const readFile = async file => {
    if (!file) return
    const extension = file.name.split('.').pop()?.toLowerCase()
    if (!['xlsx', 'xls', 'csv', 'pdf', 'png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
      setMessage({ type: 'error', text: 'Use Excel, CSV, PDF, PNG, JPG ou WEBP.' })
      return
    }
    try {
      if (['png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
        setMessage({ type: 'success', text: 'Lendo a imagem por OCR. Isso pode levar alguns minutos…' })
        const text = await recognizeImage(file)
        const rows = parseVariableText(text)
        if (!rows.length) throw new Error('A imagem foi lida, mas o formato da tabela ainda não foi reconhecido. O problema está no leitor, não na qualidade da foto.')
        await registerNewCatalogItems(rows)
        const imageModelId = selectedId === 'dalpar' ? 'dalpar' : 'personalizado'
        setSelectedId(imageModelId)
        setSourceOrderTotal(0)
        setOrderDetails(rows)
        setWorkbookRows([['EAN', 'Quantidade'], ...rows.map(row => [row.EAN, row.Quantidade])])
        setFileName(file.name)
        setSheetName(imageModelId === 'dalpar' ? 'Pedido Dalpar lido por OCR — revisar' : 'Imagem lida por OCR — revisar')
        setHeaderRow(1)
        setEanColumn('EAN')
        setQuantityColumn('Quantidade')
        setPackageColumn('')
        setQuantityMode('direct')
        setMessage({ type: 'success', text: `Imagem lida: ${rows.length} itens. Confira a prévia antes de baixar.` })
        return
      }

      if (extension === 'pdf') {
        setMessage({ type: 'success', text: 'Lendo e identificando o PDF…' })
        let modelId
        let rows
        let documentTotal = 0
        try {
          const parsed = await readPdfOrder(file)
          modelId = parsed.modelId
          rows = parsed.rows
          documentTotal = parsed.documentTotal || 0
        } catch (pdfError) {
          setMessage({ type: 'success', text: 'PDF escaneado detectado. Iniciando OCR…' })
          const text = await recognizeScannedPdf(file)
          rows = parseVariableText(text)
          modelId = 'personalizado'
          if (!rows.length) throw pdfError
        }
        await registerNewCatalogItems(rows)
        setSelectedId(modelId)
        setSourceOrderTotal(documentTotal)
        setOrderDetails(rows)
        setWorkbookRows([['EAN', 'Quantidade'], ...rows.map(row => [row.EAN, row.Quantidade])])
        setFileName(file.name)
        setSheetName('Pedido extraído do PDF')
        setHeaderRow(1)
        setEanColumn('EAN')
        setQuantityColumn('Quantidade')
        setPackageColumn('')
        setQuantityMode('direct')
        setMessage({ type: 'success', text: `PDF identificado e carregado: ${rows.length} itens encontrados.` })
        return
      }

      const data = await file.arrayBuffer()
      const workbook = XLSX.read(data, { type: 'array', cellDates: true })
      const firstSheet = workbook.SheetNames[0]
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1, defval: '', raw: false })
      setOrderDetails([])
      setSourceOrderTotal(0)
      setWorkbookRows(rows)
      setFileName(file.name)
      setSheetName(firstSheet)
      setMessage({ type: 'success', text: `Arquivo carregado: ${rows.length} linhas encontradas.` })
    } catch (readError) {
      setMessage({ type: 'error', text: readError.message || 'Não foi possível ler o arquivo. Verifique se ele não está corrompido.' })
    }
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
    XLSX.writeFile(workbook, `${baseName}_${suffix}.xlsx`)
    setMessage({ type: 'success', text: `Planilha convertida com ${converted.length} itens da indústria selecionada.` })
  }

  const exportCompleteOrder = () => {
    if (!completeItems.length) {
      setMessage({ type: 'error', text: 'O pedido completo está disponível após carregar um PDF reconhecido.' })
      return
    }

    const groups = [
      { code: 'RECKITT', label: 'RECKITT' },
      { code: 'LOREAL', label: "L'ORÉAL" },
      { code: '3M', label: '3M' },
      { code: 'NAO_IDENTIFICADA', label: 'ITENS NOVOS / A REVISAR' },
    ]
    const sheetRows = [['EAN', 'Item', 'Quantidade', 'Valor total do item']]
    const sectionRows = []

    for (const group of groups) {
      if (sheetRows.length > 1) sheetRows.push(['', '', '', ''])
      sectionRows.push(sheetRows.length)
      sheetRows.push([group.label, '', '', ''])
      for (const item of completeItems.filter(product => product.Industria === group.code)) {
        sheetRows.push([item.EAN, item.Item, item.Quantidade, item.ValorTotal])
      }
    }

    const sheet = XLSX.utils.aoa_to_sheet(sheetRows)
    sheet['!cols'] = [{ wch: 18 }, { wch: 58 }, { wch: 14 }, { wch: 20 }, { wch: 22 }, { wch: 16 }]
    sheet['!merges'] = sectionRows.map(row => ({ s: { r: row, c: 0 }, e: { r: row, c: 3 } }))
    for (let row = 1; row < sheetRows.length; row += 1) {
      const cell = `D${row + 1}`
      if (typeof sheetRows[row][3] === 'number' && sheet[cell]) sheet[cell].z = 'R$ #,##0.00'
    }

    const totals = {
      RECKITT: completeItems.filter(item => item.Industria === 'RECKITT').reduce((sum, item) => sum + item.ValorTotal, 0),
      LOREAL: completeItems.filter(item => item.Industria === 'LOREAL').reduce((sum, item) => sum + item.ValorTotal, 0),
      '3M': completeItems.filter(item => item.Industria === '3M').reduce((sum, item) => sum + item.ValorTotal, 0),
      A_REVISAR: completeItems.filter(item => item.Industria === 'NAO_IDENTIFICADA').reduce((sum, item) => sum + item.ValorTotal, 0),
    }
    const itemsTotal = completeItems.reduce((sum, item) => sum + item.ValorTotal, 0)
    const grandTotal = sourceOrderTotal || itemsTotal
    const documentAdjustment = sourceOrderTotal ? sourceOrderTotal - itemsTotal : 0
    const summaryRows = [
      ['RESUMO DO PEDIDO', 'Valor'],
      ['Reckitt', totals.RECKITT],
      ["L'Oréal", totals.LOREAL],
      ['3M', totals['3M']],
      ['Itens novos / a revisar', totals.A_REVISAR],
    ]
    if (Math.abs(documentAdjustment) >= 0.005) summaryRows.push(['Ajuste conforme total impresso no documento', documentAdjustment])
    summaryRows.push(['TOTAL DO PEDIDO', grandTotal])
    XLSX.utils.sheet_add_aoa(sheet, summaryRows, { origin: 'E1' })

    for (let row = 2; row <= summaryRows.length; row += 1) {
      const cell = `F${row}`
      if (sheet[cell]) sheet[cell].z = 'R$ #,##0.00'
    }

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, 'Pedido completo')
    const baseName = fileName.replace(/\.[^.]+$/, '') || 'pedido'
    XLSX.writeFile(workbook, `${baseName}_pedido_completo.xlsx`)
    setMessage({ type: 'success', text: `Pedido completo gerado com ${completeItems.length} itens e resumo de valores.` })
  }

  const readComparisonFile = async (file, kind) => {
    if (!file) return
    const extension = file.name.split('.').pop()?.toLowerCase()
    if (!['xlsx', 'xls', 'csv', 'pdf', 'png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
      setMessage({ type: 'error', text: 'Use Excel, CSV, PDF, PNG, JPG ou WEBP.' })
      return
    }

    setComparisonLoading(kind)
    setMessage({ type: 'success', text: `Lendo o ${kind === 'budget' ? 'orçamento' : 'pedido'}…` })
    try {
      let rows
      if (['png', 'jpg', 'jpeg', 'webp'].includes(extension)) {
        rows = parseVariableText(await recognizeImage(file))
      } else if (extension === 'pdf') {
        try {
          rows = (await readPdfOrder(file)).rows
        } catch {
          rows = parseVariableText(await recognizeScannedPdf(file))
        }
      } else {
        rows = await parseSpreadsheetOrder(file)
      }

      if (!rows?.length) throw new Error('Nenhum item foi reconhecido neste arquivo.')
      await registerNewCatalogItems(rows)
      if (kind === 'budget') {
        setBudgetRows(rows)
        setBudgetFileName(file.name)
      } else {
        setComparisonOrderRows(rows)
        setOrderFileName(file.name)
      }
      setMessage({ type: 'success', text: `${kind === 'budget' ? 'Orçamento' : 'Pedido'} carregado com ${rows.length} itens.` })
    } catch (error) {
      setMessage({ type: 'error', text: error.message || 'Não foi possível ler o arquivo.' })
    } finally {
      setComparisonLoading('')
    }
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
    const sheetRows = [['EAN', 'Item', 'Quantidade', 'Valor total do item', 'Indústria']]
    const sectionRows = []

    for (const group of groups) {
      if (sheetRows.length > 1) sheetRows.push(['', '', '', '', ''])
      sectionRows.push(sheetRows.length)
      sheetRows.push([group.label, '', '', '', ''])
      for (const item of pedido.filter(product => product.Industria === group.code)) {
        sheetRows.push([item.EAN, item.Item, item.Quantidade, item.ValorTotal, group.label])
      }
    }

    sheetRows.push(['', '', '', '', ''])
    sectionRows.push(sheetRows.length)
    sheetRows.push(['ITENS EXCLUÍDOS DO PEDIDO', '', '', '', ''])
    sheetRows.push(['EAN', 'Item', 'Quantidade no orçamento', 'Valor no orçamento', 'Indústria'])
    for (const item of excluidos) {
      const industryLabel = item.Industria === 'LOREAL' ? "L'ORÉAL" : item.Industria === 'NAO_IDENTIFICADA' ? 'A REVISAR' : item.Industria
      sheetRows.push([item.EAN, item.Item, item.Quantidade, item.ValorTotal || 0, industryLabel])
    }
    if (!excluidos.length) sheetRows.push(['Nenhum item excluído', '', '', '', ''])

    sheetRows.push(['', '', '', '', ''])
    sectionRows.push(sheetRows.length)
    sheetRows.push(['RESUMO DA PERDA DO PEDIDO', '', '', '', ''])
    sheetRows.push(['Indústria', 'Quantidade retirada', '', 'Valor perdido', ''])
    for (const group of groups) {
      const loss = comparisonLoss.byIndustry[group.code]
      sheetRows.push([group.label, loss.quantidade, '', loss.valor, ''])
    }
    sheetRows.push(['TOTAL DA PERDA', comparisonLoss.quantidade, '', comparisonLoss.valor, ''])

    const sheet = XLSX.utils.aoa_to_sheet(sheetRows)
    sheet['!cols'] = [{ wch: 18 }, { wch: 58 }, { wch: 23 }, { wch: 22 }, { wch: 18 }, { wch: 18 }]
    sheet['!merges'] = sectionRows.map(row => ({ s: { r: row, c: 0 }, e: { r: row, c: 4 } }))
    for (let row = 1; row < sheetRows.length; row += 1) {
      const cell = `D${row + 1}`
      if (typeof sheetRows[row][3] === 'number' && sheet[cell]) sheet[cell].z = 'R$ #,##0.00'
    }

    const totals = {
      RECKITT: pedido.filter(item => item.Industria === 'RECKITT').reduce((sum, item) => sum + item.ValorTotal, 0),
      LOREAL: pedido.filter(item => item.Industria === 'LOREAL').reduce((sum, item) => sum + item.ValorTotal, 0),
      '3M': pedido.filter(item => item.Industria === '3M').reduce((sum, item) => sum + item.ValorTotal, 0),
      A_REVISAR: pedido.filter(item => item.Industria === 'NAO_IDENTIFICADA').reduce((sum, item) => sum + item.ValorTotal, 0),
    }
    const grandTotal = pedido.reduce((sum, item) => sum + item.ValorTotal, 0)
    XLSX.utils.sheet_add_aoa(sheet, [
      ['RESUMO DO PEDIDO', 'Valor'],
      ['Reckitt', totals.RECKITT],
      ["L'Oréal", totals.LOREAL],
      ['3M', totals['3M']],
      ['Itens novos / a revisar', totals.A_REVISAR],
      ['TOTAL DO PEDIDO', grandTotal],
    ], { origin: 'F1' })
    for (const cell of ['G2', 'G3', 'G4', 'G5', 'G6']) {
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

  const resetFile = () => {
    setWorkbookRows([])
    setOrderDetails([])
    setSourceOrderTotal(0)
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
          <div><strong>Automatização</strong><span>CBN Distribuidora</span></div>
        </div>

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
            <button className="ghost-button" onClick={saveCurrentModel}><Settings2 size={17} /> Salvar configuração</button>
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

          {selectedId === 'confronto' ? (
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
                  <p>{budgetFileName || 'Arquivo que foi enviado ao cliente'}</p>
                  <input ref={budgetFileRef} type="file" accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp" hidden onChange={event => readComparisonFile(event.target.files[0], 'budget')} />
                  <button className="ghost-button" onClick={() => budgetFileRef.current?.click()} disabled={comparisonLoading === 'budget'}>
                    <UploadCloud size={17} /> {comparisonLoading === 'budget' ? 'Lendo…' : budgetRows.length ? 'Trocar orçamento' : 'Selecionar orçamento'}
                  </button>
                  {budgetRows.length > 0 && <button className="comparison-remove" onClick={() => resetComparison('budget')}>Remover</button>}
                  {budgetRows.length > 0 && <strong>{budgetRows.length} itens lidos</strong>}
                </div>
                <div className={`comparison-upload ${comparisonOrderRows.length ? 'loaded' : ''}`}>
                  <span className="comparison-step">2</span>
                  <FileSpreadsheet size={28} />
                  <h3>Pedido</h3>
                  <p>{orderFileName || 'Arquivo devolvido pelo cliente'}</p>
                  <input ref={orderFileRef} type="file" accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp" hidden onChange={event => readComparisonFile(event.target.files[0], 'order')} />
                  <button className="ghost-button" onClick={() => orderFileRef.current?.click()} disabled={comparisonLoading === 'order'}>
                    <UploadCloud size={17} /> {comparisonLoading === 'order' ? 'Lendo…' : comparisonOrderRows.length ? 'Trocar pedido' : 'Selecionar pedido'}
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
              className={`dropzone ${isDragging ? 'dragging' : ''}`}
              onDragOver={event => { event.preventDefault(); setIsDragging(true) }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={event => { event.preventDefault(); setIsDragging(false); readFile(event.dataTransfer.files[0]) }}
              onClick={() => fileRef.current?.click()}
            >
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.webp" onChange={event => readFile(event.target.files[0])} hidden />
              <div className="upload-icon"><UploadCloud size={34} /></div>
              <h2>Envie a planilha do cliente</h2>
              <p>Arraste o arquivo para cá ou clique para escolher</p>
              <span>Excel, CSV, PDF, foto ou print</span>
              <button>Selecionar arquivo</button>
            </div>
          ) : (
            <>
              <div className="file-card">
                <div className="file-icon"><FileSpreadsheet size={25} /></div>
                <div><strong>{fileName}</strong><span>Aba: {sheetName} · {workbookRows.length} linhas</span></div>
                <button onClick={resetFile}><RotateCcw size={16} /> Trocar arquivo</button>
              </div>

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
                      <thead><tr><th>EAN</th><th>Item</th><th>Quantidade</th><th>Valor total</th><th>Indústria</th></tr></thead>
                      <tbody>
                        {detailedPreview.slice(0, 12).map((row, index) => (
                          <tr key={index}>
                            <td>{row.EAN}</td>
                            <td>{row.Item}</td>
                            <td>{row.Quantidade}</td>
                            <td>{row.ValorTotal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
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
