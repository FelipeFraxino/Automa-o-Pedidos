import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import {
  FileSpreadsheet, UploadCloud, Download, Settings2, Plus, Trash2,
  CheckCircle2, AlertCircle, ChevronRight, Database, RotateCcw, LogOut, Mail
} from 'lucide-react'
import { supabase } from './supabase'
import './styles.css'

const BASE_MODELS = [
  { id: 'piraquara', name: 'Rede Piraquara', description: 'Quantidade direta; ignora a seção Trocas Pendentes', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'ouro-branco', name: 'Ouro Branco', description: 'Quantidade final calculada por Embalagem × Qtde', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'multiply', fixed: true },
  { id: 'adega-brasil', name: 'WG Adega Brasil', description: 'Quantidade direta; embalagem é apenas a apresentação do produto', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'dalpar', name: 'Dalpar', description: 'Pedido recebido como imagem; separação por marca e produto', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'flex-cbn', name: 'Orçamento Flex CBN', description: 'Prefixos RB, LO e SB identificam a indústria', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: true },
  { id: 'personalizado', name: 'Modelo personalizado', description: 'Configure livremente para outras redes e clientes', headerRow: 1, eanColumn: '', quantityColumn: '', packageColumn: '', quantityMode: 'direct', fixed: false },
]

const normalize = value => String(value ?? '').trim().toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')

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

function LoginScreen() {
  const [email, setEmail] = useState('')
  const [sent, setSent] = useState(false)
  const [error, setError] = useState('')
  const [sending, setSending] = useState(false)

  const requestLink = async event => {
    event.preventDefault()
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
          : 'Não foi possível enviar o acesso agora. Verifique a configuração do e-mail no Supabase e tente novamente.'
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
            <span>Informe seu e-mail. Você receberá um link seguro para entrar, sem senha.</span>
            <form onSubmit={requestLink}>
              <label>E-mail
                <div className="email-field"><Mail size={18} /><input type="email" required value={email} onChange={event => setEmail(event.target.value)} placeholder="seuemail@exemplo.com" /></div>
              </label>
              {error && <div className="login-error">{error}</div>}
              <button disabled={sending}>{sending ? 'Enviando…' : 'Enviar acesso por e-mail'}</button>
            </form>
          </>
        ) : (
          <div className="login-sent">
            <CheckCircle2 size={32} />
            <strong>Confira seu e-mail</strong>
            <span>Enviamos o link de acesso para {email}.</span>
            <button onClick={() => setSent(false)}>Usar outro e-mail</button>
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
  const [sheetName, setSheetName] = useState('')
  const [headerRow, setHeaderRow] = useState(1)
  const [eanColumn, setEanColumn] = useState('')
  const [quantityColumn, setQuantityColumn] = useState('')
  const [packageColumn, setPackageColumn] = useState('')
  const [quantityMode, setQuantityMode] = useState('direct')
  const [message, setMessage] = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const fileRef = useRef(null)

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
        const ordered = cleanQuantity(row[quantityIndex])
        const pack = quantityMode === 'multiply' ? cleanQuantity(row[packageIndex]) : 1
        const quantity = ordered !== null && pack !== null ? ordered * pack : null
        return { EAN: cleanEan(row[eanIndex]), Quantidade: quantity }
      })
      .filter(item => item.EAN && item.Quantidade !== null && item.Quantidade !== 0)
  }, [workbookRows, headerRow, eanColumn, quantityColumn, packageColumn, quantityMode, headers.join('|')])

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

  const readFile = async file => {
    if (!file) return
    const extension = file.name.split('.').pop()?.toLowerCase()
    if (!['xlsx', 'xls', 'csv'].includes(extension)) {
      setMessage({ type: 'error', text: 'Use um arquivo Excel (.xlsx ou .xls) ou CSV.' })
      return
    }
    try {
      const data = await file.arrayBuffer()
      const workbook = XLSX.read(data, { type: 'array', cellDates: true })
      const firstSheet = workbook.SheetNames[0]
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1, defval: '', raw: false })
      setWorkbookRows(rows)
      setFileName(file.name)
      setSheetName(firstSheet)
      setMessage({ type: 'success', text: `Arquivo carregado: ${rows.length} linhas encontradas.` })
    } catch {
      setMessage({ type: 'error', text: 'Não foi possível ler essa planilha. Verifique se o arquivo não está corrompido.' })
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
      description: 'Modelo personalizado editável',
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
    const sheet = XLSX.utils.json_to_sheet(converted, { header: ['EAN', 'Quantidade'] })
    sheet['!cols'] = [{ wch: 18 }, { wch: 14 }]
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, 'Pedido')
    const baseName = fileName.replace(/\.[^.]+$/, '') || 'pedido'
    XLSX.writeFile(workbook, `${baseName}_convertido.xlsx`)
    setMessage({ type: 'success', text: `Planilha convertida com ${converted.length} itens.` })
  }

  const resetFile = () => {
    setWorkbookRows([])
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

          {!workbookRows.length ? (
            <div
              className={`dropzone ${isDragging ? 'dragging' : ''}`}
              onDragOver={event => { event.preventDefault(); setIsDragging(true) }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={event => { event.preventDefault(); setIsDragging(false); readFile(event.dataTransfer.files[0]) }}
              onClick={() => fileRef.current?.click()}
            >
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={event => readFile(event.target.files[0])} hidden />
              <div className="upload-icon"><UploadCloud size={34} /></div>
              <h2>Envie a planilha do cliente</h2>
              <p>Arraste o arquivo para cá ou clique para escolher</p>
              <span>Excel .xlsx, .xls ou CSV</span>
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
                  <div><span>3</span><div><h2>Prévia do resultado</h2><p>O arquivo final terá somente EAN e Quantidade.</p></div></div>
                  <b>{converted.length} itens válidos</b>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead><tr><th>EAN</th><th>Quantidade</th><th>Status</th></tr></thead>
                    <tbody>
                      {converted.slice(0, 8).map((row, index) => (
                        <tr key={index}><td>{row.EAN}</td><td>{row.Quantidade}</td><td><span className="valid">Pronto</span></td></tr>
                      ))}
                      {!converted.length && <tr><td colSpan="3" className="empty">Selecione as colunas para gerar a prévia.</td></tr>}
                    </tbody>
                  </table>
                </div>
                {converted.length > 8 && <div className="more-rows">Mais {converted.length - 8} itens serão incluídos no arquivo.</div>}
              </div>

              <div className="action-bar">
                <div><strong>Arquivo padronizado</strong><span>Colunas: EAN e Quantidade, sem casas decimais.</span></div>
                <button className="primary-button" onClick={exportFile} disabled={!converted.length}>
                  <Download size={19} /> Converter e baixar
                </button>
              </div>
            </>
          )}
        </section>
      </main>
    </div>
  )
}

createRoot(document.getElementById('root')).render(<React.StrictMode><AuthenticatedApp /></React.StrictMode>)
