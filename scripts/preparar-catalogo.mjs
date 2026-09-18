import fs from 'node:fs'
import path from 'node:path'
import * as XLSX from 'xlsx'

const input = process.argv[2]
if (!input) {
  console.error('Uso: npm run preparar-catalogo -- "caminho/para/códigos que trabalho.xlsx"')
  process.exit(1)
}

const normalize = value => String(value ?? '').trim()
const digits = value => normalize(value).replace(/\D/g, '')

const industryOf = (_code, product) => {
  const produto = normalize(product).toUpperCase()
  // A sigla interna pode permanecer igual depois de uma mudança de indústria.
  // A importação segue a classificação da planilha-mãe, nunca o prefixo do código.
  if (produto.startsWith('RCK-')) return 'RECKITT'
  if (produto.startsWith('LOR-')) return 'LOREAL'
  if (produto.startsWith('SCB-') || produto.startsWith('3M-')) return '3M'
  return 'NAO_IDENTIFICADA'
}

const escapeCsv = value => {
  const text = String(value ?? '')
  return /[;"\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

const workbook = XLSX.readFile(input, { raw: true })
const sheet = workbook.Sheets[workbook.SheetNames[0]]
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true })

const output = rows
  .map(row => {
    const codigo = row.CODIGO ?? row.Código ?? row.codigo
    const produto = row.PRODUTO ?? row.Produto ?? row.produto
    const ean = digits(row['COD BARRAS'] ?? row['CÓD BARRAS'] ?? row.EAN ?? row.ean)
    return {
      codigo_cbn: normalize(codigo),
      produto: normalize(produto),
      ean,
      industria: industryOf(codigo, produto),
      tipo_registro: 'A_REVISAR',
      validado: false,
      origem: 'codigos_que_trabalho',
    }
  })
  .filter(row => row.codigo_cbn || row.produto || row.ean)

const headers = ['codigo_cbn', 'produto', 'ean', 'industria', 'tipo_registro', 'validado', 'origem']
const csv = [
  headers.join(';'),
  ...output.map(row => headers.map(header => escapeCsv(row[header])).join(';')),
].join('\n')

const outputDir = path.resolve('dados-privados')
fs.mkdirSync(outputDir, { recursive: true })
const outputPath = path.join(outputDir, 'catalogo-produtos.csv')
fs.writeFileSync(outputPath, '\uFEFF' + csv, 'utf8')

const totals = output.reduce((acc, row) => {
  acc[row.industria] = (acc[row.industria] || 0) + 1
  return acc
}, {})

console.log(`Catálogo criado em: ${outputPath}`)
console.log(`Registros: ${output.length}`)
console.table(totals)
