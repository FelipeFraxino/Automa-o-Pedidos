export const PIRAQUARA_DEFAULT_INSTRUCTION = 'execute todos os pedidos novos da Rede Piraquara em ordem cronológica, começando pelo mais antigo.'

export const emptyPiraquaraForm = () => ({ date: '', quantity: '', order: 'oldest', selection: '', notes: '' })

export function buildPiraquaraInstruction(form) {
  const date = String(form.date || '').trim()
  const quantity = String(form.quantity || '').trim()
  const selection = String(form.selection || '').trim()
  const notes = String(form.notes || '').trim().replace(/^Execução pedidos Piraquara:\s*/i, '')
  const hasStructuredSelection = Boolean(date || quantity || selection)
  if (!hasStructuredSelection) return notes || PIRAQUARA_DEFAULT_INSTRUCTION

  const formattedDate = /^\d{4}-\d{2}-\d{2}$/.test(date)
    ? `${date.slice(8, 10)}/${date.slice(5, 7)}/${date.slice(2, 4)}`
    : date
  const count = Number(quantity)
  const order = form.order === 'newest' ? 'mais recente' : 'mais antigo'
  const target = quantity && Number.isInteger(count) && count > 0
    ? `os ${count} ${count === 1 ? 'pedido' : 'pedidos'}`
    : 'os pedidos'
  let instruction = `execute ${target}${formattedDate ? ` do dia ${formattedDate}` : ''} começando pelo ${order}.`
  if (selection) instruction += ` Seleção específica: ${selection.replace(/[.\s]+$/, '')}.`
  if (notes) instruction += ` ${notes}`
  return instruction
}

export function buildPiraquaraCommand(instruction, requestId) {
  const command = `Execução pedidos Piraquara: ${instruction?.trim() || PIRAQUARA_DEFAULT_INSTRUCTION}`
  return requestId ? `${command} Solicitação no painel: ${requestId}.` : command
}
