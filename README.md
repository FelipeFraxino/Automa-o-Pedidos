# Automatização CBN

Sistema web para converter planilhas de pedidos de diferentes clientes para o padrão:

| EAN | Quantidade |
| --- | --- |
| Código de barras | Quantidade inteira |

## Funcionalidades da primeira versão

- Modelos para Rede Piraquara, Ouro Branco e Adega Brasil;
- Modelo personalizado e criação de modelos adicionais;
- Importação de arquivos XLSX, XLS e CSV;
- Seleção da linha de cabeçalho;
- Detecção inicial de colunas de EAN e quantidade;
- Prévia dos registros convertidos;
- Exportação para Excel com as colunas `EAN` e `Quantidade`;
- Configurações salvas localmente;
- Estrutura preparada para integração com Supabase.

## Executar localmente

```bash
npm install
npm run dev
```

## Supabase

1. Execute o arquivo `supabase/schema.sql` no SQL Editor do Supabase.
2. Copie `.env.example` para `.env.local`.
3. Preencha `VITE_SUPABASE_URL` e `VITE_SUPABASE_ANON_KEY`.
4. Reinicie a aplicação.

A tabela usa políticas para usuários autenticados. A autenticação será ligada na próxima etapa, quando o projeto Supabase estiver disponível.

## Próxima etapa

Enviar exemplos reais das planilhas da Rede Piraquara, Ouro Branco e Adega Brasil, além da planilha final desejada. Com esses arquivos, serão implementadas as regras exatas de identificação e conversão automática de cada modelo.
