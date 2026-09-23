# Automatização CBN

Sistema para converter pedidos comerciais em Excel, CSV, PDF ou imagem, separar os produtos por indústria e gerar arquivos prontos para digitação.

## Indústrias

- Reckitt
- L'Oréal
- 3M

## Modelos mapeados

- Rede Piraquara
- Ouro Branco
- WG Adega Brasil
- Dalpar
- Orçamento Flex CBN
- Modelo personalizado editável

As regras observadas nos arquivos reais estão em [docs/MAPEAMENTO-DOS-ARQUIVOS.md](docs/MAPEAMENTO-DOS-ARQUIVOS.md).

## O que já funciona

- Interface web responsiva;
- Importação de XLSX, XLS e CSV;
- Criação e edição de modelos;
- Seleção da linha de cabeçalho e colunas;
- Quantidade direta;
- Multiplicação de embalagem por quantidade para o Ouro Branco;
- Prévia do resultado;
- Exportação com EAN e quantidade;
- Configurações locais.

## Em desenvolvimento

- Leitura de PDF digital;
- OCR de imagens e PDFs escaneados;
- Importação privada da base “códigos que trabalho”;
- Separação automática de Reckitt, L'Oréal e 3M;
- Tela de revisão para leituras incertas;
- Comparação entre orçamento e pedido;
- Geração dos três arquivos separados;
- Autenticação e armazenamento no Supabase.

## Executar localmente

```bash
npm install
npm run dev
```

## Configurar o Supabase

1. Abra o SQL Editor do projeto Supabase.
2. Execute [supabase/schema.sql](supabase/schema.sql).
3. Copie `.env.example` para `.env.local`.
4. Preencha `VITE_SUPABASE_URL` e `VITE_SUPABASE_ANON_KEY`.
5. Reinicie a aplicação.

Não envie a base de produtos, pedidos ou PDFs de clientes ao repositório enquanto ele estiver público. Esses dados serão armazenados de forma privada no Supabase.


## Execução assistida Piraquara no Work

Use `Execução pedidos Piraquara: [instrução específica]` no ChatGPT Work com este projeto aberto. A instrução após os dois-pontos define o recorte e a ordem daquela execução. O painel também registra a solicitação e copia esse comando, preservando a observação quando ele é copiado novamente. A solicitação no painel, por si só, não processa pedidos.

Procedimento completo: [Execução pedidos Piraquara](docs/EXECUCAO-PEDIDOS-PIRAQUARA.md). O Work segue [AGENTS.md](AGENTS.md) e para antes de finalizar qualquer compra.

## Documentação

- [Mapeamento dos arquivos](docs/MAPEAMENTO-DOS-ARQUIVOS.md)
- [Prompt mestre](docs/PROMPT-MESTRE.md)
- [Execução pedidos Piraquara](docs/EXECUCAO-PEDIDOS-PIRAQUARA.md)
- [Documento do projeto](Automatiza%C3%A7%C3%A3o%20CBN.md)
