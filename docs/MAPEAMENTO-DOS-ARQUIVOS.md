# Mapeamento dos arquivos e regras

## Base de produtos: códigos que trabalho

A planilha possui as colunas `CODIGO`, `PRODUTO` e `COD BARRAS`, com 1.627 registros.

| Indústria | Registros | Identificação principal |
| --- | ---: | --- |
| Reckitt | 678 | Código `RB...` ou produto `RCK-...` |
| L'Oréal | 487 | Código `LO...` ou produto `LOR-...` |
| 3M | 462 | Código `SB...` ou produto `SCB-...` / `3M-...` |

A base mistura produtos, serviços, materiais promocionais, ponta de gôndola, displays e outros cadastros internos. Existem 94 registros com código de barras inferior a oito dígitos. Por isso:

1. O sistema usa a base para confrontar EAN, código CBN, produto e indústria.
2. Começar por 789 ou 790 aumenta a probabilidade de ser produto, mas não é regra absoluta.
3. Códigos de outros países ou fabricantes também podem ser produtos válidos.
4. Registros curtos ou descrições de material/serviço ficam como `A_REVISAR`.
5. A confirmação manual alimentará o campo `tipo_registro`: PRODUTO, MATERIAL ou SERVICO.

## Orçamento Flex CBN

Entrada: PDF digital com texto.

Campos úteis: sequência, código CBN, descrição, EAN, unidade, quantidade e preço.

Classificação direta:

- `RB` / `RCK-`: Reckitt.
- `LO` / `LOR-`: L'Oréal.
- `SB` / `SCB-`: 3M.

A quantidade é direta. O orçamento poderá ser comparado com o pedido devolvido por EAN para identificar itens mantidos, alterados, descartados e incluídos.

## Pedido Dalpar

Entrada habitual: imagem ou captura de tela.

Campos úteis: código, código de barras, descrição, marca e quantidade.

A indústria pode ser reconhecida pela marca:

- Elseve, Niely, Garnier, Colorama e L'Oréal: L'Oréal.
- Veja, Vanish, Finish, Harpic, SBP e outras marcas Reckitt: Reckitt.
- Scotch-Brite, Scotch, Nexcare, Post-it e outras marcas 3M: 3M.

A quantidade é direta. Como a origem é uma imagem, todos os EANs e quantidades extraídos por OCR devem passar por uma tela de revisão.

## Pedido Ouro Branco

Entrada: PDF digital.

Campos: Código, Cód. Barras, Descrição, Emb., Qtde, valor unitário e total.

Regra obrigatória:

`quantidade_final = Embalagem × Qtde`

Exemplo: Embalagem 3 e Qtde 2 geram 6 unidades.

Ignorar completamente a seção `Trocas Pendentes`.

## Pedido Piraquara

Entrada: PDF digital.

Campos: Código de barras, descrição, marca, quantidade, valor unitário e valor do pedido.

A quantidade é direta. A seção `TROCAS PENDENTES` e todas as linhas seguintes não pertencem ao pedido e devem ser ignoradas.

O arquivo pode misturar Reckitt, L'Oréal e 3M. A separação usa primeiro o catálogo e depois marca/descrição.

## Pedido WG Adega Brasil

Entrada: PDF escaneado ou foto convertida em PDF.

Campos: Cód. Barras, Cód. Fábrica, código, descrição, unidade, embalagem e quantidade.

A coluna `QUANTIDADE` já representa a quantidade final. O texto da coluna `EMBALAGEM`, como `01X25G`, descreve a apresentação e não deve multiplicar a quantidade.

Como o PDF é uma imagem, a extração exige OCR e revisão dos EANs.

## Saídas

### Reckitt

Arquivo Excel com exatamente duas colunas:

| EAN | quantidade |
| --- | --- |
| Código sem pontuação e preservando zeros à esquerda | Número inteiro |

Não incluir descrição, marca, preço, formatação decimal ou outras colunas.

### L'Oréal e 3M

Na primeira fase, gerar arquivos separados por indústria com EAN e quantidade para facilitar a digitação no Flex/CBS. O formato definitivo poderá ser ajustado quando houver uma planilha-padrão de importação dessas indústrias.

## Ordem de classificação

1. EAN encontrado no catálogo validado.
2. Código CBN/prefixo: RB, LO ou SB.
3. Prefixo do produto: RCK, LOR, SCB ou 3M.
4. Marca e palavras da descrição.
5. Caso não seja possível classificar, marcar `NAO_IDENTIFICADA` e exigir revisão.

## Validações

- Preservar EAN como texto e manter zeros à esquerda.
- Remover espaços, pontos e sufixos decimais do EAN.
- Não limitar produto aos prefixos 789 e 790.
- Quantidade final deve ser inteira e maior que zero.
- EAN duplicado na mesma indústria deve ser somado.
- OCR com EAN inexistente no catálogo deve exigir revisão.
- Nunca enviar dados de clientes ou pedidos ao repositório público.
