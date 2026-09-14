# Automatização CBN

## Objetivo

Transformar pedidos recebidos em Excel, CSV, PDF ou imagem em arquivos organizados por indústria, reduzindo a digitação manual e os erros de EAN e quantidade.

O sistema deve atender Reckitt, L'Oréal e 3M e permitir a criação de novos modelos de cliente.

## Fluxo do sistema

1. O usuário escolhe ou deixa o sistema detectar o modelo do arquivo.
2. O sistema extrai EAN, descrição, marca, código CBN, embalagem e quantidade.
3. Os itens são confrontados com a base “códigos que trabalho”.
4. O sistema calcula a quantidade final de acordo com o modelo.
5. Os itens são classificados como Reckitt, L'Oréal, 3M ou não identificados.
6. Uma tela apresenta itens incertos para revisão.
7. O sistema soma EANs repetidos dentro da mesma indústria.
8. São gerados arquivos separados para digitação.

## Base “códigos que trabalho”

A base analisada contém 1.627 registros:

- 678 Reckitt;
- 487 L'Oréal;
- 462 3M.

Ela também contém materiais promocionais, displays, serviços e códigos internos. Por isso, o catálogo terá uma situação para cada registro: produto, material, serviço ou a revisar.

A identificação da indústria usa:

- Reckitt: código CBN `RB...` ou produto `RCK-...`;
- L'Oréal: código CBN `LO...` ou produto `LOR-...`;
- 3M: código CBN `SB...` ou produto `SCB-...` / `3M-...`.

O EAN não precisa começar por 789 ou 790. Produtos válidos com outros prefixos devem ser aceitos.

## Regras dos modelos

### Ouro Branco

Quantidade final = Embalagem × Qtde.

Ignorar a seção Trocas Pendentes.

### Piraquara

Usar a quantidade direta.

Ignorar a seção Trocas Pendentes e todas as linhas seguintes.

### WG Adega Brasil

Usar a quantidade direta.

A embalagem, como `01X25G`, descreve a apresentação e não deve multiplicar a quantidade.

### Dalpar

Entrada recebida como imagem.

Usar código de barras, descrição, marca e quantidade. A leitura por OCR exige revisão antes da exportação.

### Orçamento Flex CBN

Entrada em PDF digital.

Os prefixos RB, LO e SB permitem separar diretamente Reckitt, L'Oréal e 3M. A quantidade é direta.

## Saída Reckitt

A planilha deve conter exatamente:

| EAN | quantidade |
| --- | --- |
| Código de barras como texto | Quantidade inteira |

Não incluir descrição, preço, marca ou qualquer coluna adicional.

## Comparação de arquivos

O sistema terá um modo para comparar o orçamento enviado com o pedido devolvido pelo comprador. A comparação será feita por EAN e mostrará:

- mantido;
- quantidade alterada;
- descartado;
- incluído.

## Segurança

O repositório GitHub está público. Arquivos de clientes, base de produtos e dados comerciais não devem ser gravados nele. O código fica no GitHub; os dados privados ficam no Supabase.

## Situação atual

Concluído:

- interface inicial;
- modelos principais;
- modelo personalizado;
- importação de Excel/CSV;
- configuração das colunas;
- regra de multiplicação do Ouro Branco;
- exportação EAN/quantidade;
- mapeamento dos arquivos reais;
- estrutura do banco Supabase;
- prompt mestre.

Próximas etapas:

1. Conectar o projeto Supabase à aplicação.
2. Importar a base de produtos de forma privada.
3. Implementar autenticação.
4. Implementar PDF e OCR.
5. Criar a tela de revisão.
6. Implementar separação e downloads por indústria.
7. Implementar comparação de orçamento e pedido.
8. Publicar a aplicação.
