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

## Confronto de arquivos

O sistema possui um campo específico chamado “Confronto de arquivos”, com dois envios:

1. Orçamento: arquivo original enviado ao cliente.
2. Pedido: arquivo devolvido pelo comprador.

Os dois campos aceitam Excel, CSV, PDF digital, PDF escaneado, foto ou print.

A comparação é feita pelo EAN exato, preservando todos os dígitos. O resultado gera uma planilha com:

- pedido final agrupado na sequência Reckitt, L'Oréal, 3M e itens novos / a revisar;
- colunas EAN, Item, Quantidade, Valor total do item e Indústria;
- resumo do valor de Reckitt, L'Oréal, 3M, itens a revisar e total geral do pedido;
- seção “Itens excluídos do pedido” logo abaixo do pedido final;
- quantidade e valor que cada item excluído possuía no orçamento;
- relatório da perda por indústria, com quantidade retirada e valor perdido;
- quantidade total retirada e valor total perdido do pedido.

Quando o pedido não informa preço, o sistema calcula o valor usando o preço unitário identificado no orçamento. Um item é considerado excluído somente quando aparece no orçamento e não aparece no pedido. EANs novos presentes apenas no pedido permanecem na planilha e são marcados para revisão; eles nunca são descartados automaticamente.

## Segurança

O repositório GitHub está público. Arquivos de clientes, base de produtos e dados comerciais não devem ser gravados nele. O código fica no GitHub; os dados privados ficam no Supabase.

## Situação atual

Última atualização: 15 de setembro de 2026.

Concluído:

- interface inicial do conversor;
- modelos Piraquara, Ouro Branco, WG Adega Brasil, Dalpar e Flex CBN;
- modelo personalizado editável;
- importação de Excel e CSV;
- leitura inicial de PDFs digitais;
- identificação automática dos modelos Piraquara, Ouro Branco e Flex CBN;
- bloqueio automático da seção “Trocas Pendentes” na leitura de PDF;
- configuração das colunas;
- regra de multiplicação do Ouro Branco;
- exportação com as colunas EAN e Quantidade;
- mapeamento inicial dos arquivos reais;
- estrutura do banco no Supabase;
- políticas de segurança RLS;
- conexão da aplicação com o Supabase;
- importação privada dos 1.627 registros da base “códigos que trabalho”;
- classificação inicial entre produtos válidos e códigos para revisão;
- autenticação principal por e-mail e senha;
- criação e alteração de senha dentro da conta autenticada;
- link por e-mail mantido somente como alternativa de recuperação;
- endereço de retorno configurado para o GitHub Pages;
- aplicação publicada no GitHub Pages;
- mensagem mais clara para limite temporário de envio de e-mails;
- prompt mestre e documentação inicial;
- módulo “Confronto de arquivos” com campos separados para Orçamento e Pedido;
- comparação por EAN exato, reaproveitamento do preço do orçamento quando o pedido não informa valor e exportação dos itens excluídos;
- relatório de perda do pedido por indústria, com quantidade retirada e valor perdido, além dos totais gerais.

Endereço da aplicação:

- https://felipefraxino.github.io/Automa-o-Pedidos/

Situação em teste:

- definição da primeira senha pelo usuário dentro da sessão já autenticada;
- novo acesso com e-mail e senha após encerrar a sessão;
- link por e-mail reservado para recuperação de acesso.

Próximas etapas:

1. Testar a leitura dos PDFs digitais de Piraquara, Ouro Branco e Flex CBN.
2. Corrigir as regras específicas identificadas nos testes.
3. Testar e aperfeiçoar o OCR do Modelo variável para Adega Brasil e Dalpar.
4. Criar a tela de revisão dos itens identificados.
5. Implementar separação e downloads por indústria.
6. Implementar comparação entre orçamento enviado e pedido recebido.
7. Automatizar futuramente a digitação no sistema da indústria, mantendo conferência humana antes do faturamento.
8. Implementar o módulo de análise de MSL após o recebimento das planilhas de metas e levantamento.
9. Depois que o conversor e os testes estiverem estáveis, iniciar a automação da digitação do pedido no portal/sistema da CBN.

## Próxima fase: automação da digitação no portal CBN

Depois de concluir o conversor, os testes com clientes reais, a revisão dos itens e a separação por indústria, a próxima fase do projeto será automatizar também a digitação do pedido no portal/sistema da CBN.

Fluxo desejado:

1. O conversor prepara o pedido final, já validado, com códigos/EANs e quantidades corretas.
2. Um agente de navegador ou automação abre o portal da CBN.
3. Faz login com a conta autorizada do usuário.
4. Navega até a tela de novo pedido.
5. Seleciona o cliente correto.
6. Digita os itens e quantidades automaticamente.
7. Sinaliza códigos não encontrados, divergências ou mensagens de erro para revisão.
8. Para antes da confirmação final, faturamento ou envio definitivo.
9. O usuário faz a conferência final e autoriza manualmente a conclusão do pedido.

Prioridade de implementação:

- Primeiro avaliar uso do ChatGPT Work/Cloud Browser para ensinar e repetir o fluxo de navegação no portal.
- Se o fluxo for estável e recorrente, avaliar automação própria com Playwright, Selenium ou ferramenta RPA.
- Antes de automatizar cliques, verificar se o sistema da CBN oferece API ou integração oficial para lançamento de pedidos, pois isso seria mais confiável do que automação visual.
- Manter aprovação humana obrigatória antes de faturar/enviar enquanto a automação estiver em validação.

Objetivo final desta fase:

Pedido recebido → conversor organiza e valida → automação digita no portal CBN → usuário confere → usuário autoriza o faturamento/envio.

## Modelo variável

O Modelo variável reúne as regras de leitura dos modelos existentes e aceita Excel, CSV, PDF digital, PDF escaneado, foto e print.

A saída completa deve:

- copiar os códigos de barras sem alterar nenhum dígito;
- identificar e agrupar Reckitt, L'Oréal e 3M;
- manter as colunas EAN, Item, Quantidade e Valor total do item;
- mostrar os valores por indústria e o valor total do pedido;
- exigir conferência quando a entrada for lida por OCR;
- nunca excluir um item novo ou não identificado do pedido completo;
- colocar itens sem classificação no bloco “Itens novos / a revisar”;
- incluir o valor desses itens no total geral;
- cadastrar novos EANs automaticamente no Supabase como “A revisar”, sem validá-los definitivamente.

## Regra de atualização dos modelos

Toda melhoria geral implementada em qualquer modelo deve ser incluída também no Modelo variável.

Quando um erro for identificado no Modelo variável, a correção deve ser analisada e aplicada aos modelos específicos que utilizem a mesma regra, incluindo Piraquara, Ouro Branco, WG Adega Brasil, Dalpar e Flex CBN.

As correções devem preservar as particularidades de cada cliente. Uma regra específica, como multiplicação de embalagem no Ouro Branco, não deve ser aplicada aos demais modelos sem correspondência.

Após cada alteração:

1. atualizar o código dos modelos aplicáveis;
2. testar o modelo onde o erro apareceu;
3. verificar os demais modelos afetados;
4. atualizar este arquivo de andamento;
5. confirmar a publicação no GitHub Pages.

## Futuro módulo de análise de MSL

O sistema deverá confrontar a planilha de metas de MSL com a planilha de levantamento/venda do cliente.

O módulo deverá:

- identificar quais produtos fazem parte do MSL;
- marcar visualmente esses itens na planilha de levantamento;
- mostrar EANs já vendidos e ainda não vendidos;
- indicar lojas que ainda precisam ser positivadas;
- acompanhar a cobertura do mix por loja;
- apresentar percentual realizado, pendências e evolução da meta;
- permitir filtros por loja, produto, EAN e período.

## Validação dos totais do Flex CBN

O leitor deve aceitar valores com separadores de milhar e decimal usados pelo PDF do Flex, incluindo exemplos como `2,366.3`.

Na planilha de pedido completo:

- todos os valores de itens reconhecidos entram nos subtotais das indústrias;
- o campo “Vl Total” impresso no PDF é usado como controle do total geral;
- qualquer diferença entre a soma das linhas e o “Vl Total” é exibida como “Ajuste conforme total impresso no documento”;
- o total não pode ser reduzido silenciosamente quando uma linha possui formatação numérica diferente;
- no pedido 82226070807284, o valor final de referência é R$ 15.424,11.


## Planilha CBN com pedido em unidades

No Confronto de arquivos, planilhas com os cabeçalhos abaixo devem ser reconhecidas automaticamente:

- EAN: código de barras;
- PRODUTO: descrição;
- PEDIDO EM UND: quantidade pedida;
- CBN: preço unitário.

O valor total de cada item é calculado por PEDIDO EM UND × CBN. O arquivo real CBN (6).xlsx possui 32 itens válidos, 1.012 unidades e valor calculado de R$ 11.002,82.


## Carregamento de múltiplos arquivos

Todos os modelos aceitam vários arquivos no mesmo pedido.

O usuário pode:

- selecionar várias fotos, PDFs ou planilhas de uma vez;
- carregar um arquivo inicialmente e usar “Adicionar mais arquivos”;
- repetir a inclusão quantas vezes for necessário;
- limpar o conjunto completo e começar novamente.

Os arquivos são reunidos antes da prévia e da exportação. Quando o mesmo EAN aparece mais de uma vez, as quantidades e os valores são somados em uma única linha, preservando o código de barras exato.

No Confronto de arquivos, o campo Orçamento e o campo Pedido aceitam múltiplos arquivos de forma independente. Todos os arquivos do Orçamento formam uma base única, e todos os arquivos do Pedido formam o pedido final único. O confronto, os valores por indústria, o total do pedido e o relatório de perda consideram o conjunto completo.


## Etapa explícita de processamento

Nos modelos de conversão, selecionar arquivos não executa mais o serviço imediatamente.

Fluxo:

1. selecionar ou arrastar um ou vários arquivos;
2. conferir na tela os nomes dos arquivos selecionados;
3. clicar em “Processar arquivos”;
4. aguardar a leitura;
5. conferir a prévia e baixar o resultado.

Quando nenhum item é reconhecido, o sistema não pode mostrar uma confirmação verde. Deve apresentar uma mensagem de erro e manter os arquivos selecionados para nova tentativa ou ajuste do leitor. Ao usar “Adicionar mais arquivos”, os novos arquivos também aguardam a confirmação em “Processar arquivos”.


## OCR de fotos giradas da Adega Brasil

O Modelo variável e o modelo WG Adega Brasil devem aceitar fotos do pedido feitas de lado.

A leitura:

- testa automaticamente as orientações da imagem;
- prioriza a posição que reconhece mais EANs e cabeçalhos;
- remove linhas longas da grade antes do OCR;
- executa duas formas de segmentação de texto na melhor orientação;
- reconhece o padrão Cód. Barras, Descrição, Embalagem, Quantidade, P. Final e T. Líquido;
- usa a quantidade direta, sem multiplicar pela embalagem;
- usa o total da linha e o preço final como conferência da quantidade;
- confronta pequenas falhas de caracteres do OCR com os EANs existentes na base;
- junta todas as fotos carregadas antes da prévia e da exportação.


### Correção do fundo cinza nas fotos

Em fotos feitas sobre mesa ou superfície escura, a remoção da grade deve considerar somente pixels muito escuros. O fundo cinza não pode ser interpretado como uma linha da tabela, pois isso apagava a imagem antes do OCR.

Páginas seguintes da Adega Brasil podem não repetir o nome da empresa. O leitor também identifica o modelo pelo padrão de embalagem, como 01X500ML, junto das linhas de produto.


## Controle de totais e colunas completas

Atualização de 15/09/2026:

- a planilha Reckitt/Reppos permanece com exatamente duas colunas: `EAN` e `Quantidade`;
- todas as planilhas completas passam a usar: `EAN`, `COD.`, `Item`, `Quantidade`, `Valor unitário` e `Valor total do item`;
- o Confronto de arquivos também preserva essas informações e acrescenta a coluna de indústria;
- `COD.` representa o código interno enviado pelo comprador; no pedido da WG Adega Brasil, é lido da coluna `COD.` ao lado do código de barras;
- quando o arquivo não possuir código interno, a célula de `COD.` fica vazia, sem inventar códigos;
- o valor unitário deve ser o valor enviado pelo comprador; quando o arquivo trouxer somente quantidade e total, ele pode ser calculado por total dividido pela quantidade;
- em fotos e PDFs escaneados, o campo `Total líquido` impresso no documento é usado como conferência obrigatória do total geral;
- se o OCR não reconhecer todas as linhas, a diferença permanece visível como `Diferença de itens/valores não reconhecidos no OCR`, sem reduzir silenciosamente o pedido;
- para o pedido WG Adega Brasil nº 208501, o total líquido de referência é `R$ 8.541,18`.


### Revisão do OCR da Adega Brasil — 15/09/2026

- o sistema testa todas as rotações da foto antes de escolher a melhor leitura;
- as fotos são ampliadas para aumentar a leitura das 25 linhas do pedido;
- duas passagens de OCR do mesmo item não podem mais duplicar a quantidade nem o valor;
- EANs com um caractere inserido ou omitido pelo OCR são confrontados com o catálogo por distância de edição;
- a prévia possui o campo editável `Total líquido impresso no pedido (R$)`, preenchido automaticamente quando possível;
- esse campo serve como conferência final quando o OCR confundir algum dígito do total;
- para o pedido nº 208501, deve ser informado `8541,18` caso a leitura automática não preencha corretamente.


## Nome automático dos arquivos da Rede Piraquara

Atualização de 17/09/2026:

- o leitor identifica o campo `Destino` no cabeçalho do PDF;
- o leitor identifica o `CNPJ/CPF` da loja dentro de `Dados entrega`, e não o CNPJ da CBN;
- o nome usa os quatro últimos dígitos numéricos do CNPJ da loja;
- o destino é convertido para o número da loja com dois dígitos: `001 → 01`, `002 → 02`, `016 → 16`;
- padrão do arquivo Reckitt/Reppos: `CCCC-LL.xlsx`, em que `CCCC` é o final do CNPJ e `LL` é a loja;
- exemplo real: CNPJ `23.668.138/0014-83` e Destino `016` geram `1483-16.xlsx`;
- a planilha completa do mesmo pedido usa `1483-16-pedido-completo.xlsx`;
- o Reckitt/Reppos continua contendo exatamente as colunas `EAN` e `Quantidade`;
- a leitura continua ignorando tudo após `TROCAS PENDENTES`.
