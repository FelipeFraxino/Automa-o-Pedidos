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

Última atualização: 14 de setembro de 2026.

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
- autenticação por link enviado ao e-mail;
- endereço de retorno configurado para o GitHub Pages;
- aplicação publicada no GitHub Pages;
- mensagem mais clara para limite temporário de envio de e-mails;
- prompt mestre e documentação inicial.

Endereço da aplicação:

- https://felipefraxino.github.io/Automa-o-Pedidos/

Situação em teste:

- primeiro acesso por e-mail;
- limite temporário do provedor de e-mail gratuito do Supabase após várias solicitações.

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
- exigir conferência quando a entrada for lida por OCR.

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
