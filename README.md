# Automação de Pedidos

Projeto para padronizar e reorganizar planilhas de pedidos comerciais.

## Objetivo

Receber planilhas de diferentes clientes, identificar automaticamente o formato de cada uma e convertê-las para o padrão final definido pelo usuário.

## Clientes e formatos de origem

### Rede Piraquara

Status: aguardando o envio da planilha de exemplo.

Nesta seção serão registrados:

- Como identificar a planilha;
- Nome e posição das colunas;
- Estrutura dos produtos e quantidades;
- Particularidades do arquivo;
- Regras necessárias para reorganização.

### Ouro Branco

Status: aguardando o envio da planilha de exemplo.

Nesta seção serão registrados:

- Como identificar a planilha;
- Nome e posição das colunas;
- Estrutura dos produtos e quantidades;
- Particularidades do arquivo;
- Regras necessárias para reorganização.

### Adega Brasil

Status: aguardando o envio da planilha de exemplo.

Nesta seção serão registrados:

- Como identificar a planilha;
- Nome e posição das colunas;
- Estrutura dos produtos e quantidades;
- Particularidades do arquivo;
- Regras necessárias para reorganização.

## Padrão de saída

Status: aguardando o envio da planilha-padrão.

Regras já informadas:

1. A primeira coluna deve se chamar exatamente `EAN`.
2. Abaixo de `EAN` devem ficar os códigos de barras dos produtos, em sequência.
3. A segunda coluna deve se chamar exatamente `Quantidade`.
4. Abaixo de `Quantidade` deve ficar a quantidade correspondente a cada EAN.
5. As quantidades devem ser apresentadas sem vírgulas ou casas decimais.

As demais colunas, regras de limpeza, ordenação e validação serão definidas após o recebimento da planilha-padrão.

## Fluxo planejado

1. Receber as planilhas de exemplo da Rede Piraquara, Ouro Branco e Adega Brasil.
2. Analisar e documentar como cada formato pode ser identificado.
3. Receber a planilha-padrão de saída.
4. Registrar todas as regras de transformação e validação.
5. Desenvolver a automação.
6. Testar a conversão usando exemplos reais.
7. Comparar, quando necessário, a sugestão de pedido enviada com a planilha devolvida pelo comprador, destacando itens descartados, incluídos e quantidades alteradas.

## Situação atual

Estrutura inicial do projeto criada. Aguardando os quatro arquivos de referência:

- Planilha da Rede Piraquara;
- Planilha do Ouro Branco;
- Planilha da Adega Brasil;
- Planilha-padrão com o resultado desejado.
