# Comando permanente: Execução pedidos Piraquara

Definido por Felipe em 23/09/2026. Este é o procedimento operacional do ChatGPT Work para a Rede Piraquara. O painel do Gestão de Pedidos apenas registra uma solicitação e copia o comando; a execução assistida depende do Work e de uma sessão autorizada nos serviços.

## Acionamento e prioridade

```text
Execução pedidos Piraquara: [instrução específica]
```

Tudo após os dois-pontos vale somente para aquela execução e determina data, quantidade, ordem, protocolos inicial/final e observações. Se não houver recorte, o padrão do painel é executar os pedidos novos em ordem cronológica, do mais antigo ao mais recente. Não confundir “primeiros”, “últimos” e ordem de exibição do Gmail: conferir a data/hora e o protocolo de cada PDF antes de selecionar. Uma instrução operacional mais recente de Felipe prevalece sobre esta rotina e sobre exemplos antigos. A proibição de efetivar pedidos só muda mediante autorização clara e específica de Felipe.

Exemplos válidos:

- `Execução pedidos Piraquara: execute os 3 pedidos do dia 16/09/26 começando pelo mais antigo.`
- `Execução pedidos Piraquara: comece pelo protocolo 092689848356 e execute os próximos 2 pedidos.`

Se o comando incluir `Solicitação no painel: <id>`, usar o identificador para acompanhar a solicitação registrada sem alterar o recorte descrito antes dele. O texto da observação deve continuar disponível no botão “Copiar comando” mesmo depois de registrar e atualizar a página.

## Procedimento obrigatório por pedido

1. Acessar o Gmail CBN Piraquara conectado e localizar exatamente os PDFs que atendem ao comando. Conferir protocolos, datas e ordem. Não processar pedidos além do recorte.
2. Abrir cada PDF original. Usar o CNPJ do cliente em **Dados entrega**, nunca o CNPJ da CBN no cabeçalho. Identificar `Destino` no cabeçalho como número da loja: `001 → 01`, `002 → 02`, `016 → 16`. Conferir ambos antes de entrar no Reppos. Baixar o PDF original.
3. No **Gestão de Pedidos – CBN Distribuidora**, escolher **Rede Piraquara**, anexar o PDF e acionar **Processar arquivos**. Conferir a prévia. Usar quantidade direta da coluna `Qtde`; números na descrição, como “Leve 3 Pague 2”, não substituem a quantidade pedida. Desconsiderar `TROCAS PENDENTES` e tudo depois dela. Reconciliar os itens e o valor da tabela principal com o PDF.
4. Gerar **pelo próprio conversor** a saída **Reckitt/Reppos**, sem conversão local ou improvisada salvo autorização expressa de Felipe. Aplicar a classificação pela planilha-mãe: EAN exato presente no pedido tem prioridade; sem EAN, consultar o código interno cadastrado. Sigla RB, marca e descrição isoladas não confirmam Reckitt, especialmente após a divisão Vestase. Itens novos/incertos ficam para revisão e não entram automaticamente na saída Reppos.
5. Conferir que a planilha Reppos tem somente `EAN` e `Quantidade`, quantidades inteiras e EANs numéricos sem zeros artificiais à esquerda. O nome usa quatro dígitos finais do CNPJ da loja e o destino em dois dígitos: `CCCC-LL.xlsx`. Não misturar arquivos de lojas diferentes.
6. Abrir o **Reppos no navegador na nuvem**. Seguir a demonstração operacional dos vídeos fornecidos por Felipe para iniciar atendimento, localizar cliente, importar planilha, verificar itens, deixar carrinho preparado, encerrar o atendimento e iniciar o próximo. Selecionar e confirmar o cliente pelo CNPJ capturado do PDF; semelhança de nome não basta.
7. Após importar, comparar EANs e quantidades reconhecidos com a saída do conversor, verificar que os produtos estão **no carrinho do cliente correto**, e só então encerrar a etapa daquele cliente conforme o vídeo. Não misturar atendimentos ou carrinhos. Repetir na ordem escolhida.

Os vídeos enviados anteriormente ao Work são a referência principal para os cliques e transições operacionais do Reppos. Consultar a versão disponível antes de atuar. Caso os vídeos não estejam acessíveis e uma etapa relevante não possa ser verificada na interface atual, parar nesse ponto e solicitar a referência necessária; não presumir que uma navegação inventada corresponda à demonstração. Combinar os vídeos com as regras mais recentes deste projeto e a instrução específica após os dois-pontos.

## Limite e critério de conclusão

**Nunca clicar em “Finalizar compra” nem confirmar, transmitir, faturar ou efetivar pedido.** Deixar os produtos no carrinho para conferência de Felipe. Só uma autorização posterior, clara e específica, pode mudar esse limite.

Um pedido só é “executado” depois que os itens estiverem efetivamente carregados no carrinho correto do Reppos. E-mail localizado, PDF baixado, conversão concluída ou planilha gerada são etapas intermediárias. Não marcar como concluído no painel nem relatar loja executada antes da verificação do carrinho. Se houver pendência, manter o status e o motivo objetivos.

## Comunicação

Executar com mínima comunicação. Não narrar cliques. Progresso curto apenas quando útil, por exemplo: “2 de 3 pedidos concluídos. Processando o último.” Tentar diagnosticar e corrigir erros reversíveis antes de interromper. Interromper somente por ação/autenticação indispensável de Felipe, ambiguidade que possa trocar o pedido/cliente, erro sem solução autônoma ou ação irreversível que exija autorização. Ao depender de Felipe, informar **Problema**, **O que preciso fazer** e **Onde fazer**. Nunca pedir senha ou código em texto no chat; usar o fluxo seguro de autenticação disponível.

## Relatório final

Listar apenas as lojas cujos itens foram verificados no carrinho correto, usando o número obtido de `Destino`. Formato:

```text
Execução concluída
Pedidos executados — lojas: 01, 02, 03
Pendências: nenhuma.
```

Se houver falhas, listar separadamente `Não concluídos — lojas: 04` e `Motivo loja 04: [motivo objetivo]`. Se nenhum carrinho foi carregado, escrever `Pedidos executados — lojas: nenhuma`. Não chamar a execução inteira de concluída quando houver pedidos pendentes.

## Privacidade e persistência

O repositório é público: não gravar nele PDFs, planilhas comerciais, CNPJs de pedidos, credenciais ou vídeos privados. Esta documentação contém apenas o procedimento. Atualizar estas regras e a instrução do Work quando Felipe modificar o fluxo. Uma nova sessão do Work deve abrir o projeto/repositório para ler `AGENTS.md` e este documento; uma mensagem isolada fora do contexto do projeto não garante acesso automático a estes arquivos.
