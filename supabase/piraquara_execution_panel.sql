-- Aplicada no projeto Supabase em 23/09/2026.
-- O painel lê somente resultados operacionais confirmados pelo Work.
alter table public.piraquara_execucoes
  add column if not exists pedidos_solicitados integer,
  add column if not exists loja_atual text,
  add column if not exists lojas_executadas text[] not null default '{}'::text[],
  add column if not exists pendencias jsonb not null default '{}'::jsonb,
  add column if not exists problema_acao text,
  add column if not exists orientacao_acao text;

alter table public.piraquara_execucoes
  add constraint piraquara_pedidos_solicitados_positivo
  check (pedidos_solicitados is null or pedidos_solicitados > 0);

comment on column public.piraquara_execucoes.lojas_executadas is
  'Somente lojas com itens verificados no carrinho correto do Reppos; nunca apenas PDF ou planilha convertida.';
comment on column public.piraquara_execucoes.pendencias is
  'Objeto loja -> motivo objetivo das lojas sem carrinho validado.';
