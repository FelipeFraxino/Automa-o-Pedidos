-- Estrutura inicial do projeto Automatização CBN
create extension if not exists pgcrypto;

create table if not exists public.modelos_planilha (
  id text primary key,
  user_id uuid default auth.uid(),
  nome text not null,
  descricao text,
  tipo_arquivo text[] not null default array['xlsx','xls','csv'],
  linha_cabecalho integer not null default 1,
  coluna_ean text,
  coluna_quantidade text,
  coluna_embalagem text,
  regra_quantidade text not null default 'direct'
    check (regra_quantidade in ('direct','multiply')),
  marcador_inicio text,
  marcador_fim text,
  configuracao jsonb not null default '{}'::jsonb,
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now()
);

alter table public.modelos_planilha add column if not exists user_id uuid default auth.uid();
alter table public.modelos_planilha add column if not exists tipo_arquivo text[] not null default array['xlsx','xls','csv'];
alter table public.modelos_planilha add column if not exists coluna_embalagem text;
alter table public.modelos_planilha add column if not exists regra_quantidade text not null default 'direct';
alter table public.modelos_planilha add column if not exists marcador_inicio text;
alter table public.modelos_planilha add column if not exists marcador_fim text;
alter table public.modelos_planilha add column if not exists configuracao jsonb not null default '{}'::jsonb;

create table if not exists public.catalogo_produtos (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid(),
  codigo_cbn text,
  produto text not null,
  ean text not null,
  industria text not null check (industria in ('RECKITT','LOREAL','3M','NAO_IDENTIFICADA')),
  tipo_registro text not null default 'A_REVISAR'
    check (tipo_registro in ('PRODUTO','MATERIAL','SERVICO','A_REVISAR')),
  validado boolean not null default false,
  origem text default 'codigos_que_trabalho',
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now(),
  unique (user_id, codigo_cbn, ean)
);

create index if not exists catalogo_produtos_ean_idx on public.catalogo_produtos (ean);
create index if not exists catalogo_produtos_codigo_idx on public.catalogo_produtos (codigo_cbn);
create index if not exists catalogo_produtos_industria_idx on public.catalogo_produtos (industria);

create table if not exists public.processamentos (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid(),
  modelo_id text references public.modelos_planilha(id) on delete set null,
  nome_arquivo text not null,
  tipo_arquivo text not null,
  status text not null default 'RECEBIDO'
    check (status in ('RECEBIDO','EXTRAINDO','REVISAO','CONCLUIDO','ERRO')),
  total_itens integer not null default 0,
  total_reckitt integer not null default 0,
  total_loreal integer not null default 0,
  total_3m integer not null default 0,
  mensagem_erro text,
  criado_em timestamptz not null default now(),
  concluido_em timestamptz
);

create table if not exists public.itens_processados (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid(),
  processamento_id uuid not null references public.processamentos(id) on delete cascade,
  linha_origem integer,
  ean text,
  descricao text,
  marca text,
  codigo_cbn text,
  embalagem numeric,
  quantidade_pedida numeric,
  quantidade_final integer,
  industria text check (industria in ('RECKITT','LOREAL','3M','NAO_IDENTIFICADA')),
  confianca numeric check (confianca between 0 and 1),
  requer_revisao boolean not null default false,
  dados_origem jsonb not null default '{}'::jsonb,
  criado_em timestamptz not null default now()
);

create table if not exists public.comparacoes_pedido (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid(),
  processamento_orcamento_id uuid references public.processamentos(id) on delete cascade,
  processamento_pedido_id uuid references public.processamentos(id) on delete cascade,
  ean text not null,
  quantidade_orcada integer,
  quantidade_pedida integer,
  resultado text not null check (resultado in ('MANTIDO','ALTERADO','DESCARTADO','INCLUIDO')),
  criado_em timestamptz not null default now()
);

alter table public.modelos_planilha enable row level security;
alter table public.catalogo_produtos enable row level security;
alter table public.processamentos enable row level security;
alter table public.itens_processados enable row level security;
alter table public.comparacoes_pedido enable row level security;

drop policy if exists "modelos_do_usuario" on public.modelos_planilha;
create policy "modelos_do_usuario" on public.modelos_planilha
for all to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid());

drop policy if exists "catalogo_do_usuario" on public.catalogo_produtos;
create policy "catalogo_do_usuario" on public.catalogo_produtos
for all to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid());

drop policy if exists "processamentos_do_usuario" on public.processamentos;
create policy "processamentos_do_usuario" on public.processamentos
for all to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid());

drop policy if exists "itens_do_usuario" on public.itens_processados;
create policy "itens_do_usuario" on public.itens_processados
for all to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid());

drop policy if exists "comparacoes_do_usuario" on public.comparacoes_pedido;
create policy "comparacoes_do_usuario" on public.comparacoes_pedido
for all to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid());
