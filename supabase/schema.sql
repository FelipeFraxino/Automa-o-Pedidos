create table if not exists public.modelos_planilha (
  id text primary key,
  nome text not null,
  descricao text,
  linha_cabecalho integer not null default 1,
  coluna_ean text,
  coluna_quantidade text,
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now()
);

alter table public.modelos_planilha enable row level security;

create policy "Usuários autenticados podem ler modelos"
on public.modelos_planilha for select
to authenticated
using (true);

create policy "Usuários autenticados podem criar modelos"
on public.modelos_planilha for insert
to authenticated
with check (true);

create policy "Usuários autenticados podem atualizar modelos"
on public.modelos_planilha for update
to authenticated
using (true)
with check (true);

create policy "Usuários autenticados podem excluir modelos"
on public.modelos_planilha for delete
to authenticated
using (true);
