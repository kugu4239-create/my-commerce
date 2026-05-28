-- 29CM 할인율 계산기 — 마지막 업로드 파일 1개만 저장 (예시 파일 미리보기용)
-- Supabase SQL Editor 에 이 파일 전체를 붙여 넣고 실행하세요.

create table if not exists public.calc_last_file (
  id          smallint     primary key default 1,
  filename    text         not null,
  content_b64 text         not null,
  uploaded_at timestamptz  not null default now(),
  constraint calc_last_file_single_row check (id = 1)
);

-- RLS: 앱은 anon 키를 사용하므로 읽기/쓰기 허용 (단일 행만 존재하므로 안전)
alter table public.calc_last_file enable row level security;

drop policy if exists "calc_last_file_select" on public.calc_last_file;
drop policy if exists "calc_last_file_insert" on public.calc_last_file;
drop policy if exists "calc_last_file_update" on public.calc_last_file;

create policy "calc_last_file_select"
  on public.calc_last_file for select
  using (true);

create policy "calc_last_file_insert"
  on public.calc_last_file for insert
  with check (id = 1);

create policy "calc_last_file_update"
  on public.calc_last_file for update
  using (id = 1)
  with check (id = 1);
