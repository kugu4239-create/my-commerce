-- 오프라인 세일율 계산기 — 전역 상품 마스터 파일 (대표상품코드 기반 옵션 확장용)
-- 마지막 업로드 파일 1개만 저장. 모달 열 때 자동 로드됩니다.
-- Supabase SQL Editor 에 이 파일 전체를 붙여 넣고 실행하세요.

create table if not exists public.offline_global_file (
  id          smallint     primary key default 1,
  filename    text         not null,
  content_b64 text         not null,
  uploaded_at timestamptz  not null default now(),
  constraint offline_global_file_single_row check (id = 1)
);

-- RLS: 앱은 anon 키를 사용하므로 읽기/쓰기 허용 (단일 행만 존재하므로 안전)
alter table public.offline_global_file enable row level security;

drop policy if exists "offline_global_file_select" on public.offline_global_file;
drop policy if exists "offline_global_file_insert" on public.offline_global_file;
drop policy if exists "offline_global_file_update" on public.offline_global_file;
drop policy if exists "offline_global_file_delete" on public.offline_global_file;

create policy "offline_global_file_select"
  on public.offline_global_file for select
  using (true);

create policy "offline_global_file_insert"
  on public.offline_global_file for insert
  with check (id = 1);

create policy "offline_global_file_update"
  on public.offline_global_file for update
  using (id = 1)
  with check (id = 1);

create policy "offline_global_file_delete"
  on public.offline_global_file for delete
  using (id = 1);
