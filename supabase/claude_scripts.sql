-- Claude 자동화 스크립트 저장 테이블
-- 데이터 입력 > 자동화 스크립트 탭에서 작성한 내용을 어디서든 볼 수 있도록 보관합니다.
-- Supabase 대시보드 > SQL Editor 에서 1회 실행하세요.

create table if not exists public.claude_scripts (
  id         text primary key default 'default',
  name       text not null default 'claude-script.txt',
  content    text not null default '',
  updated_at timestamptz not null default now()
);

-- updated_at 자동 갱신 트리거
create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_claude_scripts_updated_at on public.claude_scripts;
create trigger trg_claude_scripts_updated_at
  before update on public.claude_scripts
  for each row execute function public.set_updated_at();

-- RLS: 앱은 anon 키를 사용하므로 읽기/쓰기를 허용합니다.
alter table public.claude_scripts enable row level security;

drop policy if exists "claude_scripts_all" on public.claude_scripts;
create policy "claude_scripts_all" on public.claude_scripts
  for all
  using (true)
  with check (true);
