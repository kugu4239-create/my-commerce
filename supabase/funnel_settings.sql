-- 채널 유입 분석 설정 — '고객 동선 조회'의 마지막 조회 휴대폰번호를 단일 행으로 보관.
-- 페이지 진입 시 이 행을 읽어 입력값을 복원하고, 조회(blur/Enter) 시 업서트한다.
-- Supabase 대시보드 > SQL Editor 에 이 파일 전체를 붙여 넣고 1회 실행하세요.

create table if not exists public.funnel_settings (
  id          smallint     primary key default 1,
  last_phone  text         not null default '',
  updated_at  timestamptz  not null default now(),
  constraint funnel_settings_single_row check (id = 1)
);

-- RLS: 앱은 anon 키를 사용하므로 단일 행 읽기/쓰기 허용
alter table public.funnel_settings enable row level security;

drop policy if exists "funnel_settings_all" on public.funnel_settings;
create policy "funnel_settings_all"
  on public.funnel_settings for all
  using (true)
  with check (true);
