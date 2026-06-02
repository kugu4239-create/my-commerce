-- GMV 계산기 목표 설정 — 월 고정금액·목표 이익금·채널 수수료율을 단일 행으로 보관.
-- GMV 계산기가 마운트 시 이 행을 읽어 입력값을 복원하고, 변경 시 업서트한다(가장 최근값 유지).
-- Supabase SQL Editor 에 이 파일 전체를 붙여 넣고 실행하세요.

create table if not exists public.gmv_settings (
  id            smallint     primary key default 1,
  fixed_cost    bigint       not null default 0,   -- 월 고정금액(원)
  target_profit bigint       not null default 0,   -- 목표 이익금(원)
  fee_rates     jsonb        not null default '{"자사몰":3,"29CM":28,"오프라인 스토어":28}'::jsonb,
  updated_at    timestamptz  not null default now(),
  constraint gmv_settings_single_row check (id = 1)
);

-- RLS: 앱은 anon 키를 사용하므로 단일 행 읽기/쓰기 허용
alter table public.gmv_settings enable row level security;

drop policy if exists "gmv_settings_select" on public.gmv_settings;
drop policy if exists "gmv_settings_insert" on public.gmv_settings;
drop policy if exists "gmv_settings_update" on public.gmv_settings;

create policy "gmv_settings_select"
  on public.gmv_settings for select
  using (true);

create policy "gmv_settings_insert"
  on public.gmv_settings for insert
  with check (id = 1);

create policy "gmv_settings_update"
  on public.gmv_settings for update
  using (id = 1)
  with check (id = 1);
