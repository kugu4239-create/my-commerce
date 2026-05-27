-- ============================================================
-- 프로모션 플로우 테이블 — promotions / submit_promotions / hidden_promo_log
-- 이미 존재해도 안전하게 재실행 가능 (빠진 컬럼만 보충).
-- Supabase 대시보드 > SQL Editor 에서 실행하세요.
-- ============================================================

-- 1) 등록된 프로모션
create table if not exists public.promotions (
  id            bigint primary key,
  name          text,
  platform      text,
  start_date    text,   -- "YYYY-MM-DDTHH:mm" 문자열 (timestamp 아님)
  end_date      text,
  memo          text,
  content       text,
  files         jsonb default '[]'::jsonb,
  discount_plan jsonb default '{}'::jsonb,
  pinned_products jsonb default '[]'::jsonb   -- [{name, memo}] 임팩트 분석 전/후 비교용 핀셋 상품
);

alter table public.promotions add column if not exists name          text;
alter table public.promotions add column if not exists platform      text;
alter table public.promotions add column if not exists start_date    text;
alter table public.promotions add column if not exists end_date      text;
alter table public.promotions add column if not exists memo          text;
alter table public.promotions add column if not exists content       text;
alter table public.promotions add column if not exists files         jsonb default '[]'::jsonb;
alter table public.promotions add column if not exists discount_plan jsonb default '{}'::jsonb;
alter table public.promotions add column if not exists pinned_products jsonb default '[]'::jsonb;
alter table public.promotions add column if not exists submit_date    text;   -- 프로모션 제출일

-- 2) 제출해야 하는 프로모션
create table if not exists public.submit_promotions (
  id      bigint primary key,
  title   text,
  content text,
  eod     text
);

alter table public.submit_promotions add column if not exists title   text;
alter table public.submit_promotions add column if not exists content text;
alter table public.submit_promotions add column if not exists eod     text;

-- 3) 가려진 종료 프로모션 (가리기 기능) — 기기 간 동기화용
create table if not exists public.hidden_promo_log (
  id        bigint primary key,
  hidden_at text,
  data      jsonb default '{}'::jsonb
);

-- 4) RLS (앱은 anon 키 사용 → 읽기/쓰기 허용)
alter table public.promotions       enable row level security;
alter table public.submit_promotions enable row level security;
alter table public.hidden_promo_log  enable row level security;

drop policy if exists "promotions_all" on public.promotions;
create policy "promotions_all" on public.promotions
  for all using (true) with check (true);

drop policy if exists "submit_promotions_all" on public.submit_promotions;
create policy "submit_promotions_all" on public.submit_promotions
  for all using (true) with check (true);

drop policy if exists "hidden_promo_log_all" on public.hidden_promo_log;
create policy "hidden_promo_log_all" on public.hidden_promo_log
  for all using (true) with check (true);
