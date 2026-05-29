-- 할인율 계산기 — 사용자 업로드 공급가 오버라이드 테이블
-- 인벤토리에 없는 상품에도 공급가를 직접 매칭하기 위해 사용
-- Supabase SQL Editor 에 이 파일 전체를 붙여 넣고 실행하세요.

create table if not exists public.calc_supply_override (
  id            bigserial    primary key,
  product_name  text         not null,
  norm_name     text         not null unique,
  supply_price  integer      not null default 0,
  updated_at    timestamptz  not null default now()
);

create index if not exists calc_supply_override_norm_idx on public.calc_supply_override(norm_name);

-- RLS: anon 키로 읽기·쓰기·삭제 모두 허용 (단일 사용자/팀 환경)
alter table public.calc_supply_override enable row level security;

drop policy if exists "calc_supply_override_all" on public.calc_supply_override;
create policy "calc_supply_override_all"
  on public.calc_supply_override for all
  using (true)
  with check (true);
