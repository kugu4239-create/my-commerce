-- 가격 데이터베이스 / 할인율 계산기 공급가 오버라이드 테이블
-- 인벤토리에 없는 상품에도 정상가(판매가)·공급가를 직접 매칭하기 위해 사용.
-- 인벤토리 업로더의 "가격 DB" 모드와 할인율 계산기 공급가 업로드가 함께 이 테이블을 사용한다.
-- 이익률 계산(베타)의 정상가/원가 소스로도 활용된다 (인벤토리 스냅샷보다 우선).
-- Supabase SQL Editor 에 이 파일 전체를 붙여 넣고 실행하세요.

create table if not exists public.calc_supply_override (
  id            bigserial    primary key,
  product_name  text         not null,
  norm_name     text         not null unique,
  supply_price  integer      not null default 0,
  selling_price integer      not null default 0,
  updated_at    timestamptz  not null default now()
);

-- 기존 테이블에 정상가(판매가) 컬럼 추가 (이미 있으면 무시)
alter table public.calc_supply_override
  add column if not exists selling_price integer not null default 0;

create index if not exists calc_supply_override_norm_idx on public.calc_supply_override(norm_name);

-- RLS: anon 키로 읽기·쓰기·삭제 모두 허용 (단일 사용자/팀 환경)
alter table public.calc_supply_override enable row level security;

drop policy if exists "calc_supply_override_all" on public.calc_supply_override;
create policy "calc_supply_override_all"
  on public.calc_supply_override for all
  using (true)
  with check (true);
