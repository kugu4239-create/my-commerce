-- 카페24 상품코드 데이터베이스 — 상품명 기준으로 카페24(자사몰) 상품코드를 영구 보관.
-- 데이터 입력 > 인벤토리 > "카페24 상품코드" 업로더가 이 테이블에 업서트한다.
-- norm_name(정규화 상품명) 유니크 → 같은 상품을 다시 올리면 마지막 값으로 덮어쓴다(last-write-wins).
-- SKU Risk Bubble 엑셀 다운로드의 "카페24 상품코드" 매칭 소스로 사용된다.
-- Supabase SQL Editor 에 이 파일 전체를 붙여 넣고 실행하세요.

create table if not exists public.cafe24_product_codes (
  id            bigserial    primary key,
  product_name  text         not null,
  norm_name     text         not null unique,
  product_code  text         not null default '',
  updated_at    timestamptz  not null default now()
);

create index if not exists cafe24_product_codes_norm_idx on public.cafe24_product_codes(norm_name);

-- RLS: anon 키로 읽기·쓰기·삭제 모두 허용 (단일 사용자/팀 환경)
alter table public.cafe24_product_codes enable row level security;

drop policy if exists "cafe24_product_codes_all" on public.cafe24_product_codes;
create policy "cafe24_product_codes_all"
  on public.cafe24_product_codes for all
  using (true)
  with check (true);
