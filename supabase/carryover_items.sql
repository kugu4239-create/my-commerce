-- ============================================================
-- 시즌 캐리오버 추출기 데이터 저장소
-- 인벤토리 스냅샷과 달리 재고 0 상품을 포함한 전체 상품 목록을 저장.
-- 업로드할 때마다 전체 교체(단일 최신본 유지). 시즌 캐리오버 페이지의
-- 우선 데이터 소스로 사용된다 (없으면 인벤토리 스냅샷 폴백).
-- Supabase 대시보드 > SQL Editor 에서 1회 실행하세요.
-- ============================================================

create table if not exists public.carryover_items (
  id                      bigserial primary key,
  product_code            text default '',
  product_name            text not null,
  option_name             text default '',
  selling_price           integer default 0,
  supply_price            integer default 0,
  current_stock_qty       integer default 0,
  first_inbound_date      date,
  first_inbound_qty       integer default 0,
  cumulative_inbound_qty  integer default 0,
  latest_inbound_date     date,
  latest_inbound_qty      integer default 0,
  last_delivery_date      date,
  cumulative_delivery_qty integer default 0,
  uploaded_at             timestamptz default now()
);

alter table public.carryover_items enable row level security;
drop policy if exists "carryover_items_all" on public.carryover_items;
create policy "carryover_items_all" on public.carryover_items
  for all using (true) with check (true);

notify pgrst, 'reload schema';
