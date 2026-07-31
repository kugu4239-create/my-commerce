-- 리오더 계산기 — 4주 매장(오프라인) 판매량 컬럼 추가 (2026-07-31)
-- 판매 속도(예상 일판매)에 매장 판매를 더한 값을 저장/표시하기 위한 컬럼.
-- 멱등: 이미 있으면 무시. 미실행 상태여도 앱은 이 컬럼 없이 저장하도록
-- 폴백하므로 기존 동작이 깨지지 않는다 (매장 판매는 예상 일판매에 반영).

alter table public.reorder_recommendations
  add column if not exists reorder_store_monthly_sales integer default 0;
