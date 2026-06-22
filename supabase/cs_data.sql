-- CS 반품 사유 데이터 — 상품별 반품 사유를 접수일(date)과 함께 영구 보관.
-- 데이터 입력 > CS 데이터 탭의 "CS 반품 데이터 업로드" 가 이 테이블에 적재한다.
-- 대시보드 '반품 Top' 의 주요 사유 매칭 + 기간(날짜) 필터의 소스로 사용된다.
-- 이 테이블이 없으면 업로드가 localStorage 에만 남아 다른 기기/브라우저에서 사라진다.
-- Supabase 대시보드 > SQL Editor 에 이 파일 전체를 붙여 넣고 1회 실행하세요.

create table if not exists public.cs_data (
  -- 앱이 단조 증가 정수로 id 를 직접 생성해 지정한다(Date.now 기반 safe integer).
  -- 정수지만 과거 데이터(소수 id) 호환을 위해 double precision 으로 둔다.
  id             double precision primary key,
  date           text not null default '',        -- 접수일 'YYYY-MM-DD' (반품 Top 기간 필터에 사용)
  product_name   text not null default '',
  return_reason  text not null default '',
  channel        text not null default '자사몰'
);

create index if not exists cs_data_date_idx on public.cs_data(date);

-- RLS: 앱은 anon 키로 접근하므로 읽기·쓰기·삭제 모두 허용 (단일 사용자/팀 환경)
alter table public.cs_data enable row level security;

drop policy if exists "cs_data_all" on public.cs_data;
create policy "cs_data_all"
  on public.cs_data for all
  using (true)
  with check (true);
