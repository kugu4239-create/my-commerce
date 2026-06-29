-- 카페24 회원 데이터베이스 — 휴대폰번호 기준으로 자사몰(카페24) 회원과 가입일을 영구 보관.
-- '채널 퍼널' 페이지 상단의 "카페24 회원 정보 업로드" 업로더가 이 테이블에 업서트한다.
-- phone_norm(숫자만 정규화한 휴대폰) 유니크 → 같은 고객을 다시 올리면 가장 이른 가입일을 유지한다.
-- 주문 데이터(order_headers.orderer_phone)와 phone_norm 으로 매칭해 채널 유입/이동 퍼널을 산출한다.
-- Supabase 대시보드 > SQL Editor 에 이 파일 전체를 붙여 넣고 1회 실행하세요.

create table if not exists public.cafe24_members (
  id          bigserial    primary key,
  phone_norm  text         not null unique,   -- 숫자만 ('010-7168-1564' → '01071681564')
  join_date   text         not null default '', -- 회원 가입일 'YYYY-MM-DD'
  name        text         not null default '',
  grade       text         not null default '',
  updated_at  timestamptz  not null default now()
);

create index if not exists cafe24_members_phone_idx on public.cafe24_members(phone_norm);

-- RLS: anon 키로 읽기·쓰기·삭제 모두 허용 (단일 사용자/팀 환경)
alter table public.cafe24_members enable row level security;

drop policy if exists "cafe24_members_all" on public.cafe24_members;
create policy "cafe24_members_all"
  on public.cafe24_members for all
  using (true)
  with check (true);

-- ── 주문 헤더에 주문자휴대폰 컬럼 추가 (회원 매칭 키) ──────────────────
-- '데이터 입력 > 주문 배송 데이터' 업로더가 주문자휴대폰을 숫자만 정규화해 저장한다.
-- 기존 행은 빈 문자열 → 주문 CSV 를 다시 업로드하면 백필된다.
alter table public.order_headers
  add column if not exists orderer_phone text not null default '';

create index if not exists order_headers_orderer_phone_idx on public.order_headers(orderer_phone);
