-- 메인 사진 모아보기 — 사이트 목록 공유 저장 (2026-08-11)
-- 등록한 사이트의 현재 메인 화면 스크린샷을 그리드로 모아보는 페이지의
-- 사이트 목록 테이블. 미생성 상태에서도 앱은 localStorage 로 동작하지만
-- 다른 기기/사용자와 목록이 공유되지 않는다.
-- Supabase 대시보드 > SQL Editor 에 이 파일 전체를 붙여 넣고 1회 실행하세요. (멱등)

create table if not exists public.main_photo_sites (
  id          uuid primary key default gen_random_uuid(),
  name        text not null default '',
  url         text not null,
  sub_url     text not null default '',   -- 상품리스트(서브) 링크 — [상품리스트] 화면 전환용
  created_at  timestamptz not null default now()
);

-- 기존에 테이블을 이미 만든 경우 — 서브 링크 컬럼만 추가 (멱등)
alter table public.main_photo_sites
  add column if not exists sub_url text not null default '';

-- 카드 표시 순서 (드래그 정렬, 모든 사용자 공유) — 멱등
alter table public.main_photo_sites
  add column if not exists sort_order integer;

-- RLS: 앱은 anon 키로 접근 — 읽기·쓰기·삭제 모두 허용 (팀 내부 도구)
alter table public.main_photo_sites enable row level security;

drop policy if exists "main_photo_sites_all" on public.main_photo_sites;
create policy "main_photo_sites_all"
  on public.main_photo_sites for all
  using (true)
  with check (true);
