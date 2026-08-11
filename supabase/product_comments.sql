-- 상품 코멘트 — 반품 Top 상품명 클릭으로 작성/조회 (2026-08-11)
-- norm_name(정규화 상품명) 키로 저장해 기간/채널과 무관하게 같은 상품을
-- 따라다닌다. 실시간 publication 등록으로 모든 사용자에게 즉시 반영.
-- Supabase 대시보드 > SQL Editor 에 이 파일 전체를 붙여 넣고 1회 실행하세요. (멱등)

create table if not exists public.product_comments (
  id            uuid primary key default gen_random_uuid(),
  product_name  text not null,
  norm_name     text not null,
  comment       text not null,
  author        text not null default '',
  created_at    timestamptz not null default now()
);

create index if not exists product_comments_norm_idx
  on public.product_comments(norm_name);

-- RLS: 앱은 anon 키로 접근 — 읽기·쓰기·삭제 모두 허용 (팀 내부 도구)
alter table public.product_comments enable row level security;

drop policy if exists "product_comments_all" on public.product_comments;
create policy "product_comments_all"
  on public.product_comments for all
  using (true)
  with check (true);

-- 실시간 반영(모든 사용자 즉시 공유) — supabase_realtime publication 등록
do $$ begin
  alter publication supabase_realtime add table public.product_comments;
exception when duplicate_object then null; end $$;
