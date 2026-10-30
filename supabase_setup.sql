-- Chạy toàn bộ nội dung này trong Supabase > SQL Editor > New query > Run
create table if not exists public.projects (
  id uuid primary key,
  name text not null,
  customer text default '',
  payload jsonb not null default '{}'::jsonb,
  updated timestamptz not null default now()
);

create table if not exists public.uploads (
  id uuid primary key,
  pid uuid not null references public.projects(id) on delete cascade,
  name text not null,
  path text not null unique
);

-- Chỉ backend dùng Secret key mới truy cập dữ liệu; không cấp quyền công khai cho anon/authenticated.
alter table public.projects enable row level security;
alter table public.uploads enable row level security;
