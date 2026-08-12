-- Định mức review board — Supabase schema
-- Run this once in the Supabase SQL editor (or via `supabase db push`).
-- The FastAPI app (api/db.py) talks to this table through PostgREST with the
-- SERVICE-ROLE key (SUPABASE_SERVICE_ROLE_KEY). Keep RLS disabled or manage it
-- with an allow-all policy for the service role.

create extension if not exists pgcrypto;

create table if not exists public.dinh_muc_orders (
  id uuid primary key default gen_random_uuid(),
  order_id text not null,            -- số đơn hàng (YCSX / JSON)
  customer text,
  product_name text,
  product_code text,
  qty numeric,
  so_mau_in int,
  family text,                       -- 'opp' | 'paper_kp'
  stage text not null default 'thong_tin'
      check (stage in ('thong_tin', 'ke_toan', 'qc', 'dinh_muc', 'chuan_bi')),
  order_json jsonb not null default '{}'::jsonb,    -- canonical order dict (round-trip source)
  fields_json jsonb not null default '[]'::jsonb,   -- computed BOM (per-product fields)
  summary_json jsonb not null default '{}'::jsonb,  -- compact write-back summary
  warnings jsonb not null default '[]'::jsonb,
  reviewer text,                     -- lightweight identity (name), demo-grade
  accepted_at timestamptz,
  accepted_by text,
  reject_reason text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists dinh_muc_orders_stage_idx on public.dinh_muc_orders (stage);
create index if not exists dinh_muc_orders_created_idx on public.dinh_muc_orders (created_at desc);

-- Service role bypasses RLS; without a policy, PostgREST returns 401/empty.
alter table public.dinh_muc_orders enable row level security;
create policy "service_role full access"
  on public.dinh_muc_orders
  for all
  to service_role
  using (true)
  with check (true);
