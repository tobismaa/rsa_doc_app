-- Asset Manager Supabase schema
-- Run this in Supabase SQL Editor.

create table if not exists public.assets (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  tag text not null unique,
  category text not null default 'Computer equipment',
  custodian text,
  location text,
  purchase_cost numeric(14, 2) not null default 0,
  book_value numeric(14, 2) not null default 0,
  purchase_date date,
  status text not null default 'Active'
    check (status in ('Active', 'Service due', 'Review', 'Disposal')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists assets_set_updated_at on public.assets;

create trigger assets_set_updated_at
before update on public.assets
for each row
execute function public.set_updated_at();

alter table public.assets enable row level security;

-- Prototype policies: useful while building without user login.
-- Tighten these when you add authentication and roles.
drop policy if exists "Allow public asset reads" on public.assets;
create policy "Allow public asset reads"
on public.assets for select
to anon
using (true);

drop policy if exists "Allow public asset inserts" on public.assets;
create policy "Allow public asset inserts"
on public.assets for insert
to anon
with check (true);

drop policy if exists "Allow public asset updates" on public.assets;
create policy "Allow public asset updates"
on public.assets for update
to anon
using (true)
with check (true);

insert into public.assets
  (name, tag, category, custodian, location, purchase_cost, book_value, purchase_date, status)
values
  ('MacBook Pro 16', 'IT-2026-0041', 'Computer equipment', 'Product Team', 'Lagos HQ', 3200, 2680, '2026-02-12', 'Active'),
  ('Toyota Hilux', 'FLT-2024-0012', 'Vehicle', 'Operations', 'Abuja Office', 44500, 31900, '2024-08-04', 'Service due'),
  ('ERP License', 'SW-2025-0008', 'Software', 'Finance', 'Cloud', 60000, 48000, '2025-01-10', 'Active'),
  ('Warehouse Racking', 'PLT-2023-0027', 'Plant', 'Logistics', 'Ikeja Warehouse', 28000, 19400, '2023-05-18', 'Review')
on conflict (tag) do nothing;
