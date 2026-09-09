-- Asset Manager authenticated-access policies
-- Run this once in the Supabase SQL Editor for the Asset Manager project.

alter table public.assets enable row level security;

drop policy if exists "Allow public asset reads" on public.assets;
drop policy if exists "Allow public asset inserts" on public.assets;
drop policy if exists "Allow public asset updates" on public.assets;

drop policy if exists "Authenticated users can read assets" on public.assets;
create policy "Authenticated users can read assets"
on public.assets for select
to authenticated
using (true);

drop policy if exists "Authenticated users can create assets" on public.assets;
create policy "Authenticated users can create assets"
on public.assets for insert
to authenticated
with check (true);

drop policy if exists "Authenticated users can update assets" on public.assets;
create policy "Authenticated users can update assets"
on public.assets for update
to authenticated
using (true)
with check (true);
