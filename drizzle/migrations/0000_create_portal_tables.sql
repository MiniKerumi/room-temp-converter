create type public.app_role as enum ('admin', 'staff');

create table public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null,
  full_name text not null default '',
  created_at timestamptz not null default now()
);

create table public.user_roles (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete cascade not null,
  role app_role not null,
  unique (user_id, role)
);

create table public.inspection_submissions (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid(),
  form_data jsonb not null,
  created_at timestamptz not null default now()
);

grant select, insert, update, delete on public.profiles to authenticated;
grant all on public.profiles to service_role;
grant select on public.user_roles to authenticated;
grant all on public.user_roles to service_role;
grant select, insert, update, delete on public.inspection_submissions to authenticated;
grant all on public.inspection_submissions to service_role;

alter table public.profiles enable row level security;
alter table public.user_roles enable row level security;
alter table public.inspection_submissions enable row level security;

create or replace function public.has_role(_user_id uuid, _role app_role)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.user_roles
    where user_id = _user_id and role = _role
  )
$$;

create policy "Users read own profile" on public.profiles
  for select to authenticated using (auth.uid() = id);
create policy "Admins read all profiles" on public.profiles
  for select to authenticated using (public.has_role(auth.uid(), 'admin'));

create policy "Users read own roles" on public.user_roles
  for select to authenticated using (auth.uid() = user_id);
create policy "Admins read all roles" on public.user_roles
  for select to authenticated using (public.has_role(auth.uid(), 'admin'));

create policy "Users insert own submissions" on public.inspection_submissions
  for insert to authenticated with check (auth.uid() = user_id);
create policy "Users read own submissions" on public.inspection_submissions
  for select to authenticated using (auth.uid() = user_id);
create policy "Admins read all submissions" on public.inspection_submissions
  for select to authenticated using (public.has_role(auth.uid(), 'admin'));