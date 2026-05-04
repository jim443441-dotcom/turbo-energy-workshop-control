-- Turbo Energy Workshop Control - Supabase SQL schema
-- Run this once in Supabase: SQL Editor > New query > paste all > Run.

create extension if not exists pgcrypto;

create table if not exists public.fleet_machines (
  id text primary key default gen_random_uuid()::text,
  fleet_no text unique not null,
  machine_type text,
  make_model text,
  reg_no text,
  department text,
  location text,
  hours numeric default 0,
  mileage numeric default 0,
  status text default 'Available',
  service_interval_hours numeric default 250,
  last_service_hours numeric default 0,
  next_service_hours numeric default 0,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.breakdowns (
  id text primary key default gen_random_uuid()::text,
  fleet_no text,
  category text,
  fault text,
  reported_by text,
  assigned_to text,
  start_time timestamptz,
  end_time timestamptz,
  downtime_hours numeric default 0,
  status text default 'Open',
  cause text,
  action_taken text,
  delay_reason text,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.repairs (
  id text primary key default gen_random_uuid()::text,
  fleet_no text,
  job_card_no text,
  fault text,
  assigned_to text,
  start_time timestamptz,
  end_time timestamptz,
  hours_worked numeric default 0,
  status text default 'In progress',
  parts_used text,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.services (
  id text primary key default gen_random_uuid()::text,
  fleet_no text,
  service_type text,
  scheduled_hours numeric default 0,
  due_date date,
  completed_date date,
  completed_hours numeric default 0,
  status text default 'Due',
  technician text,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.spares (
  id text primary key default gen_random_uuid()::text,
  part_no text,
  description text,
  section text,
  machine_type text,
  stock_qty numeric default 0,
  min_qty numeric default 0,
  supplier text,
  lead_time_days numeric default 0,
  order_status text default 'In stock',
  requested_by text,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.spares_orders (
  id text primary key default gen_random_uuid()::text,
  part_no text,
  description text,
  qty numeric default 1,
  requested_by text,
  priority text default 'Normal',
  status text default 'Requested',
  machine_fleet_no text,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.personnel (
  id text primary key default gen_random_uuid()::text,
  name text,
  section text,
  role text,
  phone text,
  employment_status text default 'Active',
  leave_start date,
  leave_end date,
  leave_reason text,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.work_logs (
  id text primary key default gen_random_uuid()::text,
  person_name text,
  fleet_no text,
  section text,
  work_date date,
  hours_worked numeric default 0,
  work_type text,
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.tyres (
  id text primary key default gen_random_uuid()::text,
  serial_no text,
  fleet_no text,
  position text,
  brand text,
  size text,
  fitment_date date,
  fitment_hours numeric default 0,
  expected_life_hours numeric default 5000,
  estimated_replacement_hours numeric default 0,
  status text default 'Fitted',
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.batteries (
  id text primary key default gen_random_uuid()::text,
  serial_no text,
  fleet_no text,
  volts text,
  cca text,
  fitment_date date,
  warranty_end date,
  expected_life_months numeric default 18,
  status text default 'Fitted',
  notes text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create table if not exists public.photo_logs (
  id text primary key default gen_random_uuid()::text,
  fleet_no text,
  linked_type text,
  linked_id text,
  caption text,
  file_name text,
  image_data text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- Prototype policies: easy testing from browser with the anon key.
-- For production, replace this with Supabase Auth + restricted RLS policies.
alter table public.fleet_machines enable row level security;
alter table public.breakdowns enable row level security;
alter table public.repairs enable row level security;
alter table public.services enable row level security;
alter table public.spares enable row level security;
alter table public.spares_orders enable row level security;
alter table public.personnel enable row level security;
alter table public.work_logs enable row level security;
alter table public.tyres enable row level security;
alter table public.batteries enable row level security;
alter table public.photo_logs enable row level security;

do $$
declare
  t text;
begin
  foreach t in array array['fleet_machines','breakdowns','repairs','services','spares','spares_orders','personnel','work_logs','tyres','batteries','photo_logs'] loop
    execute format('drop policy if exists "prototype_select_%1$s" on public.%1$I', t);
    execute format('drop policy if exists "prototype_insert_%1$s" on public.%1$I', t);
    execute format('drop policy if exists "prototype_update_%1$s" on public.%1$I', t);
    execute format('drop policy if exists "prototype_delete_%1$s" on public.%1$I', t);
    execute format('create policy "prototype_select_%1$s" on public.%1$I for select to anon, authenticated using (true)', t);
    execute format('create policy "prototype_insert_%1$s" on public.%1$I for insert to anon, authenticated with check (true)', t);
    execute format('create policy "prototype_update_%1$s" on public.%1$I for update to anon, authenticated using (true) with check (true)', t);
    execute format('create policy "prototype_delete_%1$s" on public.%1$I for delete to anon, authenticated using (true)', t);
  end loop;
end $$;

create index if not exists idx_fleet_no on public.fleet_machines (fleet_no);
create index if not exists idx_breakdowns_fleet on public.breakdowns (fleet_no);
create index if not exists idx_repairs_fleet on public.repairs (fleet_no);
create index if not exists idx_services_fleet on public.services (fleet_no);
create index if not exists idx_tyres_serial on public.tyres (serial_no);
create index if not exists idx_batteries_serial on public.batteries (serial_no);
