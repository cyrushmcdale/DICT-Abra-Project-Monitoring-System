-- DICT Abra Project Monitoring System — Supabase schema + seed data
-- Run this once in the Supabase SQL Editor (Project -> SQL Editor -> New query).

create table if not exists programs (
  id text primary key,
  code text not null,
  name text not null,
  color text not null,
  bg text not null,
  created_at timestamptz default now()
);

create table if not exists events (
  id text primary key,
  program_id text references programs(id) on delete cascade,
  name text not null,
  start_date date,
  end_date date,
  location text,
  personnel text,
  description text,
  status text default 'Target',
  created_at timestamptz default now()
);

create table if not exists employees (
  id text primary key,
  name text not null,
  role text,
  created_at timestamptz default now()
);

create table if not exists notes (
  program_id text primary key references programs(id) on delete cascade,
  text text,
  updated_at timestamptz default now()
);

-- Row Level Security: single shared office login, so any authenticated
-- user may read/write everything (no per-user data isolation needed).
alter table programs enable row level security;
alter table events enable row level security;
alter table employees enable row level security;
alter table notes enable row level security;

create policy "Authenticated users can manage programs" on programs
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "Authenticated users can manage events" on events
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "Authenticated users can manage employees" on employees
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "Authenticated users can manage notes" on notes
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

-- Seed data (same defaults the app previously shipped with in localStorage mode)
insert into programs (id, code, name, color, bg) values
  ('ilcdb', 'ILCDB', 'ICT Literacy and Competency Development Bureau', '#1a56db', '#ebf5ff'),
  ('dtc',   'DTC',   'Digital Transformation Center', '#0e9f6e', '#f3faf7'),
  ('spark', 'SPARK', 'Strengthening the Philippine Workforce through Adaptive and Responsive Digital Knowledge', '#d97706', '#fffbeb'),
  ('cyber', 'CYBER', 'Cybersecurity Awareness', '#e02424', '#fdf2f2'),
  ('wifi',  'WIFI',  'Free Wi-Fi for All', '#7e3af2', '#f5f3ff')
on conflict (id) do nothing;

insert into events (id, program_id, name, start_date, end_date, location, personnel, description, status) values
  ('ilcdb_1', 'ilcdb', 'Barangay IT Literacy Training', '2026-06-03', '2026-06-05', 'Barangay Hall, San Isidro, Abra', 'Cyrush Mcdale, Jasmine A.', 'Conduct basic computing skills development courses and spreadsheet processing orientations for local barangay personnel.', 'Target'),
  ('ilcdb_2', 'ilcdb', 'Sangguniang Bayan Upskilling Caravan', '2026-07-12', '2026-07-12', 'Provincial Capitol, Bangued, Abra', 'Provincial Project Officer', 'Advanced office management software application coaching sessions for local council executives.', 'Accomplished'),
  ('dtc_1', 'dtc', 'Tech Hub Walk-in Advisory Support', '2026-06-08', '2026-06-08', 'Tech Hub, Bangued, Abra', 'Engr. R. Corpuz, J. Corpuz', 'Deploy local hub hardware setups and troubleshoot community technical requests.', 'Target'),
  ('spark_1', 'spark', 'Digital Scholarship Enrollment Drive', '2026-08-15', '2026-08-15', 'DICT Regional Office, Bangued, Abra', 'Cyrush Mcdale, Technical Support staff', 'Onboard qualifying scholars into advanced responsive platform programs.', 'Target'),
  ('cyber_1', 'cyber', 'Provincial Cybersecurity Awareness Seminar', '2026-06-17', '2026-06-19', 'Provincial Capitol, Bangued, Abra', 'Cyrush Mcdale, Cybersecurity Team', 'Strategic threat mitigation training sessions for key municipal heads.', 'Target'),
  ('wifi_1', 'wifi', 'Free Wi-Fi Access Site Testing', '2026-06-25', '2026-06-25', 'Municipal Hall, Tayum, Abra', 'Free Wi-Fi Field Engineers', 'Validate network connectivity metrics and perform hardware routine signal verification.', 'Target')
on conflict (id) do nothing;

insert into employees (id, name, role) values
  ('emp_1', 'Cyrush Mcdale', 'Provincial Project Officer'),
  ('emp_2', 'Jasmine A.', 'Field Training Coordinator'),
  ('emp_3', 'Engr. R. Corpuz', 'Technical Support Engineer'),
  ('emp_4', 'J. Corpuz', 'Technical Support Staff'),
  ('emp_5', 'Cybersecurity Team Lead', 'Cybersecurity Specialist'),
  ('emp_6', 'Free Wi-Fi Field Engineer', 'Network Engineer')
on conflict (id) do nothing;
