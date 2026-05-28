-- =====================================================
--  MyClass — Supabase Schema
--  วิธีใช้: เปิด Supabase → SQL Editor → วางและรัน
-- =====================================================

-- ── Students ─────────────────────────────────────────
create table if not exists students (
  id          uuid primary key default gen_random_uuid(),
  student_no  integer,
  full_name   text not null,
  level       text,
  room        text,
  avatar      text,
  created_at  timestamptz default now()
);

-- ── Classes ──────────────────────────────────────────
create table if not exists classes (
  id           uuid primary key default gen_random_uuid(),
  class_name   text not null,
  subject_name text,
  color        text default '#2563EB',
  created_at   timestamptz default now()
);

-- ── Class ↔ Students ─────────────────────────────────
create table if not exists class_students (
  id         uuid primary key default gen_random_uuid(),
  class_id   uuid references classes(id) on delete cascade,
  student_id uuid references students(id) on delete cascade,
  unique(class_id, student_id)
);

-- ── Teaching Sessions ─────────────────────────────────
create table if not exists teaching_sessions (
  id           uuid primary key default gen_random_uuid(),
  class_id     uuid references classes(id) on delete cascade,
  teach_date   date default current_date,
  topic        text,
  objective    text,
  activities   text,
  highlights   text,
  problems     text,
  improvements text,
  reflection   text,
  status       text default 'draft',
  created_at   timestamptz default now()
);

-- ── Attendance ────────────────────────────────────────
create table if not exists attendance (
  id         uuid primary key default gen_random_uuid(),
  session_id uuid references teaching_sessions(id) on delete cascade,
  student_id uuid references students(id) on delete cascade,
  status     text not null
);

-- ── Behavior Logs ─────────────────────────────────────
create table if not exists behavior_logs (
  id           uuid primary key default gen_random_uuid(),
  session_id   uuid references teaching_sessions(id) on delete cascade,
  student_id   uuid references students(id) on delete cascade,
  behavior     text,
  points       integer default 0,
  icon         text,
  created_at   timestamptz default now()
);

-- ── Assessments ───────────────────────────────────────
create table if not exists assessments (
  id         uuid primary key default gen_random_uuid(),
  session_id uuid references teaching_sessions(id) on delete cascade,
  student_id uuid references students(id) on delete cascade,
  level      text,
  note       text
);

-- ── Session Images ────────────────────────────────────
create table if not exists session_images (
  id         uuid primary key default gen_random_uuid(),
  session_id uuid references teaching_sessions(id) on delete cascade,
  image_url  text not null,
  created_at timestamptz default now()
);

-- ── Profiles ─────────────────────────────────────────
create table if not exists profiles (
  id            uuid primary key default gen_random_uuid(),
  teacher_name  text,
  school_name   text,
  subject_group text,
  avatar_url    text,
  created_at    timestamptz default now()
);

-- ── Disable RLS (MVP — เปิดทีหลังถ้าต้องการ Auth) ───
alter table students          disable row level security;
alter table classes           disable row level security;
alter table class_students    disable row level security;
alter table teaching_sessions disable row level security;
alter table attendance        disable row level security;
alter table behavior_logs     disable row level security;
alter table assessments       disable row level security;
alter table session_images    disable row level security;
alter table profiles          disable row level security;
