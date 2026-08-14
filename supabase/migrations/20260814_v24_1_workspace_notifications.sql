-- ================================================================
-- SISTEMA CAV V24.1
-- Google Workspace Directory + Notification Outbox + Weekly Digest
-- Ejecutar UNA VEZ en Supabase SQL Editor.
-- ================================================================

create extension if not exists pgcrypto;

-- ------------------------------------------------
-- 1. Profesores: vínculo estable con Google Workspace
-- ------------------------------------------------
alter table public.profesores
    add column if not exists workspace_user_id text,
    add column if not exists workspace_primary_email text,
    add column if not exists workspace_active boolean not null default false,
    add column if not exists workspace_org_unit text,
    add column if not exists workspace_match_method text,
    add column if not exists workspace_last_sync timestamptz;

create unique index if not exists profesores_workspace_user_id_uidx
    on public.profesores(workspace_user_id)
    where workspace_user_id is not null;

create index if not exists profesores_workspace_email_idx
    on public.profesores(lower(workspace_primary_email));

-- ------------------------------------------------
-- 2. Copia local del Directory de Workspace
-- ------------------------------------------------
create table if not exists public.workspace_users (
    google_id text primary key,
    primary_email text not null unique,
    full_name text,
    given_name text,
    family_name text,
    org_unit_path text,
    suspended boolean not null default false,
    archived boolean not null default false,
    is_admin boolean not null default false,
    present_in_directory boolean not null default true,
    raw jsonb not null default '{}'::jsonb,
    synced_at timestamptz not null default now()
);

create index if not exists workspace_users_email_idx
    on public.workspace_users(lower(primary_email));

create index if not exists workspace_users_name_idx
    on public.workspace_users(lower(full_name));

create table if not exists public.workspace_groups (
    google_id text primary key,
    email text not null unique,
    name text,
    description text,
    direct_members_count integer,
    present_in_directory boolean not null default true,
    raw jsonb not null default '{}'::jsonb,
    synced_at timestamptz not null default now()
);

create index if not exists workspace_groups_email_idx
    on public.workspace_groups(lower(email));

create table if not exists public.workspace_group_members (
    group_google_id text not null
        references public.workspace_groups(google_id)
        on delete cascade,
    member_google_id text,
    member_email text not null,
    role text,
    type text,
    status text,
    synced_at timestamptz not null default now(),
    primary key (group_google_id, member_email)
);

create index if not exists workspace_group_members_email_idx
    on public.workspace_group_members(lower(member_email));

-- ------------------------------------------------
-- 3. Historial de sincronización
-- ------------------------------------------------
create table if not exists public.workspace_sync_log (
    id uuid primary key default gen_random_uuid(),
    source text not null default 'scheduled',
    status text not null default 'running'
        check (status in ('running','success','error')),
    users_count integer not null default 0,
    groups_count integer not null default 0,
    members_count integer not null default 0,
    linked_professors_count integer not null default 0,
    error text,
    started_at timestamptz not null default now(),
    finished_at timestamptz
);

create index if not exists workspace_sync_log_started_idx
    on public.workspace_sync_log(started_at desc);

-- ------------------------------------------------
-- 4. Cola de notificaciones
-- ------------------------------------------------
create table if not exists public.notification_outbox (
    id uuid primary key default gen_random_uuid(),
    type text not null,
    professor_id bigint references public.profesores(id) on delete set null,
    reservation_id bigint,
    recipient_email text not null,
    subject text not null,
    html_body text not null,
    metadata jsonb not null default '{}'::jsonb,
    dedupe_key text,
    status text not null default 'pending'
        check (status in ('pending','sending','sent','error','cancelled')),
    attempts integer not null default 0,
    available_at timestamptz not null default now(),
    error text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    sent_at timestamptz
);

create unique index if not exists notification_outbox_dedupe_uidx
    on public.notification_outbox(dedupe_key)
    where dedupe_key is not null;

create index if not exists notification_outbox_pending_idx
    on public.notification_outbox(status, available_at)
    where status = 'pending';

create index if not exists notification_outbox_professor_idx
    on public.notification_outbox(professor_id, created_at desc);

-- ------------------------------------------------
-- 5. Control de resumen semanal
-- ------------------------------------------------
create table if not exists public.weekly_digest_log (
    id uuid primary key default gen_random_uuid(),
    professor_id bigint not null
        references public.profesores(id)
        on delete cascade,
    week_start date not null,
    recipient_email text not null,
    reservations_count integer not null default 0,
    status text not null default 'queued'
        check (status in ('queued','sent','error','skipped')),
    error text,
    created_at timestamptz not null default now(),
    sent_at timestamptz,
    unique (professor_id, week_start)
);

create index if not exists weekly_digest_week_idx
    on public.weekly_digest_log(week_start desc);

-- ------------------------------------------------
-- 6. Permisos para backend de confianza
-- ------------------------------------------------
grant select, insert, update, delete on table public.workspace_users to service_role;
grant select, insert, update, delete on table public.workspace_groups to service_role;
grant select, insert, update, delete on table public.workspace_group_members to service_role;
grant select, insert, update, delete on table public.workspace_sync_log to service_role;
grant select, insert, update, delete on table public.notification_outbox to service_role;
grant select, insert, update, delete on table public.weekly_digest_log to service_role;

-- La aplicación actual usa backend con credencial de servidor.
-- Habilitamos RLS para evitar exposición accidental por clientes públicos.
alter table public.workspace_users enable row level security;
alter table public.workspace_groups enable row level security;
alter table public.workspace_group_members enable row level security;
alter table public.workspace_sync_log enable row level security;
alter table public.notification_outbox enable row level security;
alter table public.weekly_digest_log enable row level security;

-- service_role omite RLS. No se crean políticas para anon/authenticated.
-- Así estas tablas quedan privadas para el backend institucional.
