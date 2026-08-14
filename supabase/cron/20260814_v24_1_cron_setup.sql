-- ================================================================
-- V24.1 - CRON
-- IMPORTANTE:
-- 1) Reemplaza TU_PROJECT_REF.
-- 2) Ejecuta primero la migración principal.
-- 3) Despliega las 3 Edge Functions.
-- 4) Crea los secretos de Edge Functions.
--
-- Este archivo usa Supabase Vault para NO dejar el service role
-- escrito dentro del job de cron.
-- ================================================================

create extension if not exists pg_cron;
create extension if not exists pg_net;
create extension if not exists supabase_vault;

-- ------------------------------------------------
-- A. Guarda URL y service role una sola vez en Vault
-- ------------------------------------------------
-- REEMPLAZA estas dos líneas antes de ejecutar.
select vault.create_secret(
    'https://TU_PROJECT_REF.supabase.co',
    'cav_project_url',
    'URL del proyecto para cron CAV'
);

select vault.create_secret(
    'PEGA_AQUI_TU_SERVICE_ROLE_SOLO_EN_SQL_EDITOR',
    'cav_service_role',
    'Service role para invocar Edge Functions internas'
);

-- ------------------------------------------------
-- B. Procesar correos pendientes cada minuto
-- ------------------------------------------------
select cron.schedule(
    'cav-process-notifications',
    '* * * * *',
    $$
    select net.http_post(
        url := (
            select decrypted_secret
            from vault.decrypted_secrets
            where name = 'cav_project_url'
        ) || '/functions/v1/process-notifications',
        headers := jsonb_build_object(
            'Content-Type', 'application/json',
            'Authorization', 'Bearer ' || (
                select decrypted_secret
                from vault.decrypted_secrets
                where name = 'cav_service_role'
            ),
            'apikey', (
                select decrypted_secret
                from vault.decrypted_secrets
                where name = 'cav_service_role'
            )
        ),
        body := '{}'::jsonb
    );
    $$
);

-- ------------------------------------------------
-- C. Resumen semanal: se revisa cada 15 minutos.
-- La Edge Function SOLO actúa lunes 07:30-07:44 America/Santiago.
-- Esto evita problemas por horario de invierno/verano.
-- ------------------------------------------------
select cron.schedule(
    'cav-weekly-digest',
    '*/15 * * * *',
    $$
    select net.http_post(
        url := (
            select decrypted_secret
            from vault.decrypted_secrets
            where name = 'cav_project_url'
        ) || '/functions/v1/weekly-digest',
        headers := jsonb_build_object(
            'Content-Type', 'application/json',
            'Authorization', 'Bearer ' || (
                select decrypted_secret
                from vault.decrypted_secrets
                where name = 'cav_service_role'
            ),
            'apikey', (
                select decrypted_secret
                from vault.decrypted_secrets
                where name = 'cav_service_role'
            )
        ),
        body := '{}'::jsonb
    );
    $$
);

-- ------------------------------------------------
-- D. Directory Workspace cada 6 horas
-- ------------------------------------------------
select cron.schedule(
    'cav-workspace-sync',
    '17 */6 * * *',
    $$
    select net.http_post(
        url := (
            select decrypted_secret
            from vault.decrypted_secrets
            where name = 'cav_project_url'
        ) || '/functions/v1/workspace-sync',
        headers := jsonb_build_object(
            'Content-Type', 'application/json',
            'Authorization', 'Bearer ' || (
                select decrypted_secret
                from vault.decrypted_secrets
                where name = 'cav_service_role'
            ),
            'apikey', (
                select decrypted_secret
                from vault.decrypted_secrets
                where name = 'cav_service_role'
            )
        ),
        body := '{"source":"cron_6h"}'::jsonb
    );
    $$
);

-- ------------------------------------------------
-- Diagnóstico
-- ------------------------------------------------
select jobid, jobname, schedule, active
from cron.job
where jobname like 'cav-%'
order by jobname;
