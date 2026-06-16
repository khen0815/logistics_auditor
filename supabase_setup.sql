create extension if not exists "pgcrypto";

create table if not exists clients (
    id uuid primary key default gen_random_uuid(),
    client_name text not null,
    created_at timestamp with time zone not null default now()
);

alter table clients add column if not exists industry text;
alter table clients add column if not exists primary_courier text;
alter table clients add column if not exists region text;
alter table clients add column if not exists monthly_order_volume_estimate numeric;
alter table clients add column if not exists monthly_courier_spend_estimate numeric;

create table if not exists audits (
    id uuid primary key default gen_random_uuid(),
    client_id uuid not null references clients(id) on delete cascade,
    audit_month text not null,
    audit_year int not null,
    total_loss_zar numeric not null,
    dead_stock_value_zar numeric not null,
    created_at timestamp with time zone not null default now()
);

alter table audits add column if not exists direct_recovery_zar numeric not null default 0;
alter table audits add column if not exists operational_exposure_zar numeric not null default 0;
alter table audits add column if not exists packaging_leak_zar numeric not null default 0;
alter table audits add column if not exists financial_anomaly_zar numeric not null default 0;
alter table audits add column if not exists avg_confidence_score numeric not null default 0;
alter table audits add column if not exists dispute_ready_count int not null default 0;
alter table audits add column if not exists manual_review_count int not null default 0;
alter table audits add column if not exists leakage_rate numeric not null default 0;
alter table audits add column if not exists courier_provider text;
alter table audits add column if not exists rows_raw int not null default 0;
alter table audits add column if not exists rows_valid int not null default 0;
alter table audits add column if not exists rows_skipped int not null default 0;

create table if not exists audit_findings (
    id uuid primary key default gen_random_uuid(),
    audit_id uuid references audits(id) on delete cascade,
    finding_category text not null,
    evidence_tier text not null,
    confidence_score numeric not null default 0,
    recovery_type text not null,
    waybill_id text,
    order_id text,
    direct_recovery_zar numeric not null default 0,
    operational_exposure_zar numeric not null default 0,
    explanation text,
    recommended_action text,
    created_at timestamp with time zone not null default now()
);

create table if not exists dispute_outcomes (
    id uuid primary key default gen_random_uuid(),
    finding_id uuid references audit_findings(id) on delete cascade,
    status text not null default 'not_submitted',
    credited_amount_zar numeric not null default 0,
    rejected_reason text,
    notes text,
    created_at timestamp with time zone not null default now(),
    updated_at timestamp with time zone not null default now()
);

create table if not exists client_rate_cards (
    id uuid primary key default gen_random_uuid(),
    client_id uuid references clients(id) on delete cascade,
    courier text,
    service_level text,
    origin_zone text,
    destination_zone text,
    min_weight_kg numeric not null default 0,
    max_weight_kg numeric not null default 999999,
    base_rate_zar numeric not null default 0,
    per_kg_rate_zar numeric not null default 0,
    minimum_charge_zar numeric not null default 0,
    fuel_surcharge_pct numeric not null default 0,
    vat_included boolean not null default true,
    rounding_increment_kg numeric not null default 1,
    volumetric_divisor numeric not null default 5000,
    effective_from date,
    effective_to date,
    source_name text,
    created_at timestamp with time zone not null default now()
);
