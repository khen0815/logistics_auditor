create extension if not exists "pgcrypto";

create table if not exists clients (
    id uuid primary key default gen_random_uuid(),
    client_name text not null,
    created_at timestamp with time zone not null default now()
);

create table if not exists audits (
    id uuid primary key default gen_random_uuid(),
    client_id uuid not null references clients(id) on delete cascade,
    audit_month text not null,
    audit_year int not null,
    total_loss_zar numeric not null,
    dead_stock_value_zar numeric not null,
    created_at timestamp with time zone not null default now()
);
