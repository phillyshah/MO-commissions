-- Package A additions (additive only; 0001_core.sql is not edited).

-- IMAP cursor. Deliberately NOT a config_version row: the cursor moves every minute and is written by
-- the worker alone, so it has no business under dual control or in the audit chain.
create table ingest_state (
  key         text primary key,                    -- "<imap_user>/<folder>"
  last_uid    bigint not null default 0,
  uidvalidity bigint,                              -- when the server resets this, the cursor restarts
  updated_at  timestamptz not null default now()
);

-- VEN-3 dual control for vendor edits. A proposed change is a diff the proposer cannot apply alone
-- while config_current.dual_control_enabled is true; a different person approves and the diff is
-- applied then. With the switch off (Andy acting alone before cutover) the row is written already
-- applied, so the history reads the same either way.
create table vendor_change (
  id          uuid primary key default gen_random_uuid(),
  vendor_id   uuid references vendor(id),          -- null for a proposed NEW vendor
  diff        jsonb not null,                      -- {field: {from, to}} ; for a new vendor {field: {from: null, to}}
  proposed_by uuid not null references ap_user(id),
  proposed_at timestamptz not null default now(),
  approved_by uuid references ap_user(id),
  applied_at  timestamptz,
  rejected_at timestamptz,
  note        text
);
create index on vendor_change (vendor_id);
create index on vendor_change (applied_at) where applied_at is null and rejected_at is null;

-- What the parser container said about the file: hidden-text flags, page count, PDF metadata,
-- extraction attempts. Kept on the invoice because the screen reads it and the AP review page shows it.
alter table invoice add column parse_detail jsonb not null default '{}';

-- The plain-text body, so screening can look for a billing-platform account id and urgency wording
-- without re-opening the raw .eml. Untrusted text (CAP-6): escaped wherever it is shown.
alter table message add column body_text text;
