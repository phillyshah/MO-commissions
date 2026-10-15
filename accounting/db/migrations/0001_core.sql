-- Maxx Accounting Portal — core schema (Stage 1).
-- Money is bigint cents. Nothing here is deleted by the application (DEL-1); the audit tables refuse
-- UPDATE/DELETE at the trigger level. SAP columns arrive in Stage 2 as additive migrations.

create extension if not exists pgcrypto;

-- ---------------------------------------------------------------------------
-- People. Identity is the MaxxDash Supabase project; this table is the portal's view of a user
-- (portal role, manager, backup), keyed by the Supabase subject.
-- ---------------------------------------------------------------------------
create type portal_role as enum
  ('clerk', 'supervisor', 'controller', 'approver', 'cfo', 'compliance', 'admin', 'auditor');

create table ap_user (
  id             uuid primary key default gen_random_uuid(),
  subject        uuid not null unique,          -- auth.users.id in the MaxxDash project
  email          text not null unique,
  display_name   text not null,
  roles          portal_role[] not null default '{}',
  status         text not null default 'active' check (status in ('active', 'frozen', 'disabled')),
  manager_id     uuid references ap_user(id),
  backup_id      uuid references ap_user(id),
  last_directory_sync timestamptz,
  created_at     timestamptz not null default now()
);

create table portal_session (
  id             uuid primary key default gen_random_uuid(),
  user_id        uuid not null references ap_user(id),
  token_hash     bytea not null unique,            -- sha256 of the cookie value; cookie is never stored
  refresh_token_enc bytea not null,                -- Supabase refresh token, encrypted with the field key
  access_token_exp timestamptz not null,
  device_hash    text not null,
  ip_country     text,
  created_at     timestamptz not null default now(),
  last_seen_at   timestamptz not null default now(),
  last_revalidated_at timestamptz not null default now(),
  absolute_expiry timestamptz not null,
  revoked_at     timestamptz,
  revoked_reason text
);
create index on portal_session (user_id) where revoked_at is null;

create table passkey (
  id             uuid primary key default gen_random_uuid(),
  user_id        uuid not null references ap_user(id),
  credential_id  bytea not null unique,
  public_key     bytea not null,
  sign_count     bigint not null default 0,
  attestation    jsonb,
  transports     text[],
  label          text,
  created_at     timestamptz not null default now(),
  last_used_at   timestamptz,
  status         text not null default 'active' check (status in ('active', 'revoked'))
);

-- One-time codes the Controller hands out out-of-band before a passkey can be enrolled (ID-4).
create table enrollment_code (
  id             uuid primary key default gen_random_uuid(),
  user_id        uuid not null references ap_user(id),
  code_hash      bytea not null,
  issued_by      uuid not null references ap_user(id),
  issued_at      timestamptz not null default now(),
  expires_at     timestamptz not null,
  used_at        timestamptz
);

create table signing_session (
  id             uuid primary key default gen_random_uuid(),
  user_id        uuid not null references ap_user(id),
  passkey_id     uuid not null references passkey(id),
  portal_session_id uuid not null references portal_session(id),
  started_at     timestamptz not null default now(),
  last_activity_at timestamptz not null default now(),
  expires_at     timestamptz not null,
  device_hash    text not null,
  ended_at       timestamptz
);

-- "Freeze my access" tokens (NOT-5): POST-only, hashed, 14 days, can only reduce access.
create table freeze_token (
  id             uuid primary key default gen_random_uuid(),
  user_id        uuid not null references ap_user(id),
  token_hash     bytea not null unique,
  created_at     timestamptz not null default now(),
  expires_at     timestamptz not null,
  used_at        timestamptz
);

-- ---------------------------------------------------------------------------
-- Configuration under dual control (ADM-1). A row is proposed, then approved by a different
-- person, and takes effect the next business day. `config_current` is the view the app reads.
-- ---------------------------------------------------------------------------
create table config_version (
  id             bigserial primary key,
  key            text not null,                    -- step_up_amount_cents, dual_control_enabled, ...
  value          jsonb not null,
  proposed_by    uuid not null references ap_user(id),
  proposed_at    timestamptz not null default now(),
  approved_by    uuid references ap_user(id),
  approved_at    timestamptz,
  effective_from timestamptz,
  note           text
);
create view config_current as
  select distinct on (key) key, value, effective_from
  from config_version
  where approved_at is not null and effective_from <= now()
  order by key, effective_from desc;

-- ---------------------------------------------------------------------------
-- Vendor master (VEN-1, VEN-2). Portal-owned; SAP stays the system of record for bank data.
-- ---------------------------------------------------------------------------
create table vendor (
  id             uuid primary key default gen_random_uuid(),
  sap_vendor_number text unique,
  legal_name     text not null,
  name_key       text not null,                    -- normalized for matching
  tin_enc        bytea,                            -- field-encrypted
  tin_status     text,                             -- verified | pending | mismatch
  remit_to       jsonb,                            -- {name, line1, line2, city, state, postal, country}
  country        text,
  default_sponsor_id uuid references ap_user(id),
  backup_sponsor_id  uuid references ap_user(id),
  default_gl     text,
  default_cost_center text,
  payment_terms  text,
  hcp            boolean not null default false,
  quality_critical boolean not null default false,
  asl_status     text,                             -- approved | conditional | not_approved | n/a
  fcpa_agent     boolean not null default false,
  intercompany   boolean not null default false,
  recurring_schedule text,                         -- monthly | quarterly | annual | null
  recurring_tolerance_pct numeric(5,2),
  po_required    boolean not null default false,
  phone          text,
  phone_since    date,                             -- QUA-2: usable for callback after 90 days unchanged
  status         text not null default 'active' check (status in ('active', 'blocked', 'setup')),
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now()
);

-- VEN-2: keyed HMAC of (routing + account) plus last 4; last-change date drives the 30/90-day rules.
create table vendor_bank_fingerprint (
  id             uuid primary key default gen_random_uuid(),
  vendor_id      uuid not null references vendor(id),
  fingerprint    bytea not null,
  last4          text not null,
  source         text not null,                    -- ap_verified_invoice | imported_vendor_list | sap
  verified_by    uuid references ap_user(id),
  effective_from timestamptz not null default now(),
  superseded_at  timestamptz
);
create index on vendor_bank_fingerprint (vendor_id) where superseded_at is null;

create table vendor_domain (
  id             uuid primary key default gen_random_uuid(),
  vendor_id      uuid not null references vendor(id),
  pattern        text not null,                    -- domain or exact address, lower-case
  status         text not null default 'proposed' check (status in ('proposed', 'approved', 'revoked')),
  added_by       uuid not null references ap_user(id),
  approved_by    uuid references ap_user(id),
  created_at     timestamptz not null default now(),
  unique (vendor_id, pattern)
);

create table billing_platform (
  id             text primary key,                 -- bill_com, quickbooks, coupa, ariba, ...
  display_name   text not null,
  sending_domains text[] not null,
  account_id_regex text                            -- how to find the vendor's account id in the message
);

create table billing_platform_map (
  id             uuid primary key default gen_random_uuid(),
  platform_id    text not null references billing_platform(id),
  vendor_id      uuid not null references vendor(id),
  account_id     text not null,
  added_by       uuid not null references ap_user(id),
  approved_by    uuid references ap_user(id),
  created_at     timestamptz not null default now(),
  unique (platform_id, account_id)
);

create table blocked_sender (
  pattern        text primary key,
  reason         text,
  blocked_by     uuid references ap_user(id),
  blocked_at     timestamptz not null default now()
);

-- ---------------------------------------------------------------------------
-- Messages and files (INT-1). The raw message is stored unaltered in the object store; this row is
-- the index. `trust` records whether it came through the Maxx Microsoft 365 redirect.
-- ---------------------------------------------------------------------------
create table message (
  id             uuid primary key default gen_random_uuid(),
  imap_uid       bigint,
  message_id     text,
  raw_ref        text not null,                    -- object key of the .eml
  raw_sha256     bytea not null,
  from_addr      text,
  from_name      text,
  original_sender text,                            -- for internal forwards: the true sender
  forwarded_by   text,
  subject        text,
  received_at    timestamptz not null,             -- Date / first Received, not our clock
  ingested_at    timestamptz not null default now(),
  headers        jsonb not null,
  auth_results   jsonb,                            -- parsed Authentication-Results / ARC
  trust          text not null check (trust in ('forwarded_by_maxx', 'direct', 'portal_own', 'bounce')),
  classification text,                             -- invoice | statement | credit_memo | bank_change | inquiry | spam_other
  classifier_detail jsonb,
  unique (message_id, raw_sha256)
);
create index on message (ingested_at);

create table attachment (
  id             uuid primary key default gen_random_uuid(),
  message_id     uuid not null references message(id),
  filename       text,
  content_type   text,
  size_bytes     bigint not null,
  sha256         bytea not null,
  ref            text not null,
  risky          boolean not null default false,
  risk_reason    text
);
create index on attachment (sha256);

-- ---------------------------------------------------------------------------
-- Invoices (lifecycle table in the spec). One row per invoice; several may come from one message.
-- ---------------------------------------------------------------------------
create type invoice_state as enum (
  'RECEIVED', 'CAPTURED', 'AP_REVIEW', 'QUARANTINED', 'VENDOR_SETUP', 'PENDING_APPROVAL',
  'QUESTIONED', 'APPROVED', 'ENTERED',
  'DIVERTED', 'DUPLICATE', 'FRAUD', 'REJECTED', 'VOID'
);

create table invoice (
  id             uuid primary key default gen_random_uuid(),
  portal_number  bigserial unique,                 -- human-readable AP-000123
  state          invoice_state not null default 'RECEIVED',
  message_id     uuid references message(id),
  attachment_id  uuid references attachment(id),
  source         text not null default 'email' check (source in ('email', 'upload')),
  split_index    int,                              -- INT-3: position within a multi-invoice file
  split_confirmed boolean not null default true,
  page_from      int, page_to int,
  file_sha256    bytea,
  vendor_id      uuid references vendor(id),
  vendor_match   jsonb,                            -- how the vendor was matched (keys, confidence)
  -- header fields (CAP-1). Each has a confidence in `field_confidence`.
  vendor_name_raw text,
  invoice_number text,
  invoice_number_norm text,
  invoice_date   date,
  due_date       date,
  currency       text,
  gross_cents    bigint,
  tax_cents      bigint,
  line_total_cents bigint,
  po_number      text,
  remit_to       jsonb,
  bank_last4     text,
  bank_fingerprint bytea,
  field_confidence jsonb not null default '{}',
  content_hash   bytea,                            -- SIG-3: sha256(original file || approved fields)
  tags           text[] not null default '{}',     -- recurring, hcp, quality_critical, fcpa_agent, intercompany
  soft_signals   text[] not null default '{}',
  diverted_reason text,
  linked_invoice_id uuid references invoice(id),   -- EXC-1: corrected invoice links the old one
  received_at    timestamptz not null,
  sla_paused_at  timestamptz,
  sla_paused_total interval not null default '0',
  routed_at      timestamptz,
  approved_at    timestamptz,
  entered_at     timestamptz,
  sap_document_number text,
  entered_by     uuid references ap_user(id),
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now()
);
create index on invoice (state);
create index on invoice (vendor_id, invoice_number_norm);
create index on invoice (file_sha256);

create table invoice_line (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     uuid not null references invoice(id),
  line_no        int not null,
  description    text,
  quantity       numeric(14,4),
  uom            text,
  unit_price_cents bigint,
  amount_cents   bigint,
  gl             text,
  cost_object    text,
  po_item        text,
  tax_code       text
);

-- Pages rendered by the parser (DATA-2: approvers only ever see these PNGs).
create table invoice_render (
  invoice_id     uuid not null references invoice(id),
  page           int not null,
  ref            text not null,
  width          int, height int,
  primary key (invoice_id, page)
);

-- Field edits by AP (AP_REVIEW -> CAPTURED, and the "any edit voids signatures" rule).
create table invoice_correction (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     uuid not null references invoice(id),
  user_id        uuid not null references ap_user(id),
  diff           jsonb not null,
  created_at     timestamptz not null default now()
);

create table invoice_question (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     uuid not null references invoice(id),
  asked_by       uuid not null references ap_user(id),
  question       text not null,
  asked_at       timestamptz not null default now(),
  answered_by    uuid references ap_user(id),
  answer         text,
  answered_at    timestamptz
);

-- ---------------------------------------------------------------------------
-- Routing (APR-1). The matrix is versioned; each approval stores the version it was judged under.
-- ---------------------------------------------------------------------------
create table matrix_version (
  id             bigserial primary key,
  rules          jsonb not null,                   -- [{min_cents, max_cents, gl_prefix, cost_center, roles:[...], sequential:bool}]
  proposed_by    uuid not null references ap_user(id),
  proposed_at    timestamptz not null default now(),
  approved_by    uuid references ap_user(id),
  approved_at    timestamptz,
  effective_from timestamptz
);

create table approval_request (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     uuid not null references invoice(id),
  matrix_version_id bigint references matrix_version(id),
  assignee_id    uuid not null references ap_user(id),
  role           text not null,                    -- sponsor | threshold | compliance | quality | cfo
  tier           int not null default 1,           -- APR-5: same tier runs in parallel
  required_amount_cents bigint,
  created_at     timestamptz not null default now(),
  due_at         timestamptz not null,             -- business-day clock to the approver
  reminded_at    timestamptz,
  escalated_manager_at timestamptz,
  escalated_controller_at timestamptz,
  status         text not null default 'open' check (status in ('open', 'signed', 'rejected', 'reassigned', 'voided')),
  closed_at      timestamptz,
  reassigned_to  uuid references ap_user(id),
  reassigned_by  uuid references ap_user(id),
  reassign_reason text
);
create index on approval_request (assignee_id) where status = 'open';
create index on approval_request (invoice_id);

create table delegation (
  id             uuid primary key default gen_random_uuid(),
  from_user_id   uuid not null references ap_user(id),
  to_user_id     uuid not null references ap_user(id),
  starts_on      date not null,
  ends_on        date not null,
  created_at     timestamptz not null default now(),
  revoked_at     timestamptz,
  check (from_user_id <> to_user_id)
);

-- ---------------------------------------------------------------------------
-- Signatures (SIG-*). Insert-only; voids are new rows.
-- ---------------------------------------------------------------------------
create table attestation_text (
  id             text primary key,                 -- att-v1
  text           text not null,
  effective_from timestamptz not null default now()
);
insert into attestation_text (id, text) values
  ('att-v1', 'I have reviewed this invoice and approve it for payment under my authority.');

create table approval_signature (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     uuid not null references invoice(id),
  approval_request_id uuid references approval_request(id),
  signer_id      uuid not null references ap_user(id),
  signer_role    text not null,
  decision       text not null check (decision in ('approve', 'reject')),
  reason         text,
  content_hash   bytea not null,
  payload        jsonb not null,                   -- JCS-canonical payload as signed
  payload_sha256 bytea not null,
  passkey_id     uuid references passkey(id),
  signing_session_id uuid references signing_session(id),
  assertion      jsonb,                            -- WebAuthn assertion (null for in-session click under step-up)
  seal           bytea not null,                   -- signing-service signature over payload_sha256
  seal_key_id    text not null,
  attestation_text_id text not null references attestation_text(id),
  matrix_version_id bigint references matrix_version(id),
  signed_at      timestamptz not null default now(),
  status         text not null default 'valid' check (status in ('valid', 'voided', 'withdrawn')),
  supersedes     uuid references approval_signature(id)   -- the record this one voids/withdraws
);
create index on approval_signature (invoice_id);

-- ---------------------------------------------------------------------------
-- Quarantine (QUA-*).
-- ---------------------------------------------------------------------------
create table quarantine_case (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     uuid references invoice(id),
  message_id     uuid references message(id),
  triggers       text[] not null,
  detail         jsonb,
  state          text not null default 'open' check (state in ('open', 'released', 'rejected', 'reported')),
  opened_at      timestamptz not null default now(),
  reviewer_id    uuid references ap_user(id),
  second_reviewer_id uuid references ap_user(id),
  callback_source text,                            -- contract | onboarding_file | vendor_master_90d
  callback_number text,
  callback_note  text,
  decision_at    timestamptz,
  decision_reason text,
  escalated_cfo_at timestamptz
);
create index on quarantine_case (state);

-- ---------------------------------------------------------------------------
-- Notifications (NOT-*).
-- ---------------------------------------------------------------------------
create table notification (
  id             uuid primary key default gen_random_uuid(),
  user_id        uuid references ap_user(id),
  to_addr        text not null,
  type           text not null,                    -- pending | reminder | manager_copy | digest | enroll | freeze | heartbeat | controller_daily
  template_version text not null,
  message_id_header text,                          -- so the ingestor can drop our own mail
  related_invoice_id uuid references invoice(id),
  sent_at        timestamptz,
  delivery_status text not null default 'queued',
  error          text,
  created_at     timestamptz not null default now()
);

-- ---------------------------------------------------------------------------
-- Audit: insert-only, hash-chained (AUD-1). Updates and deletes raise.
-- ---------------------------------------------------------------------------
create table event_log (
  seq            bigserial primary key,
  at             timestamptz not null default now(),
  invoice_id     uuid references invoice(id),
  from_state     invoice_state,
  to_state       invoice_state,
  event_type     text not null,                    -- transition | signature | config | access | security | ...
  actor_id       uuid references ap_user(id),
  actor_type     text not null check (actor_type in ('user', 'system', 'controller', 'second_person')),
  reason_code    text,
  payload        jsonb not null default '{}',
  prev_hash      bytea not null,
  hash           bytea not null
);
create index on event_log (invoice_id);

create table security_event (
  id             uuid primary key default gen_random_uuid(),
  at             timestamptz not null default now(),
  type           text not null,
  user_id        uuid references ap_user(id),
  source         text,
  detail         jsonb not null default '{}',
  event_seq      bigint references event_log(seq)  -- also chained into event_log
);

create table anchor (
  id             bigserial primary key,
  at             timestamptz not null default now(),
  head_seq       bigint not null,
  head_hash      bytea not null,
  object_key     text,
  status         text not null default 'pending' check (status in ('pending', 'written', 'failed')),
  error          text
);

create or replace function refuse_change() returns trigger language plpgsql as $$
begin
  raise exception 'table % is insert-only', tg_table_name using errcode = '42501';
end $$;

create trigger event_log_insert_only before update or delete on event_log
  for each row execute function refuse_change();
create trigger approval_signature_insert_only before update or delete on approval_signature
  for each row execute function refuse_change();
create trigger security_event_insert_only before update or delete on security_event
  for each row execute function refuse_change();
create trigger message_insert_only before delete on message
  for each row execute function refuse_change();
create trigger attachment_insert_only before delete on attachment
  for each row execute function refuse_change();
create trigger invoice_no_delete before delete on invoice
  for each row execute function refuse_change();

-- Daily counts for the INT-4 heartbeat.
create view forwarded_volume_by_day as
  select date_trunc('day', ingested_at) as day, trust, count(*) as n
  from message group by 1, 2;

-- Seed billing platforms the shadow run will see.
insert into billing_platform (id, display_name, sending_domains, account_id_regex) values
  ('bill_com',   'Bill.com',    '{hq.bill.com,bill.com}',                 '(?i)account\s*(?:id|#|number)?[:\s]+([A-Z0-9-]{6,})'),
  ('quickbooks', 'QuickBooks',  '{intuit.com,notification.intuit.com}',   '(?i)customer\s*(?:id|#)?[:\s]+([A-Z0-9-]{4,})'),
  ('coupa',      'Coupa',       '{coupahost.com,coupa.com}',              '(?i)supplier\s*(?:id|#)?[:\s]+([A-Z0-9-]{4,})'),
  ('ariba',      'SAP Ariba',   '{ansmtp.ariba.com,ariba.com}',           '(?i)AN[0-9]{10,}'),
  ('tipalti',    'Tipalti',     '{tipalti.com}',                          null),
  ('stampli',    'Stampli',     '{stampli.com}',                          null);
