-- Package B additions (approvals, signing, quarantine). Additive only; 0001_core.sql is untouched.

-- WebAuthn challenges live server-side so the browser can never choose what it signs. A row is
-- written when options are issued and consumed (used_at) exactly once on verification; anything
-- older than two minutes is dead. For a SIGNING challenge the challenge IS the payload hash and the
-- payload is kept beside it, because the payload carries server_time and must be the same bytes at
-- verification as when the approver was shown the options.
create table webauthn_challenge (
  id                  uuid primary key default gen_random_uuid(),
  user_id             uuid not null references ap_user(id),
  kind                text not null check (kind in ('register', 'session', 'sign')),
  challenge           bytea not null,
  approval_request_id uuid references approval_request(id),
  payload             jsonb,
  created_at          timestamptz not null default now(),
  used_at             timestamptz
);
create index on webauthn_challenge (user_id, kind) where used_at is null;

-- QUA-2: a release that needs a second person is parked on the case until that person confirms on
-- the same screen. The first reviewer's callback record is stored at that point so the second person
-- sees exactly what they are confirming.
alter table quarantine_case add column release_pending_at timestamptz;
