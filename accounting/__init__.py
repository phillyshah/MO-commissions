"""Maxx accounting portal — Stage 1 (invoice intake, screening, routing, passkey approval).

Layout
------
config.py            environment settings (one place; every module imports `settings`)
db/                  psycopg pool, migration runner, SQL migrations (insert-only audit tables)
lifecycle.py         the invoice state machine and the hash-chained event log
identity/            MaxxDash (Supabase Auth) sign-in, JWKS verification, server-side sessions,
                     portal access checks, Traefik ForwardAuth endpoint
ingest/              IMAP poller for accounts@90ten.life, trusted-path check, raw store, classifier, splitter
parse/               client for the no-network parser container (rasterize + text + hidden-text flags)
capture/             vision extractor (ported from usage/app/pipeline), deterministic validators
vendors/             vendor master, bank fingerprint, allowed domains, billing-platform map, lookalike check
screening/           hard triggers (quarantine) and soft signals (banner only)
routing/             approval matrix, chains, auto-release rule
approvals/           WebAuthn enrollment/assertion, signing sessions, JCS payload, signing-service client
notify/              SMTP sender (accounts@90ten.life), templates, reminder scheduler, freeze link
reports/             approved queue + CSV, month-end received-not-entered, dashboard counts, audit export
web/                 FastAPI routers and Jinja templates (MaxxDash design rules: 14px min, 44px targets)
audit/               15-minute chain-head anchoring to S3 Object Lock, verifier
"""

APP_NAME = "Maxx Accounting Portal"
APP_VERSION = "0.1.0"
