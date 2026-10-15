# MO-commissions → Maxx accounting codebase

Two apps live here:

- `app.py` — the Flask commission-statement generator (`mo-commissions.90ten.life`). Unchanged; see the lower half of README.md.
- `accounting/` — the **Maxx Accounting Portal** (`accounting.90ten.life`): invoice intake from `accounts@90ten.life`, screening, capture, vendor routing, passkey-signed approvals, approved queue for SAP entry. Stage 1 of `docs/ACCOUNTING_SPEC.md` (environment-applied) — the original spec is `docs/AP_AUTOMATION_PORTAL_ORIGINAL_SPEC.md`.

## Read before touching `accounting/`
- `docs/ACCOUNTING_SPEC.md` §0 (what changed from the original spec and why) and §3 (identity).
- `accounting/CONTRACTS.md` — module boundaries and the cross-module function signatures.
- `docs/ACCOUNTING_RUNBOOK.md` — what Andy sets up by hand (mailbox, DNS, M365 redirect rule, Supabase, AWS) and how to deploy.
- `infra/signing-lambda/README.md` — the signing service in Lenny's AWS account.

## Rules that already have a reason behind them
- **Identity is MaxxDash, nothing else.** The portal verifies the Supabase JWT against the MaxxDash project's JWKS *and* calls `/auth/v1/user` (a ban only shows up there). Access is the `portal_access` grant in MaxxDash, not a role. The portal never holds the MaxxDash service-role key.
- **Sessions are server-side.** Opaque `__Host-` cookie, 30 min idle / 8 h absolute / hourly re-check. Never a refresh cookie in the browser; never read a `maxxdash-*` cookie.
- **Every state change goes through `lifecycle.transition()`.** It is the only writer of `invoice.state`, and it chains an `event_log` row. `event_log`, `approval_signature`, `security_event` are insert-only at the trigger level — tests that need a clean DB disable the trigger, truncate, re-enable (see `tests/accounting/conftest.py`).
- **Money is integer cents.** Never float; format at the edge (`web/templating.py`).
- **`ap_user.roles` is a `portal_role[]`.** Select it as `roles::text[]` or psycopg hands you the literal `'{clerk,admin}'` string — that bug shipped once in review and would have hidden the whole nav.
- **Model output never lowers a control** (CAP-5). The extractor proposes; validators and screening decide. Anthropic only — no OpenRouter for invoices (bank details, no BAA).
- **Emails carry no amounts, no vendor names, no invoice links** (NOT-1..4). The only state-changing link is "Freeze my access", POST-only.
- **Approvals fail closed** when the signing service is unreachable. Dev signer only when `ACCOUNTING_ENV=test`.
- **Design floors from MaxxDash:** 14 px minimum text, 12 px captions, 44 px targets, high contrast, "?" help is a real button opening a `<dialog>`. Every data page shows "As of …" and the confidential footer (base.html).

## Checks
From the repo root with the local Postgres 16 up (`pg_ctlcluster 16 main start`; DSN in `accounting/config.py` default):
- `ACCOUNTING_ENV=test python3 -m accounting.db.migrate`
- `ACCOUNTING_ENV=test SECURE_COOKIES=false python3 -m pytest tests/accounting -q` (run serially; the suite shares one database)
- `docker compose -f docker-compose.accounting.yml config`

## Branch
Active dev branch: `claude/sleepy-ramanujan-qhexzx`
