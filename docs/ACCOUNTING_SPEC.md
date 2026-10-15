# Accounting Portal — Stage 1 Implementation Spec (environment-applied)

Oct 14, 2026 · Derived from "AP Automation Portal — System Spec" (Oct 13) and a review of the `maxxdash`, `usage` and `MO-commissions` repos. Where this document and the original spec disagree on an implementation detail, this document wins; the original keeps authority on controls, lifecycle and scope.

Stage 2 (SAP) is unchanged and not covered here.

---

## 0. What changed, and why

| Spec said | Environment says | Decision here |
| --- | --- | --- |
| "A new AP role in Max Dash" (ID-2) | `user_profiles.role` is a single enum value. A sales manager who sponsors a vendor cannot also hold an `ap_user` role. `maxx_accounting` already exists and means "manages commission statements". | Portal access is a **grant**, not a role: new table `public.portal_access(user_id, portal)`, managed from the admin panel. Any MaxxDash user, whatever their role, can be granted the accounting portal. SoD is per invoice, as the spec already says. |
| Portal verifies Supabase tokens against signing keys (ID-1) | Nothing in MaxxDash verifies a JWT itself; it calls GoTrue `/auth/v1/user`. Asymmetric signing keys are reportedly on (Andy, Oct 14) but unverified in code. | Portal verifies the access JWT against the project JWKS **and** calls `/auth/v1/user` on sign-in, on every invoice page load and before every signature. The JWT alone does not reflect a ban; the `/user` call does. |
| Fresh sign-in every session, 8 h absolute, 30 min idle (ID-6) | Supabase access tokens live 1 hour; MaxxDash keeps a 30-day refresh cookie to bridge that. | Portal keeps a **server-side session** (opaque host-only cookie → `portal_session` row holding the encrypted refresh token). It refreshes silently until idle or absolute expiry, then the row is deleted. No MaxxDash cookie is read, no refresh cookie is ever set in the browser. |
| Leaver disable via "a ban in Supabase" (ID-5) | The admin panel has create / set role / set password / delete. No ban. The only way today is the Supabase console. | Add a **Disable** action to the MaxxDash admin panel (`PUT /auth/v1/admin/users/{id}` with `ban_duration`, via `lib/admin-service-role.ts`). Portal re-validates every open session hourly; a banned user's `/user` call fails → session killed, open items reassigned. |
| One purpose-built capture API first; LLM vision only if accuracy < 92 % (D7) | `usage/app/pipeline/` already has a production vision extractor: PDF→JPEG, preprocessing, per-field confidence, Anthropic (BAA) with escalation, retry classification. | **Flip D7.** Start with the usage pipeline on Anthropic as the one extractor. Add a commercial capture API only if the 200-invoice test is under 92 %. CAP-5 (a model can add a flag, never clear one) is unchanged. |
| Isolated parsing VM (INT-5) | One Hostinger VPS, Docker Compose, Traefik. No hypervisor access. | For testing and the shadow run: a dedicated **parser container** with `network_mode: none`, read-only root, non-root user, `pids_limit`, memory/CPU caps, tmpfs scratch. It receives bytes over a Unix socket and returns PNG + text. This is weaker than a VM and is recorded as a known gap; the production VPS (PLAT-2) revisits it. |
| Signing service "on a separate host under a different cloud account" (SIG-0) | Everything is on one Hostinger account that Andy administers. | A new **AWS account owned by Lenny Knobbs** (system owner). Lambda function URL + KMS asymmetric key (ECDSA P-256) + S3 bucket with Object Lock for audit anchors. Andy gets an IAM user that can only invoke the function and write anchors. Nobody on the VPS can read the key. |
| Mailbox read by IMAP, sends by SMTP (INT-4, NOT-3) | Mail for 90ten.life is on Hostinger. Both existing apps send through `smtp.hostinger.com:465` from one shared mailbox (apparently `andybot@phillyshah.com`). No IMAP code exists anywhere. | New Hostinger mailbox **`accounts@90ten.life`**, used by the portal only. IMAP `imap.hostinger.com:993`, SMTP `smtp.hostinger.com:465`. `EMAIL_FROM` must equal `SMTP_USER` (Hostinger refuses otherwise). The existing shared mailbox is not touched. |
| Portal at accounting.90ten.life | Traefik on the VPS already routes `dashboard.90ten.life`, `mobileapp.dashboard.90ten.life` and `usage.90ten.life` via the external `proxy` network and the `letsencrypt` resolver. | Testing and shadow run on the **same VPS** (A record + compose labels, half a day). Own VPS before cutover, per PLAT-2. |
| Nest commissions and usage under one login | Commission statements already live **inside MaxxDash** (Commissions tab, gated by `maxx_accounting`). `MO-commissions` is a separate public Flask generator on bare-metal Nginx. `usage` has **no sign-in at all** and runs on the service-role key. | The portal landing page is a tile menu. Commissions → MaxxDash Commissions tab. Usage → `usage.90ten.life`, placed behind **Traefik ForwardAuth** pointing at the portal's `/auth/verify`, which gives it sign-in without touching its code. `MO-commissions` either moves into compose behind the same ForwardAuth or stays as is (decision below). |

---

## 1. Settled decisions (restated with environment detail)

- **D1** SAP integration is Stage 2. v1 ends at an approved queue; AP keys into SAP and records the document number.
- **D3 / D4** Passkey tap per signing session (15 min idle / 60 max) under $25k; every approval at or above $25k and every HCP approval. Server seal from the Lenny-owned AWS account. Step-up amount is a config row under dual control.
- **D5 / ID-8** Identity is MaxxDash (Supabase Auth, project the dashboard already uses) for testing and the shadow run. Microsoft 365 sign-in is a Gate 2 condition before approvals become binding. The identity adapter is one module; swapping it is configuration.
- **D6** Bank fingerprint per vendor, captured from the first AP-verified invoice or imported from AP's vendor list. Stored field-encrypted in the portal DB. Never touches SAP in v1.
- **D7 (flipped)** One extractor: the usage vision pipeline on Anthropic. Deterministic validators decide; the model only proposes.
- **D9** One company code. Meril intercompany, credit memos, other company codes → manual queues.
- **D10** Build vs hybrid at Gate 1. Having a working extractor and a working Traefik/compose/Supabase pattern moves the lean from hybrid to **build (~65 %)**; the labor math in the original spec still holds, so a vendor demo is still worth running.
- **People:** Lenny Knobbs — system owner and AWS account owner. Corey Perine — backup Controller. Andy — MaxxDash admin, builder, grants portal access alone for now (recorded).

---

## 2. Target architecture

```
                 Microsoft 365 (Maxx)
                 accounting mailbox ──redirect rule──▶ accounts@90ten.life (Hostinger IMAP)
                                                                │
   Hostinger VPS 72.62.174.193 ── docker compose, network: proxy ─┼──────────────────────────┐
   │                                                            ▼                          │
   │  Traefik (existing)       ┌─ accounting-ingestor (Python) ─ polls IMAP every 60 s     │
   │   dashboard.90ten.life    │        │ raw .eml + attachments → object store (hashed)    │
   │   usage.90ten.life ◀──────┤        ▼                                                  │
   │   accounting.90ten.life   │  accounting-parser (Python, network_mode: none)            │
   │                           │        │ PNG per page + text + hidden-text flags          │
   │                           │        ▼                                                  │
   │                           ├─ accounting-worker (Python) ─ usage vision pipeline,       │
   │                           │     validators, classifier, routing, reminders, anchors   │
   │                           │                                                           │
   │                           └─ accounting-web (Next.js, copied shell from MaxxDash)      │
   │                                 sign-in, dashboard, approvals, WebAuthn, admin, CSV    │
   └───────────────────────────────────────────────────────────────────────────────────────┘
                 │                                   │                          │
                 ▼                                   ▼                          ▼
   Supabase project "maxx-accounting"    Supabase project "maxxdash"     AWS (Lenny's account)
   Postgres (portal DB), Storage          Auth only: JWKS, /auth/v1/user  Lambda sign+seal, KMS key,
   (originals, PNGs), PITR                user_profiles, portal_access    S3 Object Lock anchors
```

**Trust zones on one VPS.** This is not the three-host layout the spec draws. What holds the line in testing: the parser has no network; the seal key is off-box; the audit chain is anchored off-box; cookies are host-only so a MaxxDash or usage cookie never reaches the portal. What does not hold yet: a root compromise of the VPS exposes the portal DB credentials and the mailbox credentials. Own VPS and secrets vault before cutover (PLAT-1/2).

**One stack, one repo (decision Oct 14).** The portal lives in the `MO-commissions` repository as the Python package `accounting/` — that repo becomes the accounting codebase. FastAPI + Jinja server-rendered pages + a little vanilla JS (WebAuthn needs it), the same shape as `usage`. The MaxxDash login page, design tokens and type floors (55+ audience, 14 px minimum, 44 px targets, `--blue`/`--orange`, "?" help buttons as real buttons) are copied as HTML/CSS; the ingestion and extraction code is lifted from `usage/app/pipeline`. The web process and the worker are the same image with different commands; the parser is its own image. The existing Flask statement generator stays in the repo untouched and gets a Dockerfile so it can move behind Traefik.

---

## 3. Identity and sessions (replaces ID-1 to ID-7)

**ID-1** The portal talks to the **MaxxDash Supabase project** for identity only:
- Sign-in: `POST /auth/v1/token?grant_type=password` (same call as `apps/dashboard/app/auth/session/route.ts`, same rate limiter pattern from `lib/rate-limit.ts`).
- Verification: access JWT checked against `https://<maxxdash-project>.supabase.co/auth/v1/.well-known/jwks.json` (cached, re-fetched on unknown `kid`), then `GET /auth/v1/user` with the bearer token. Both must pass.
- Pre-build check: open the MaxxDash project → Authentication → JWT Signing Keys and confirm the current key is ECC (P-256) or RSA, not the legacy HS256 shared secret. If it is still HS256, rotate to an asymmetric key first; the dashboard is unaffected because it never verifies locally.

**ID-2** Access = a row in **`public.portal_access`** (MaxxDash DB):
```sql
create table public.portal_access (
  user_id    uuid not null references auth.users(id) on delete cascade,
  portal     text not null check (portal in ('accounting')),
  granted_by uuid not null references auth.users(id),
  granted_at timestamptz not null default now(),
  primary key (user_id, portal)
);
-- RLS: a user reads own rows; admins read all. Writes only via RPC.
create function public.admin_set_portal_access(target uuid, portal text, grant_it boolean) ...
  security definer -- gated by admin_assert_may_manage(target); maxx_sys_admin only
revoke all on function public.admin_set_portal_access from public, anon;
```
The portal reads `portal_access` and `user_profiles(role, display_name)` with the user's own token under RLS, exactly as MaxxDash reads profiles. No portal access → generic 404 at `accounting.90ten.life`. One grant covers clerk, controller and approver; the portal's own `ap_user` table carries the portal role (`clerk | supervisor | controller | approver | cfo | compliance | admin | auditor`) and the manager/backup fields, which MaxxDash does not have.

**ID-3** Background checks: JWKS verify + `/auth/v1/user` on sign-in, on every invoice page load, before every signature, and hourly for every open session. Failure → session row deleted, user signs in again.

**ID-4** Passkeys enrolled in the portal (WebAuthn, `userVerification: required`, RP ID `accounting.90ten.life`). Enrollment needs a live session plus a one-time code the Controller gives out-of-band. Enrollment emails the user and the Controller.

**ID-5** Leavers: Andy clicks **Disable** in the MaxxDash admin panel (new action) and removes the portal grant. Portal sees it within the hour and reassigns open items to the backup sponsor. Daily email to the Controller lists every grant, removal and disable.

**ID-6** Sessions: opaque `__Host-ap_session` cookie (host-only by construction, Secure, HttpOnly, SameSite=Strict, no Max-Age so it dies with the browser). Server row: user, encrypted refresh token, created, last seen, device hash, IP country. 30 min idle, 8 h absolute, new device or country → re-sign-in. The portal never reads `maxxdash-*` cookies and never sets a refresh cookie.

**ID-7** Break-glass unchanged (CFO enables; under step-up only; CFO co-approves; Controller reviews).

**Traefik ForwardAuth for the sibling apps.** The portal exposes `GET /auth/verify`: 200 with `X-Portal-User` when the session is valid, otherwise a 302 to the portal login with `next=` the page the person wanted. Traefik middleware `ap-auth` on `usage.accounting.90ten.life` and `commissions.accounting.90ten.life` sends every request there first; usage keeps its service-role key server-side and simply stops being reachable anonymously. Two cookies are issued at sign-in: the `__Host-ap_session` cookie that opens the portal itself (never leaves `accounting.90ten.life`), and a second opaque `ap_sso` cookie scoped to `Domain=accounting.90ten.life` that only `/auth/verify` accepts — it cannot open the portal, and it is never scoped to `.90ten.life`, so MaxxDash and website cookies still never reach us. The old `mo-commissions.90ten.life` name 301s to the new one.

---

## 4. Mail (replaces INT-4, NOT-3, MAIL-1 with specifics)

**Mailbox.** `accounts@90ten.life` on Hostinger, created in hPanel → Emails. IMAP on. App password stored only in the VPS `.env` for the `accounting-ingestor` container.

**Forwarding.** Exchange admin → Mail flow → Rules: on the Maxx accounting mailbox, **"Redirect the message to accounts@90ten.life"** (not "forward", not "forward as attachment"). Redirect keeps the original `From`, `Authentication-Results`, `ARC-*` and `DKIM-Signature` headers. Test with one message and check the raw source before trusting it.

**Trusting the path (INT-4).** The ingestor accepts a message only when the last `Received:` hop before Hostinger is from Microsoft 365 (`*.outbound.protection.outlook.com`, SPF-aligned to Maxx's tenant) and the `Authentication-Results` header names Maxx's tenant. Anything else — including mail sent straight to `accounts@90ten.life` — goes to QUARANTINED with trigger `direct-sender`. Daily heartbeat: expected forwarded volume (set after week 1) vs actual; outside range → email the Controller.

**Sending (NOT-3).** All portal mail from `accounts@90ten.life` via `smtp.hostinger.com:465`, implicit TLS. Reuse `usage/app/email.py` (Python) — it already does SMTP_SSL and returns a reason instead of raising. One fixed template set, fixed footer ("This mailbox never sends approval buttons or invoice links. If you receive one, forward it to security."). Portal's own outbound `Message-ID`s are recorded so the ingestor drops them, bounces and auto-replies.

**Domain auth (MAIL-1).** Before the first portal email:
1. SPF for 90ten.life includes `_spf.mail.hostinger.com` (hPanel shows the exact record).
2. DKIM enabled for 90ten.life in hPanel → Emails → DNS/DKIM.
3. DMARC: check the current `_dmarc.90ten.life` record. Move to `p=reject` only after confirming nothing else sends `@90ten.life` mail (MaxxDash and usage send from `@phillyshah.com`, so they are unaffected).
4. Confirm Hostinger's per-mailbox daily send limit in hPanel. Expected portal volume is well under 500/day; if the limit is lower than that, use a second mailbox for digests.

**Where DNS lives.** Not recorded in any repo. Expected to be Hostinger since VPS and mail are there; confirm in hPanel → Domains before week 1.

---

## 5. Data (replaces the v1 part of "Data model")

**Portal database: a new Supabase project `maxx-accounting`** (Pro, PITR add-on → RPO under 15 min, matching the NFR). The portal is the only client; it connects with its own service key and does authorization itself (the user is already identified by the MaxxDash token). No portal table is reachable from the browser. MaxxDash's project is **not** touched beyond `portal_access` and the admin Disable action.

**Object storage:** Supabase Storage in the same project, two private buckets: `originals` (raw `.eml`, attachments; written once, never updated, SHA-256 in the DB row) and `renders` (PNG pages served to approvers through the portal, never the original file).

**Core tables (v1):** `message`, `invoice`, `invoice_line`, `vendor`, `vendor_domain`, `billing_platform_map`, `routing_rule`, `matrix_version`, `ap_user`, `passkey`, `signing_session`, `approval_signature`, `attestation_text`, `notification`, `quarantine_case`, `security_event`, `event_log`, `portal_session`, `config` (step-up amount, dual-control flags, heartbeat range). Columns as in the original spec minus every `sap_*` field, which arrive in Stage 2 as additive migrations.

**Rules carried over from MaxxDash's database rules, because they have already bitten:**
- Money crosses API boundaries as **text**, never JSON numbers.
- Every function gets `revoke all ... from public, anon;`.
- New enum values get their own migration.
- `event_log` and `approval_signature` are insert-only: a trigger raises on UPDATE/DELETE, and the DB role the portal uses has no privilege to drop triggers.

**Audit chain (AUD-1):** `event_log.hash = sha256(prev_hash || canonical row)`. Every 15 minutes the worker PUTs the chain head to `s3://maxx-ap-anchors/<date>/<ts>.json` in Lenny's account with Object Lock (compliance mode, 7 years). A nightly job in the same account (Lambda on a schedule) re-walks the chain from the last anchor and emails the Controller on any gap.

**Retention (RET-1):** originals and signatures 7 years; Storage objects in `originals` are never deleted by any code path; a lifecycle rule in S3 mirrors `originals` nightly as the off-site copy.

---

## 6. Capture (replaces CAP-1 to CAP-4 with the existing pipeline)

Reused as-is from `usage/app/pipeline/`: `pdf.py` (page → JPEG), `preprocess.py`, `vision.py` (per-field value + confidence, provider switch, escalation model), `transient.py`, `tracer.py`. Rewritten: the prompt, the field schema (vendor, invoice number, date, due date, currency, gross, tax, PO number, remit-to, bank last 4 if printed, line totals) and `confidence.py` rules.

**Provider:** `VISION_PROVIDER=anthropic` only. OpenRouter is **not** used in the portal — invoices carry bank details and sometimes personal payee data, and there is no BAA or data-processing agreement with OpenRouter. (Different call from `usage`, where Andy accepted it on cost; recorded here so it reads as a decision.)

**Validators (deterministic, decide routing):** arithmetic ties, date sanity (not future, not > 1 year old), currency matches vendor, duplicate normalized invoice number across portal records, vendor match by (TIN, remit-to, sender domain) with the billing-platform allowlist, hidden-text detection on the extracted text layer (white-on-white, < 4 pt, off-page, "ignore previous instructions"-class strings) — any hit is a hard quarantine trigger (CAP-5).

**Accuracy gate (CAP-4):** 200 labeled invoices from the baseline weeks; header accuracy ≥ 92 % on repeat vendors or a second extractor is added.

---

## 7. Approvals and signature (SIG-0 to SIG-7, environment-specific)

**Signing service (SIG-0):** AWS Lambda in Lenny's account, function URL with IAM auth. Input: the WebAuthn assertion, the client data, the canonical JCS payload. The Lambda verifies the assertion against the passkey public key (which it reads from a DynamoDB table the portal writes at enrollment, so the portal cannot swap keys silently — enrollment writes are logged there too), recomputes the SHA-256, compares it to the challenge, then signs with a KMS asymmetric key (`ECC_NIST_P256`, `SIGN_VERIFY`). Key policy: only the Lambda role may `kms:Sign`; nobody may export; Andy's IAM user may `lambda:InvokeFunctionUrl` and `s3:PutObject` on the anchors bucket, nothing else. Portal fails closed if the Lambda is unreachable.

**Payload (SIG-1), attestation text (SIG-2), content-hash voiding (SIG-3), immutable records (SIG-5), chain rules (SIG-6):** unchanged. **SIG-7** nightly re-verification runs in the AWS account (it holds the public key), not on the VPS.

**UI:** approval screen is a MaxxDash-style page: rendered PNG left, extracted fields right, flags banner, vendor history, Approve / Reject (reason) / Reassign. Attestation text beside the button. Minimum 14 px, 44 px targets, `<HelpButton>` pattern for "Why do I need to tap my passkey?".

---

## 8. Notifications and reminders

Unchanged in substance (NOT-1 to NOT-5, REM-1 to REM-4): generic email, dashboard-home link only, D+2 reminder, D+3 manager copy (manager comes from `ap_user.manager_id`, set by Andy at grant time), D+4 Controller reassign, weekly digest. "Freeze my access" link: POST-only, hashed 128-bit token, 14 days, reduces access only.

Scheduler: APScheduler inside `accounting-worker`, business-day aware (US Federal holidays table), America/New_York.

---

## 9. Portal landing page (the "nesting")

`accounting.90ten.life/` after sign-in shows tiles the user is entitled to:

| Tile | Goes to | Gate |
| --- | --- | --- |
| Invoices (AP) | `/invoices` in the portal | `portal_access` + `ap_user` row |
| Commission statements | `https://dashboard.90ten.life` Commissions tab | MaxxDash already gates by role; the tile shows when `user_profiles.role` is `maxx_accounting`, an admin, or a commission reader |
| Usage tickets | `https://usage.accounting.90ten.life` (behind ForwardAuth) | `portal_access` |
| Statement generator | `https://commissions.accounting.90ten.life` (MO-commissions behind ForwardAuth) | `portal_access` and role `maxx_accounting` or admin |

The two sibling apps keep their own code and deploys. Moving MO-commissions means a small Dockerfile (gunicorn + LibreOffice) and compose labels; its bare-metal Nginx unit and the stale `deploy.sh` paths are retired.

---

## 10. Security controls

Everything in the original "Security controls" section stands (hard triggers, soft signals, billing-platform allowlist, forwarded-invoice rule, QUA-1 to QUA-5, shadow-run gate ≤ 3 %). Environment-specific additions:

- **PLAT-1** Secrets: VPS `.env` for testing; move to Docker secrets or a vault on the production VPS. MFA on Hostinger, Supabase, GitHub and AWS accounts now, not at cutover.
- **PLAT-2** Shared VPS for testing/shadow run; dedicated VPS before cutover. Host-only cookies from day one.
- **DATA-1** TIN and bank fingerprint columns encrypted with `pgsodium` (available in Supabase) using a key the worker holds; the DB alone cannot decrypt them.
- **DATA-2** Approvers only ever see PNGs from `renders`; `originals` is served to AP through the parser container, never directly.
- **LLM-1** The vision call receives images and text from the parser only; no tools, no URLs, no write path.
- **Backups** Supabase PITR plus a nightly `pg_dump` to the S3 bucket in Lenny's account; quarterly restore test actually performed.

---

## 11. Build order (Stage 1, ~14–16 weeks, weeks 0–3 overlap the baseline)

**Week 0 — setup Andy does by hand (no code):**
1. Create `accounts@90ten.life` on Hostinger; enable IMAP; note the app password.
2. Confirm SPF/DKIM for 90ten.life in hPanel; read the current DMARC record; confirm the daily send limit; confirm where 90ten.life DNS is hosted.
3. Microsoft 365 redirect rule from the accounting mailbox; send one test; check headers.
4. A record `accounting.90ten.life → 72.62.174.193`.
5. MaxxDash Supabase → JWT Signing Keys: confirm asymmetric; note the JWKS URL.
6. New Supabase project `maxx-accounting` (Pro + PITR).
7. Lenny creates the AWS account (root MFA), creates an IAM user for Andy with the narrow policy above. Builder never holds root.
8. Decide: sibling apps move under `*.accounting.90ten.life` (recommended) or stay on their hostnames with a ticket hand-off.

**Weeks 1–2 — MaxxDash side (one PR):** `portal_access` table + RPC, admin-panel grant toggle and Disable action, `lib/roles.ts` helper, tests (`role-scope-check.sh`, roles test, admin-gate test). Version bump, What's New.

**Weeks 2–4 — portal skeleton:** compose with four services, Traefik labels, Next.js shell copied from MaxxDash (`login/page.tsx`, `modal.tsx`, `help.tsx`, design tokens), sign-in route with JWKS + `/user` check, server-side session, 404 for non-granted users, landing tiles, ForwardAuth endpoint.

**Weeks 3–6 — ingest and capture:** IMAP poller, raw store + hash, trusted-path check, parser container, usage pipeline port, validators, classifier, event log + hash chain + anchors.

**Weeks 5–8 — vendor master, routing, quarantine:** vendor table and import from AP's list (with bank fingerprint), domains and platform map, matrix v1, auto-release rule, quarantine screen with callback record and second-person release.

**Weeks 7–10 — approvals:** WebAuthn enrollment, signing sessions, Lambda signing service, approval screen, chain completion, voiding on content change, approved queue + CSV, SAP document number entry.

**Weeks 9–10 — notifications, reminders, digests, freeze link, month-end received-not-entered report, dashboard tiles (RPT-1).**

**Weeks 10–12 — shadow run** (≥ 300 invoices, observe-only). **Gate 2:** quarantine ≤ 3 %, header accuracy ≥ 92 %, 90 % of known-vendor invoices routable within 1 business day, Microsoft 365 sign-in live.

**Weeks 12–16 — cutover by segment, dedicated VPS, pen test, restore test, day-30 acceptance.**

---

## 12. Open questions still open after this pass

- Who at Maxx can create the Microsoft 365 mail-flow rule (needs Exchange admin)?
- Is Lenny also Portal Admin, or is that a separate person? (The owner should not approve his own config proposals.)
- Sibling-app hostnames: move under `*.accounting.90ten.life` or keep and use a ticket hand-off?
- Annual invoice count and PO / recurring / intercompany / HCP split (targets move if the 3–6k assumption is wrong).
- Do any clinical invoices contain PHI? (Anthropic BAA covers it; the question decides whether approvers' rendered PNGs need redaction.)
- Entra tenant: which Maxx tenant, and who can register the OIDC app for Gate 2?
