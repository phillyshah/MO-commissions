# Module contracts (Stage 1)

Cross-module calls use exactly these signatures. Each module owns its own package; nobody edits
another package's files or `db/migrations/0001_core.sql` (add `0002_<area>.sql` etc. if a column is
genuinely missing — additive only). Templates live under `web/templates/<area>/…` and extend
`base.html`. All state changes go through `accounting.lifecycle.transition()`; everything else worth
auditing goes through `accounting.lifecycle.log_event()`. Money is integer cents everywhere.
`conn` is a psycopg connection inside a transaction from `accounting.db.pool.connection()`.

## Already written (do not rewrite)
- `accounting/config.py` → `settings`
- `accounting/db/pool.py` → `connection()`, `fetch_one/fetch_all/execute`
- `accounting/lifecycle.py` → `transition(conn, invoice_id, to_state, *, reason_code, actor_type, actor_id=None, payload=None)`,
  `log_event(conn, *, event_type, actor_type, actor_id=None, invoice_id=None, from_state=None, to_state=None, reason_code=None, payload=None) -> seq`,
  `canonical(obj) -> bytes` (JCS-style), `verify_chain()`, `chain_head(conn)`
- `accounting/identity/*` → `sessions.CurrentUser` (fields: id, subject, email, display_name, roles, status, session_id, device_hash; `.has(*roles)`),
  `sessions.background_check(user) -> bool`, `sessions.revoke_all_for_user(conn, user_id, reason)`,
  `identity.router.issue_freeze_token(conn, user_id) -> raw_token`, `crypto.encrypt/decrypt/hmac_fingerprint/sha256`
- `accounting/web/deps.py` → `current_user`, `require(*roles)`, `require_active`
- `accounting/web/templating.py` → `render(request, name, ctx, status_code=200)`; filters `money`, `local`, `day`, `age`
- `accounting/web/routers.py` imports: `accounting.web.invoices.router`, `accounting.web.uploads.router`, `accounting.web.admin.router`,
  `accounting.approvals.router.router`, `accounting.quarantine.router.router`, `accounting.vendors.router.router`, `accounting.reports.router.router`
- `accounting/main.py` calls `accounting.reports.dashboard.home_tiles(user) -> dict(greeting, as_of, banners:[{kind,text}], work:[{title,count,tone,text,href}], systems:[{title,text,href,external}])`

## Package A — store, ingest, parse, capture, screening, vendors, pipeline
- `accounting/store.py`: `put(bucket: str, key: str, data: bytes, content_type: str) -> str(ref)`, `get(ref) -> bytes`, `exists(ref)`. Buckets: `originals`, `renders`. Local backend writes under `settings.store_local_dir`; s3 backend uses boto3-compatible endpoint settings.
- `accounting/ingest/intake.py`: `ingest_raw_message(raw: bytes, *, imap_uid: int | None, source: str = "email") -> list[UUID]` — stores raw, parses headers, decides `trust`, classifies, stores attachments (risky flag), splits into invoice rows in state RECEIVED (one per invoice; `split_confirmed=False` when the split is a guess), writes quarantine_case + QUARANTINED for direct senders / risky files, DIVERTED for statements/credit memos/bank-change requests, returns new invoice ids. Idempotent on (message_id, raw_sha256). Drops the portal's own mail (notification.message_id_header), bounces and auto-replies.
- `accounting/ingest/imap_poller.py`: `poll_once() -> int` (new messages ingested; tracks last UID in a `config_version`-free table of its own: add `0002_ingest_state.sql`).
- `accounting/ingest/heartbeat.py`: `check_daily() -> dict` compares yesterday's forwarded volume with settings range; returns {ok, n, text}.
- `accounting/parse/service.py`: the parser container entrypoint (`python -m accounting.parse.service`) — Unix-socket server; request {bytes, content_type, filename} → {pages:[png bytes], text, page_count, hidden_text_flags:[...], risky:bool, risk_reason}. Enforces settings.parser_* limits. Uses pypdfium2 + Pillow; images (jpg/png) are a one-page document. No network.
- `accounting/parse/client.py`: `parse(data: bytes, content_type: str, filename: str) -> ParseResult`; in test mode (no socket) runs the service in-process.
- `accounting/capture/extractor.py`: `extract(pages_png: list[bytes], text_hint: str) -> dict[field -> {value, confidence}]` via Anthropic (port usage/app/pipeline/vision.py: JSON-only prompt, defensive parse, escalation model, transient retry). Fields: vendor_name, invoice_number, invoice_date, due_date, currency, gross, tax, po_number, remit_to{...}, bank_last4, line_total, lines[].
- `accounting/capture/validators.py`: `validate(invoice_row, extracted) -> list[Failure(code, field, text)]`.
- `accounting/vendors/match.py`: `match_vendor(conn, *, sender_addr, vendor_name, tin, remit_to) -> MatchResult(vendor_id|None, method, confidence, detail)` using vendor_domain, billing_platform_map(+account id in text), name_key, remit-to.
- `accounting/vendors/fingerprint.py`: `fingerprint(routing: str, account: str) -> (bytes, last4)`; `current(conn, vendor_id) -> row|None`; `record(conn, vendor_id, fp, last4, source, verified_by)`.
- `accounting/vendors/lookalike.py`: `is_lookalike(sender_domain, known_domains) -> (bool, reason)` — edit distance ≤2, homoglyphs, added words.
- `accounting/screening/run.py`: `screen(conn, invoice_id) -> Screen(hard:[codes], soft:[codes], detail)`; hard list per spec; if hard → open quarantine_case and transition to QUARANTINED.
- `accounting/pipeline.py`: `process_invoice(invoice_id) -> str(final_state)` — RECEIVED → screen → parse → extract → validate → vendor match → (CAPTURED | AP_REVIEW | QUARANTINED | VENDOR_SETUP); on CAPTURED calls `accounting.routing.route.route(conn, invoice_id)`. `process_pending(limit=20) -> int`.
- `accounting/vendors/router.py` + templates `web/templates/vendors/`: list, detail (history, fingerprint last4 + since, domains, platform map), new/edit (dual control: proposer ≠ approver when `config_current.dual_control_enabled` is true; the admin may act alone when false), CSV import (`vendors/import`), domain approve, VENDOR_SETUP queue ("Add this vendor" from an invoice).

## Package B — routing, approvals, signing, quarantine
- `accounting/routing/route.py`: `route(conn, invoice_id) -> str(state)`. Order: ASL/quality hard stops (→ AP_REVIEW), auto-release (`auto_release.eligible(conn, inv)` → APPROVED, reason `auto_release_recurring`, payload rule+variance), else build chain from `matrix.current(conn)` + vendor sponsor (+Compliance for HCP) → approval_request rows (tier parallel) → PENDING_APPROVAL. Never route to the requester/corrector (SOD-1) or someone who edited the vendor (SOD-2); skip self; apply active delegation.
- `accounting/routing/auto_release.py`: recurring vendor, prior ENTERED/APPROVED invoice for a different service period, within ±tolerance, no soft signals, no quarantine history, no bank change in 90 days, not new, not HCP.
- `accounting/routing/matrix.py`: `current(conn) -> (version_id, rules)`; `propose(conn, rules, by)`, `approve(conn, version_id, by)` (dual control, effective next business day); `required_roles(rules, inv) -> list[(role, tier)]`.
- `accounting/routing/reassign.py`: `reassign(conn, request_id, to_user, by, reason)`, `escalate_due(conn, now) -> list[action]` (D+2 remind, D+3 manager copy, D+4 controller) — returns notification intents `{type, user_id, invoice_id}` for the notify package to send; the leaver sweep `reassign_open_items_for(conn, user_id)`.
- `accounting/approvals/payload.py`: `content_hash(conn, invoice_id) -> bytes` (sha256(original file bytes hash || canonical approved fields)), `build_payload(conn, invoice_id, signer, role, decision, attestation_id) -> dict`, `payload_hash(payload) -> bytes`.
- `accounting/approvals/webauthn_flow.py`: registration/authentication options + verification using the `webauthn` package; RP from settings; stores passkey rows; enrollment requires a valid `enrollment_code`.
- `accounting/approvals/signing_sessions.py`: `start(conn, user, passkey_id, portal_session_id, device_hash)`, `active(conn, user) -> row|None` (15 idle / 60 max), `needs_tap(conn, inv, user) -> bool` (amount ≥ step-up from config_current, or HCP, or no active session).
- `accounting/approvals/signing_service.py`: `seal(payload_hash: bytes, assertion: dict | None, passkey_public_key: bytes | None) -> Seal(signature, key_id)`. Production: POST to settings.signing_service_url with SigV4 (hand-rolled or botocore); fails closed. Dev (settings.dev_seal_private_key_pem set AND is_test): local ECDSA P-256.
- `accounting/approvals/sign.py`: `approve(conn, user, request_id, assertion|None) -> signature_id`, `reject(conn, user, request_id, reason, assertion|None)`; checks SOD, background_check, content hash unchanged, chain completion → `transition(..., 'APPROVED', 'chain_complete')` and calls `accounting.notify.send.queue(conn, type='approved_for_ap', user_id=None, invoice_id=...)`.
- `accounting/approvals/void.py`: `void_all_for_invoice(conn, invoice_id, reason, by)` (SIG-3) — new rows with status voided + supersedes.
- `accounting/approvals/router.py` + `web/templates/approvals/`: `/approvals` (mine, oldest first, both clocks), `/approvals/{request_id}` (PNG pages, fields, flags banner, vendor history, attestation, Approve/Reject/Reassign/Ask a question), WebAuthn JSON endpoints under `/approvals/webauthn/*`, `/passkeys` (list, enroll with code, revoke), `/delegations`. Calls `sessions.background_check` on every invoice page load and before signing.
- `accounting/quarantine/service.py`: `release(conn, case_id, by, second_by|None, callback_source, callback_number, note)`, `reject(conn, case_id, by, reason)`, `report(conn, case_id, by)` (blocks sender domain, queues `notify` type `security_report`), `escalate_stale(conn)` (2 business days → CFO). Release for sender/bank cases needs callback + second person; any release > $10k needs second person.
- `accounting/quarantine/router.py` + `web/templates/quarantine/`: Controller screen per QUA-1.
- `infra/signing-lambda/handler.py` + `infra/signing-lambda/README.md`: Lambda that verifies the WebAuthn assertion (public key passed in request, matched against a DynamoDB registry the portal writes at enrollment), recomputes the payload hash, and signs with KMS; IAM policy snippets for Andy's user and Lenny's root.

## Package C — notify, reminders, reports, worker, audit, invoices UI, uploads, admin UI
- `accounting/business_days.py`: `add_business_days(dt, n)`, `business_days_between(a, b)`, `is_business_day(d)`; US federal holidays; America/New_York.
- `accounting/notify/smtp.py`: `send(to, subject, text, *, headers) -> (ok, reason, message_id)` — SMTP_SSL 465, From = settings.email_from; never raises.
- `accounting/notify/templates.py`: `TEMPLATE_VERSION`, builders for: pending, reminder, manager_copy, controller_reassign, digest, enroll_alert, freeze_link, heartbeat_alert, controller_daily, security_report, approved_for_ap, vendor_setup. No amounts/vendor/invoice links; footer text per NOT-4; freeze link per NOT-5.
- `accounting/notify/send.py`: `queue(conn, *, type, user_id, invoice_id=None, to=None, context=None) -> notification_id`; `flush(limit=50) -> int` sends queued rows, records message_id_header.
- `accounting/notify/reminders.py`: `run(now) -> int` — uses `routing.reassign.escalate_due`, queues notifications, stamps approval_request reminded/escalated columns; `weekly_digest()`; `controller_daily()`.
- `accounting/audit/anchors.py`: `anchor_now() -> row` (S3 PutObject with ObjectLockMode COMPLIANCE, 7y retention; local file in test); `accounting/audit/verifier.py`: `verify_since_last_anchor() -> (ok, detail)`.
- `accounting/worker.py`: APScheduler BlockingScheduler: imap poll (settings.imap_poll_seconds), `pipeline.process_pending`, `notify.send.flush` (30 s), `reminders.run` (hourly), heartbeat (08:00), anchors (15 min), nightly verify + SIG-7 re-verify, hourly session revalidation sweep, `quarantine.service.escalate_stale`, weekly digest (Mon 07:00), controller daily (07:30).
- `accounting/reports/dashboard.py`: `home_tiles(user)`; `nav_counts(user)`.
- `accounting/reports/queue.py`: approved queue (oldest first), CSV export columns: portal_number, vendor, sap_vendor_number, invoice_number, invoice_date, due_date, currency, gross, tax, po_number, gl, cost_center, approved_at, approvers; `record_sap_document(conn, invoice_id, number, by)` → ENTERED.
- `accounting/reports/month_end.py`: received-not-entered as of a cutoff (RPT-2/CUT-2), with Controller sign-off row (log_event).
- `accounting/reports/audit_export.py`: full event history of one invoice as CSV/JSON (RPT-3) incl. signature verification status.
- `accounting/reports/router.py` + `web/templates/reports/`: queue by status, AP workload, approver backlog, touchless rate, quarantine count, cycle time by stage, aging by sponsor, month-end, audit export. Every page shows "As of <time>" and the confidential footer (base.html does the footer).
- `accounting/web/invoices.py` + `web/templates/invoices/`: list with filters, detail (pages, fields with confidence, timeline from event_log, questions), AP review edit (writes invoice_correction, recomputes norm fields, calls `approvals.void.void_all_for_invoice` if any signature exists, transitions AP_REVIEW→CAPTURED then `routing.route.route`), mark duplicate/divert, confirm split, answer question (QUESTIONED→PENDING_APPROVAL), record SAP document number, VENDOR_SETUP hand-off link to vendors.
- `accounting/web/uploads.py`: `/upload` manual upload (INT-6) → `ingest.intake.ingest_raw_message(raw_eml_built_from_upload, imap_uid=None, source='upload')`.
- `accounting/web/admin.py` + `web/templates/admin/`: users (roles, manager, backup, status, freeze/unfreeze, issue enrollment code), config proposals/approvals (step_up_amount_cents, dual_control_enabled, soft_signals_enabled), matrix versions (propose/approve JSON rules), delegations, security events, daily grants report, break-glass toggle (CFO).
