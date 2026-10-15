# AP Automation Portal — System Spec

Oct 13, 2026 · @Andy

## Executive summary

Build an AP portal that reads the shared invoice inbox, screens every message for fraud, routes each invoice to the vendor's sponsor, collects a passkey-verified signed approval in a dashboard, and gives AP a clean approved queue to key into SAP. **v1 does not connect to SAP** (Maxx runs on Meril's S/4HANA with no Maxx-owned license; write access is unconfirmed). Target: median receipt-to-entered-in-SAP from about 10 business days to 5 or less.

The lag is mostly approval latency, not SAP entry (Andy: more than half of the 10 days is the approval process). Invoices sit in an inbox waiting for someone to notice them, then sit in a sponsor's inbox waiting for a sign-off nobody is tracking. The portal turns every invoice into a tracked record with a clock, an owner and an escalation path from the moment it lands.

How it works, in order:

1. **Ingest and screen.** Every message to the AP inbox becomes a record. Files are parsed in an isolated sandbox. Anything that looks like phishing or fraud is quarantined and sent to the controller before any approver sees it.
2. **Capture and match.** One capture API extracts header fields, deterministic validators check them, and the vendor is matched against a vendor list kept in the portal (sponsor, backup, GL defaults, bank fingerprint) that AP loads and maintains.
3. **Route and approve.** The sponsor gets a generic email, signs in through Max Dash, and signs the approval in the dashboard with a passkey tap. Reminder at D+2, manager escalation at D+3, Controller reassigns at D+4. Recurring invoices within ±2% of the last one skip the sponsor.
4. **Hand off to AP.** Approved invoices appear in an approved queue (and CSV export). AP keys each into SAP as today and records the SAP document number in the portal, which closes the clock.
5. **Report.** Cycle time by stage, aging by sponsor, quarantine and exception rates, and a received-not-entered list for month-end.

**Stage 2 (SAP integration)** is designed in this document but deferred: parking and posting invoices in SAP, reading vendor and payment status. It starts only when Meril confirms Maxx can have write access. Sections that describe it are marked Stage 2.

Two rounds of independent review (security and controls, then integration and spec consistency) changed the design materially; see Decisions and Review history. The biggest changes: passkey signatures instead of server-only sealing, a bank fingerprint for the quarantine rules, and a smaller v1 scope.

Recommendation on vehicle: decide at Gate 1 (week 3) between a hybrid (buy capture and the approval workflow shell; build the compliance pre-filter and signing layer) and a full build. My lean is hybrid, 55% confidence, because the security requirements add real operating load for one builder and the labor savings at around 3,000 invoices a year are about the size of a subscription. This document doubles as the requirements checklist for the vendor demo.

Confidence: the 50% cycle-time target is medium-high, because approvals are over half the lag and the portal attacks that directly with tracking and escalation; it is medium if sponsor SLAs are advisory. The earlier −75% manual-input target cannot be reached in v1 because AP still keys into SAP; v1 aims at about −40% blended (low confidence), and −75% on repeat-vendor invoices needs Stage 2.

## Decisions locked for v1 and scope

These are the decisions the reviews said would force a redesign if left open. Each has a default. Reverse one now, not after build starts.

| ID | Decision | Why |
| --- | --- | --- |
| D1 | SAP integration is out of v1 (Stage 2). v1 ends at an approved queue: AP keys approved invoices into SAP by hand and records the SAP document number in the portal. When Stage 2 is built, the design is park-then-post (route to be confirmed with Meril: LIV or FI invoice), and the portal never sets or clears payment blocks. | Maxx is on Meril's S/4HANA with no own license and LIV believed inactive, so write access is unconfirmed, while approvals are over half the lag. For Stage 2: one posting path, no reversals, and no posting user that also releases payment. |
| D2 | Stage 2 only. SAP is never called from the internet. Either an SAP job pulls from the portal over HTTPS with mutual TLS, or a relay agent inside the SAP network polls the portal outbound. | RFC and OData must not face the internet, and a VPS can't otherwise reach SAP. Chosen with Basis in Gate 0. |
| D3 | A passkey tap is the signature gesture: once per signing session (15 minutes idle, 60 maximum) for approvals under the step-up amount, on every approval at or above it, and on every HCP approval. A server seal adds integrity and time, with its key held in an account the builder doesn't control. | Silent-only approval lets one stolen session cookie sign every invoice under the threshold. A server-only seal is the server vouching for itself. |
| D4 | The step-up amount is a fixed, admin-configured dollar figure. It starts at $25k, is reset once from baseline data (near the 90th percentile of invoice amounts), and changes only under dual control. | A threshold that floats on its own is not auditable. |
| D5 | Testing: sign-in uses Max Dash accounts (Supabase Auth) and the portal verifies their tokens. Production: Microsoft 365 (Entra) sign-in, which adds MFA, manager and active status. One identity adapter, so the swap is configuration. The portal never stores passwords. Passkeys are enrolled in the portal. | Max Dash has password-only login, no MFA and no manager field (repo review). Acceptable for testing only; Microsoft 365 sign-in is a gate before production use. |
| D6 | The portal keeps a bank fingerprint per vendor, captured from the first AP-verified invoice or imported from AP's vendor list, and never reads or writes SAP in v1. A later difference triggers quarantine. Vendor creation and bank changes stay in SAP. | The bank-mismatch quarantine trigger needs a baseline to compare against; without SAP it comes from AP-verified data. |
| D7 | v1 uses one purpose-built capture API plus deterministic validators. An LLM-vision second extractor is added only if the 200-invoice test shows header accuracy under 92% on repeat vendors. | Less to build, and deterministic checks, not model agreement, protect routing. |
| D8 | Primary KPI is receipt-to-entered-in-SAP (median, business days; AP records the entry). v1 manual-input target is about -40% blended (low confidence, to be set from the time-and-motion study); -75% repeat-vendor and -55% blended need Stage 2. | AP still keys into SAP in v1, so keying effort does not fall. |
| D9 | v1 handles one company code. Meril intercompany invoices, credit memos and other company codes are diverted to manual queues. Currency and tax coding stay with AP at SAP entry. | Keeps v1 to routing and approval. |
| D10 | Build vs hybrid is decided at Gate 1 (week 3), after a vendor quote and demo against the must-have checklist. | Security and operations load for one builder. |

**In v1:** intake with splitting and duplicate checks; isolated parsing; one capture API plus validators; vendor match and routing by vendor; auto-release for recurring invoices under the auto-release rules; an approved queue and CSV export for AP (AP keys into SAP and records the document number); dashboard approval with passkey, signature and seal; Max Dash sign-in; generic notifications (NOT-1 to NOT-7); hard-trigger quarantine with soft signals as a banner only; HCP tag with Compliance co-sign and hard stops; event log with off-box anchoring; month-end report and core dashboard.

**Out of v1 (P2 or later):** SAP integration of any kind (parking, posting, vendor and payment-status reads, PO and goods-receipt matching: Stage 2); partial approval by line (a non-PO invoice has no lines, and the signature binds to the whole invoice); bulk approval; Slack/Teams; invoice deep links; vendor status replies; per-vendor correction memory; leaderboard (a manual report instead); automated FMV band checks (Compliance co-sign instead); FCPA certification tracking (manual tag plus Compliance); customs and transfer-pricing reconciliation (not an AP function); Open Payments extract (P2, but HCP tags are captured from day one so the data exists); intercompany and credit-memo automation; multiple company codes; cost-center routing; risk scoring on IP and geography; spreadsheet imports into SAP. Each item pulled into v1 adds roughly 1 to 3 weeks (my estimate, low confidence).

## Current state

Andy reports more than half of the 10 days is the approval process (the panel estimated roughly half); capture and SAP entry together are 2–3 days. These are the panel's estimates, not measured — baseline for 4 weeks before go-live or the 50% claim can't be proven.

| Stage | Est. days | Confidence |
| --- | --- | --- |
| Sitting in shared inbox before anyone triages | 1–2 | Med |
| Capture, coding, vendor lookup, finding the sponsor | 1 | Med |
| Waiting on sponsor approval | 4–5 | Med-High |
| Rework: wrong coding, disputes, "is this ours?", missing PO | 1–2 | Med |
| Batching for SAP entry (weekly cadence typical) | 1–2 | Med |

Failure modes in the as-is flow:

- No invoice is a tracked object until a person opens the email. Nothing has a clock, so nothing is late.
- Sponsor approval is requested by forwarded email, so it is invisible to AP once sent and invisible to the sponsor once scrolled past.
- Invoices are not in SAP until approved, so month-end accruals for received-but-unapproved invoices are guesswork, and SAP's duplicate check can't see them.
- Every invoice gets a sponsor approval, including PO-backed and recurring ones where the approval already happened at PO or contract time.
- Weekly batching to SAP adds up to a week of dead time to invoices approved on a Monday.
- The shared inbox is the attack surface for bank-detail-change fraud; there is no structural control between "new remit-to on an invoice" and payment.

## Target state

Every known-vendor invoice is routed to its sponsor within one business day of receipt and reaches AP as an approved, signed record. The portal (at accounting.90ten.life) is the system of record for status and approval; SAP stays the system of record for money, and in v1 AP keys approved invoices into SAP by hand. Stage 2 adds parking and posting.

1. **Receive (D0).** Every message to the AP inbox becomes a record, with or without an attachment: link-only invoices, forwards, replies. Originals and headers are stored immutably and each file is hashed.
2. **Screen and capture (D0).** Files are parsed in an isolated sandbox. The screen looks for phishing and fraud signals; hard triggers quarantine the invoice for the controller (see Security controls). One capture API returns header fields, and deterministic validators check footing, vendor match, duplicates (portal records), and date and currency sanity. Failures go to AP review.
3. **Classify (D0).** Vendor matched against the portal vendor list by tax ID, remit address and sender domain. Tags: recurring, HCP, quality-critical, FCPA agent. Statements, credit memos, Meril intercompany invoices and other company codes are diverted to manual queues. An unknown vendor stops in VENDOR\_SETUP until AP adds it (sponsor, SAP vendor number, callback-verified bank); the one-day target starts when the vendor is in the list.
4. **Route (D0 to D1).** Auto-release applies first (recurring invoice within ±2% of the last, one per service period). Everything else goes to the sponsor plus threshold approvers, in parallel unless the matrix requires an order, plus Compliance for any HCP payee.
5. **Approve and sign (D1 to D3).** The approver gets a generic email, signs in to the dashboard through Max Dash, passes the background checks, taps a passkey, reviews and signs. Each step in the chain signs separately.
6. **Hand off.** When the last signature lands the invoice moves to the approved queue. AP is notified and can download a CSV of approved invoices.
7. **Enter in SAP (AP, manual).** AP keys the invoice into SAP and records the SAP document number in the portal. That ends the portal's cycle clock.
8. **Pay.** SAP payment runs are unchanged, a human step with Treasury.

Timing budget against the 5-day target (business days):

| Stage | Budget | Clock holder |
| --- | --- | --- |
| Receive to routed (known vendors) | 1 or less | System plus AP exception queue |
| Routed to signed, auto-release | 0 | Nobody |
| Routed to signed, sponsor path | 3 or less, median target 1.5 | Sponsor, then manager |
| Signed to entered in SAP | 1 or less | AP |
| Receipt to entered, single-approver sponsor path | 5 or less |  |

Chains with more than one approver run in parallel, so the budget doesn't multiply. A chain the matrix forces into sequence adds up to 3 days per extra tier, and those invoices are reported separately.

Recurring invoices that auto-release skip the sponsor entirely. Guess: 10 to 25% of invoices (low confidence; measure it in the baseline). PO-matched auto-release needs SAP data and is Stage 2.

## Invoice lifecycle

Every invoice is in exactly one state. Terminal states: DIVERTED, DUPLICATE, FRAUD, REJECTED, ENTERED, VOID (PAID is added in Stage 2). No transition outside this table is valid, and each one writes an event\_log row with from-state, to-state, actor, actor type and a reason code.

| From | To | Trigger | Logged |
| --- | --- | --- | --- |
| RECEIVED | DIVERTED or DUPLICATE | Classifier rule (both classifiers agree) or file-hash match | Class, matched record |
| RECEIVED | CAPTURED | Capture succeeds and validators pass | Extracted fields |
| RECEIVED | AP\_REVIEW | A validator fails | Failing fields |
| RECEIVED, CAPTURED or AP\_REVIEW | QUARANTINED | A hard trigger fires | Signals fired |
| QUARANTINED | CAPTURED | Controller release (second person where required) | Reason, callback record |
| QUARANTINED | FRAUD | Controller rejects | Reason, blocked sender |
| AP\_REVIEW | CAPTURED | AP clerk edit | Field diffs |
| CAPTURED | VENDOR\_SETUP | Vendor not in the portal vendor list | Proposed sponsor |
| VENDOR\_SETUP | CAPTURED | AP adds the vendor and sponsor; callback verified | SAP vendor number, verifier |
| CAPTURED | APPROVED | Auto-release rule | Rule ID, variance |
| CAPTURED | PENDING\_APPROVAL | Routing | Chain, assignees |
| PENDING\_APPROVAL | QUESTIONED | Approver asks a question | Question text |
| QUESTIONED | PENDING\_APPROVAL | AP answers | Answer |
| PENDING\_APPROVAL | APPROVED | Last signature in the chain | Signature IDs |
| PENDING\_APPROVAL | REJECTED | Signed rejection | Reason |
| Any state before ENTERED | AP\_REVIEW | Edit to amount, vendor, remit data or coding | Diff; all signatures voided |
| APPROVED | ENTERED | AP keys the invoice into SAP and records the document number | SAP document number, AP user |
| Any non-terminal state | VOID | Controller with a second person | Reason |

After ENTERED, corrections are handled in SAP by AP with the Controller's approval and noted in the portal. The portal reverses nothing in v1.

Two clocks. The invoice SLA clock pauses while an invoice is QUESTIONED. Sponsor aging, and anything derived from it, never pauses; an out-of-office sponsor with no delegate counts against that sponsor.

## Review history and what changed

The spec went through two rounds of independent review. Round 1 (Oct 13) was a four-expert panel: AP controller, SAP architect, medical-device CFO and automation architect. Round 2 (Oct 14) was four fresh reviewers given only the document: application security, internal controls and fraud, integration and infrastructure feasibility, and spec completeness. Round 2 found that the original spec was not buildable as written; the table below is what changed.

**Round 1 findings and their status**

| Finding | Source | Status |
| --- | --- | --- |
| Get the invoice into SAP on receipt, approve afterward | All four | Adopted as park-then-post (D1). The payment-block variant was dropped in round 2. |
| Remove sponsor approval for PO-matched and recurring invoices | Controller, CFO | Adopted with conditions (APR-4) |
| Spreadsheet import is a dead end | Controller, SAP architect | Adopted: API calls through the SAP adapter (D2) |
| Surgeon, royalty and commission payments should not start as emailed invoices | CFO | Adopted as an intake tag plus Compliance co-sign; self-billing is outside this project |
| Meril intercompany bypasses sponsor routing | CFO | Adopted: diverted to a manual queue in v1 |
| Open Payments tagging, QMSR hard stop | CFO | Adopted (tag and hard stops in v1, extract in P2) |
| Duplicate-number normalization, 2-way and 3-way match, bank-change callback | Controller | Adopted (duplicate rules, APR-4, bank process) |
| Signed single-use email links and the Safe Links pre-click risk | Automation architect | Superseded: no approval authority in email at all |
| 75% manual-input target overstated | CFO vs automation architect | Resolved: -75% repeat-vendor, -55% blended (D8) |

**Round 2 findings and fixes**

| Finding | Severity | Fix |
| --- | --- | --- |
| "The portal can never pay" was false: posting makes an invoice payable in the next payment run, and the same SAP user posted and released | Critical | Reworded as a posting boundary; posting is gated on a verified signature chain; payment runs stay human with Treasury review |
| The server-side seal was the server vouching for itself; anyone with root on the VPS could obtain a valid seal | Critical | Passkey assertion per signing; signing service and KMS account outside the portal host; fails closed |
| "Bank data is never mirrored" disabled the bank-mismatch and recent-bank-change controls | Critical | Read-only bank fingerprint from SAP (VEN-2, D6) |
| Silent-only approval: one stolen session cookie signs everything under the threshold; MFA was P2 while step-up was v1 | High | Passkey tap per signing session (D3, ID-4) |
| Posting model mixed FI and LIV; no released function module clears a payment block | High | LIV park-then-post (D1) with a Gate 0 SAP spike to confirm function modules in your release |
| A VPS cannot reach SAP over the internet | High | Pull job or relay inside the SAP network (D2) |
| Quarantine triggers would flag QuickBooks, Bill.com, Ariba and similar platforms, running 10 to 25% instead of 2% | High | Triggers narrowed, per-platform allowlists, 3% shadow-run gate |
| Auto-release could bypass quarantine, bank-change and new-vendor controls; a fake goods receipt releases a PO invoice | High | Auto-release rules (APR-4) |
| Audit-log chain can be rewritten by the database owner | High | Insert-only tables, 15-minute anchors to write-once storage, off-box verifier (AUD-1) |
| Parser exploits on the app host; prompt injection steering routing | High | Isolated parsing VM, no URL fetching, LLM output never lowers a control (INT-5, CAP-5) |
| Compromised shared mailbox can plant messages with forged headers | Med-high | Mailbox scoping and trusted-header rule (INT-4) |
| Delegation, reassignment and role abuse; self-approval loopholes | Med-high | DEL-1, SOD-4 to SOD-7, ADM-1, role table |
| v1 scope too large for one builder (reviewer estimate: 9 to 12 months) | High | Scope lock; schedule 18 to 22 weeks after the cuts |
| 75% blended arithmetic doesn't hold; sequential chains blow the 4.5-day budget | High | D8; parallel chains |
| Closed-period auto-roll misstated cutoff | Medium | Controller decides; accrual listing (CUT-1 to CUT-4) |
| About 18 internal contradictions (reminder timing, step-up threshold, Slack in v1, leftover email-link language, soft-signal rule, partial approval) | Medium | Resolved in this revision |

**Where reviewers disagreed, and the call.** Number of extractors: one in v1 (D7). Recovery point: 15 minutes. Step-up threshold: fixed (D4). Build vs buy: decided at Gate 1, with my lean moved from build to hybrid.

## Functional requirements

Each requirement has an ID and a tag: **(v1)** ships in the first release, **(P2)** is designed for but built later, **(Out)** is explicitly not built (it goes to a manual queue or stays in SAP). Anything untagged does not exist. Acceptance for every v1 item is a test the Controller can watch run.

**Intake**

- **INT-1 (v1)** Every message to the shared inbox becomes a record with its original stored unaltered (raw message, attachments, headers, receive time). Nothing is deleted or auto-discarded; non-invoices are classified and kept.
- **INT-2 (v1)** Classifier labels each message: invoice, statement, credit memo, remittance/bank-change request, vendor inquiry, spam/other. Credit memos and anything unlabeled go to the manual queue (D9).
- **INT-3 (v1)** Multi-invoice PDFs and multi-attachment emails are split into one record per invoice; a split a human has not confirmed is flagged.
- **INT-4 (v1)** Ingestion reads only accounts@90ten.life, which receives a forwarded copy of the accounting mailbox from Maxx's Microsoft 365 through a mail-flow rule that preserves the original headers. The portal has no access to Maxx's Microsoft 365. It accepts mail only from the Maxx forwarding path (verified by sending IP range or connector, plus ARC where present) and trusts only the Authentication-Results stamped by Maxx's Microsoft 365, never headers a sender could forge. Anything sent directly to accounts@ from elsewhere is quarantined. A daily heartbeat compares forwarded volume to the expected range and alerts the Controller if the rule stops or floods.
- **INT-5 (v1)** Attachments are opened only in an isolated parsing VM with no access to the network, SAP or secrets: limits of 60 s CPU, 1 GB memory, 50 pages, 100 MB, and 100:1 decompression. No URLs in an email or PDF are ever fetched by the system. Reviewers see a rasterized PNG served from a cookieless origin, never the original file inline. Risky file types (macro Office files, executables, scripts, archives with executables) are quarantined.
- **INT-6 (v1)** Manual upload for invoices that arrive on paper, by portal download, or by another channel; it follows the same pipeline.
- **INT-7 (P2)** Vendor portal / EDI / e-invoicing feeds for the biggest vendors.

**Capture**

- **CAP-1 (v1)** One capture API plus deterministic validators extracts header fields: vendor, invoice number, date, due date, currency, gross, tax, PO number, remit-to, line totals. Line-level extraction is out of v1 for non-PO invoices; PO invoices match at header level plus open PO value.
- **CAP-2 (v1)** Every extracted field carries a confidence; validators check arithmetic (lines sum to total), date sanity, currency vs vendor, and PO existence.
- **CAP-3 (v1)** Fields below threshold go to a one-screen correction view showing the image beside the fields; corrections feed per-vendor learning (the 80% repeat vendors get better each cycle).
- **CAP-4 (v1)** Header accuracy is measured on a labeled set of 200 invoices before go-live; if below 92% a second extractor is added (D7).
- **CAP-5 (v1)** Model output can never lower a control. A model may add a flag, never clear one, never approve, never pick a payee or bank account. Hidden-text detection (white-on-white, tiny fonts, text outside the page, instructions aimed at an AI) is deterministic and is itself a quarantine trigger.
- **CAP-6 (v1)** Extracted text is treated as untrusted data everywhere it is displayed or passed on (escaped on output, never executed, never used as a prompt for an action).

**Vendor master and routing**

- **VEN-1 (v1)** Vendor record: SAP vendor number, legal name, TIN status, remit-to, default sponsor, backup sponsor, default GL/cost center, payment terms, HCP flag, ASL status, allowed sender domains, billing-platform mappings, recurring schedule if any.
- **VEN-2 (v1)** The portal keeps a bank fingerprint per vendor (keyed HMAC, last 4, last-change date), captured from the first AP-verified invoice or imported from AP's vendor list. It never reads or writes SAP in v1. Any invoice whose bank or remit-to differs from the fingerprint, or any request to change bank details, is quarantined (see Security).
- **VEN-3 (v1)** Sponsor, routing rules, allowed domains and recurring schedules are changed through dual control (one proposes, a different person approves); for now the Max Dash admin (Andy) may assign vendor approvers alone, and dual control is a config switch to turn on before cutover; changes are versioned and take effect the next business day.
- **VEN-4 (v1)** New vendors are never auto-routed: they go to AP with a request for sponsor and SAP vendor creation; the invoice waits parked in the portal, not in SAP.
- **VEN-5 (v1)** ASL / QMSR hard stop: a vendor that supplies product or service affecting product quality and is not Approved cannot be routed past AP; Quality is notified.
- **VEN-6 (v1)** HCP flag drives extra handling: signature tap on every approval, Compliance copy, Open Payments tagging at posting.
- **VEN-7 (P2)** Mapping of sponsors to SAP cost-center owners (CSKS-VERAK\_USER) to catch sponsor drift automatically.

**Approvals**

- **APR-1 (v1)** Each invoice routes to the vendor's sponsor; amount/GL/cost-center rules from the approval matrix add further signers (chain). Matrix is configured by Portal Admin under dual control and versioned; each approval stores the matrix version it was judged under.
- **APR-2 (v1)** Approver sees the invoice image, extracted fields, PO/GR match result, vendor history, and any flags, and chooses Approve, Reject (reason required), or Reassign (to a delegate, never to self).
- **APR-3 (v1)** Approving requires the signature step in Sign-in and signature below; nobody can approve by replying to an email.
- **APR-4 (v1)** v1 auto-release (no human approval) applies only to a recurring invoice within ±2% of the prior one, one per service period. PO-matched auto-release is Stage 2, because it needs goods-receipt data from SAP. Never auto-release: quarantined or soft-signal invoices, vendors with a bank change in the last 90 days, new vendors, HCP vendors.
- **APR-5 (v1)** Parallel chains: independent signers (for example cost-center owner and Quality) are asked at the same time; the invoice posts when all are done.
- **APR-6 (v1)** Overlap rule: if one person is required more than once on a chain, one signature covers it only when the amounts and bases are identical; otherwise they sign each.
- **APR-7 (v1)** Leaver or role change: HR / Max Dash signals trigger reassignment of open items to the backup sponsor within 1 hour.
- **APR-8 (v1)** Delegation: an approver can delegate for a date range; limits cannot exceed their own, and a delegate cannot approve their own expense or sub-delegate.
- **APR-9 (Out)** Partial approval (approving some lines or a lower amount). Disagreement goes back to AP as a reject with reason; the vendor is asked for a corrected invoice or credit memo.

**Reminders**

- **REM-1 (v1)** D0: generic email that something needs review on the dashboard (no invoice details, no links to the invoice). D+2: reminder to the approver. D+3: copy to the approver's manager. D+4: Controller reassigns or decides.
- **REM-2 (v1)** Two clocks: the business-day clock to the approver and an overall invoice-age clock for AP; the weekly status email reports both, and exceptions are visible to the Controller daily.
- **REM-3 (v1)** Weekly email per user: counts waiting on them, oldest age, and delinquent items by vendor; no amounts for HCP items. Dashboard shows full detail.
- **REM-4 (v1)** No Slack or SMS channel in v1; email only; Slack/Teams are out of v1.

**Exceptions**

- **EXC-1 (v1)** Reject / dispute: reason code, vendor notified by AP (never auto-sent by the system), invoice parked with status DISPUTED; a corrected invoice is a new record linked to the old.
- **EXC-2 (v1)** Duplicates: handled by the hard-block rules in the SAP spec (DUP-1).
- **EXC-3 (v1)** Missing PO / no GR: AP asks the sponsor; the invoice stays RECEIVED and ages on the AP clock.
- **EXC-4 (v1)** Quarantine: see Security; a quarantined invoice cannot be approved or sent to SAP until the Controller releases it.

**SAP interface and reporting**

- **SAP-1 (v1)** In v1 there is no connection to SAP. AP gets an approved queue and CSV export, keys the invoice into SAP, and records the SAP document number in the portal (a one-field entry that ends the cycle clock). Stage 2 connects through the adapter interface in the Data model section.
- **SAP-2 (P2)** Posting happens only after the signature chain is complete (D1); the portal never sets or clears payment blocks.
- **SAP-3 (P2)** Payment status (cleared / open) is read back nightly and shown on the invoice.
- **RPT-1 (v1)** Dashboard: queue by status, oldest first; AP workload; approver backlog; touchless rate; quarantine count.
- **RPT-2 (v1)** Month-end received-not-posted listing for accruals (CUT-2).
- **RPT-3 (v1)** Exportable audit report of any invoice's full event history.

## Sign-in, approval signature and notifications

Approvals run through the dashboard only. Emails carry no links to invoices and no approval buttons (v1 decision). The identity comes from Max Dash; the legally meaningful act is a passkey tap that signs exactly what the approver saw.

**What changed from the first draft.** The first draft proposed "silent authentication" as the approval check. Review showed that silent re-authentication only proves a browser cookie is still alive, so anyone with a stolen session could approve. It is kept as a background check (is the Max Dash session still valid), but it is not the signature. The signature is a passkey tap, once per signing session, which costs an approver about one tap per sitting.

### Sign-in

- **ID-1 (v1)** Sign-in goes through one identity adapter. Testing: the portal verifies Max Dash (Supabase Auth) tokens against the project's signing keys. Production: OIDC (authorization code + PKCE) against Microsoft 365 (Entra). The portal stores no passwords and no shared secrets that unlock approvals.
- **ID-2 (v1)** Roles and the approver list come from the identity adapter (an explicit AP role added to Max Dash for testing; Entra groups and manager in production). Access needs a new AP role created in Max Dash. Users without it see no AP tile, and a direct visit returns a generic not-found page (no hint the system exists). One AP role covers both approver and invoice creator; SOD-1 is enforced per invoice, so nobody can approve an invoice they captured or corrected, refreshed at every sign-in and at least hourly for open sessions.
- **ID-3 (v1)** Background checks: the portal re-validates the Max Dash session by a backend check (token signature, expiry, account not banned or disabled; never an iframe), on each page load of an invoice and before every signature. If it fails, the user signs in again; nothing is lost.
- **ID-4 (v1)** Passkey enrollment happens in the portal. A new passkey needs Max Dash sign-in plus a one-time code delivered out-of-band by the Controller (in person, phone call, or chat to a known number); each enrollment alerts the user and the Controller. Two passkeys per approver are recommended (phone and laptop).
- **ID-5 (v1)** Leavers: an account disable (a ban in Supabase during testing, an Entra disable in production) or HR event immediately revokes sessions and passkeys and reassigns open items within 1 hour (APR-7). The Max Dash admin (currently Andy) grants and removes the AP role; there is no backup admin for now (Andy's decision), and every grant or removal is reported to the Controller daily. A Max Dash role alone never makes anyone an approver: approvers must also be assigned as vendor sponsors in the portal, which Andy does alone for now (Andy's decision).
- **ID-6 (v1)** Sessions: no persistence. Every new session needs a fresh sign-in (the portal ignores any existing Max Dash cookie and never offers remember-me); the session ends when the browser closes, 8 hours absolute, 30 minutes idle; a new device or IP country change forces re-sign-in.
- **ID-7 (v1)** Max Dash outage fallback: if Max Dash is down for more than 4 hours, the CFO may enable a break-glass mode in which only invoices under the step-up amount are approved, each needing a CFO co-approval, and every action is logged and reviewed by the Controller afterward. No password fallback.

**ID-8 (v1) Identity phases.** Testing and the shadow run use Max Dash accounts. The shadow run is observe-only: nothing is posted to SAP and no approval is binding. Max Dash has no MFA, so Microsoft 365 sign-in (MFA, manager, active status, leaver disable) must be live before cutover, when approvals become binding and the portal posts to SAP. This is a Gate 2 condition. The identity adapter is the only code that knows which provider is in use.

### Why a passkey tap, and how much friction

Step-up rules (D3, D4): a signing session lasts 15 minutes idle / 60 minutes maximum. Within it, an approver taps their passkey once at the start and then approves invoices under the step-up amount with a click that is still bound to a server-side seal. At or above the step-up amount (starts at $25,000; reset once near the 90th percentile of baseline, dual control), and for every HCP invoice, each approval needs its own tap. In practice a sponsor with a handful of invoices a week taps once or twice a week; this is the balance between security and usability.

&#91;embedded content: approval flow · 3 checks, 1 signature\]

Nothing in an email grants authority. It comes from the live Max Dash session, the assignee match and the signature.

### Approval signature

A signature is a record that this person, on this device, approved exactly this content at this time. Mechanism: WebAuthn assertion (userVerification required) over a challenge that is the SHA-256 of an RFC 8785 (JCS) canonical JSON payload. A separate signing service verifies the assertion and recomputes the payload hash itself.

- **SIG-0 (v1)** The signing service runs on a separate host, under a different cloud account than the one the builder administers, with its sealing key in a KMS that account owner controls (ECDSA P-256 or Ed25519). The portal application cannot produce a seal on its own. If the service is unreachable or verification fails, the approval fails closed.
- **SIG-1 (v1)** Payload contains: invoice ID, content hash (original file + extracted fields as approved), vendor SAP number, bank fingerprint reference, gross amount and currency, GL / cost center, PO reference, approval-matrix version, signer user ID, signer role, attestation-text ID, and server time.
- **SIG-2 (v1)** Attestation text is versioned and shown next to the button ("I have reviewed this invoice and approve it for payment under my authority"); the text ID is in the payload.
- **SIG-3 (v1)** Content-hash binding: if any field or the file changes after an approval, every approval on that invoice is voided and re-requested.
- **SIG-4 (v1)** The seal (service signature over the verified payload) is stored with the assertion, public key ID and certificate chain of the passkey attestation where available.
- **SIG-5 (v1)** Each approval creates an immutable approval\_signature record; withdrawals and voids are new records, never edits.
- **SIG-6 (v1)** Chain rules: a signer cannot appear twice with different roles on one invoice; the requester, vendor-master editor, and approver of one invoice must be different people (see SoD); chains complete only when every required signature is sealed against the same content hash.
- **SIG-7 (v1)** Signature verification is part of the nightly integrity job: every sealed approval is re-verified against the stored payload and the public key.
- **SIG-8 (Out)** Bulk approval (one tap for many invoices). It breaks "approved exactly what they saw". Revisit in v2 only for recurring invoices within ±2% under the step-up amount.

**Regulatory position.** This is an internal control signature meant to hold up in an audit. It is not claimed to be a 21 CFR Part 11 electronic signature for regulated records: AP invoices are not records required by FDA predicate rules, so Part 11 does not apply to them. If any invoice is later used as a regulated quality record, run it through a Part 11 assessment first.

### Notifications (v1: generic)

- **NOT-1 (v1)** When an invoice needs a person, email: "You have a pending approval in the AP dashboard" with a link to the dashboard home. No amount, vendor, or invoice link.
- **NOT-2 (v1)** Reminder schedule per REM-1; weekly digest per REM-3.
- **NOT-3 (v1)** All portal email (approval notices, reminders, digests) is sent only from accounts@90ten.life, with SPF, DKIM and DMARC p=reject on 90ten.life, always in the same format, and never asks for a password or a payment action. Approvers are told this is the only sender. A reply to it is ingested as ordinary correspondence with no authority, and the ingestor ignores the portal's own outbound mail, bounces and auto-replies so nothing loops.
- **NOT-4 (v1)** Users are trained once: "the portal's email never has a button to approve or an invoice link; if one does, forward it to security." Every email carries a footer that says so.
- **NOT-5 (v1)** The one email link that changes state is "Freeze my access": POST-only, a 128-bit random token stored hashed, valid 14 days, no sign-in needed, and it can only reduce access (revoke sessions and passkeys, pause the user's approvals). The Controller gets a daily digest of freezes. Nothing else in email changes state.
- **NOT-6 (P2)** Deep links to the invoice, Slack/Teams notifications, and mobile push.

### Data additions

- **user**: id, Max Dash subject, email, roles, status, manager, backup, last directory sync.
- **passkey**: id, user, credential ID, public key, attestation, created, last used, status.
- **signing\_session**: id, user, passkey, started, last activity, device fingerprint, expires.
- **approval\_signature**: id, invoice, content hash, payload, assertion, seal, key ID, attestation\_text ID, matrix\_version, signed at, status (valid / voided / withdrawn).
- **attestation\_text**: id, text, effective date.
- **notification**: id, user, type, template version, sent time, delivery status.
- **matrix\_version**: id, rules, proposer, approver, effective date.
- The approval record has no token and no channel field (links and channels are not an authority).

## Security controls and suspicious-invoice screening

Friction budget: an approver sees one tap per sitting and nothing else new. Almost all the security work is on the AP and Controller side and in infrastructure, where friction costs nothing. Threat model, in order of likelihood: (1) fake or altered invoices and bank-change requests from outside, (2) a compromised genuine vendor mailbox, (3) stolen approver session, (4) insider fraud (colluding sponsor and AP), (5) attack on the portal itself.

### Suspicious-invoice screening

Any invoice that looks like phishing or payment fraud is quarantined and sent to the Controller. A quarantined invoice cannot be approved, signed, or sent to SAP.

&#91;embedded content: quarantine flow · 1 screen, 1 controller decision\]

Only invoices that pass the screen, or that the controller releases, ever reach a sponsor.

**Hard triggers** (any one quarantines the invoice)

- Sender domain is a lookalike of a vendor's known domain (edit distance, homoglyphs, added words).
- Display name matches a known vendor but the address is not on that vendor's allowed list and is not an allowlisted billing platform (see below).
- DMARC fails and the sender's domain publishes quarantine or reject (a domain with no DMARC policy does not trigger by itself).
- Bank or remit-to details differ from the fingerprint, or the message asks to change bank details.
- Vendor bank details changed in the last 30 days.
- Same invoice number as an earlier invoice with a different amount or bank.
- Payee is a person, crypto wallet, gift card, or in a different country from the vendor record.
- Risky file type (macro Office file, executable, script, archive with executable, password-protected archive).
- Hidden text or instructions aimed at an AI reader (deterministic detection, CAP-5).

**Soft signals** (banner only, never block; they add a banner to the approver's screen and a count to the Controller report, behind a config flag): urgency language, unusual hour, first invoice from a new contact at a known vendor, amount far above this vendor's history, round amounts, new remit-to wording, PDF metadata mismatch. Soft signals never lower a hard trigger. A soft-signal invoice is excluded from auto-release (APR-4).

**Billing-platform allowlist.** Many vendors send through platforms (Bill.com, QuickBooks, Coupa, SAP Ariba and similar), so the platform's sending domain is allowed only when mapped to a vendor by platform plus the vendor's account ID in the message. Without the account ID match it is treated as an unlisted sender.

**Forwarded invoices.** An invoice forwarded internally by an employee is treated by the original sender, not the forwarder; the forwarder is recorded. If the original sender cannot be determined, it is an unlisted sender.

**Controller workflow**

- **QUA-1 (v1)** The Controller sees one screen: the rendered invoice, why it triggered, vendor history, and the bank-fingerprint comparison. Actions: Release, Reject, Report.
- **QUA-2 (v1)** Release for sender or bank cases requires a callback to a number from the contract or vendor onboarding file, or the vendor-master phone number unchanged for at least 90 days. The number in the email is never used. A second person confirms every sender/bank release and every release above $10,000.
- **QUA-3 (v1)** After release the invoice is never auto-released and still goes through the full approval chain.
- **QUA-4 (v1)** Report: the message is forwarded to the security mailbox and the sender domain is blocked.
- **QUA-5 (v1)** The Controller has a named backup; a quarantined invoice unreviewed for 2 business days escalates to the CFO.

**Honest limit.** If a genuine vendor's mailbox is compromised and the attacker sends a real-looking invoice with unchanged bank details, no sender check catches it; the controls that do are the sponsor's review of the work received, PO/GR matching, and the bank-change process. The portal does not claim to stop it.

**Shadow-run gate.** Before go-live, the triggers run in observe-only mode for at least 300 invoices; they go live only if quarantine is at or below 3% of invoices and every trigger is explainable. A trigger that quarantines more than 1% on its own is re-tuned.

### Other security controls

- **Posting boundary.** In v1 the portal has no connection to SAP at all. In Stage 2 it can park and post invoices only after a complete signature chain and cannot change payment blocks, vendor bank data, or payment runs. Payment release stays with Treasury in SAP (F110 proposal reviewed there). This replaces the earlier "portal can never pay" claim, which was too strong.
- **PLAT-1 Platform.** Private network, WAF, rate limiting, secrets in a vault, no standing admin keys, MFA for all builder and admin access, patching SLA of 14 days for critical issues, penetration test before go-live and annually.
- **PLAT-2 Isolation.** Parsing VM, SAP adapter, and signing service in separate trust zones (see Architecture); the parsing VM has no outbound network. Hosting: the 90ten.life mailbox, website, Max Dash and the VPS share one Hostinger service today. The AP portal gets its own VPS (not shared with other apps), session cookies are host-only (no Domain=.90ten.life, so Max Dash and website cookies never reach the portal), and the signing service, audit anchors and off-site backups live outside Hostinger in an account the builder does not control, so one Hostinger compromise cannot forge approvals or rewrite history.
- **AUD-1 Audit log.** Insert-only, hash-chained event log; the chain head is anchored every 15 minutes to S3 Object Lock in an account the builder does not control; an off-box verifier checks the chain and the anchors daily and alerts on any gap.
- **DATA-1 Data protection.** Encryption at rest and in transit; TINs and bank fingerprints field-encrypted; original files in an encrypted immutable store; least-privilege roles; production data never copied to development.
- **DATA-2 Email safety.** Display of original emails is text/PNG only; attachments are never downloadable by approvers (AP downloads through the sandbox only).
- **MAIL-1 Domain.** DMARC p=reject on the company domain and on the portal's sending domain.
- **LLM-1 Models.** Any model sees only extracted text and images from the sandbox, has no tools, no network, and no write access to anything.

### Security data additions

- **quarantine\_case**: id, invoice, triggers, state, reviewer, second reviewer, callback source, decision, time.
- **vendor\_domain**: vendor, domain or address, status, added by, approved by, date.
- **billing\_platform\_map**: platform, vendor, account ID, added by, approved by.
- **security\_event**: id, type, user, source, detail, time (also hash-chained into the audit log).

## Data model and SAP import spec

**Stage 2 material.** v1 has no SAP connection (see Decisions). The core tables below apply to v1 except the SAP fields and the adapter, which are Stage 2; duplicate checks in v1 cover portal records only, and AP still relies on SAP's own duplicate warning at entry. RPT-2 and CUT-2/CUT-3 (month-end received-not-entered listing) work in v1 from portal data. The Gate 0 spike, field contract and write-back below are the Stage 2 plan and may change (for example to FI invoice parking) once Meril confirms what access Maxx can have.

The portal owns workflow and evidence; SAP owns the books. The portal reaches SAP only through one adapter interface (D2: SAP is never called from the internet; SAP pulls, or a relay agent inside the SAP network makes the calls). SAP-side function names below must be verified against the customer's release (ECC vs S/4) in the Gate 0 spike before anything else is built.

**Core tables**

| Table | Key contents |
| --- | --- |
| `message` | id, raw store ref, headers, auth results, received time, classification |
| `invoice` | id, state (see lifecycle), message, file hash, vendor\_id, header fields, currency, amounts, PO ref, tags, portal ID, sap\_rbkp\_belnr, sap\_bkpf\_belnr, sap\_posting\_date, paid status (read-only), cutoff flags |
| `invoice_line` | invoice, line no, description, qty, UoM, unit price, amount, GL, cost object, PO item, tax code (PO invoices; non-PO lines are P2) |
| `vendor` | SAP number + read-only subset (name, TIN status, terms, bank fingerprint); portal-only: sponsor, backup, default GL/cost center, po\_required, recurring schedule, tolerance, tags (HCP, ASL, FCPA-agent, clinical, intercompany) |
| `routing_rule` / `matrix_version` | rules, versions, proposer, approver, effective date |
| `approval_signature` | see Sign-in and signature |
| `quarantine_case` | see Security |
| `event_log` | invoice, from\_state, to\_state, actor, actor\_type, payload, prev\_hash, hash. Insert-only, hash-chained, anchored (AUD-1) |

**SAP adapter interface** (the only code that knows SAP): `park(invoice)`, `update_parked(invoice)`, `delete_parked(portal_id)`, `post_parked(portal_id)`, `lookup_by_portal_id(portal_id)`, `read_document(belnr)`, `read_payment_status(belnr)`, vendor master delta (name, terms, TIN status, bank fingerprint, phone), open items and cleared items. Every write is idempotent via the portal ID; before any retry the adapter looks the portal ID up so it never creates a second document.

**Gate 0 spike (weeks 0–3)** acceptance: against a SAP sandbox, (1) park a PO invoice and a non-PO invoice via LIV; (2) post the parked document; (3) delete a parked document; (4) look up by portal ID; (5) read payment status; (6) negative tests: duplicate portal ID rejected, closed period rejected, blocked vendor rejected, one-time vendor rejected, missing GR rejected, tolerance breach parked with block (and the portal does not clear it), user without authorization rejected. If any of 1–6 fails, the vehicle decision (D10) changes before more is built.

**SAP field contract (park via LIV)**

| Portal field | SAP field | Rule |
| --- | --- | --- |
| vendor number | LIFNR | must exist, not blocked, not one-time |
| invoice date | BLDAT | from invoice |
| posting date | BUDAT | portal proposes; SAP period must be open (else Controller decides, CUT-1) |
| reference | XBLNR | invoice number as printed; if over 16 characters, truncated plus a hash suffix; full number kept in the portal |
| portal ID | ZUONR / ALLOC\_NMBR | idempotency key and audit link |
| amount, currency | WRBTR, WAERS | foreign currency at the SAP daily rate; no conversion in the portal |
| tax | MWSKZ | vendor default tax code (US); anything else goes to manual queue |
| PO / item | EBELN, EBELP | PO invoices only; GR-based or SES required |
| G/L, cost center | GLACCOUNTDATA | non-PO only; requires direct G/L posting enabled in LIV customizing |
| withholding | WITHT | 1099 vendors per vendor master |
| payment block | none | the portal never sets or clears payment blocks |
| forbidden fields |  | one-time vendor data, EMPFB, BVTYP, payment method, house bank: never written by the portal |

**Pre-park validations.** Vendor exists and is not blocked; PO open with quantity left; currency matches PO/vendor; arithmetic ties; period open; no unresolved quarantine; signature chain complete and sealed against the current content hash; all required fields present.

**Duplicate control.**

- **DUP-1** Hard block: same normalized invoice number (strip spaces, leading zeros, case) from the same vendor, or from any vendor sharing a TIN or bank fingerprint. Flag only (AP review): same amount and date, same PO and amount, same amount within 30 days. Search the portal and SAP open and cleared items (BSIK/BSAK) for 24 months before parking.

**Write-back.** After posting: store RBKP and BKPF document numbers, posting date, and SAP user. Nightly: read cleared items to update paid status and date. The portal does not store or depend on a payment run ID.

**Cutoff and period-end**

- **CUT-1** If the SAP posting period is closed, the invoice is not posted; the Controller decides (post to the current period, or reopen). The portal never forces a period.
- **CUT-2** On workday +1 after month end, the portal produces a received-not-posted listing (received before cutoff, not yet posted) for the accrual entry; the Controller signs it.
- **CUT-3** Invoices received but unposted at month end are shown with the receive date, so expense can be recognized in the right period.
- **CUT-4** Confirm with the external auditors that park-with-accrual satisfies their cutoff testing before go-live.

**Scope and resubmission.** One company code (D9). Credit memos, Meril intercompany invoices and other company codes are diverted to manual queues and excluded from the KPIs. A rejected or disputed invoice that is resubmitted is a new record linked to the old; the duplicate check treats the earlier one as superseded only when AP marks it so.

## Controls, compliance and audit

The portal adds a layer of workflow control on top of SAP's own authorizations. It does not replace SAP controls, and nobody in the chain (including the Portal Admin and the builder) is assumed incapable of wrongdoing; the design limits what each role can do alone and makes what they do visible.

**Roles**

| Role | Can | Cannot |
| --- | --- | --- |
| AP Clerk | Ingest, correct captured fields, route, request sponsor, park in SAP | Approve, edit vendor master or bank data, release quarantine, run payments |
| AP Supervisor | Everything the clerk can; approve routing-rule proposals; reassign | Sign as approver on same invoice; release quarantine alone |
| Controller | Release/reject quarantine (with second person); approve config changes (second person); cutoff decisions | Be sole approver of an invoice she or he handled in AP; edit vendor bank data |
| Approver (sponsor) | Review and sign invoices routed to them; delegate | Approve own expenses; change routing; see bank data beyond fingerprint |
| CFO | Co-approve above-limit invoices; break-glass mode; escalated quarantine | Edit config alone |
| Compliance | See HCP / FCPA / ASL items; block a vendor; receive Open Payments tags | Approve for payment |
| Treasury | Review the F110 proposal and run payments in SAP | Approve invoices; park invoices |
| Portal Admin | Propose config (matrix, domains, routes); manage user access | Approve own proposals; approve invoices; read the signing key |
| Builder | Deploy code via reviewed pipeline | Standing production access; signing-service account access; sign anything |
| SAP posting user (technical) | Park and post via the adapter | Set payment blocks; edit vendor master; run F110 |
| Master-data staff (in SAP) | Create vendors, change bank data under SAP dual control | Approve invoices for those vendors |

**Segregation of duties**

- **SOD-1** The person who captures or corrects an invoice cannot approve it.
- **SOD-2** The vendor sponsor cannot create or edit that vendor's master data or bank data.
- **SOD-3** The person who changes bank data in SAP cannot release a quarantine or approve invoices for that vendor within 30 days.
- **SOD-4** A configuration proposer cannot approve it (dual control).
- **SOD-5** Treasury (payment run) cannot park or sign invoices.
- **SOD-6** The CEO and CFO are approvers like anyone else; approving their own expenses goes to the other.
- **SOD-7** The SoD matrix is checked automatically each night against SAP and Max Dash role data; conflicts alert the Controller.
- **DEL-1** Nothing is deleted. Corrections, voids and withdrawals are new records; in v1 nothing is deleted at all; in Stage 2 the only deletion allowed is of a parked SAP document before posting, logged with reason.
- **ADM-1** Configuration (approval matrix, domains, thresholds, step-up amount) uses dual control, is versioned, takes effect the next business day, and the matrix version is stored with each approval.

**Bank-change process.** The portal never changes bank data. A bank change is requested by the vendor, verified by callback (QUA-2 rules for the number), entered in SAP by master-data staff under SAP's dual control, and the first payment to the new account is held for about 5 days. The vendor is notified at the previous contact on file. The portal sees the change only as a new fingerprint, which triggers the 30-day quarantine rule.

**Fraud controls.** Quarantine and hard triggers (Security); callback rules; payee must match vendor master; HCP and FCPA-agent tagging; ASL hard stop; daily Controller digest of exceptions; automated SoD check.

**Medical-device-specific hooks**

| Hook | Trigger | Action |
| --- | --- | --- |
| Open Payments (Sunshine Act) | Payee has NPI or HCP tag | Auto-tag at line level with nature of payment; excluded under \~$13 / \~$135 aggregate per year (indexed); annual report due March 31; reconcile reported total to GL. Meril's own payments to US physicians are also reportable and sit outside this portal — flag to Compliance. |
| HCP contracts and FMV | Any HCP payee | Contract on file and in date; rate within FMV band (ortho consulting typically $300–600/hr, med); timesheet or deliverable attached; YTD vs. contract cap. Sales can never sponsor or approve. |
| Anti-Kickback / POD | Distributor or agent payee | Flag if any surgeon holds ownership in the distributor (OIG 2013 POD alert). Commissions should be self-billed from Maxx data, not invoiced in. |
| FCPA (foreign agents, 60+ countries) | FCPA-agent tag | Due diligence and anti-corruption certification current; bank account in agent's own name in its operating country; no third-party payees or offshore accounts; commission >\~15% flagged to Compliance. |
| QMSR / ISO 13485 purchasing controls | Quality-critical vendor tag | Off-ASL supplier → flag, not stop (controls apply at acceptance, not payment). Quality-critical vendor with **no PO** → hard stop, route to QA as potential CAPA. |
| Customs / transfer pricing | Intercompany (Meril) or customs broker | Reconcile declared customs value per entry to Meril invoice transfer price; TP true-ups require CBP reconciliation; HTS classification verified given 2025 India tariff changes (low-med on device treatment). |

**Audit evidence.** For any invoice the portal can show: the original message and file, extracted values and corrections with who made them, the routing rule and matrix version, each signature with its payload and verification result, quarantine decisions with callback source, the SAP document numbers, and the hash-chain verification status. Auditors get a read-only role.

**Retention (RET-1).** Original files, event log and signatures are kept 7 years in write-once storage (Object Lock); TINs are field-encrypted. Privacy and deletion requests are handled by redaction of non-required personal data, never by deleting invoice records.

## Architecture and build options

**Recommendation: hybrid, about 55% confidence (D10).** Buy the commodity layers (capture, mailbox ingestion, workflow shell) if a vendor passes the must-have checklist below; build the thin parts that carry the risk (signing service, SAP adapter, quarantine rules, audit log). Decide at Gate 1 (week 3), after vendor demos. Pure build is the fallback if no vendor passes.

**Vendor must-have checklist.** Any bought component must: support mailbox-scoped ingestion; allow approval via our own signing service (or run approvals outside the vendor tool); export all data and the audit trail; not require SAP to be reachable from the internet; support SSO with Max Dash; provide SOC 2 Type II; permit our isolated parsing or show equivalent sandboxing; keep our data out of model training; give us the hash-chained log or an export we can chain.

**Components and trust zones**

| Component | Job | Trust zone |
| --- | --- | --- |
| Mail ingestor | Reads accounts@90ten.life (fed by a Maxx Microsoft 365 forwarding rule), stores raw messages | Edge (internet-facing, no secrets beyond mailbox scope) |
| Parse sandbox VM | Opens attachments, rasterizes, extracts text | Sandbox (no network, no secrets) |
| Encrypted immutable store | Original messages and files | Data (write-once) |
| Capture API | Header extraction | External or sandbox-fed; no authority |
| Validator | Deterministic checks, hidden-text detection | Application |
| Vendor master and rules | Sponsors, routes, matrix, domains | Application (dual-control writes) |
| Dashboard / API | User interface, queues | Application (Max Dash SSO) |
| WebAuthn service | Passkey enrollment and assertions | Application |
| Signing service + KMS | Verifies assertion, recomputes hash, seals | Separate account and host (builder has no access) |
| SAP adapter | Stage 2 only. Parks, posts, reads | Inside SAP network or SAP-initiated pull (D2) |
| Scheduler / notifications | Reminders, digests | Application |
| Audit log, anchor, verifier | Hash chain, 15-min anchors to Object Lock, daily verification | Separate account, off-box verifier |
| Quarantine queue | Controller review | Application |

**Labor math.** Handling an invoice by hand costs roughly 15 minutes of keying, routing and chasing. At 3,000 invoices a year that is about 750 hours, roughly the price of a commercial AP automation subscription. So buying is worth it only if it also delivers the lag and control goals; the case for building rests on our unusual approval and signature requirements, not on cost.

**Buy vs build at \~3,000–6,000 invoices/yr**

| Option | Annual cost | Confidence | Fit |
| --- | --- | --- | --- |
| Stampli | $18–40k + $5–15k implementation | Med | Closest match to the sponsor-approval workflow; SAP connectors; no HCP/ASL logic |
| Medius | $40–100k + $25–60k implementation | Low-Med | Most SAP-native; heavy for this size |
| Tipalti | $20–50k + payment fees | Low-Med | Built around payment execution Maxx doesn't need if SAP pays |
| Bill.com | $8–25k | Med | Weak SAP integration; poor fit |
| Build | $1–3k APIs/hosting + about 12 wks to a shadow-run-ready build, 18–22 wks hardened, 4–8 hrs/wk upkeep | High on cost, Med on timeline | Owns the compliance layer; owns the maintenance burden |

**Where buying wins**: capture accuracy out of the box, mailbox plumbing, fast start. **Where building wins**: the signing model, SAP specifics, and our quarantine and audit design. A bought tool that insists on its own approval step or on email-link approval fails the checklist.

**Gates**

- **Gate 0 (week 3):** Stage 2 only. The SAP spike passes (Data model section) and Meril confirms write access. v1 does not depend on it.
- **Gate 1 (week 3):** build vs hybrid decision, vendor demos scored against the checklist, and the blocking inputs below answered.
- **Gate 2 (week 12):** shadow-run results meet go-live gates (KPIs section) before any cutover.

**Non-functional requirements.**

- Availability 99.5% in business hours; planned maintenance outside them.
- RPO 15 minutes or better; RTO 8 hours; quarterly restore test, with a restore actually performed.
- Dashboard p95 response under 2 seconds.
- Capacity: 10,000 invoices a year, 50 approvers, 7-year retention.
- If the portal is down, invoices continue to accumulate in the mailbox and are ingested later; AP can still work in SAP directly, and ingestion after an outage preserves original receive time.

## KPIs and measurement

**Primary KPI:** median time from receipt in the inbox to entered in SAP (AP records the SAP document number in the portal), at or below 50% of baseline. Everything else supports it.

**Baseline (measured before build, weeks 0–4).** At least 250 invoices over 4 weeks, measured from the inbox timestamp to SAP's entry timestamp (CPUDT/CPUTM), not the posting date (BUDAT). Touches per invoice come from a time-and-motion study of 100 invoices. The "10 days" figure is a starting belief, not the baseline.

| Measure | Definition | Target |
| --- | --- | --- |
| Receipt to entered, median | Inbox time to CPUDT/CPUTM | ≤ 50% of baseline |
| Receipt to entered, P90 | same | ≤ 7 business days |
| Manual touches, repeat-vendor segment | Human actions per invoice | v1: ≤ 55% of baseline (−45%); −75% needs Stage 2 |
| Manual touches, blended | All eligible invoices | v1: ≤ 60% of baseline (−40%); −55% needs Stage 2 |
| Touchless rate | Invoices reaching routing with no human edit | ≥ 50% |
| Quarantine rate | Invoices quarantined | ≤ 3% |
| Approval time | Notification to signature | reported by approver and sponsor |

**Go-live gates (shadow run).** At least 300 invoices processed in observe-only mode alongside the old process; at least 90% of known-vendor invoices ready to route within 1 business day; header accuracy at least 92%; quarantine at or below 3%.

**Day-30 acceptance.** Thirty days after cutover of the first segment, the Controller signs off if the median, P90, touch and touchless targets above are met.

**Cohort rules.** All invoices count, including slow ones, except those diverted by design (Meril intercompany, credit memos, other company codes: DIVERTED) and confirmed fraud (FRAUD). Quarantined legitimate invoices count.

**Instrumentation.** Every timestamp comes from the event log; a weekly report shows stage-by-stage time and touches per invoice, so a miss is traceable to a stage.

## Rollout, risks and open questions

**Phases** (Stage 1: 14–16 weeks end to end, low-medium confidence; Stage 2 is scheduled separately)

1. **Weeks 0–3: baseline and policy.** Start the 4-week baseline; collect the sponsor for every active vendor; allowed domains and billing-platform map; approval matrix; signing-service spike.
2. **Gate 1 (week 3):** build vs hybrid decision from vendor demos, plus the blocking inputs below.
3. **Weeks 3–10: build.** Ingestion, screening, capture, vendor list, routing, dashboard, sign-in at accounting.90ten.life, signatures, notifications, quarantine, approved queue and CSV.
4. **Weeks 10–12: shadow run.** At least 300 invoices alongside the old process.
5. **Weeks 12–14: cutover by segment.** Recurring vendors first, then the rest. Microsoft 365 sign-in must be live (ID-8).
6. **Weeks 14–16: harden.** Penetration test, restore test, tuning, day-30 acceptance.
7. **Stage 2 (SAP integration)** starts only when Meril confirms write access: SAP spike, then adapter build, roughly 6 to 8 weeks (low confidence).

**Migration rules.** MIG-1 in-flight invoices finish in the old process; MIG-2 the portal ingests only mail received after the cutover timestamp; MIG-3 in v1 the duplicate check covers the portal only, so AP still relies on SAP's own duplicate warning at entry.

**Definition of done for any feature.** Passes the negative-path tests; the event log verifiably shows the run; and the Controller signs off.

**Risks**

| Risk | Likelihood | Impact | Mitigation |
| --- | --- | --- | --- |
| Stage 2: Meril does not grant SAP write access, or LIV is not usable | Medium | High | v1 does not depend on it; fallback is a daily file of approved invoices for AP or Meril to load |
| Low PO discipline, most spend non-PO | Medium | High | Measure at baseline; targets are re-based if confirmed |
| Approvers do not engage (tap fatigue, ignoring emails) | Medium | Medium | One tap per sitting; D+3 manager escalation; D+4 reassign; weekly report |
| Quarantine volume overwhelms the Controller | Medium | Medium | Narrow hard triggers; 3% shadow-run gate; soft signals banner only |
| Compromised genuine vendor mailbox | Low | High | Cannot be fully stopped by the portal; sponsor review, PO/GR matching, bank-change process |
| Signing-key or account compromise | Low | Very high | Separate account the builder does not control; KMS; fail closed; daily verification |
| Builder is a single point of failure | Medium | High | Documentation, second engineer on call, quarterly restore test, no standing admin |
| Max Dash outage | Low | Medium | Break-glass mode (ID-7) |
| Scope creep | High | Medium | v1 scope list in Decisions; changes need Controller and Andy approval |

**Open questions**

Blocking inputs, ranked by how much rework a wrong answer causes:

- [ ] Answered Oct 14: S/4HANA, running on Meril's infrastructure with no Maxx-owned license. Logistics Invoice Verification is believed inactive. Andy is checking whether Maxx can get write access at all. Still open: who at Meril approves an integration user and test client, and which posting route (LIV or FI invoice) Meril will allow (blocks Stage 2)
- [ ] Annual invoice count, and the PO / recurring / intercompany / HCP split? (Everything above assumes 3,000–6,000/yr; the targets move if that's wrong.)
- [ ] Does Maxx have PO discipline today, or is most spend non-PO? This decides whether the project is a portal or a procurement reform.
- [ ] System owner: Lenny Knobbs (lenny.knobbs@maxxortho.com), named by Andy on Oct 14. Still open: is Lenny also the Portal Admin, or is that a separate person (the owner should not approve his own config proposals) (blocks week 3). Backup Controller: Corey Perine (corey.perine@maxxortho.com), named Oct 14.
- [ ] Are surgeon consulting, royalties and distributor commissions currently invoiced in via the shared inbox? If yes, pulling them out is a separate workstream.
- [ ] Does Meril's group close deadline constrain the month-end report timing?
- [ ] Any objection to the hybrid lean (about 55%)? It is decided at Gate 1 (week 3) either way.

* [ ] Annual invoice volume, PO share, billing-platform share, and how concentrated spend is in the top 50 vendors.
* [ ] Stage 2, SAP: is direct G/L posting enabled in LIV; is there a PO release strategy; how timely is goods receipt?
* [ ] Max Dash findings (repo review, Oct 14): users sign on at dashboard.90ten.life with email and password against Supabase Auth. It is not an identity provider (no OIDC/SAML, no tokens other apps can verify), has no MFA, has roles for sales and distributor users but no manager field, and no admin screen to disable a user (a ban can be set in Supabase). 30-day sessions. So SSO means either (A) the AP portal signs in through Microsoft 365 (Entra), which gives MFA, manager and active status for free, with a Max Dash link into the portal; or (B) the AP portal verifies Supabase tokens and we add AP roles, a manager field and a disable flow to Max Dash. Decision (Andy, Oct 14): option B for testing, migrate to A (Microsoft 365) before production. Token checking uses Supabase public-key signing, confirmed enabled (Andy, Oct 14). Andy is the Max Dash admin who creates the AP role and removes users. No backup admin for now (Andy's decision)?
* [ ] Mailbox is Microsoft 365; the portal reads accounts@90ten.life, fed by a forwarding rule (Andy, Oct 14). Still to confirm: Hostinger's daily sending limit and DKIM/DMARC setup for 90ten.life (mail is on Hostinger, so the portal reads by IMAP and sends by SMTP); that the Microsoft 365 rule keeps the original headers (redirect, not forward-as-attachment); and how to identify Maxx's forwarder (IP range or connector).
* [ ] Who owns the KMS / signing account and who holds root on the hosting? It must not be the builder, and it cannot be Hostinger: the signing service needs a separate AWS or Azure account.
* [ ] Do any clinical invoices contain PHI?
* [ ] Stage 2: who in Treasury reviews the F110 proposal?
* [ ] Do you accept Stage 1 touch targets of about −40% blended, with −75% (repeat vendors) and −55% (blended) only after Stage 2?
* [ ] Stage 2: who holds FB60 / MIRO / FB02 / XK02 / FK08 / F110 access today (SoD baseline)?
* [ ] Are HCP payments made through Amex or expense reports as well?
* [ ] Stage 2: do the auditors accept park-with-accrual for cutoff?
