# Maxx Accounting Portal — runbook (Stage 1, week 0 to first sign-in)

This is spec §11 "Week 0" turned into steps with the exact commands. Do them in order; each one
assumes the ones above it are done.

**Assumptions written into every command below**

- The VPS is the shared Hostinger box: `root@72.62.174.193` (the same one that runs
  `dashboard.90ten.life` and `usage.90ten.life`).
- This repo is cloned on it at **`/root/MO-commissions`**. If it is somewhere else, change the `cd`
  lines; nothing else cares.
- Traefik is already running there, attached to the Docker network named `proxy`, with an entrypoint
  called `websecure` and a certificate resolver called `letsencrypt`. That is what the usage and
  MaxxDash compose files use, so it is true today.
- Where a value has to come from a console (a password, a project reference), the command is written
  with a shell variable set on the line above it. Fill the `export` line, then paste the command
  unchanged.

Signing-service and AWS steps live in **`infra/signing-lambda/README.md`** (written separately); this
file only points at it.

---

## 1. DNS: A records for the portal and the two sibling names

Where DNS for `90ten.life` is hosted is not recorded anywhere. Check hPanel → **Domains → 90ten.life →
DNS / Nameservers** first. If the nameservers are Hostinger's, add the records in hPanel → **Domains →
DNS Zone**; otherwise add them wherever the nameservers point.

Add three **A** records, TTL 300 (raise it later):

| Name | Type | Points to |
| --- | --- | --- |
| `accounting` | A | `72.62.174.193` |
| `commissions.accounting` | A | `72.62.174.193` |
| `usage.accounting` | A | `72.62.174.193` |

`mo-commissions.90ten.life` and `usage.90ten.life` already exist and stay (they become 301s later).

Check from anywhere:

```bash
dig +short accounting.90ten.life
dig +short commissions.accounting.90ten.life
dig +short usage.accounting.90ten.life
```

All three must print `72.62.174.193` before step 8, or Traefik's certificate request fails and the
browser shows a certificate warning.

---

## 2. Mailbox: accounts@90ten.life on Hostinger

1. hPanel → **Emails → 90ten.life → Create email account**: name `accounts`, a long generated
   password. Store it in the password manager; it is the IMAP *and* SMTP password (Hostinger has no
   separate app passwords).
2. In the mailbox settings confirm **IMAP is enabled**. Hostinger's defaults: IMAP `imap.hostinger.com`
   port 993 (SSL), SMTP `smtp.hostinger.com` port 465 (SSL). These are already the defaults in
   `.env.accounting.example`.
3. Note the **daily sending limit** shown for the mailbox (spec §4 point 4). Expected portal volume is
   well under 500/day; if the limit shown is lower, plan a second mailbox for digests — nothing in
   step 7 changes, only `EMAIL_FROM`/`SMTP_USER` for the digest sender later.
4. Confirm the mailbox works from a laptop before trusting it on the server:

```bash
# IMAP: should print "* OK" and then a LOGIN OK line
export ACCOUNTS_MAILBOX_PASSWORD='the-mailbox-password-from-hPanel'
printf 'a1 LOGIN accounts@90ten.life %s\r\na2 LIST "" "*"\r\na3 LOGOUT\r\n' "$ACCOUNTS_MAILBOX_PASSWORD" \
  | openssl s_client -quiet -connect imap.hostinger.com:993 2>/dev/null
```

Nothing else sends as `@90ten.life` today (MaxxDash and usage send from `@phillyshah.com`), so this
mailbox is the only sender on the domain.

---

## 3. SPF, DKIM, DMARC for 90ten.life

Run these from anywhere:

```bash
dig +short TXT 90ten.life
dig +short TXT _dmarc.90ten.life
dig +short TXT default._domainkey.90ten.life
```

What you want to see, and what to do if you don't:

- **SPF** (first command): a record starting `v=spf1` that contains `include:_spf.mail.hostinger.com`.
  hPanel → **Emails → 90ten.life → DNS / Email Configuration** shows the exact record to add if it
  is missing. One SPF record only — if there are two `v=spf1` lines, merge them.
- **DKIM** (third command): Hostinger's selector is not always `default`. hPanel → **Emails → 90ten.life
  → DKIM** shows the selector and the `TXT` value; if it says, for example, `hostingermail-a`, the
  check is `dig +short TXT hostingermail-a._domainkey.90ten.life`. Enable DKIM there if it is off.
- **DMARC** (second command): read whatever is there. Do **not** move to `p=reject` yet — first run the
  portal for a week with `p=none` or `p=quarantine` and confirm every message from `accounts@` passes.
  If there is no record at all, add `_dmarc.90ten.life TXT "v=DMARC1; p=none; rua=mailto:accounts@90ten.life"`
  so you at least get reports.

Send one message from the new mailbox (webmail in hPanel) to a Gmail address and open "Show original":
SPF, DKIM and DMARC must each say PASS.

---

## 4. Microsoft 365: redirect the accounting mailbox to accounts@90ten.life

This needs someone with **Exchange admin** on Maxx's tenant (open question in the spec — find out who
before you start).

1. Exchange admin center → **Mail flow → Rules → + Add a rule → Create a new rule**.
2. Name: `AP portal — redirect accounting mailbox`.
3. **Apply this rule if**: *The recipient is…* → the Maxx accounting mailbox.
4. **Do the following**: *Redirect the message to…* → `accounts@90ten.life`.
   It must be **Redirect**. Not "Forward", not "Forward as attachment", not a mailbox forwarding
   setting. Redirect keeps the original `From`, `DKIM-Signature`, `Authentication-Results` and
   `ARC-*` headers; forwarding rewrites them and the portal would quarantine everything.
5. Save, enable, wait ~15 minutes.

**Header test.** Send one email from an outside address (Gmail is fine) to the Maxx accounting
mailbox. When it arrives in `accounts@90ten.life`, open it in Hostinger webmail → *View source* (or
pull it over IMAP) and check:

- the last `Received:` line *before* Hostinger's own names a host ending in
  `.outbound.protection.outlook.com`;
- there is an `Authentication-Results:` header whose first token (the authserv-id) is Maxx's tenant
  domain. That token is what goes in `FORWARDER_AUTHSERV_ID` in step 7 — the example file assumes
  `maxxortho.com`; if the header says something else (for example `maxxortho.onmicrosoft.com`), use
  what the header says;
- `From:` is still the Gmail sender, not the Maxx mailbox.

If any of those is false, the rule is a forward, not a redirect. Fix it before going further.

---

## 5. MaxxDash Supabase: confirm asymmetric signing keys, collect the two identity values

Supabase dashboard → the **MaxxDash** project (the one `dashboard.90ten.life` uses):

1. **Project Settings → JWT Signing Keys** (older UI: *Authentication → JWT*). The current key must be
   **ECC (P-256)** or **RSA**. If it shows the legacy *HS256 shared secret*, click *Rotate / Create
   asymmetric key* and make it current. The dashboard is unaffected — it never verifies tokens
   itself. The portal **cannot** sign anyone in while HS256 is current.
2. **Project Settings → API**: copy *Project URL* and the *anon public* key. These become
   `MAXXDASH_SUPABASE_URL` and `MAXXDASH_SUPABASE_ANON_KEY`. Never the `service_role` key.
3. Check the JWKS endpoint answers and contains a key with `"kty": "EC"` or `"kty": "RSA"`:

```bash
export MAXXDASH_SUPABASE_URL='https://your-maxxdash-project-ref.supabase.co'   # Project Settings -> API -> Project URL
curl -s "$MAXXDASH_SUPABASE_URL/auth/v1/.well-known/jwks.json"
```

An empty `{"keys":[]}` means the project is still on HS256.

4. The MaxxDash-side PR (weeks 1–2 in the spec) adds `public.portal_access` and the admin-panel grant
   toggle. Until it has shipped, the grant row has to be inserted by hand in **step 10**.

---

## 6. New Supabase project `maxx-accounting`

Supabase dashboard → **New project**: name `maxx-accounting`, region **us-east-1** (same as the AWS
side), plan **Pro**, and under *Add-ons* turn on **Point in Time Recovery** (RPO under 15 minutes is a
requirement). Store the database password in the password manager.

Collect:

- **Connect** (top bar) → *URI* under *Session pooler* (port 5432). This is `ACCOUNTING_DATABASE_URL`.
- **Project Settings → Storage → S3 Connection**: the endpoint
  (`https://<ref>.supabase.co/storage/v1/s3`) and *New access key* → id + secret. These are
  `ACCOUNTING_STORE_S3_ENDPOINT`, `ACCOUNTING_STORE_S3_ACCESS_KEY`, `ACCOUNTING_STORE_S3_SECRET_KEY`.
- **Storage → New bucket**, twice: `originals` and `renders`, both **private**.

For the first days you can leave `ACCOUNTING_STORE_BACKEND=local` (files land in the
`accounting-store` Docker volume) and switch to `s3` once the buckets exist; nothing written during
testing needs to survive the switch.

Turn on MFA for your Supabase account now (PLAT-1), not at cutover.

---

## 7. AWS account (Lenny) and the signing service

Follow **`infra/signing-lambda/README.md`**. It covers: Lenny creating the account with root MFA, the
KMS key, the Lambda and its function URL, the DynamoDB passkey registry, the `maxx-ap-anchors` bucket
with Object Lock, and the narrow IAM policy for Andy's user. Out of it you get:

- `SIGNING_SERVICE_URL`, `AWS_REGION`, `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY`, `ANCHOR_BUCKET`.

The portal starts and runs without these; approvals just fail closed (see §15). So this step can run
in parallel with 8–12.

---

## 8. First deploy on the VPS

```bash
ssh root@72.62.174.193
```

Clone (first time only) and set up the environment file:

```bash
cd /root
git clone https://github.com/phillyshah/MO-commissions.git
cd /root/MO-commissions
cp .env.accounting.example .env.accounting
chmod 600 .env.accounting
```

Generate the field-encryption key and put it in the password manager as well as the file:

```bash
python3 -c "import os,base64;print(base64.b64encode(os.urandom(32)).decode())"
```

Edit the file and fill in every blank you collected in steps 2, 5, 6 (and 7 when ready):

```bash
nano /root/MO-commissions/.env.accounting
```

Minimum to boot: `ACCOUNTING_DATABASE_URL`, `ACCOUNTING_FIELD_KEY`, `MAXXDASH_SUPABASE_URL`,
`MAXXDASH_SUPABASE_ANON_KEY`, `IMAP_PASSWORD`, `SMTP_PASSWORD`, `CONTROLLER_EMAIL`.

Confirm the Traefik network exists and the compose file resolves:

```bash
docker network inspect proxy --format '{{.Name}} {{.Driver}}'
cd /root/MO-commissions && make accounting-config > /dev/null && echo "compose OK"
```

Build and start:

```bash
cd /root/MO-commissions && make accounting-deploy
make accounting-ps
make accounting-logs SERVICE=accounting-web
```

`accounting-web` logs `Uvicorn running on http://0.0.0.0:8000`; the first certificate takes up to a
minute. `accounting-worker` will complain about the database until step 9 — expected.

Every later deploy is the same one line: `cd /root/MO-commissions && make accounting-deploy`.

---

## 9. Apply the database migrations

```bash
cd /root/MO-commissions && make accounting-migrate
```

Prints one `applied  0001_core.sql` line per file. Re-running is safe; it prints nothing when there is
nothing pending. `make accounting-migrate-status` lists what is still pending.

Then restart the worker so it stops retrying against an empty schema:

```bash
cd /root/MO-commissions && docker compose -f docker-compose.accounting.yml restart accounting-worker
```

---

## 10. The first admin

Order matters: the **grant** in the MaxxDash project first, then the **role** in the portal database.
The portal looks the person up by their MaxxDash `auth.users.id` (the "subject").

**10a. Find the subject.** Supabase → **MaxxDash** project → *Authentication → Users* → the person →
copy *User UID*. Or in that project's SQL editor:

```sql
select id, email from auth.users where email = 'andylash22@gmail.com';
```

(Use whichever address the admin signs in to MaxxDash with.)

**10b. Grant portal access (MaxxDash project).** Once the MaxxDash PR has shipped, this is the toggle
in the admin panel. Before that, in the **MaxxDash** project's SQL editor, with both UUIDs being the
admin's own id (self-granted, recorded in the spec as the accepted state for now):

```sql
insert into public.portal_access (user_id, portal, granted_by)
values ('00000000-0000-0000-0000-000000000000', 'accounting', '00000000-0000-0000-0000-000000000000')
on conflict do nothing;
```

Replace both zero UUIDs with the subject from 10a. If `portal_access` does not exist yet, the
MaxxDash PR has not been merged; the portal will show a 404 to everyone until it is.

**10c. Give the row the admin and controller roles (portal database).** On the VPS, from inside the
web container so it uses the same connection string as the app:

```bash
export AP_SUBJECT='00000000-0000-0000-0000-000000000000'   # the User UID from 10a
export AP_EMAIL='andylash22@gmail.com'                       # the same person's MaxxDash email
cd /root/MO-commissions && docker compose -f docker-compose.accounting.yml exec \
  -e AP_SUBJECT="$AP_SUBJECT" -e AP_EMAIL="$AP_EMAIL" accounting-web python - <<'PY'
import os
from accounting.db.pool import connection, close_pool
sql = """
insert into ap_user (subject, email, display_name, roles, status)
values (%s, %s, %s, '{admin,controller}'::portal_role[], 'active')
on conflict (subject) do update
   set roles = '{admin,controller}'::portal_role[], status = 'active'
returning id, email, roles;
"""
with connection() as conn:
    print(conn.execute(sql, (os.environ["AP_SUBJECT"], os.environ["AP_EMAIL"], os.environ["AP_EMAIL"])).fetchone())
close_pool()
PY
```

The same statement, for the `maxx-accounting` project's SQL editor if you prefer it there (paste the
two values in place of the quoted ones):

```sql
insert into ap_user (subject, email, display_name, roles, status)
values ('00000000-0000-0000-0000-000000000000', 'andylash22@gmail.com', 'andylash22@gmail.com',
        '{admin,controller}'::portal_role[], 'active')
on conflict (subject) do update
   set roles = '{admin,controller}'::portal_role[], status = 'active';
```

It is safe to run before *or* after the person's first sign-in: sign-in creates the row with no roles
if it is missing and only refreshes email/display name if it exists. Later users get their roles from
the portal's admin screen, not SQL.

---

## 11. Smoke test

From a laptop:

```bash
curl -I https://accounting.90ten.life/healthz
curl -s https://accounting.90ten.life/healthz
curl -I https://accounting.90ten.life/auth/verify
```

Expect `200` with `{"ok":true,"version":"0.1.0"}` on the first two, `302` to `/login` on the third (no cookie =
not signed in; Traefik's ForwardAuth passes that redirect straight to the browser). Also check the response headers
carry `cache-control: no-store` and `strict-transport-security`.

In a browser:

1. Open `https://accounting.90ten.life/` → redirected to `/login`.
2. Sign in with the MaxxDash email and password from step 10. Landing page with tiles.
3. Admin → Users → yourself → *Issue enrollment code*. Then *Passkeys → Enrol*, enter the code, tap
   the device's passkey. Both you and `CONTROLLER_EMAIL` receive an enrollment email from
   `accounts@90ten.life` — that is also the first proof SMTP works.
4. Sign out, sign in again, confirm the passkey is listed.

If sign-in fails with "Unable to sign in" and the worker log shows a JWT error, go back to step 5
(HS256 still current). If the page is a 404 after a correct password, the `portal_access` row is
missing (step 10b).

---

## 12. Run one worker job by hand

```bash
cd /root/MO-commissions && make accounting-worker-once JOB=heartbeat
```

That is `python -m accounting.worker --once heartbeat` in a throwaway worker container with the same
environment. Other job names: `imap`, `pipeline`, `notify`, `reminders`, `anchors`, `verify`,
`sessions`, `digest`, `controller_daily` (the names are listed in `accounting/worker.py`).
`imap` is the one to run after step 4's header test: it should ingest the test message and either
create an invoice row or quarantine it with trigger `direct-sender`.

---

## 13. Backups and the restore test

**What exists without doing anything:** Supabase PITR on `maxx-accounting` (step 6) and, once the
AWS side is up, a chain-head anchor in S3 every 15 minutes.

**Nightly `pg_dump` to the anchors bucket** (spec §10). On the VPS, with the AWS values already in
`.env.accounting`:

```bash
mkdir -p /root/backups/accounting
cat > /root/backups/accounting/dump.sh <<'SH'
#!/bin/bash
set -euo pipefail
set -a; . /root/MO-commissions/.env.accounting; set +a
STAMP=$(date -u +%Y%m%dT%H%M%SZ)
OUT=/root/backups/accounting/maxx-accounting-$STAMP.sql.gz
docker run --rm -e PGCONNECT_TIMEOUT=15 postgres:16 pg_dump --no-owner --no-privileges "$ACCOUNTING_DATABASE_URL" | gzip > "$OUT"
docker run --rm -v /root/backups/accounting:/b -e AWS_ACCESS_KEY_ID -e AWS_SECRET_ACCESS_KEY -e AWS_REGION \
  amazon/aws-cli s3 cp "/b/$(basename "$OUT")" "s3://$ANCHOR_BUCKET/pg_dump/$(basename "$OUT")"
find /root/backups/accounting -name '*.sql.gz' -mtime +7 -delete
SH
chmod 700 /root/backups/accounting/dump.sh
( crontab -l 2>/dev/null; echo '15 3 * * * /root/backups/accounting/dump.sh >> /root/backups/accounting/dump.log 2>&1' ) | crontab -
```

Run it once now and check the log: `/root/backups/accounting/dump.sh && tail /root/backups/accounting/dump.log`.
Andy's IAM user needs `s3:PutObject` on the bucket for the `pg_dump/` prefix too — the README in
`infra/signing-lambda/` is where that policy lives.

**Restore test (quarterly, actually performed).** Into a scratch database, never the live one:

```bash
docker run -d --name ap-restore-test -e POSTGRES_PASSWORD=scratch postgres:16
sleep 5
LATEST=$(ls -t /root/backups/accounting/*.sql.gz | head -1)
gunzip -c "$LATEST" | docker exec -i ap-restore-test psql -U postgres -d postgres -v ON_ERROR_STOP=1
docker exec ap-restore-test psql -U postgres -d postgres -c "select count(*) as invoices from invoice;" \
  -c "select count(*) as events from event_log;"
docker rm -f ap-restore-test
```

Write the date and the two counts into the Controller's daily report thread. That is the evidence
the restore was done, not that the backup exists.

---

## 14. Day-to-day

| Need | Command |
| --- | --- |
| Deploy the latest `main` | `cd /root/MO-commissions && make accounting-deploy` |
| Apply new migrations after a deploy | `make accounting-migrate` |
| Logs, all services | `make accounting-logs` |
| Logs, one service | `make accounting-logs SERVICE=accounting-worker` |
| Status | `make accounting-ps` |
| Restart web+worker (new `.env.accounting` values) | `make accounting-restart` |
| Stop everything (volumes kept) | `make accounting-down` |

A changed `.env.accounting` needs `make accounting-restart` (or `up -d`), not a rebuild.

---

## 15. When the signing service is down

Symptoms: approvers see "The signing service is not reachable. Nothing was lost — your approval was
not recorded; try again later." The worker log shows `signing_service` errors; anchors may also stop.

What is true while it is down:

- **Nothing is signed.** The portal fails closed: no approval, rejection or chain completion is
  recorded without a seal from Lenny's account. There is no local fallback in production.
- **Nothing is lost.** Invoices keep arriving, parsing, routing and sitting in PENDING_APPROVAL. The
  approver re-opens the item and approves again once the service is back.
- **The audit chain is intact**; it is only the off-box *anchor* that is late. The nightly verifier
  in AWS will report the gap between the last anchor and the first one after recovery, which is the
  expected shape of an outage, not tampering.

What to do:

1. `curl -sS -o /dev/null -w '%{http_code}\n' "$SIGNING_SERVICE_URL"` from the VPS (403 = reachable but
   unsigned request, i.e. fine; timeout or 5xx = AWS-side problem).
2. AWS console → Lambda → the function → *Monitor*; CloudWatch logs say why. Only Lenny or someone he
   has given access to can fix an IAM or KMS policy problem — Andy's user cannot.
3. Tell the Controller (`CONTROLLER_EMAIL`) that approvals are paused and why; D+2/D+3 reminders keep
   firing, which is correct.
4. After recovery: `make accounting-worker-once JOB=anchors` then `JOB=verify`, and read the result.

Do not set `DEV_SEAL_PRIVATE_KEY_PEM` on the VPS to "get through the day". The code ignores it outside
test mode, and a seal that is not from Lenny's key is not a seal.

---

## 16. Shared VPS caveats

One Hostinger box runs MaxxDash, usage and now this portal, all under one root account that Andy
administers. The spec (§2 "Trust zones", PLAT-1/2) accepts this **for testing and the shadow run
only**. What that means in practice:

- A root compromise of the VPS exposes `.env.accounting`: the portal DB connection string, the
  mailbox password, the Anthropic key, Andy's AWS access key. It does **not** expose the KMS seal
  key (off-box, non-exportable) or let anyone rewrite anchored history.
- The parser container is a process boundary, not a VM. It has no network, a read-only root, no
  capabilities and hard caps, which is the strongest thing Compose can do on this box. Recorded as a
  known gap.
- Secrets are in a plain file. MFA on Hostinger, Supabase, GitHub and AWS goes on **now**.
- The portal shares Traefik, the `proxy` network and the Docker daemon with the other two apps: a
  mistake in one compose file can take the others down. Run `make accounting-config` before every
  deploy that touched the YAML.

**"Own VPS before cutover"** (PLAT-2) means: before any approval becomes binding, the four portal
services move to a dedicated VPS with its own Traefik, `.env.accounting` moves into Docker secrets or
a vault, and `accounting.90ten.life`'s A record moves with it. MaxxDash and usage stay where they are.
Nothing in this repo has to change for that move except the hostname in the DNS zone; the compose
file is the same.

---

## 17. Moving MO-commissions behind Traefik

Today the statement generator runs bare-metal: `commission-app.service` starts gunicorn on
`127.0.0.1:5001` and `nginx-commissions.conf` fronts it at `mo-commissions.90ten.life` on port 80
(certbot added the 443 block on the box). `deploy.sh` still refers to `/opt/commission-app` and
`commissions.phillyshah.com`; it is stale and is not used.

The `commissions-generator` service in `docker-compose.accounting.yml` replaces all of that: same
`app.py`, LibreOffice inside the image, Traefik terminates TLS, the portal's ForwardAuth gates it, and
the old name 301s to `commissions.accounting.90ten.life`.

**Cutover, on the VPS, after step 11 works:**

```bash
# 1. See what currently owns ports 80/443 — Traefik must have both for the 301 and the certificate.
ss -ltnp | grep -E ':(80|443) '

# 2. Stop and disable the bare-metal unit (the container is already running from step 8).
systemctl disable --now commission-app
systemctl status commission-app --no-pager

# 3. Retire the nginx site. If nginx serves nothing else on this box, stop nginx too so Traefik can
#    bind port 80; if it does serve something else, remove only this site.
rm -f /etc/nginx/sites-enabled/commissions.phillyshah.com /etc/nginx/sites-enabled/nginx-commissions.conf
nginx -t && systemctl reload nginx
# ...or, when nothing else uses nginx:
# systemctl disable --now nginx

# 4. Prove the redirect and the gate.
curl -sI https://mo-commissions.90ten.life/ | grep -iE '^(HTTP|location)'
curl -sI https://commissions.accounting.90ten.life/ | grep -iE '^(HTTP|location)'
```

Expected: the first prints `301` with `location: https://commissions.accounting.90ten.life/`; the
second prints `302` with a `location:` on the portal's login page because there is no portal cookie. Open the
new name in the browser you signed into the portal with: the generator's three-step page appears.

Then delete `/opt/mo-commission-app` (or `/opt/commission-app`, whichever exists) once the new
container has produced one real statement bundle, and remove the `cleanup.sh` crontab line — job
folders now live in the `commissions-outputs` volume. Prune it occasionally:

```bash
cd /root/MO-commissions && docker compose -f docker-compose.accounting.yml exec commissions-generator \
  find /app/outputs -mindepth 1 -maxdepth 1 -type d -mmin +1440 -exec rm -rf {} +
```

**How the sibling hosts know you are signed in.** The portal's own cookie is `__Host-ap_session`, which a
browser sends only to `accounting.90ten.life`. At sign-in the portal also sets a second cookie, `ap_sso`,
scoped to `Domain=accounting.90ten.life`, so `commissions.accounting.*` and `usage.accounting.*` present it
to `/auth/verify`. That cookie opens nothing on the portal itself and is never scoped to `.90ten.life`.
Nothing to decide here; it is in the code (`accounting/identity/sessions.py`).

The same steps apply to usage later: swap the labels on `labels-api` in `/root/usage/docker-compose.yml`
for the block commented at the bottom of `docker-compose.accounting.yml`, then
`cd /root/usage && docker compose up -d`.
