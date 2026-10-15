# Maxx Orthopedics — MO Commission Tools

## This repo now holds two apps

**MO Commission Tools** (`app.py`, `process_commissions.py`) — the Flask statement generator
described in the rest of this file. It is unchanged. It can still run bare-metal under
`commission-app.service`, and it can now also run as the `commissions-generator` container in
`docker-compose.accounting.yml` (built by `Dockerfile.commissions`), where it sits behind Traefik at
`commissions.accounting.90ten.life` and behind the accounting portal's sign-in.

**Maxx Accounting Portal, Stage 1** (the `accounting/` Python package) — invoice intake from
`accounts@90ten.life`, screening, vendor matching, approval routing, passkey-signed approvals and an
approved queue for AP. FastAPI + Jinja, Postgres in the `maxx-accounting` Supabase project, identity
from the MaxxDash Supabase project, a no-network parser container, and a signing service in an AWS
account Lenny owns. Runs from `docker-compose.accounting.yml` with `make accounting-deploy`.

- Setting it up and running it: [docs/ACCOUNTING_RUNBOOK.md](docs/ACCOUNTING_RUNBOOK.md)
- How the modules fit together: [accounting/CONTRACTS.md](accounting/CONTRACTS.md)
- Environment variables: [.env.accounting.example](.env.accounting.example)

---

Internal tool for generating distributor commission statements and manager reports.

**Live at:** https://mo-commissions.90ten.life

## What It Does

Three-step workflow:

1. **Manager Split** — Upload the Summary .xlsx → one workbook + PDF per manager
2. **Distributor Tab Generator** — Upload the Summary .xlsx → one tab per distributor + Summary sheet
3. **PDF Generator** — Upload the Step 2 workbook → one PDF per distributor, bundled as ZIP

## Input File Requirements

The Summary .xlsx must contain:
- **MasterLog** sheet — header row with Manager, Hospital, Distrib Code, Comm $, PO. Pay Date and Title above header.
- **Surgeon Lookup** sheet — Distrib Code, Distributor, Contact, Vendor Code (last column)
- **Template** sheet — branded commission statement with placeholders

## Tech Stack

- **Python 3 / Flask** — backend processing
- **openpyxl** — Excel file generation
- **LibreOffice Calc** (headless) — PDF conversion
- **Gunicorn** — production WSGI server
- **Nginx** — reverse proxy + SSL termination (Let's Encrypt)

## Deployment (Hostinger VPS)

```bash
cd /opt/mo-commission-app
git pull origin main
sudo systemctl restart mo-commission-app
```

## Local Development

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
# Open http://localhost:5001
```

## Versioning

Version tracked in `version.py`. Displayed in the site footer.
- Increment by 0.1 for minor changes
- Increment by 1.0 for major changes

## File Structure

```
MO-commissions/
├── app.py                    # Flask app + all processing logic
├── process_commissions.py    # Shared helpers (Step 1 & 2 core logic)
├── version.py                # App version number
├── requirements.txt
├── gunicorn.conf.py          # Production server config
├── mo-commission-app.service # systemd service
├── nginx-commissions.conf    # Nginx site config
├── static/
│   └── maxx_logo.png         # Maxx Orthopedics logo
├── templates/
│   └── index.html            # Web UI (single page, 3 steps)
├── uploads/                  # Temporary upload storage
└── outputs/                  # Generated files (job-based, auto-cleaned)
```
