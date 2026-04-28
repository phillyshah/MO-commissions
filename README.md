# Maxx Orthopedics — MO Commission Tools

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
