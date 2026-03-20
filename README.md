# Maxx Health — Commission Statement Generator

Internal tool for generating distributor commission statements from invoice worksheets.

## What It Does

1. Upload a Commission Worksheet (`.xlsx`) containing:
   - **Invoice List** sheet — all invoices grouped by distributor code
   - **Dist Lookup** sheet — distributor code → name → contact mapping
   - **Trauma** sheet (optional) — trauma-specific invoices

2. The app automatically:
   - Detects the sales month/year from the Invoice List
   - Creates a **Summary** sheet with all distributors and totals
   - Creates **individual distributor tabs** with Maxx branding, logo, formatted headers, footer
   - Generates **individual PDFs** for each distributor (landscape, print-ready)
   - Packages all PDFs into a downloadable **zip file**

3. Download the completed workbook and/or the PDF bundle.

## Tech Stack

- **Python 3 / Flask** — backend processing
- **openpyxl** — Excel file generation
- **LibreOffice Calc** (headless) — PDF conversion
- **Gunicorn** — production WSGI server
- **Nginx** — reverse proxy + SSL termination

## Deployment (Hostinger Ubuntu VPS)

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USER/commission-app.git /opt/commission-app

# 2. Run the deploy script
cd /opt/commission-app
sudo bash deploy.sh

# 3. Point DNS: commissions.phillyshah.com → your VPS IP

# 4. Enable SSL
sudo certbot --nginx -d commissions.phillyshah.com
```

## Local Development

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

## Commission Processor (process_commissions.py)

Standalone script that takes the completed commission workbook (the `.xlsx` produced by the web app) and splits it into one manager-specific workbook each, with distributor subtotals inserted automatically.

```bash
python process_commissions.py <input.xlsx>
```

**What it does:**

- Locates the `Summary` sheet and dynamically finds the header row (no hardcoded positions)
- Strips legacy subtotal rows from prior runs
- Produces one output file per manager: `{Manager}-{input_filename}.xlsx`
- Each output contains only that manager's rows, with a subtotal row (bold, `SUM` formula) inserted after each consecutive distributor group
- Column widths and all cell styles are preserved

**Output files** are written to the same directory as the input file.

## File Structure

```
commission-app/
├── app.py                    # Flask app + all processing logic
├── requirements.txt
├── gunicorn.conf.py          # Production server config
├── commission-app.service    # systemd service
├── nginx-commissions.conf    # Nginx site config
├── deploy.sh                 # One-command server setup
├── cleanup.sh                # Cron job for old file cleanup
├── static/
│   └── maxx_logo.png         # Maxx Health logo
├── templates/
│   └── index.html            # Web UI
├── uploads/                  # Temporary upload storage
└── outputs/                  # Generated files (auto-cleaned)
```
