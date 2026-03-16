#!/bin/bash
# ═══════════════════════════════════════════════════════════════════════════════
# Maxx Commission App — Deploy to Hostinger Ubuntu VPS
# Run as root: sudo bash deploy.sh
# ═══════════════════════════════════════════════════════════════════════════════

set -e

APP_DIR="/opt/commission-app"
REPO_DIR="$APP_DIR"

echo "═══ 1. Installing system dependencies ═══"
apt-get update
apt-get install -y python3 python3-venv python3-pip nginx libreoffice-calc certbot python3-certbot-nginx

echo "═══ 2. Setting up application directory ═══"
mkdir -p $APP_DIR
mkdir -p /var/log/commission-app

# If deploying from git, clone here:
# git clone https://github.com/YOUR_USER/commission-app.git $APP_DIR
# For now, assumes files are already in $APP_DIR

echo "═══ 3. Creating Python virtual environment ═══"
cd $APP_DIR
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
deactivate

echo "═══ 4. Setting permissions ═══"
chown -R www-data:www-data $APP_DIR
chmod -R 755 $APP_DIR
chown -R www-data:www-data /var/log/commission-app

echo "═══ 5. Installing systemd service ═══"
cp $APP_DIR/commission-app.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable commission-app
systemctl restart commission-app

echo "═══ 6. Configuring Nginx ═══"
cp $APP_DIR/nginx-commissions.conf /etc/nginx/sites-available/commissions.phillyshah.com
ln -sf /etc/nginx/sites-available/commissions.phillyshah.com /etc/nginx/sites-enabled/
nginx -t
systemctl reload nginx

echo "═══ 7. Setting up SSL with Let's Encrypt ═══"
echo "Run this after DNS is pointing to this server:"
echo "  sudo certbot --nginx -d commissions.phillyshah.com"

echo ""
echo "═══ DONE ═══"
echo "App running at http://commissions.phillyshah.com"
echo ""
echo "Useful commands:"
echo "  systemctl status commission-app    # Check status"
echo "  journalctl -u commission-app -f    # View logs"
echo "  systemctl restart commission-app   # Restart"
