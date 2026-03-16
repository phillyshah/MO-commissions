#!/bin/bash
# Clean up output files older than 24 hours
# Add to crontab: 0 */6 * * * /opt/commission-app/cleanup.sh
find /opt/commission-app/outputs -mindepth 1 -maxdepth 1 -type d -mmin +1440 -exec rm -rf {} +
