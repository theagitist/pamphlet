#!/bin/bash
# Failsafe cron sweep: deletes files older than 12 minutes from storage.
# The 12-minute threshold avoids racing with the 10-minute setTimeout in Node.
#
# Install: crontab -e
#   */5 * * * * /home/theagitist/apps/pamphlet/deploy/cron-sweep.sh
#
STORAGE_ROOT="${PAMPHLET_STORAGE_ROOT:-/dev/shm/pamphlet}"

if [ -d "$STORAGE_ROOT" ]; then
  find "$STORAGE_ROOT" -type f -mmin +12 -delete 2>/dev/null
fi
