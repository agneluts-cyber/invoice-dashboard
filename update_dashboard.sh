#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON="/usr/bin/python3"
LOG="$SCRIPT_DIR/scrape.log"

echo "========================================" >> "$LOG"
echo "[$(date '+%Y-%m-%d %H:%M:%S')] update_dashboard.sh started" >> "$LOG"

cd "$SCRIPT_DIR"

$PYTHON "$SCRIPT_DIR/scrape_invoices.py" >> "$LOG" 2>&1
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Scraper exited with code $EXIT_CODE" >> "$LOG"
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Update completed successfully" >> "$LOG"
fi
