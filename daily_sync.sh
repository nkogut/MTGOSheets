#!/bin/bash

# Syncs daily MTGO card prices and updates tracker spreadsheets
# Schedule this with a few minutes delay from login
# Update PYEXE to point to your python executable if needed

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export PYTHONPATH="$DIR/src"
export PYEXE="$DIR/venv/bin/python"

MONTH_DIR=$(date +%Y_%m)
DATE_STR=$(date +%Y_%m_%d)

LOG_DIR="$DIR/logs/$MONTH_DIR"
LOG_FILE="$LOG_DIR/update_prices_log_$DATE_STR.txt"

mkdir -p "$LOG_DIR"

# Only run if there is no log for today yet
if [ ! -f "$LOG_FILE" ]; then
    START_TIME=$(date +"%T")
    echo "Begin updating" >> "$LOG_FILE"
    echo "Start time: $START_TIME" >> "$LOG_FILE"

    "$PYEXE" -u -m mtgo_sheets.updater >> "$LOG_FILE" 2>&1
    "$PYEXE" -u -m mtgo_sheets.tracker >> "$LOG_FILE" 2>&1

    echo "Updating complete" >> "$LOG_FILE"
    END_TIME=$(date +"%T")
    echo "End time: $END_TIME" >> "$LOG_FILE"
fi
