#!/bin/bash
# auto-sync.sh — watches /Users/I319202/ba-xmas-event-2026 for changes and
# auto-commits + pushes to origin/main every 30 seconds if anything changed.

REPO_DIR="/Users/I319202/ba-xmas-event-2026"
LOG="$REPO_DIR/.autosync.log"

cd "$REPO_DIR" || { echo "$(date): ERROR — repo dir not found" >> "$LOG"; exit 1; }

echo "$(date): auto-sync started (pid $$)" >> "$LOG"

while true; do
  # Check for any uncommitted changes (tracked or new untracked files)
  CHANGED=$(git status --porcelain 2>/dev/null)

  if [ -n "$CHANGED" ]; then
    TIMESTAMP=$(date '+%Y-%m-%d %H:%M:%S')
    git add -A
    MSG="auto-sync: $TIMESTAMP"
    git commit -m "$MSG" >> "$LOG" 2>&1

    if git push origin main >> "$LOG" 2>&1; then
      echo "[$TIMESTAMP] Pushed to origin/main" >> "$LOG"
    else
      echo "[$TIMESTAMP] Push failed — will retry next cycle" >> "$LOG"
    fi
  fi

  sleep 30
done
