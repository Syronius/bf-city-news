#!/bin/sh
set -e

mkdir -p /app/data

# Seed default files into the volume on first run only (never overwrite
# existing data, so scraped articles and edited city lists survive redeploys).
for f in cities.json articles.db; do
  if [ ! -e "/app/data/$f" ] && [ -e "/app/seed/$f" ]; then
    cp "/app/seed/$f" "/app/data/$f"
  fi
done

exec python app.py
