#!/usr/bin/env bash
set -euo pipefail

echo "=== Starting Ilorin East 01 at $(date) ==="
node api-mode.js --file "safaari excel/Ilorin East 01.csv" --password "Pdp@2026" --concurrency 50 --start-row 304
echo "=== Finished Ilorin East 01 at $(date) ==="

echo "=== Starting Ilorin East 02 at $(date) ==="
node api-mode.js --file "safaari excel/Ilorin East 02.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Ilorin East 02 at $(date) ==="

echo "=== Starting Ilorin South 02 at $(date) ==="
node api-mode.js --file "safaari excel/Ilorin South 02.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Ilorin South 02 at $(date) ==="

echo "=== Starting Irepodun 01 at $(date) ==="
node api-mode.js --file "safaari excel/Irepodun 01.csv" --password "Pdp@2026" --concurrency 50
echo "=== Finished Irepodun 01 at $(date) ==="

echo "=== Starting Irepodun 02 at $(date) ==="
node api-mode.js --file "safaari excel/Irepodun 02.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Irepodun 02 at $(date) ==="

echo "=== Starting Isin at $(date) ==="
node api-mode.js --file "safaari excel/Isin.csv" --password "Pdp@2026" --concurrency 50
echo "=== Finished Isin at $(date) ==="

echo "=== Starting Kaiama 01 at $(date) ==="
node api-mode.js --file "safaari excel/Kaiama 01.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Kaiama 01 at $(date) ==="
