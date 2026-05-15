#!/usr/bin/env bash
set -euo pipefail

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

echo "=== Starting Kaiama 02 at $(date) ==="
node api-mode.js --file "safaari excel/Kaiama 02.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Kaiama 02 at $(date) ==="

echo "=== Starting Moro 01 at $(date) ==="
node api-mode.js --file "safaari excel/Moro 01.csv" --password "Pdp@2026" --concurrency 50
echo "=== Finished Moro 01 at $(date) ==="

echo "=== Starting Moro 02 at $(date) ==="
node api-mode.js --file "safaari excel/Moro 02.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Moro 02 at $(date) ==="

echo "=== Starting Offa 01 at $(date) ==="
node api-mode.js --file "safaari excel/Offa 01.csv" --password "Pdp@2026" --concurrency 50
echo "=== Finished Offa 01 at $(date) ==="

echo "=== Starting Offa 02 at $(date) ==="
node api-mode.js --file "safaari excel/Offa 02.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Offa 02 at $(date) ==="

echo "=== Starting Oke-ero at $(date) ==="
node api-mode.js --file "safaari excel/Oke-ero.csv" --password "Pdp@2026" --concurrency 50
echo "=== Finished Oke-ero at $(date) ==="

echo "=== Starting Oyun 01 at $(date) ==="
node api-mode.js --file "safaari excel/Oyun 01.csv" --password "Pdp@2026" --concurrency 50
echo "=== Finished Oyun 01 at $(date) ==="

echo "=== Starting Oyun 02 at $(date) ==="
node api-mode.js --file "safaari excel/Oyun 02.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Oyun 02 at $(date) ==="

echo "=== Starting Patigi 01 at $(date) ==="
node api-mode.js --file "safaari excel/Patigi 01.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Patigi 01 at $(date) ==="

echo "=== Starting Patigi o2 at $(date) ==="
node api-mode.js --file "safaari excel/Patigi o2.xlsx" --password "Pdp@2026" --concurrency 50
echo "=== Finished Patigi o2 at $(date) ==="

echo "=== Starting Baruten at $(date) ==="
node api-mode.js --file "safaari excel/Baruten.csv" --password "Pdp@2026" --concurrency 50
echo "=== Finished Baruten at $(date) ==="
