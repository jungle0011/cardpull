# cardpull

Bulk screenshot automation for PDP membership ID cards.

## Local run

```bash
npm install
npm run install-browsers
npm start
```

Open `http://localhost:3000`, upload `.xlsx`, `.xls`, or `.csv` files, choose the email column, then start processing.

## Bulk settings

```bash
CONCURRENCY=5 npm start
```

`CONCURRENCY` defaults to `5`. Each worker waits 3 seconds between member logins. The app skips members that already have a card file in the output folder, writes failures to `failed.txt`, and saves `progress.json` as the run advances.

## Docker run

```bash
docker build -t cardpull .
docker run --rm -p 3000:3000 cardpull
```

## Render

Create a Render Web Service from this repository and choose Docker as the runtime. Render will use the included `Dockerfile`; the app listens on `PORT`, uses system Chromium at `/usr/bin/chromium`, and writes output to `/tmp/output`.

## Output

Local runs write cards to `output/`, write failed rows to `output/failed.txt`, save progress to `output/progress.json`, and create ZIP downloads under `output/_zips/`.
npm install && npm start

node generate-duplex.js

zip -r oyun-output.zip output
cd /workspaces/cardpull
zip -r ilorin-east02-output.zip output

## Safari Completion Checklist

- Oyun 01: sent.
- Oyun 02A: sent.
- Oyun 02B: sent. Oyun completed and sent.
- Moro 01: sent.
- Moro 02A: sent.
- Moro 02B: sent. Moro 02 completed and sent.
- Ilorin South 02A: sent.
- Ilorin South 02B: sent.
- Ilorin East 02A: sent.
- Ilorin East 02B: sent.
- Ilorin South 01: sent.
- Ilorin East 01: in progress / retry-heavy. Was previously parked because of many invalid logins.
- Ilorin East 03: pending.
- Ilorin East 04: pending.
- Ilorin West 001: pending. Continue from row 1185.
- Ilorin West 003: pending.
- Ilorin West 004A: sent.
- Ilorin West 004B: remaining/retry.
- Ilorin West 005A: sent.
- Ilorin West 005B: sent.
- Ilorin West 006: sent.
- Ilorin West 007A: sent.
- Ilorin West 007B: remaining/retry.
- Baruten A: sent.
- Baruten B: sent.
- Isin: sent.
- Offa 01: completed and sent.
- Offa 02A: sent.
- Offa 02B: sent.
- Ifelodun 01: sent.
- Ifelodun 02: sent.
- Irepodun 01: sent.
- Irepodun 02A: sent.
- Irepodun 02B: not planned for now.
- Kaiama 01A: sent.
- Kaiama 01B: sent. Kaiama 01 completed and sent.
- Kaiama 02: sent.
- Oke-eroA: sent.
- Oke-eroB: sent.
- Patigi 01A: sent.
- Patigi 01B: sent. Patigi 01 completed and sent.
- Patigi 02: sent.

Uploaded PDF conversion-only work such as ASA1&2, EDU, EKITI, Ilorin East uploaded PDFs, and PDP ID uploaded PDFs are not counted as generated-zone completion.
