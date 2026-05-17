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
- Oyun 02B: duplex generated, ready to send. 49 Oyun 02 phones still not represented in downloaded material.
- Moro 01: sent.
- Moro 02A: sent. Moro 02B remaining/retry.
- Ilorin South 02A: sent.
- Ilorin South 02B: sent.
- Ilorin East 02A: sent.
- Ilorin East 02B: sent.
- Baruten A: sent.
- Baruten B: sent.
- Isin: sent.
- Offa 01A: sent. Offa 01B remaining/retry.
- Ifelodun 01: pending.
- Ifelodun 02: pending.
- Kaiama 01A: duplex generated, ready to send. Kaiama 01B remaining/retry.
- Kaiama 02: running locally on this Mac.
- Patigi 01A: duplex generated, ready to send. Patigi 01B remaining/retry.
- Patigi 02: running on sageverselab Codespace.

Uploaded PDF conversion-only work such as ASA1&2, EDU, EKITI, Ilorin East uploaded PDFs, and PDP ID uploaded PDFs are not counted as generated-zone completion.
