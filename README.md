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
