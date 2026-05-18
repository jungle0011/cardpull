const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const zlib = require("zlib");
const XLSX = require("xlsx");
const Papa = require("papaparse");
const pLimit = require("p-limit");
const QRCode = require("qrcode");
const { Agent } = require("undici");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");

const API_BASE_URL = "https://api.pdpnigeria.org";
const APP_BASE_URL = "https://pdpnigeria.org";
const LOGIN_URL = `${API_BASE_URL}/api/auth/login`;
const ME_URL = `${API_BASE_URL}/api/auth/me`;
const PDP_BROWSER_HEADERS = {
  Accept: "application/json, text/plain, */*",
  "Accept-Language": "en-US,en;q=0.9",
  "Accept-Encoding": "gzip, deflate, br",
  Connection: "keep-alive",
  "Cache-Control": "no-cache",
  Pragma: "no-cache",
  DNT: "1",
  Origin: APP_BASE_URL,
  Referer: `${APP_BASE_URL}/login`,
  "User-Agent":
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
  "sec-ch-ua": '"Chromium";v="148", "Google Chrome";v="148", "Not)A;Brand";v="99"',
  "sec-ch-ua-mobile": "?0",
  "sec-ch-ua-platform": '"macOS"',
  "sec-fetch-dest": "empty",
  "sec-fetch-mode": "cors",
  "sec-fetch-site": "same-site",
};
const FETCH_DISPATCHER = new Agent({
  connections: 128,
  keepAliveTimeout: 30000,
  keepAliveMaxTimeout: 120000,
  pipelining: 1,
});
const OUTPUT_DIR = path.join(__dirname, "output");
const FAILED_PATH = path.join(OUTPUT_DIR, "failed.txt");
const RATE_LIMITED_PATH = path.join(OUTPUT_DIR, "rate-limited.txt");
const SKIPPED_PATH = path.join(OUTPUT_DIR, "api-mode-skipped.txt");
const REQUEST_RETRIES = 10;
let requestSpacingMs = 500;
let nextRequestAt = 0;
let rateLimitedUntil = 0;
const CARD_WIDTH = 420;
const CARD_HEIGHT = 264;
const GREEN = rgb(0, 0.54, 0.25);
const RED = rgb(0.9, 0.06, 0.06);
const DARK = rgb(0.12, 0.14, 0.16);
const MUTED = rgb(0.42, 0.46, 0.5);
const LIGHT = rgb(0.95, 0.95, 0.95);
let pdpLogoBytesPromise = null;

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.file || !args.password) {
    printUsage();
    process.exit(1);
  }

  const concurrency = parsePositiveInt(args.concurrency, 50);
  const startRow = parsePositiveInt(args["start-row"], 1);
  requestSpacingMs = parseNonNegativeInt(args["request-delay"], 500);
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const phoneData = await readPhones(args.file, { startRow });
  const phones = phoneData.phones;
  if (phones.length === 0) {
    throw new Error("No valid phone numbers found in the input file.");
  }

  const existing = new Set(fs.readdirSync(OUTPUT_DIR));
  const pending = phones.filter((phone) => !existing.has(`${phone}.pdf`));
  const alreadyDone = phones.length - pending.length;
  const pendingSet = new Set(pending);
  const previouslyFailedSet = intersectSet(readStatusPhones(FAILED_PATH), pendingSet);
  const retryAvailableSet = subtractSet(
    intersectSet(readStatusPhones(RATE_LIMITED_PATH), pendingSet),
    previouslyFailedSet
  );
  const freshPending = pending.length - previouslyFailedSet.size - retryAvailableSet.size;
  const unsavedTotal = pending.length;

  console.log(`Loaded ${phoneData.sourceTotal} unique phone numbers from ${args.file}`);
  console.log(`Zone total: ${phoneData.sourceTotal}`);
  if (startRow > 1) {
    console.log(`Skipped before start row: ${phoneData.skippedBeforeStartRow}`);
  }
  console.log(`In scope this run: ${phones.length}`);
  console.log(`Already done: ${alreadyDone}`);
  console.log(`Unsaved total: ${unsavedTotal}`);
  console.log(`Previously failed in unsaved: ${previouslyFailedSet.size}`);
  console.log(`Retry available in unsaved: ${retryAvailableSet.size}`);
  console.log(`Fresh pending in unsaved: ${freshPending}`);
  console.log(`Queued this run: ${pending.length}`);
  console.log(`Accounted total: ${alreadyDone + unsavedTotal}`);
  console.log(`Concurrency: ${concurrency}`);
  console.log(`Start row: ${startRow}`);
  console.log(`Request spacing: ${requestSpacingMs}ms`);

  let completed = 0;
  let saved = 0;
  let failed = 0;
  let rateLimited = 0;
  const limit = pLimit(concurrency);

  const tasks = pending.map((phone) =>
    limit(async () => {
      try {
        const member = await fetchMember(phone, args.password);
        const outputPath = path.join(OUTPUT_DIR, `${phone}.pdf`);
        await createMemberCardPdf(member, outputPath);
        saved += 1;
        console.log(`[${completed + 1}/${pending.length}] saved ${phone}.pdf`);
      } catch (error) {
        const reason = cleanReason(error.message || String(error));
        if (isRateLimitReason(reason)) {
          rateLimited += 1;
          await fsp.appendFile(RATE_LIMITED_PATH, `${new Date().toISOString()}\t${phone}\t${reason}\n`, "utf8");
          console.log(`[${completed + 1}/${pending.length}] rate-limited ${phone}: rerun will retry`);
        } else {
          failed += 1;
          await fsp.appendFile(FAILED_PATH, `${new Date().toISOString()}\t${phone}\t${reason}\n`, "utf8");
          console.log(`[${completed + 1}/${pending.length}] failed ${phone}: ${reason}`);
        }
      } finally {
        completed += 1;
      }
    })
  );

  await Promise.all(tasks);

  console.log("");
  console.log("API mode complete.");
  console.log(`Zone total: ${phoneData.sourceTotal}`);
  console.log(`In scope this run: ${phones.length}`);
  console.log(`Saved: ${saved}`);
  console.log(`Failed: ${failed}`);
  console.log(`Rate limited, retry on rerun: ${rateLimited}`);
  console.log(`Already done before run: ${alreadyDone}`);
  console.log(`Unsaved total before run: ${unsavedTotal}`);
  console.log(`Previously failed before run: ${previouslyFailedSet.size}`);
  console.log(`Retry available before run: ${retryAvailableSet.size}`);
  console.log(`Fresh pending before run: ${freshPending}`);
  console.log(`Output: ${OUTPUT_DIR}`);
}

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const item = argv[i];
    if (!item.startsWith("--")) continue;
    const key = item.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith("--")) {
      args[key] = true;
    } else {
      args[key] = next;
      i += 1;
    }
  }
  return args;
}

function printUsage() {
  console.log('Usage: node api-mode.js --file "zone.xlsx" --password "Pdp@2026" --concurrency 50 --start-row 304');
}

function parsePositiveInt(value, fallback) {
  const parsed = Number.parseInt(String(value || ""), 10);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function parseNonNegativeInt(value, fallback) {
  const parsed = Number.parseInt(String(value || ""), 10);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : fallback;
}

function readStatusPhones(filePath) {
  if (!fs.existsSync(filePath)) {
    return new Set();
  }

  const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/);
  const phones = new Set();
  for (const line of lines) {
    if (!line.trim()) continue;
    const parts = line.split("\t");
    const phone = (parts[1] || "").trim();
    if (phone) {
      phones.add(phone);
    }
  }
  return phones;
}

function intersectSet(source, filter) {
  const result = new Set();
  for (const value of source) {
    if (filter.has(value)) {
      result.add(value);
    }
  }
  return result;
}

function subtractSet(source, remove) {
  const result = new Set();
  for (const value of source) {
    if (!remove.has(value)) {
      result.add(value);
    }
  }
  return result;
}

async function readPhones(filePath, options = {}) {
  const absolutePath = path.resolve(filePath);
  if (!fs.existsSync(absolutePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const ext = path.extname(absolutePath).toLowerCase();
  const rows = ext === ".csv" ? await readCsvRows(absolutePath) : readExcelRows(absolutePath);
  const phones = [];
  const seen = new Set();
  const sourceSeen = new Set();
  const skipped = [];
  const startRow = options.startRow || 1;
  let memberRowNumber = 0;
  let sourceTotal = 0;
  let skippedBeforeStartRow = 0;

  const phoneColumnIndex = detectPhoneColumnIndex(rows);

  rows.forEach((row, index) => {
    const rowNumber = index + 1;
    const phone = findPhoneInRow(row, phoneColumnIndex);
    if (!phone) {
      skipped.push(`row ${rowNumber}: skipped - no valid phone number`);
      return;
    }
    memberRowNumber += 1;
    if (!sourceSeen.has(phone)) {
      sourceSeen.add(phone);
      sourceTotal += 1;
    }
    if (memberRowNumber < startRow) {
      skippedBeforeStartRow += 1;
      return;
    }
    if (!seen.has(phone)) {
      seen.add(phone);
      phones.push(phone);
    }
  });

  if (skipped.length > 0) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    await fsp.writeFile(SKIPPED_PATH, `${skipped.join("\n")}\n`, "utf8");
    console.log(`Skipped ${skipped.length} rows. Details: ${SKIPPED_PATH}`);
  }

  return {
    phones,
    sourceTotal,
    skippedBeforeStartRow,
  };
}

async function readCsvRows(filePath) {
  const text = await fsp.readFile(filePath, "utf8");
  const parsed = Papa.parse(text, {
    skipEmptyLines: true,
  });
  const realErrors = (parsed.errors || []).filter((error) => error.code !== "UndetectableDelimiter");
  if (realErrors.length) {
    const message = realErrors.map((error) => error.message).join("; ");
    throw new Error(`CSV parse failed: ${message}`);
  }
  return parsed.data;
}

function readExcelRows(filePath) {
  try {
    const workbook = XLSX.readFile(filePath, { cellDates: false });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
      throw new Error("Excel file has no sheets.");
    }
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    });
  } catch (error) {
    const rows = readBrokenXlsxRows(filePath);
    if (rows.length === 0) {
      throw error;
    }
    return rows;
  }
}

function readBrokenXlsxRows(filePath) {
  const parts = extractZipParts(fs.readFileSync(filePath));
  const sheetXml = parts.get("xl/worksheets/sheet1.xml");
  if (!sheetXml) return [];

  const sharedStrings = parseSharedStrings(parts.get("xl/sharedStrings.xml") || "");
  const rows = [];
  const rowPattern = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
  let rowMatch;
  while ((rowMatch = rowPattern.exec(sheetXml))) {
    const row = [];
    const cellPattern = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    let cellMatch;
    while ((cellMatch = cellPattern.exec(rowMatch[1]))) {
      const attrs = cellMatch[1];
      const body = cellMatch[2];
      const ref = readXmlAttr(attrs, "r");
      const cellType = readXmlAttr(attrs, "t");
      const columnIndex = ref ? columnNameToIndex(ref.replace(/\d+/g, "")) : row.length;
      row[columnIndex] = readCellValue(body, cellType, sharedStrings);
    }
    if (!isEmptyRow(row)) rows.push(row);
  }
  return rows;
}

function extractZipParts(buffer) {
  const offsets = [];
  for (let index = 0; index < buffer.length - 4; index += 1) {
    if (buffer.readUInt32LE(index) === 0x04034b50) {
      offsets.push(index);
    }
  }

  const parts = new Map();
  for (let index = 0; index < offsets.length; index += 1) {
    const offset = offsets[index];
    const method = buffer.readUInt16LE(offset + 8);
    const compressedSize = buffer.readUInt32LE(offset + 18);
    const nameLength = buffer.readUInt16LE(offset + 26);
    const extraLength = buffer.readUInt16LE(offset + 28);
    const name = buffer.slice(offset + 30, offset + 30 + nameLength).toString();
    const dataStart = offset + 30 + nameLength + extraLength;
    const dataEnd = compressedSize > 0
      ? dataStart + compressedSize
      : index + 1 < offsets.length ? offsets[index + 1] : buffer.length;
    const compressed = buffer.slice(dataStart, dataEnd);
    const xml = method === 8
      ? zlib.inflateRawSync(compressed, { finishFlush: zlib.constants.Z_SYNC_FLUSH }).toString("utf8")
      : compressed.toString("utf8");
    parts.set(name, xml);
  }
  return parts;
}

function parseSharedStrings(xml) {
  const strings = [];
  const pattern = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let match;
  while ((match = pattern.exec(xml))) {
    const textParts = [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((part) => decodeXml(part[1]));
    strings.push(textParts.join(""));
  }
  return strings;
}

function readCellValue(body, cellType, sharedStrings) {
  if (cellType === "inlineStr") {
    const inline = body.match(/<t\b[^>]*>([\s\S]*?)<\/t>/);
    return inline ? decodeXml(inline[1]) : "";
  }

  const valueMatch = body.match(/<v>([\s\S]*?)<\/v>/);
  if (!valueMatch) return "";
  const value = decodeXml(valueMatch[1]);
  if (cellType === "s") {
    return sharedStrings[Number(value)] || "";
  }
  return value;
}

function readXmlAttr(attrs, name) {
  const match = attrs.match(new RegExp(`${name}="([^"]*)"`));
  return match ? decodeXml(match[1]) : "";
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function columnNameToIndex(name) {
  let index = 0;
  for (const char of name) {
    index = index * 26 + (char.toUpperCase().charCodeAt(0) - 64);
  }
  return Math.max(0, index - 1);
}

function isEmptyRow(row) {
  const cells = Array.isArray(row) ? row : Object.values(row || {});
  return cells.every((value) => String(value ?? "").trim() === "");
}

function detectPhoneColumnIndex(rows) {
  const headerRows = rows.slice(0, 5);
  for (const row of headerRows) {
    const cells = Array.isArray(row) ? row : Object.values(row || {});
    const index = cells.findIndex((cell) => {
      const label = String(cell || "").trim().toLowerCase();
      return ["phone", "phone number", "mobile", "mobile number", "email", "username"].includes(label)
        || label.includes("phone")
        || label.includes("mobile");
    });
    if (index >= 0) return index;
  }
  return -1;
}

function findPhoneInRow(row, phoneColumnIndex = -1) {
  const cells = Array.isArray(row) ? row : Object.values(row || {});
  if (phoneColumnIndex >= 0) {
    return normalizePhone(cells[phoneColumnIndex]);
  }

  for (const cell of cells) {
    const phone = normalizePhone(cell);
    if (phone) return phone;
  }
  return null;
}

function normalizePhone(value) {
  if (value == null) return null;
  let text = String(value).trim();
  if (!text) return null;

  if (/e\+?/i.test(text) && /^[-+]?\d*\.?\d+e\+?\d+$/i.test(text.replace(/\s+/g, ""))) {
    const numeric = Number(text);
    if (Number.isFinite(numeric)) {
      text = numeric.toFixed(0);
    }
  }

  if (/[a-df-z]/i.test(text)) return null;

  let digits = text.replace(/\D/g, "");
  if (digits.startsWith("234") && digits.length === 13) {
    digits = `0${digits.slice(3)}`;
  }
  if (digits.length === 10) {
    digits = `0${digits}`;
  }
  if (digits.length !== 11 || !digits.startsWith("0") || digits.startsWith("00")) {
    return null;
  }
  return digits;
}

async function fetchMember(phone, password) {
  const loginResponse = await fetchWithRetry(LOGIN_URL, {
    method: "POST",
    headers: {
      ...PDP_BROWSER_HEADERS,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ email: phone, password }),
  });

  const loginText = await loginResponse.text();
  const loginJson = parseJson(loginText);
  if (!loginResponse.ok) {
    throw new Error(loginJson?.message || loginJson?.error || `Login failed with HTTP ${loginResponse.status}`);
  }

  const loginMember = extractMember(loginJson);
  if (loginMember) {
    return { ...loginMember, loginPhone: phone };
  }

  const token = extractToken(loginJson);
  if (!token) {
    throw new Error("Login succeeded but no member data or auth token was returned.");
  }

  const profileResponse = await fetchWithRetry(ME_URL, {
    headers: {
      ...PDP_BROWSER_HEADERS,
      Authorization: `Bearer ${token}`,
    },
  });
  const profileText = await profileResponse.text();
  const profileJson = parseJson(profileText);

  if (!profileResponse.ok) {
    throw new Error(profileJson?.message || profileJson?.error || `Profile failed with HTTP ${profileResponse.status}`);
  }

  const member = extractMember(profileJson);
  if (!member || typeof member !== "object") {
    throw new Error("Profile response did not include member data.");
  }
  return { ...member, loginPhone: phone };
}

async function fetchWithRetry(url, options) {
  let lastResponse;
  let lastError;

  for (let attempt = 1; attempt <= REQUEST_RETRIES; attempt += 1) {
    try {
      await waitForRequestSlot();
      const response = await fetch(url, { ...options, dispatcher: FETCH_DISPATCHER });
      if (response.status !== 429 && response.status < 500) {
        return response;
      }
      lastResponse = response;
      const retryAfter = Number.parseInt(response.headers.get("retry-after") || "", 10);
      const waitMs = Number.isInteger(retryAfter)
        ? retryAfter * 1000
        : response.status === 429
          ? Math.min(60000, 5000 * attempt) + Math.floor(Math.random() * 2000)
          : Math.min(15000, 1000 * attempt * attempt) + Math.floor(Math.random() * 1000);
      if (response.status === 429) {
        rateLimitedUntil = Math.max(rateLimitedUntil, Date.now() + waitMs);
      }
      await delay(waitMs);
    } catch (error) {
      lastError = error;
      await delay(Math.min(15000, 1000 * attempt * attempt) + Math.floor(Math.random() * 1000));
    }
  }

  if (lastResponse) {
    return lastResponse;
  }
  throw lastError || new Error(`Request failed: ${url}`);
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForRequestSlot() {
  const now = Date.now();
  const rateLimitWaitMs = Math.max(0, rateLimitedUntil - now);
  if (requestSpacingMs <= 0) {
    if (rateLimitWaitMs > 0) {
      await delay(rateLimitWaitMs);
    }
    return;
  }

  const waitMs = Math.max(0, nextRequestAt - now, rateLimitedUntil - now);
  nextRequestAt = Math.max(now, nextRequestAt, rateLimitedUntil) + requestSpacingMs;
  if (waitMs > 0) {
    await delay(waitMs);
  }
}

function parseJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function extractToken(json) {
  return json?.token
    || json?.access_token
    || json?.accessToken
    || json?.data?.token
    || json?.data?.access_token
    || json?.data?.accessToken;
}

function extractMember(json) {
  const candidates = [
    json?.member,
    json?.data?.member,
    json?.data?.user,
    json?.user,
    json?.data,
    json,
  ];
  return candidates.find(isMemberLike);
}

function isMemberLike(candidate) {
  if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
    return false;
  }
  return Boolean(
    candidate.membershipId
      || candidate.passportPhoto
      || candidate.firstName
      || candidate.lastName
      || candidate.stateOrigin
      || candidate.pollingUnit
  );
}

function isRateLimitReason(reason) {
  return /HTTP 429|rate limit|too many requests/i.test(reason);
}

async function createMemberCardPdf(member, outputPath) {
  const pdfDoc = await PDFDocument.create();
  const fonts = {
    regular: await pdfDoc.embedFont(StandardFonts.Helvetica),
    bold: await pdfDoc.embedFont(StandardFonts.HelveticaBold),
  };

  const photo = await embedDataUrlImage(pdfDoc, member.passportPhoto).catch(() => null);
  const logo = await embedPdpLogo(pdfDoc).catch(() => null);
  const qrDataUrl = await QRCode.toDataURL(verificationUrl(member), {
    margin: 1,
    width: 280,
  });
  const qr = await embedDataUrlImage(pdfDoc, qrDataUrl);

  drawFront(pdfDoc.addPage([CARD_WIDTH, CARD_HEIGHT]), fonts, member, photo, qr, logo);
  drawBack(pdfDoc.addPage([CARD_WIDTH, CARD_HEIGHT]), fonts, member, logo);

  const bytes = await pdfDoc.save();
  await fsp.writeFile(outputPath, bytes);
}

async function embedDataUrlImage(pdfDoc, dataUrl) {
  if (!dataUrl || typeof dataUrl !== "string") {
    throw new Error("Missing image data.");
  }
  const trimmed = dataUrl.trim();
  const remoteBytes = await fetchRemoteImageBytes(trimmed);
  if (remoteBytes) {
    return embedImageBytes(pdfDoc, remoteBytes, "");
  }

  const match = trimmed.match(/^data:([^;]+);base64,(.+)$/);
  const mime = match?.[1]?.toLowerCase() || "";
  const base64 = match ? match[2] : trimmed;
  const bytes = Buffer.from(base64.replace(/\s/g, ""), "base64");
  return embedImageBytes(pdfDoc, bytes, mime);
}

async function fetchRemoteImageBytes(value) {
  const urls = [];
  if (/^https?:\/\//i.test(value)) {
    urls.push(value);
  } else if (value.startsWith("/")) {
    urls.push(`${API_BASE_URL}${value}`, `${APP_BASE_URL}${value}`);
  }
  if (urls.length === 0) return null;

  for (const url of urls) {
    try {
      const response = await fetch(url, {
        dispatcher: FETCH_DISPATCHER,
        headers: {
          ...PDP_BROWSER_HEADERS,
          Accept: "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
          Referer: `${APP_BASE_URL}/dashboard`,
          "sec-fetch-dest": "image",
        },
      });
      if (!response.ok) continue;
      return Buffer.from(await response.arrayBuffer());
    } catch {
      // Try the next possible host for root-relative image paths.
    }
  }

  return null;
}

function embedImageBytes(pdfDoc, bytes, mime) {
  const isPng = mime.includes("png") || bytes.subarray(0, 4).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47]));
  const isJpg = mime.includes("jpeg") || mime.includes("jpg") || bytes.subarray(0, 2).equals(Buffer.from([0xff, 0xd8]));

  if (isPng) {
    return pdfDoc.embedPng(bytes);
  }
  if (isJpg) {
    return pdfDoc.embedJpg(bytes);
  }
  throw new Error("Unsupported image data.");
}

async function embedPdpLogo(pdfDoc) {
  if (!pdpLogoBytesPromise) {
    pdpLogoBytesPromise = fetch(`${APP_BASE_URL}/favicon.png`)
      .then((response) => {
        if (!response.ok) throw new Error(`Logo failed with HTTP ${response.status}`);
        return response.arrayBuffer();
      })
      .then((buffer) => Buffer.from(buffer));
  }
  return pdfDoc.embedPng(await pdpLogoBytesPromise);
}

function drawFront(page, fonts, member, photo, qr, logo) {
  drawCardShell(page);
  drawHeader(page, fonts, "MEMBERSHIP CARD", logo);

  page.drawText("FULL NAME", { x: 122, y: 179, size: 6, font: fonts.bold, color: MUTED });
  page.drawText(pdfText(memberName(member)), { x: 122, y: 166, size: 11, font: fonts.bold, color: DARK });

  drawField(page, fonts, "CARD NUMBER", member.membershipId || "-", 122, 145, 7, GREEN);
  drawField(page, fonts, "DATE OF BIRTH", formatDob(member.dob), 122, 119);
  drawField(page, fonts, "STATE", member.stateOrigin || "-", 122, 93);
  drawField(page, fonts, "LGA", member.lga || "-", 220, 93);
  drawField(page, fonts, "WARD", member.ward || "-", 122, 67);
  drawField(page, fonts, "POLLING UNIT", member.pollingUnit || "-", 220, 67);

  if (photo) {
    page.drawImage(photo, { x: 28, y: 82, width: 78, height: 92 });
  } else {
    page.drawRectangle({ x: 28, y: 82, width: 78, height: 92, borderColor: GREEN, borderWidth: 1, color: LIGHT });
    page.drawText("PHOTO", { x: 49, y: 124, size: 8, font: fonts.bold, color: MUTED });
  }

  page.drawImage(qr, { x: 314, y: 82, width: 72, height: 72 });
  page.drawText("Scan to verify", { x: 327, y: 72, size: 5, font: fonts.regular, color: MUTED });

  drawFooter(page, fonts, `Card ID: ${member.membershipId || "-"}`);
}

function drawBack(page, fonts, member, logo) {
  drawCardShell(page);
  drawHeader(page, fonts, "MEMBERSHIP CARD - BACK", logo);

  page.drawText("Official Authentication Advisory", {
    x: 104,
    y: 137,
    size: 13,
    font: fonts.bold,
    color: RED,
  });

  const advisory = [
    "This card must be authenticated at your respective Ward Office for",
    "confirmation. You are required to present your passport photograph for",
    "the authentication process.",
  ];
  advisory.forEach((line, index) => {
    page.drawText(line, {
      x: 70,
      y: 114 - index * 13,
      size: 8,
      font: fonts.regular,
      color: DARK,
    });
  });

  page.drawRectangle({ x: 16, y: 26, width: 388, height: 24, color: LIGHT });
  drawFooter(page, fonts, `Card ID: ${member.membershipId || "-"}`);
}

function drawCardShell(page) {
  drawFilledRoundedRect(page, 0, 0, CARD_WIDTH, CARD_HEIGHT, 16, GREEN);
  page.drawRectangle({ x: 2, y: 2, width: CARD_WIDTH - 4, height: 208, color: rgb(1, 1, 1) });
}

function drawHeader(page, fonts, subtitle, logo) {
  page.drawRectangle({ x: 2, y: 210, width: CARD_WIDTH - 4, height: 44, color: GREEN });
  if (logo) {
    page.drawImage(logo, { x: 25, y: 217, width: 30, height: 30 });
  } else {
    page.drawCircle({ x: 36, y: 231, size: 14, color: rgb(1, 1, 1), borderColor: RED, borderWidth: 1.5 });
    page.drawText("PDP", { x: 27, y: 226, size: 8, font: fonts.bold, color: GREEN });
  }
  page.drawText("PEOPLES DEMOCRATIC PARTY", { x: 68, y: 235, size: 8, font: fonts.bold, color: rgb(1, 1, 1) });
  page.drawText(subtitle, { x: 68, y: 224, size: 5, font: fonts.bold, color: rgb(1, 1, 1) });
  page.drawText("FEDERAL REPUBLIC OF", { x: 308, y: 235, size: 5, font: fonts.bold, color: rgb(1, 1, 1) });
  page.drawText("NIGERIA", { x: 340, y: 224, size: 9, font: fonts.bold, color: rgb(1, 1, 1) });
}

function drawField(page, fonts, label, value, x, y, size = 7, color = DARK) {
  page.drawText(label, { x, y, size: 5.5, font: fonts.bold, color: MUTED });
  page.drawText(pdfText(value || "-"), { x, y: y - 10, size, font: fonts.bold, color });
}

function drawFooter(page, fonts, text) {
  page.drawText(pdfText(text), {
    x: 165,
    y: 23,
    size: 5.5,
    font: fonts.bold,
    color: GREEN,
  });
}

function drawRoundedRect(page, x, y, width, height, radius, color, borderColor, borderWidth) {
  if (borderWidth && borderColor) {
    drawFilledRoundedRect(page, x, y, width, height, radius, borderColor);
    drawFilledRoundedRect(
      page,
      x + borderWidth,
      y + borderWidth,
      width - borderWidth * 2,
      height - borderWidth * 2,
      Math.max(0, radius - borderWidth),
      color
    );
    return;
  }

  drawFilledRoundedRect(page, x, y, width, height, radius, color);
}

function drawFilledRoundedRect(page, x, y, width, height, radius, color) {
  page.drawRectangle({ x: x + radius, y, width: width - radius * 2, height, color });
  page.drawRectangle({ x, y: y + radius, width, height: height - radius * 2, color });
  page.drawCircle({ x: x + radius, y: y + radius, size: radius, color });
  page.drawCircle({ x: x + width - radius, y: y + radius, size: radius, color });
  page.drawCircle({ x: x + radius, y: y + height - radius, size: radius, color });
  page.drawCircle({ x: x + width - radius, y: y + height - radius, size: radius, color });
}

function memberName(member) {
  return [member.firstName, member.lastName].filter(Boolean).join(" ") || member.loginPhone || "-";
}

function pdfText(value) {
  return String(value || "-")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/[^\x20-\x7E]/g, "")
    .trim() || "-";
}

function verificationUrl(member) {
  const cardId = member.membershipId || member.loginPhone || "";
  return `https://pdpnigeria.org/verify/${encodeURIComponent(cardId)}`;
}

function formatDob(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toISOString().slice(0, 10);
}

function cleanReason(reason) {
  return String(reason || "Unknown error").replace(/\s+/g, " ").trim();
}
