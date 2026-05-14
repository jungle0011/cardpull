const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const archiver = require("archiver");
const pLimit = require("p-limit");
const { chromium } = require("playwright");
const crypto = require("crypto");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const LOGIN_URL = "https://pdpnigeria.org/login";
const ROOT_DIR = __dirname;
const IS_RENDER = Boolean(process.env.RENDER);
const UPLOAD_DIR = process.env.UPLOAD_DIR || (IS_RENDER ? "/tmp/uploads" : path.join(ROOT_DIR, "uploads"));
const OUTPUT_DIR = process.env.OUTPUT_DIR || (IS_RENDER ? "/tmp/output" : path.join(ROOT_DIR, "output"));
const ZIP_DIR = path.join(OUTPUT_DIR, "_zips");
const RATE_LIMIT_MS = Number(process.env.RATE_LIMIT_MS || 3000);
const MEMBER_TIMEOUT_MS = Number(process.env.MEMBER_TIMEOUT_MS || 60000);
const CONCURRENCY = Number(process.env.CONCURRENCY || 5);
const MAX_RETRIES = Number(process.env.MAX_RETRIES || 3);
const SYSTEM_CHROMIUM_PATH = process.env.CHROMIUM_EXECUTABLE_PATH || "/usr/bin/chromium";
const COMMON_CARD_EXTENSIONS = [".pdf", ".png", ".jpg", ".jpeg", ".webp"];

const uploads = new Map();
const jobs = new Map();

fs.mkdirSync(UPLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });
fs.mkdirSync(ZIP_DIR, { recursive: true });

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, `${Date.now()}-${crypto.randomUUID()}${ext}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: Number(process.env.MAX_UPLOAD_BYTES || 25 * 1024 * 1024) },
});

app.use(express.json());
app.use(express.static(path.join(ROOT_DIR, "public")));

app.post("/api/upload", upload.single("spreadsheet"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Upload an Excel file first." });
    }

    const workbook = readWorkbookRows(req.file.path);
    if (workbook.headers.length === 0) {
      return res.status(400).json({ error: "No header row was found in the first worksheet." });
    }

    const uploadId = crypto.randomUUID();
    uploads.set(uploadId, {
      id: uploadId,
      filePath: req.file.path,
      originalName: req.file.originalname,
      headers: workbook.headers,
      rowCount: workbook.rows.length,
      createdAt: new Date().toISOString(),
    });

    res.json({
      uploadId,
      originalName: req.file.originalname,
      headers: workbook.headers,
      rowCount: workbook.rows.length,
    });
  } catch (error) {
    res.status(400).json({ error: error.message || "Unable to read the uploaded Excel file." });
  }
});

app.post("/api/process", async (req, res) => {
  const { uploadId, emailColumnIndex, passwordColumnIndex, startRow, endRow } = req.body;
  const meta = uploads.get(uploadId);

  if (!meta) {
    return res.status(404).json({ error: "Upload not found. Please upload the Excel file again." });
  }

  const emailIndex = Number(emailColumnIndex);
  const passwordIndex = Number(passwordColumnIndex);
  const validIndexes = new Set(meta.headers.map((header) => header.index));

  if (!validIndexes.has(emailIndex) || !validIndexes.has(passwordIndex)) {
    return res.status(400).json({ error: "Choose valid email and password columns." });
  }

  if (emailIndex === passwordIndex) {
    return res.status(400).json({ error: "Email and password columns must be different." });
  }

  const range = parseRowRange(startRow, endRow, meta.rowCount);
  if (!range.ok) {
    return res.status(400).json({ error: range.error });
  }

  const rangeTotal = range.endRow - range.startRow + 1;
  const jobId = crypto.randomUUID();
  const job = {
    id: jobId,
    uploadId,
    status: "queued",
    current: 0,
    total: rangeTotal,
    success: 0,
    failed: 0,
    message: "Queued...",
    downloadUrl: null,
    startedAt: new Date().toISOString(),
    finishedAt: null,
    error: null,
    alreadyDone: 0,
    remaining: rangeTotal,
    concurrency: CONCURRENCY,
    startRow: range.startRow,
    endRow: range.endRow,
    zipPath: path.join(ZIP_DIR, `${jobId}.zip`),
    outputPath: OUTPUT_DIR,
    failedPath: path.join(OUTPUT_DIR, "failed.txt"),
    progressPath: path.join(OUTPUT_DIR, "progress.json"),
    _failedWriteQueue: Promise.resolve(),
    _progressWriteQueue: Promise.resolve(),
    _lastProgressSaved: 0,
  };

  jobs.set(jobId, job);
  runJob(job, meta, emailIndex, passwordIndex).catch((error) => {
    job.status = "failed";
    job.error = error.message;
    job.message = "Run failed.";
    job.finishedAt = new Date().toISOString();
  });

  res.json({ jobId });
});

app.get("/api/jobs/:jobId", (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) {
    return res.status(404).json({ error: "Job not found." });
  }

  res.json(publicJob(job));
});

app.get("/download/:jobId", (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job || job.status !== "completed" || !job.zipPath || !fs.existsSync(job.zipPath)) {
    return res.status(404).send("ZIP is not ready yet.");
  }

  res.download(job.zipPath, `cardpull-${job.id}.zip`);
});

app.use((error, _req, res, _next) => {
  res.status(400).json({ error: error.message || "Request failed." });
});

app.listen(PORT, () => {
  console.log(`cardpull is running on http://localhost:${PORT}`);
});

function readWorkbookRows(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("The workbook does not contain any worksheets.");
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  });

  if (matrix.length === 0) {
    throw new Error("The first worksheet is empty.");
  }

  const headers = matrix[0].map((value, index) => ({
    index,
    name: normalizeCell(value) || `Column ${index + 1}`,
  }));

  const rows = matrix
    .slice(1)
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""));

  return { headers, rows };
}

function parseRowRange(startRowValue, endRowValue, rowCount) {
  if (rowCount < 1) {
    return { ok: false, error: "The uploaded sheet does not have any member rows." };
  }

  const startRow = Number.parseInt(String(startRowValue || "1"), 10);
  const endRow = endRowValue === "" || endRowValue == null
    ? rowCount
    : Number.parseInt(String(endRowValue), 10);

  if (!Number.isInteger(startRow) || startRow < 1) {
    return { ok: false, error: "Start Row must be 1 or higher." };
  }

  if (!Number.isInteger(endRow) || endRow < 1) {
    return { ok: false, error: "End Row must be 1 or higher." };
  }

  if (startRow > rowCount) {
    return { ok: false, error: `Start Row cannot be greater than ${rowCount}.` };
  }

  if (endRow > rowCount) {
    return { ok: false, error: `End Row cannot be greater than ${rowCount}.` };
  }

  if (startRow > endRow) {
    return { ok: false, error: "Start Row cannot be greater than End Row." };
  }

  return { ok: true, startRow, endRow };
}

async function runJob(job, meta, emailIndex, passwordIndex) {
  await fsp.mkdir(job.outputPath, { recursive: true });
  await fsp.mkdir(ZIP_DIR, { recursive: true });

  const { rows } = readWorkbookRows(meta.filePath);
  const selectedRows = rows.slice(job.startRow - 1, job.endRow);
  const members = selectedRows.map((row, index) => ({
    rowNumber: job.startRow + index + 1,
    email: normalizeCell(row[emailIndex]),
    password: normalizeCell(row[passwordIndex]),
  }));

  job.total = members.length;
  job.remaining = members.length;
  job.status = "running";
  job.message = `0 already done, processing remaining ${job.remaining}...`;
  await saveProgress(job, true);

  const limit = pLimit(Math.max(1, CONCURRENCY));
  const tasks = members.map((member) =>
    limit(async () => {
      let usedBrowser = false;

      try {
        if (!member.email || !member.password) {
          await recordFailure(job, member.email || `row-${member.rowNumber}`, "Missing email or password.");
          job.failed += 1;
          return;
        }

        const existingFile = await findExistingCardFile(job.outputPath, member.email);
        if (existingFile) {
          job.alreadyDone += 1;
          return;
        }

        usedBrowser = true;
        await processMemberWithRetries(member, job.outputPath);
        job.success += 1;
      } catch (error) {
        await recordFailure(job, member.email || `row-${member.rowNumber}`, error.message);
        job.failed += 1;
      } finally {
        job.current += 1;
        job.remaining = Math.max(job.total - job.current, 0);
        updateJobMessage(job);
        await saveProgress(job);

        if (usedBrowser) {
          await delay(RATE_LIMIT_MS);
        }
      }
    })
  );

  await Promise.all(tasks);
  await Promise.all([job._failedWriteQueue, job._progressWriteQueue]);
  await saveProgress(job, true);

  job.status = "zipping";
  job.message = "Creating ZIP...";
  await saveProgress(job, true);
  await zipFolder(job.outputPath, job.zipPath);
  job.status = "completed";
  job.message = `Completed ${job.total} members. ${job.alreadyDone} already done, ${job.success} downloaded, ${job.failed} failed.`;
  job.downloadUrl = `/download/${job.id}`;
  job.finishedAt = new Date().toISOString();
  await saveProgress(job, true);
}

async function processMemberWithRetries(member, outputPath) {
  let lastError;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt += 1) {
    try {
      await processMember(member.email, member.password, outputPath, member.rowNumber);
      return;
    } catch (error) {
      lastError = error;
      if (attempt < MAX_RETRIES) {
        await delay(1000 * attempt);
      }
    }
  }

  throw new Error(`Failed after ${MAX_RETRIES} attempts: ${lastError.message}`);
}

async function processMember(email, password, outputPath, rowNumber) {
  let browser;
  let context;
  let timedOut = false;
  const timeout = setTimeout(() => {
    timedOut = true;
    if (context) {
      context.close().catch(() => {});
    }
    if (browser) {
      browser.close().catch(() => {});
    }
  }, MEMBER_TIMEOUT_MS);

  try {
    browser = await chromium.launch(getChromiumLaunchOptions());
    context = await browser.newContext({
      viewport: { width: 1440, height: 1100 },
      acceptDownloads: true,
    });
    const page = await context.newPage();
    page.setDefaultTimeout(MEMBER_TIMEOUT_MS);

    await page.goto(LOGIN_URL, { waitUntil: "domcontentloaded", timeout: MEMBER_TIMEOUT_MS });
    await fillLoginForm(page, email, password);
    await clickLogin(page);

    const idCardButton = page
      .getByRole("button", { name: /generate\s*\/?\s*download membership id card/i })
      .first();
    const fallbackIdCardButton = page
      .locator("button", { hasText: /generate\s*\/?\s*download membership id card/i })
      .first();

    await Promise.race([
      idCardButton.waitFor({ state: "visible", timeout: MEMBER_TIMEOUT_MS }),
      fallbackIdCardButton.waitFor({ state: "visible", timeout: MEMBER_TIMEOUT_MS }),
    ]).catch(async () => {
      const loginError = await visibleText(page, /invalid|incorrect|failed|error|unauthorized/i);
      throw new Error(loginError || "Login failed or ID Card button was not found.");
    });

    if (await idCardButton.isVisible().catch(() => false)) {
      await idCardButton.click();
    } else {
      await fallbackIdCardButton.click();
    }

    await waitForCardModal(page);
    await downloadIdCard(page, email, outputPath, rowNumber);
  } finally {
    clearTimeout(timeout);
    if (context) {
      await context.close().catch(() => {});
    }
    if (browser) {
      await browser.close().catch(() => {});
    }
    if (timedOut) {
      throw new Error(`Member timed out after ${MEMBER_TIMEOUT_MS / 1000} seconds.`);
    }
  }
}

async function fillLoginForm(page, email, password) {
  const emailInput = page.locator("input[type='email']").first();
  const fallbackInput = page.locator("input").first();
  const passwordInput = page.locator("input[type='password']").first();

  if (await emailInput.isVisible().catch(() => false)) {
    await emailInput.fill(email);
  } else {
    await fallbackInput.waitFor({ state: "visible" });
    await fallbackInput.fill(email);
  }

  await passwordInput.waitFor({ state: "visible" });
  await passwordInput.fill(password);
}

async function clickLogin(page) {
  const roleButton = page.getByRole("button", { name: /login/i }).first();
  const textButton = page.locator("button", { hasText: /login/i }).first();

  if (await roleButton.isVisible().catch(() => false)) {
    await roleButton.click();
  } else {
    await textButton.waitFor({ state: "visible" });
    await textButton.click();
  }

  await page.waitForLoadState("networkidle", { timeout: 15000 }).catch(() => {});
}

async function waitForCardModal(page) {
  await page
    .locator(".modal.show, [role='dialog'], .modal, [class*='modal']")
    .first()
    .waitFor({ state: "visible", timeout: 30000 })
    .catch(() => {});

  await page
    .locator('button:has-text("Download ID Card")')
    .first()
    .waitFor({ state: "visible", timeout: 30000 });
}

async function downloadIdCard(page, email, outputDir, rowNumber) {
  const downloadButtonSelector = 'button:has-text("Download ID Card")';
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: MEMBER_TIMEOUT_MS }),
    page.click(downloadButtonSelector),
  ]);
  const ext = path.extname(download.suggestedFilename()) || ".png";
  const fileName = `${safeFileName(email) || `row-${rowNumber}`}${ext}`;
  await download.saveAs(path.join(outputDir, fileName));
}

async function visibleText(page, pattern) {
  return page.evaluate((source) => {
    const regex = new RegExp(source, "i");
    const elements = Array.from(document.body.querySelectorAll("*"));
    const match = elements.find((element) => {
      const rect = element.getBoundingClientRect();
      const style = window.getComputedStyle(element);
      return (
        rect.width > 0 &&
        rect.height > 0 &&
        style.display !== "none" &&
        style.visibility !== "hidden" &&
        regex.test(element.textContent || "")
      );
    });
    return match ? match.textContent.trim().replace(/\s+/g, " ").slice(0, 200) : "";
  }, pattern.source);
}

function zipFolder(sourceDir, destinationPath) {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(destinationPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", resolve);
    archive.on("error", reject);
    archive.pipe(output);
    archive.glob("**/*", {
      cwd: sourceDir,
      dot: true,
      ignore: ["_zips/**", "*.zip"],
    });
    archive.finalize();
  });
}

function getChromiumLaunchOptions() {
  const launchOptions = {
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  };

  if (fs.existsSync(SYSTEM_CHROMIUM_PATH)) {
    launchOptions.executablePath = SYSTEM_CHROMIUM_PATH;
  }

  return launchOptions;
}

async function findExistingCardFile(outputPath, email) {
  const baseName = safeFileName(email);
  if (!baseName) {
    return null;
  }

  for (const ext of COMMON_CARD_EXTENSIONS) {
    const filePath = path.join(outputPath, `${baseName}${ext}`);
    if (await pathExists(filePath)) {
      return filePath;
    }
  }

  return null;
}

async function pathExists(filePath) {
  try {
    await fsp.access(filePath);
    return true;
  } catch (_error) {
    return false;
  }
}

async function recordFailure(job, email, reason) {
  const cleanReason = String(reason || "Unknown error").replace(/\s+/g, " ").trim();
  job._failedWriteQueue = job._failedWriteQueue.catch(() => {}).then(() =>
    fsp.appendFile(job.failedPath, `${new Date().toISOString()}\t${email}\t${cleanReason}\n`, "utf8")
  );
  await job._failedWriteQueue;
}

async function saveProgress(job, force = false) {
  if (!force && (job.current === job._lastProgressSaved || job.current % 10 !== 0)) {
    return;
  }

  job._lastProgressSaved = job.current;
  const snapshot = {
    ...publicJob(job),
    outputPath: job.outputPath,
    failedPath: job.failedPath,
    progressPath: job.progressPath,
    updatedAt: new Date().toISOString(),
  };

  job._progressWriteQueue = job._progressWriteQueue.catch(() => {}).then(() =>
    fsp.writeFile(job.progressPath, JSON.stringify(snapshot, null, 2), "utf8")
  );
  await job._progressWriteQueue;
}

function updateJobMessage(job) {
  job.message = `${job.alreadyDone} already done, processing remaining ${job.remaining}... Processing ${job.current} of ${job.total}.`;
}

function publicJob(job) {
  return {
    id: job.id,
    status: job.status,
    current: job.current,
    total: job.total,
    success: job.success,
    failed: job.failed,
    alreadyDone: job.alreadyDone,
    remaining: job.remaining,
    concurrency: job.concurrency,
    startRow: job.startRow,
    endRow: job.endRow,
    message: job.message,
    downloadUrl: job.downloadUrl,
    startedAt: job.startedAt,
    finishedAt: job.finishedAt,
    error: job.error,
  };
}

function safeFileName(value) {
  return String(value)
    .trim()
    .replace(/[/\\?%*:|"<>]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 180);
}

function normalizeCell(value) {
  return String(value ?? "").trim();
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
