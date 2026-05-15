const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const archiver = require("archiver");
const pLimit = require("p-limit");
const { PDFDocument } = require("pdf-lib");
const { chromium } = require("playwright");
const crypto = require("crypto");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const LOGIN_URL = "https://pdpnigeria.org/login";
const CONNECTIVITY_URL = "https://pdpnigeria.org";
const ROOT_DIR = __dirname;
const IS_RENDER = Boolean(process.env.RENDER);
const UPLOAD_DIR = process.env.UPLOAD_DIR || (IS_RENDER ? "/tmp/uploads" : path.join(ROOT_DIR, "uploads"));
const OUTPUT_DIR = process.env.OUTPUT_DIR || (IS_RENDER ? "/tmp/output" : path.join(ROOT_DIR, "output"));
const ZIP_DIR = path.join(OUTPUT_DIR, "_zips");
const RATE_LIMIT_MS = Number(process.env.RATE_LIMIT_MS || 3000);
const MEMBER_TIMEOUT_MS = Number(process.env.MEMBER_TIMEOUT_MS || 60000);
const MAX_TIMEOUT_RETRIES = 2;
const RETRY_DELAY_MS = 5000;
const WORKER_STAGGER_MS = 500;
const CONNECTIVITY_TIMEOUT_MS = 10000;
const NETWORK_PAUSE_THRESHOLD = 10;
const DEFAULT_CONCURRENCY = clampNumber(Number(process.env.CONCURRENCY || 8), 3, 25);
const COMMON_CARD_EXTENSIONS = [".pdf", ".png", ".jpg", ".jpeg", ".webp"];
const PRINT_READY_FILE_NAME = "all-cards-print-ready.pdf";
const DUPLEX_PRINT_FILE_NAME = "duplex-print-ready.pdf";
const A4_WIDTH = 595;
const A4_HEIGHT = 842;
const A4_HALF_HEIGHT = A4_HEIGHT / 2;

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

app.get("/health", (_req, res) => {
  res.status(200).json({ status: "ok" });
});

app.post("/api/clear-output", async (_req, res) => {
  try {
    await clearOutputFolder();
    jobs.clear();
    res.json({ message: "Output cleared, ready for fresh run" });
  } catch (error) {
    res.status(500).json({ error: error.message || "Could not clear output." });
  }
});

app.post("/api/upload", upload.array("spreadsheet"), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: "Upload at least one Excel or CSV file first." });
    }

    const parsedFiles = [];
    const fileWarnings = [];

    for (const file of req.files) {
      try {
        const workbook = readWorkbookRows(file.path);
        if (workbook.headers.length === 0 || workbook.rows.length === 0) {
          fileWarnings.push(`${file.originalname}: no usable member rows found`);
          continue;
        }

        parsedFiles.push({
          id: crypto.randomUUID(),
          filePath: file.path,
          originalName: file.originalname,
          headers: workbook.headers,
          rowCount: workbook.rows.length,
        });
      } catch (error) {
        fileWarnings.push(`${file.originalname}: ${error.message}`);
      }
    }

    if (parsedFiles.length === 0) {
      return res.status(400).json({
        error: `No uploaded files could be read. ${fileWarnings.join(" ")}`.trim(),
      });
    }

    if (parsedFiles[0].headers.length === 0) {
      return res.status(400).json({ error: "No header row was found in the first worksheet." });
    }

    const uploadId = crypto.randomUUID();
    uploads.set(uploadId, {
      id: uploadId,
      files: parsedFiles,
      headers: parsedFiles[0].headers,
      rowCount: parsedFiles.reduce((total, file) => total + file.rowCount, 0),
      createdAt: new Date().toISOString(),
    });

    res.json({
      uploadId,
      headers: parsedFiles[0].headers,
      rowCount: parsedFiles.reduce((total, file) => total + file.rowCount, 0),
      files: parsedFiles.map((file) => ({
        id: file.id,
        originalName: file.originalName,
        rowCount: file.rowCount,
      })),
      warnings: fileWarnings,
    });
  } catch (error) {
    res.status(400).json({ error: error.message || "Unable to read the uploaded Excel or CSV file." });
  }
});

app.post("/api/preview", (req, res) => {
  const { uploadId, emailColumnIndex, fileRanges } = req.body;
  const meta = uploads.get(uploadId);

  if (!meta) {
    return res.status(404).json({ error: "Upload not found. Please upload the Excel or CSV files again." });
  }

  const emailIndex = Number(emailColumnIndex);
  const validIndexes = new Set(meta.headers.map((header) => header.index));
  if (!validIndexes.has(emailIndex)) {
    return res.status(400).json({ error: "Choose a valid email column." });
  }

  const prepared = prepareMembers(meta, emailIndex, "", fileRanges);
  if (!prepared.ok) {
    return res.status(400).json({ error: prepared.error });
  }

  res.json({
    total: prepared.members.length,
    fileCount: prepared.fileCount,
    warnings: prepared.warnings,
  });
});

app.post("/api/process", async (req, res) => {
  const { uploadId, emailColumnIndex, sharedPassword, fileRanges, concurrency } = req.body;
  const meta = uploads.get(uploadId);

  if (!meta) {
    return res.status(404).json({ error: "Upload not found. Please upload the Excel or CSV files again." });
  }

  const emailIndex = Number(emailColumnIndex);
  const password = normalizeCell(sharedPassword);
  const validIndexes = new Set(meta.headers.map((header) => header.index));

  if (!validIndexes.has(emailIndex)) {
    return res.status(400).json({ error: "Choose a valid email column." });
  }

  if (!password) {
    return res.status(400).json({ error: "Enter the shared password." });
  }

  const prepared = prepareMembers(meta, emailIndex, password, fileRanges);
  if (!prepared.ok) {
    return res.status(400).json({ error: prepared.error });
  }

  const requestedConcurrency = parseConcurrency(concurrency);
  if (!requestedConcurrency.ok) {
    return res.status(400).json({ error: requestedConcurrency.error });
  }

  const connectivity = await checkSiteConnectivity();
  if (!connectivity.ok) {
    return res.status(400).json({
      error: "Cannot reach pdpnigeria.org - check your internet connection before starting",
    });
  }

  const rangeTotal = prepared.members.length;
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
    concurrency: requestedConcurrency.value,
    fileRanges: prepared.fileRanges,
    fileCount: prepared.fileCount,
    parseWarnings: prepared.warnings,
    preparedMembers: prepared.members,
    zipPath: path.join(ZIP_DIR, `${jobId}.zip`),
    printReadyPath: path.join(OUTPUT_DIR, PRINT_READY_FILE_NAME),
    printReadyUrl: null,
    duplexPrintPath: path.join(OUTPUT_DIR, DUPLEX_PRINT_FILE_NAME),
    duplexPrintUrl: null,
    outputPath: OUTPUT_DIR,
    failedPath: path.join(ZIP_DIR, `${jobId}-failed.txt`),
    progressPath: path.join(OUTPUT_DIR, "progress.json"),
    outputFiles: [],
    failedList: [],
    networkRetryQueue: [],
    networkRetryList: [],
    networkRetrySet: new Set(),
    networkErrors: 0,
    consecutiveNetworkErrors: 0,
    pausedReason: null,
    _resumeWaiters: [],
    _statusBeforePause: null,
    _failedWriteQueue: Promise.resolve(),
    _progressWriteQueue: Promise.resolve(),
    _lastProgressSaved: 0,
  };

  jobs.set(jobId, job);
  runJob(job).catch((error) => {
    job.status = "failed";
    job.error = error.message;
    job.message = "Run failed.";
    job.finishedAt = new Date().toISOString();
  });

  res.json({ jobId });
});

app.post("/api/jobs/:jobId/resume", async (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) {
    return res.status(404).json({ error: "Job not found." });
  }

  if (job.status !== "paused") {
    return res.json(publicJob(job));
  }

  resumeJob(job);
  await saveProgress(job, true);
  res.json(publicJob(job));
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

app.get("/download/:jobId/print-ready", (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job || job.status !== "completed" || !job.printReadyPath || !fs.existsSync(job.printReadyPath)) {
    return res.status(404).send("Print-ready PDF is not ready yet.");
  }

  res.download(job.printReadyPath, PRINT_READY_FILE_NAME);
});

app.get("/download/:jobId/duplex-print", (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job || job.status !== "completed" || !job.duplexPrintPath || !fs.existsSync(job.duplexPrintPath)) {
    return res.status(404).send("Duplex print PDF is not ready yet.");
  }

  res.download(job.duplexPrintPath, DUPLEX_PRINT_FILE_NAME);
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

  if (startRow > endRow) {
    return { ok: false, error: "Start Row cannot be greater than End Row." };
  }

  return { ok: true, startRow, endRow };
}

function prepareMembers(meta, emailIndex, sharedPassword, fileRanges) {
  const rangesByFile = new Map(
    Array.isArray(fileRanges)
      ? fileRanges.map((range) => [range.fileId, range])
      : []
  );
  const seen = new Set();
  const members = [];
  const resolvedRanges = [];
  const warnings = [];

  for (const file of meta.files) {
    const requestedRange = rangesByFile.get(file.id) || {};
    const range = parseRowRange(requestedRange.startRow, requestedRange.endRow, file.rowCount);
    if (!range.ok) {
      return { ok: false, error: `${file.originalName}: ${range.error}` };
    }

    resolvedRanges.push({
      fileId: file.id,
      originalName: file.originalName,
      startRow: range.startRow,
      endRow: range.endRow,
      rowCount: file.rowCount,
    });

    let rows;
    try {
      rows = readWorkbookRows(file.filePath).rows;
    } catch (error) {
      warnings.push(`${file.originalName}: ${error.message}`);
      continue;
    }

    const selectedRows = rows.slice(range.startRow - 1, range.endRow);

    for (let index = 0; index < selectedRows.length; index += 1) {
      const email = normalizeMemberId(selectedRows[index][emailIndex]);
      const dedupeKey = email.toLowerCase();
      if (!email || seen.has(dedupeKey)) {
        continue;
      }

      seen.add(dedupeKey);
      members.push({
        rowNumber: range.startRow + index + 1,
        fileName: file.originalName,
        email,
        password: sharedPassword,
      });
    }
  }

  return {
    ok: true,
    members,
    fileRanges: resolvedRanges,
    fileCount: meta.files.length,
    warnings,
  };
}

function parseConcurrency(value) {
  const parsed = value === "" || value == null
    ? DEFAULT_CONCURRENCY
    : Number.parseInt(String(value), 10);

  if (!Number.isInteger(parsed) || parsed < 3 || parsed > 25) {
    return { ok: false, error: "Speed must be between 3 and 25 parallel downloads." };
  }

  return { ok: true, value: parsed };
}

async function runJob(job) {
  await fsp.mkdir(job.outputPath, { recursive: true });
  await fsp.mkdir(ZIP_DIR, { recursive: true });

  const members = job.preparedMembers;

  job.total = members.length;
  job.remaining = members.length;
  job.status = "running";
  job.message = `0 already done, processing remaining ${job.remaining}...`;
  await saveProgress(job, true);

  const limit = pLimit(job.concurrency);
  const tasks = members.map((member, index) =>
    limit(async () => {
      let usedBrowser = false;

      try {
        await waitIfPaused(job);

        if (!member.email || !member.password) {
          await recordFailure(job, member.email || `row-${member.rowNumber}`, "Missing email or password.");
          job.failed += 1;
          job.consecutiveNetworkErrors = 0;
          return;
        }

        const existingFile = await findExistingCardFile(job.outputPath, member.email);
        if (existingFile) {
          job.alreadyDone += 1;
          member.outputFile = existingFile;
          job.consecutiveNetworkErrors = 0;
          return;
        }

        usedBrowser = true;
        await delay((index % job.concurrency) * WORKER_STAGGER_MS);
        const savedFile = await processMemberWithRetries(member, job.outputPath);
        member.outputFile = savedFile;
        job.success += 1;
        job.consecutiveNetworkErrors = 0;
      } catch (error) {
        if (isNetworkError(error)) {
          await recordNetworkError(job, member, error);
        } else {
          await recordFailure(job, member.email || `row-${member.rowNumber}`, error.message);
          job.failed += 1;
          job.consecutiveNetworkErrors = 0;
        }
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
  await processNetworkRetryQueue(job);
  job.outputFiles = members.map((member) => member.outputFile).filter(Boolean);
  await Promise.all([job._failedWriteQueue, job._progressWriteQueue]);
  await saveProgress(job, true);

  job.status = "zipping";
  job.message = "Creating print-ready PDF and ZIP...";
  await saveProgress(job, true);
  const mergedPdfPath = await mergePdfFiles(job.outputFiles, job.printReadyPath);
  if (mergedPdfPath) {
    job.printReadyUrl = `/download/${job.id}/print-ready`;
  }
  const duplexPdfPath = await createDuplexPrintPdf(job.outputFiles, job.duplexPrintPath);
  if (duplexPdfPath) {
    job.duplexPrintUrl = `/download/${job.id}/duplex-print`;
  }
  await zipFiles(
    [...job.outputFiles, mergedPdfPath, duplexPdfPath].filter(Boolean),
    job.failedPath,
    job.zipPath
  );
  job.status = "completed";
  job.message = `Completed ${job.total} members. ${job.alreadyDone} already done, ${job.success} downloaded, ${job.failed} failed, ${job.networkErrors || 0} network error(s).`;
  job.downloadUrl = `/download/${job.id}`;
  job.finishedAt = new Date().toISOString();
  await saveProgress(job, true);
}

async function processNetworkRetryQueue(job) {
  const retryMembers = [...job.networkRetryQueue];
  if (retryMembers.length === 0) {
    return;
  }

  job.status = "network-retry";
  job.message = `Retrying ${retryMembers.length} network error(s) now that the main run is complete...`;
  await saveProgress(job, true);
  await waitForConnectivityOrPause(job);

  job.networkRetryQueue = [];
  job.networkRetrySet.clear();
  job.networkRetryList = [];
  job.networkErrors = 0;

  const limit = pLimit(job.concurrency);
  const tasks = retryMembers.map((member, index) =>
    limit(async () => {
      await waitIfPaused(job);
      await delay((index % job.concurrency) * WORKER_STAGGER_MS);

      try {
        const existingFile = await findExistingCardFile(job.outputPath, member.email);
        if (existingFile) {
          member.outputFile = existingFile;
          job.alreadyDone += 1;
          job.consecutiveNetworkErrors = 0;
          return;
        }

        const savedFile = await processMemberWithRetries(member, job.outputPath);
        member.outputFile = savedFile;
        job.success += 1;
        job.consecutiveNetworkErrors = 0;
      } catch (error) {
        if (isNetworkError(error)) {
          await recordNetworkError(job, member, error, { retryAgain: false });
        } else {
          await recordFailure(job, member.email || `row-${member.rowNumber}`, error.message);
          job.failed += 1;
          job.consecutiveNetworkErrors = 0;
        }
      } finally {
        updateJobMessage(job);
        await saveProgress(job);
        await delay(RATE_LIMIT_MS);
      }
    })
  );

  await Promise.all(tasks);
}

async function processMemberWithRetries(member, outputPath) {
  let lastError;

  for (let attempt = 1; attempt <= MAX_TIMEOUT_RETRIES + 1; attempt += 1) {
    try {
      return await processMember(member.email, member.password, outputPath, member.rowNumber);
    } catch (error) {
      lastError = error;

      if (!shouldRetryMemberError(error) || attempt > MAX_TIMEOUT_RETRIES) {
        throw error;
      }

      await delay(RETRY_DELAY_MS);
    }
  }

  throw lastError;
}

function shouldRetryMemberError(error) {
  const message = String(error?.message || error || "").toLowerCase();
  const isTimeout = message.includes("timeout") || message.includes("timed out");
  const isPermanentFailure = /invalid|incorrect|unauthorized|not found/.test(message);

  return isTimeout && !isPermanentFailure;
}

function isNetworkError(error) {
  const message = String(error?.message || error || "").toUpperCase();
  return (
    message.includes("ERR_CONNECTION") ||
    message.includes("ERR_SOCKET") ||
    message.includes("TIMED_OUT") ||
    message.includes("NET::") ||
    message.includes("ECONNREFUSED") ||
    message.includes("ENOTFOUND")
  );
}

async function recordNetworkError(job, member, error, options = {}) {
  const { retryAgain = true } = options;
  const email = member.email || `row-${member.rowNumber}`;
  const reason = String(error?.message || error || "Network error").replace(/\s+/g, " ").trim();

  await deleteMemberOutputFiles(job.outputPath, email);

  if (retryAgain && !job.networkRetrySet.has(email)) {
    job.networkRetrySet.add(email);
    job.networkRetryQueue.push(member);
    job.networkRetryList.push({ email, reason });
  } else if (!retryAgain) {
    job.networkRetryList.push({ email, reason });
  }

  job.networkErrors = job.networkRetryList.length;
  job.consecutiveNetworkErrors += 1;

  if (job.consecutiveNetworkErrors >= NETWORK_PAUSE_THRESHOLD) {
    await pauseJob(job);
  }
}

async function deleteMemberOutputFiles(outputPath, email) {
  const baseName = safeFileName(email);
  if (!baseName) {
    return;
  }

  await Promise.all(
    COMMON_CARD_EXTENSIONS.map((ext) =>
      fsp.rm(path.join(outputPath, `${baseName}${ext}`), { force: true }).catch(() => {})
    )
  );
}

async function pauseJob(job) {
  if (job.status === "paused") {
    return;
  }

  job._statusBeforePause = job.status === "network-retry" ? "network-retry" : "running";
  job.status = "paused";
  job.pausedReason = "Network issues detected, job paused. Check connection then click Resume";
  job.message = job.pausedReason;
  await saveProgress(job, true);
}

function resumeJob(job) {
  const waiters = job._resumeWaiters.splice(0);
  job.status = job._statusBeforePause || "running";
  job._statusBeforePause = null;
  job.pausedReason = null;
  job.consecutiveNetworkErrors = 0;

  if (job.status === "network-retry") {
    job.message = `Retrying ${job.networkRetryQueue.length || job.networkErrors} network error(s)...`;
  } else {
    updateJobMessage(job);
  }

  waiters.forEach((resolve) => resolve());
}

async function waitIfPaused(job) {
  while (job.status === "paused") {
    await new Promise((resolve) => {
      job._resumeWaiters.push(resolve);
    });
  }
}

async function waitForConnectivityOrPause(job) {
  while (true) {
    const connectivity = await checkSiteConnectivity();
    if (connectivity.ok) {
      return;
    }

    await pauseJob(job);
    await waitIfPaused(job);
  }
}

async function checkSiteConnectivity() {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), CONNECTIVITY_TIMEOUT_MS);

  try {
    await fetch(CONNECTIVITY_URL, {
      method: "GET",
      signal: controller.signal,
      redirect: "follow",
    });
    return { ok: true };
  } catch (error) {
    return { ok: false, error };
  } finally {
    clearTimeout(timeout);
  }
}

async function processMember(email, password, outputPath, rowNumber) {
  let browser;
  let context;
  let phase = "starting browser";
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
    phase = "creating browser context";
    context = await browser.newContext({
      viewport: { width: 1440, height: 1100 },
      acceptDownloads: true,
    });
    const page = await context.newPage();
    page.setDefaultTimeout(MEMBER_TIMEOUT_MS);

    phase = "opening login page";
    await page.goto(LOGIN_URL, { waitUntil: "domcontentloaded", timeout: MEMBER_TIMEOUT_MS });
    phase = "filling login form";
    await fillLoginForm(page, email, password);
    phase = "submitting login";
    await clickLogin(page);

    phase = "waiting for dashboard ID card button";
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
      phase = "clicking dashboard ID card button";
      await idCardButton.click();
    } else {
      phase = "clicking fallback dashboard ID card button";
      await fallbackIdCardButton.click();
    }

    phase = "waiting for card modal";
    await waitForCardModal(page);
    phase = "downloading ID card";
    return await downloadIdCard(page, email, outputPath, rowNumber);
  } finally {
    clearTimeout(timeout);
    if (context) {
      await context.close().catch(() => {});
    }
    if (browser) {
      await browser.close().catch(() => {});
    }
    if (timedOut) {
      throw new Error(`Member timed out after ${MEMBER_TIMEOUT_MS / 1000} seconds while ${phase}.`);
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
  const ext = path.extname(download.suggestedFilename()) || ".pdf";
  const fileName = `${safeFileName(email) || `row-${rowNumber}`}${ext}`;
  const filePath = path.join(outputDir, fileName);
  await download.saveAs(filePath);
  return filePath;
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

function zipFiles(filePaths, failedPath, destinationPath) {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(destinationPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", resolve);
    archive.on("error", reject);
    archive.pipe(output);

    for (const filePath of [...new Set(filePaths)]) {
      if (filePath && fs.existsSync(filePath)) {
        archive.file(filePath, { name: path.basename(filePath) });
      }
    }

    if (failedPath && fs.existsSync(failedPath) && fs.statSync(failedPath).size > 0) {
      archive.file(failedPath, { name: "failed.txt" });
    }

    archive.finalize();
  });
}

async function mergePdfFiles(filePaths, destinationPath) {
  const pdfPaths = [...new Set(filePaths)]
    .filter((filePath) => filePath && path.extname(filePath).toLowerCase() === ".pdf")
    .filter((filePath) => fs.existsSync(filePath));

  if (pdfPaths.length === 0) {
    return null;
  }

  const mergedPdf = await PDFDocument.create();

  for (const pdfPath of pdfPaths) {
    const sourceBytes = await fsp.readFile(pdfPath);
    const sourcePdf = await PDFDocument.load(sourceBytes);
    const copiedPages = await mergedPdf.copyPages(sourcePdf, sourcePdf.getPageIndices());
    for (const page of copiedPages) {
      mergedPdf.addPage(page);
    }
  }

  const mergedBytes = await mergedPdf.save();
  await fsp.writeFile(destinationPath, mergedBytes);
  return destinationPath;
}

async function createDuplexPrintPdf(filePaths, destinationPath) {
  const pdfPaths = [...new Set(filePaths)]
    .filter((filePath) => filePath && path.extname(filePath).toLowerCase() === ".pdf")
    .filter((filePath) => fs.existsSync(filePath))
    .filter((filePath) => ![PRINT_READY_FILE_NAME, DUPLEX_PRINT_FILE_NAME].includes(path.basename(filePath)));

  if (pdfPaths.length === 0) {
    return null;
  }

  const duplexPdf = await PDFDocument.create();

  for (let index = 0; index < pdfPaths.length; index += 2) {
    const memberA = await readMemberCardPages(duplexPdf, pdfPaths[index]);
    const memberB = pdfPaths[index + 1]
      ? await readMemberCardPages(duplexPdf, pdfPaths[index + 1])
      : null;

    const frontPage = duplexPdf.addPage([A4_WIDTH, A4_HEIGHT]);
    drawCardSlot(frontPage, memberA.front, "top");
    if (memberB) {
      drawCardSlot(frontPage, memberB.front, "bottom");
    }

    const backPage = duplexPdf.addPage([A4_WIDTH, A4_HEIGHT]);
    if (memberA.back) {
      drawCardSlot(backPage, memberA.back, "top");
    }
    if (memberB?.back) {
      drawCardSlot(backPage, memberB.back, "bottom");
    }
  }

  const duplexBytes = await duplexPdf.save();
  await fsp.writeFile(destinationPath, duplexBytes);
  return destinationPath;
}

async function readMemberCardPages(targetPdf, sourcePath) {
  const sourceBytes = await fsp.readFile(sourcePath);
  const sourcePdf = await PDFDocument.load(sourceBytes);
  const pageIndices = sourcePdf.getPageIndices().slice(0, 2);
  const [front, back] = await targetPdf.embedPdf(sourceBytes, pageIndices);

  return { front, back };
}

function drawCardSlot(page, embeddedPage, slot) {
  if (!embeddedPage) {
    return;
  }

  page.drawPage(embeddedPage, {
    x: 0,
    y: slot === "top" ? A4_HALF_HEIGHT : 0,
    width: A4_WIDTH,
    height: A4_HALF_HEIGHT,
  });
}

function getChromiumLaunchOptions() {
  const args = [
    "--no-sandbox",
    "--disable-setuid-sandbox",
    "--disable-dev-shm-usage",
    "--disable-gpu",
    "--no-zygote",
  ];

  const launchOptions = {
    headless: true,
    args,
  };

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
  job.failedList.push({ email, reason: cleanReason });
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
    printReadyPath: job.printReadyPath,
    duplexPrintPath: job.duplexPrintPath,
    updatedAt: new Date().toISOString(),
  };

  job._progressWriteQueue = job._progressWriteQueue.catch(() => {}).then(() =>
    fsp.writeFile(job.progressPath, JSON.stringify(snapshot, null, 2), "utf8")
  );
  await job._progressWriteQueue;
}

function updateJobMessage(job) {
  const networkText = job.networkErrors ? ` Network errors queued: ${job.networkErrors}.` : "";
  job.message = `${job.alreadyDone} already done, processing remaining ${job.remaining}... Processing ${job.current} of ${job.total}.${networkText}`;
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
    fileCount: job.fileCount,
    fileRanges: job.fileRanges,
    parseWarnings: job.parseWarnings,
    failedList: job.failedList,
    networkErrors: job.networkErrors || 0,
    networkRetryList: job.networkRetryList || [],
    pausedReason: job.pausedReason,
    message: job.message,
    downloadUrl: job.downloadUrl,
    printReadyUrl: job.printReadyUrl,
    duplexPrintUrl: job.duplexPrintUrl,
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

function normalizeMemberId(value) {
  let text = normalizeCell(value);
  if (!text) {
    return "";
  }

  text = text.replace(/\u00a0/g, " ").trim();

  if (text.includes("@")) {
    return text.replace(/\s+/g, "").toLowerCase();
  }

  const digits = text.replace(/[^\d]/g, "");
  if (!digits) {
    return "";
  }

  if (digits.length === 10 && /^[789]/.test(digits)) {
    return `0${digits}`;
  }

  if (digits.length === 13 && digits.startsWith("234")) {
    return `0${digits.slice(3)}`;
  }

  return digits;
}

function clampNumber(value, min, max) {
  if (!Number.isFinite(value)) {
    return min;
  }

  return Math.min(Math.max(value, min), max);
}

function normalizeCell(value) {
  return String(value ?? "").trim();
}

async function clearOutputFolder() {
  await fsp.mkdir(OUTPUT_DIR, { recursive: true });

  const entries = await fsp.readdir(OUTPUT_DIR, { withFileTypes: true });
  await Promise.all(
    entries.map((entry) =>
      fsp.rm(path.join(OUTPUT_DIR, entry.name), {
        recursive: true,
        force: true,
      })
    )
  );

  await fsp.mkdir(ZIP_DIR, { recursive: true });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
