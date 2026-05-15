const fs = require("fs");
const fsp = fs.promises;
const path = require("path");

const ROOT_DIR = __dirname;
const OUTPUT_DIR = process.env.OUTPUT_DIR || path.join(ROOT_DIR, "output");
const CARD_EXTENSIONS = [".pdf", ".png", ".jpg", ".jpeg", ".webp"];

async function main() {
  const failedFile = await resolveFailedFile(process.argv[2]);

  if (!failedFile) {
    console.log("No failed.txt file found.");
    console.log(`Checked ${path.join(OUTPUT_DIR, "failed.txt")} and failed logs inside ${OUTPUT_DIR}.`);
    return;
  }

  const failedText = await fsp.readFile(failedFile, "utf8");
  const failedMembers = unique(
    failedText
      .split(/\r?\n/)
      .map(parseFailedMember)
      .filter(Boolean)
  );

  let deleted = 0;
  let missing = 0;

  for (const member of failedMembers) {
    const baseName = safeFileName(member);
    let deletedForMember = false;

    for (const ext of CARD_EXTENSIONS) {
      const cardPath = path.join(OUTPUT_DIR, `${baseName}${ext}`);

      if (await pathExists(cardPath)) {
        await fsp.unlink(cardPath);
        deleted += 1;
        deletedForMember = true;
        console.log(`Deleted ${path.relative(ROOT_DIR, cardPath)}`);
      }
    }

    if (!deletedForMember) {
      missing += 1;
    }
  }

  console.log(`Failed log: ${path.relative(ROOT_DIR, failedFile)}`);
  console.log(`Failed members found: ${failedMembers.length}`);
  console.log(`Files deleted: ${deleted}`);
  console.log(`Members with no matching output file: ${missing}`);
}

async function resolveFailedFile(requestedPath) {
  if (requestedPath) {
    const absolutePath = path.resolve(ROOT_DIR, requestedPath);
    if (!(await pathExists(absolutePath))) {
      throw new Error(`Failed log not found: ${requestedPath}`);
    }
    return absolutePath;
  }

  const defaultFailedPath = path.join(OUTPUT_DIR, "failed.txt");
  if (await pathExists(defaultFailedPath)) {
    return defaultFailedPath;
  }

  const failedLogs = await findFailedLogs(OUTPUT_DIR);
  failedLogs.sort((a, b) => b.mtimeMs - a.mtimeMs);
  return failedLogs[0] ? failedLogs[0].filePath : null;
}

async function findFailedLogs(directory) {
  if (!(await pathExists(directory))) {
    return [];
  }

  const entries = await fsp.readdir(directory, { withFileTypes: true });
  const logs = [];

  for (const entry of entries) {
    const entryPath = path.join(directory, entry.name);

    if (entry.isDirectory()) {
      logs.push(...(await findFailedLogs(entryPath)));
    } else if (entry.isFile() && /(?:^failed|failed)\.txt$/i.test(entry.name)) {
      const stats = await fsp.stat(entryPath);
      if (stats.size > 0) {
        logs.push({ filePath: entryPath, mtimeMs: stats.mtimeMs });
      }
    }
  }

  return logs;
}

function parseFailedMember(line) {
  const trimmed = line.trim();
  if (!trimmed) {
    return "";
  }

  const tabParts = trimmed.split("\t").map((part) => part.trim()).filter(Boolean);
  if (tabParts.length >= 2) {
    return tabParts[1];
  }

  const emailMatch = trimmed.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  if (emailMatch) {
    return emailMatch[0];
  }

  const phoneMatch = trimmed.match(/\b(?:\+?234|0)?[789]\d{9}\b/);
  return phoneMatch ? phoneMatch[0] : "";
}

function safeFileName(value) {
  return String(value)
    .trim()
    .replace(/[/\\?%*:|"<>]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 180);
}

function unique(values) {
  return [...new Set(values)];
}

async function pathExists(filePath) {
  try {
    await fsp.access(filePath);
    return true;
  } catch (_error) {
    return false;
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
