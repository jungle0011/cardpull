const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const zlib = require("zlib");
const XLSX = require("xlsx");
const Papa = require("papaparse");
const { PDFDocument } = require("pdf-lib");

const ROOT_DIR = __dirname;
const outputDir = path.join(ROOT_DIR, "output");
const zoneDir = path.join(ROOT_DIR, "safaari excel");
const destinationDir = "/Users/user/Desktop/safari completed";
const A4_WIDTH = 595;
const A4_HEIGHT = 842;
const A4_HALF_HEIGHT = A4_HEIGHT / 2;
const IGNORED_PDFS = new Set([
  "all-cards-print-ready.pdf",
  "duplex-print-ready.pdf",
]);

const ZONE_GROUPS = [
  { name: "Ilorin East", files: ["Ilorin East 01.csv", "Ilorin East 02.xlsx"] },
  { name: "Ilorin South", files: ["Ilorin South 02.xlsx"] },
  { name: "Irepodun", files: ["Irepodun 01.csv", "Irepodun 02.xlsx"] },
  { name: "Isin", files: ["Isin.csv"] },
  { name: "Kaiama", files: ["Kaiama 01.xlsx", "Kaiama 02.xlsx"] },
  { name: "Moro", files: ["Moro 01.csv", "Moro 02.xlsx"] },
  { name: "Offa", files: ["Offa 01.csv", "Offa 02.xlsx"] },
  { name: "Oke-ero", files: ["Oke-ero.csv"] },
  { name: "Oyun", files: ["Oyun 01.csv", "Oyun 02.xlsx"] },
  { name: "Patigi", files: ["Patigi 01.xlsx", "Patigi o2.xlsx"] },
  { name: "Baruten", files: ["Baruten.csv"] },
];

async function main() {
  if (!fs.existsSync(outputDir)) {
    throw new Error(`Output folder not found: ${outputDir}`);
  }
  if (!fs.existsSync(zoneDir)) {
    throw new Error(`Zone folder not found: ${zoneDir}`);
  }

  await fsp.mkdir(destinationDir, { recursive: true });
  const outputPdfs = buildOutputPdfMap();
  const report = [];

  for (const group of ZONE_GROUPS) {
    const zonePhones = readZonePhones(group.files);
    const matchedPaths = zonePhones
      .map((phone) => outputPdfs.get(phone))
      .filter(Boolean);
    const destination = path.join(destinationDir, `${group.name}-duplex.pdf`);

    if (matchedPaths.length > 0) {
      await createDuplexPrintPdf(matchedPaths, destination);
    }

    report.push({
      zone: group.name,
      zoneMembers: zonePhones.length,
      pdfsFound: matchedPaths.length,
      missing: zonePhones.length - matchedPaths.length,
      destination: matchedPaths.length > 0 ? destination : null,
    });
  }

  console.log("Duplex generation report:");
  for (const entry of report) {
    const output = entry.destination ? ` -> ${entry.destination}` : " -> no PDFs found, skipped";
    console.log(`${entry.zone}: ${entry.pdfsFound}/${entry.zoneMembers} PDFs (${entry.missing} missing)${output}`);
  }
}

function buildOutputPdfMap() {
  const map = new Map();
  for (const fileName of fs.readdirSync(outputDir)) {
    if (!fileName.toLowerCase().endsWith(".pdf")) continue;
    if (IGNORED_PDFS.has(fileName)) continue;

    const phone = normalizePhone(path.basename(fileName, ".pdf"));
    if (phone && !map.has(phone)) {
      map.set(phone, path.join(outputDir, fileName));
    }
  }
  return map;
}

function readZonePhones(fileNames) {
  const phones = [];
  const seen = new Set();

  for (const fileName of fileNames) {
    const filePath = path.join(zoneDir, fileName);
    if (!fs.existsSync(filePath)) {
      console.warn(`Missing zone file: ${filePath}`);
      continue;
    }

    const rows = readRows(filePath);
    const phoneIndex = detectPhoneColumnIndex(rows);
    for (const row of rows) {
      const phone = findPhoneInRow(row, phoneIndex);
      if (!phone || seen.has(phone)) continue;
      seen.add(phone);
      phones.push(phone);
    }
  }

  return phones;
}

function readRows(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".csv") {
    const parsed = Papa.parse(fs.readFileSync(filePath, "utf8"), {
      skipEmptyLines: true,
    });
    return parsed.data.filter((row) => !isEmptyRow(row));
  }

  try {
    const workbook = XLSX.readFile(filePath, { cellDates: false });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) return [];
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false,
    }).filter((row) => !isEmptyRow(row));
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

function detectPhoneColumnIndex(rows) {
  for (const row of rows.slice(0, 5)) {
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

function findPhoneInRow(row, phoneColumnIndex) {
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
  if (!text || /[a-df-z]/i.test(text)) return null;

  if (/e\+?/i.test(text) && /^[-+]?\d*\.?\d+e\+?\d+$/i.test(text.replace(/\s+/g, ""))) {
    const numeric = Number(text);
    if (Number.isFinite(numeric)) text = numeric.toFixed(0);
  }

  let digits = text.replace(/\D/g, "");
  if (digits.startsWith("234") && digits.length === 13) digits = `0${digits.slice(3)}`;
  if (digits.length === 10) digits = `0${digits}`;
  return digits.length === 11 && digits.startsWith("0") ? digits : null;
}

function isEmptyRow(row) {
  const cells = Array.isArray(row) ? row : Object.values(row || {});
  return cells.every((value) => String(value ?? "").trim() === "");
}

async function createDuplexPrintPdf(filePaths, destination) {
  const duplexPdf = await PDFDocument.create();

  for (let index = 0; index < filePaths.length; index += 2) {
    const memberA = await readMemberCardPages(duplexPdf, filePaths[index]);
    const memberB = filePaths[index + 1]
      ? await readMemberCardPages(duplexPdf, filePaths[index + 1])
      : null;

    const frontPage = duplexPdf.addPage([A4_WIDTH, A4_HEIGHT]);
    drawCardSlot(frontPage, memberA.front, "top");
    if (memberB) drawCardSlot(frontPage, memberB.front, "bottom");

    const backPage = duplexPdf.addPage([A4_WIDTH, A4_HEIGHT]);
    if (memberA.back) drawCardSlot(backPage, memberA.back, "top");
    if (memberB?.back) drawCardSlot(backPage, memberB.back, "bottom");
  }

  await fsp.writeFile(destination, await duplexPdf.save());
}

async function readMemberCardPages(targetPdf, sourcePath) {
  const sourceBytes = await fsp.readFile(sourcePath);
  const sourcePdf = await PDFDocument.load(sourceBytes);
  const pageIndices = sourcePdf.getPageIndices().slice(0, 2);
  const embeddedPages = await targetPdf.embedPdf(sourceBytes, pageIndices);
  return {
    front: embeddedPages[0],
    back: embeddedPages[1] || null,
  };
}

function drawCardSlot(page, embeddedPage, slot) {
  if (!embeddedPage) return;
  page.drawPage(embeddedPage, {
    x: 0,
    y: slot === "top" ? A4_HALF_HEIGHT : 0,
    width: A4_WIDTH,
    height: A4_HALF_HEIGHT,
  });
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
