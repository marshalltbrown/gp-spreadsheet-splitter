const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

function sanitizeFileName(name) {
  return String(name)
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .slice(0, 120);
}

function normalizeHeaderText(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function deepClone(value) {
  if (value === null || value === undefined) return value;
  if (value instanceof Date) return new Date(value.getTime());
  if (Array.isArray(value)) return value.map((item) => deepClone(item));
  if (typeof value === "object") {
    const out = {};
    Object.keys(value).forEach((key) => {
      out[key] = deepClone(value[key]);
    });
    return out;
  }
  return value;
}

function getSourceXlsxFiles(sourceDir) {
  if (!fs.existsSync(sourceDir)) {
    throw new Error(`Source folder not found: ${sourceDir}`);
  }

  return fs
    .readdirSync(sourceDir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === ".xlsx")
    .map((entry) => entry.name);
}

function getCellText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || "").join("");
    }
    if (value.text !== undefined) return String(value.text);
    if (value.result !== undefined) return String(value.result);
  }
  return String(value);
}

function getWorksheetUsedRowBounds(worksheet) {
  let minRow = null;
  let maxRow = null;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    let hasValue = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (getCellText(cell.value).trim() !== "") {
        hasValue = true;
      }
    });
    if (hasValue) {
      if (minRow === null) minRow = rowNumber;
      maxRow = rowNumber;
    }
  });

  return { minRow, maxRow };
}

function findRowUsedColumnBounds(worksheet, rowNumber) {
  const row = worksheet.getRow(rowNumber);
  let startCol = null;
  let endCol = null;

  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (getCellText(cell.value).trim() === "") return;
    if (startCol === null) startCol = colNumber;
    endCol = colNumber;
  });

  return { startCol, endCol };
}

function findSectionsWithLicenseType(worksheet, minRow, maxRow) {
  const sections = [];

  for (let rowNumber = minRow; rowNumber <= maxRow; rowNumber += 1) {
    const bounds = findRowUsedColumnBounds(worksheet, rowNumber);
    if (bounds.startCol === null || bounds.endCol === null) continue;

    let licenseCol = null;
    for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
      const text = normalizeHeaderText(getCellText(worksheet.getRow(rowNumber).getCell(col).value));
      if (text === "license type") {
        licenseCol = col;
        break;
      }
    }

    if (licenseCol !== null) {
      sections.push({
        headerRow: rowNumber,
        startCol: bounds.startCol,
        endCol: bounds.endCol,
        licenseCol,
      });
    }
  }

  sections.forEach((section, index) => {
    const nextSection = sections[index + 1];
    section.dataStartRow = section.headerRow + 1;
    section.dataEndRow = nextSection ? nextSection.headerRow - 1 : maxRow;
  });

  return sections;
}

function parseMerges(worksheet) {
  const merges = worksheet.model && Array.isArray(worksheet.model.merges) ? worksheet.model.merges : [];
  return merges.map((ref) => {
    const [startAddress, endAddress = startAddress] = ref.split(":");
    const startCell = worksheet.getCell(startAddress);
    const endCell = worksheet.getCell(endAddress);
    return {
      startRow: startCell.row,
      endRow: endCell.row,
      startCol: startCell.col,
      endCol: endCell.col,
    };
  });
}

function findMergeForCell(merges, rowNumber, colNumber) {
  for (let i = 0; i < merges.length; i += 1) {
    const merge = merges[i];
    if (
      rowNumber >= merge.startRow &&
      rowNumber <= merge.endRow &&
      colNumber >= merge.startCol &&
      colNumber <= merge.endCol
    ) {
      return merge;
    }
  }
  return null;
}

function getLicenseBlocksForSection(worksheet, section, merges) {
  const blocksByLicense = new Map();
  const pendingRows = new Set();

  for (let rowNumber = section.dataStartRow; rowNumber <= section.dataEndRow; rowNumber += 1) {
    pendingRows.add(rowNumber);
  }

  for (let rowNumber = section.dataStartRow; rowNumber <= section.dataEndRow; rowNumber += 1) {
    if (!pendingRows.has(rowNumber)) continue;

    const merge = findMergeForCell(merges, rowNumber, section.licenseCol);
    let blockStart = rowNumber;
    let blockEnd = rowNumber;

    if (merge) {
      blockStart = Math.max(merge.startRow, section.dataStartRow);
      blockEnd = Math.min(merge.endRow, section.dataEndRow);
      if (rowNumber !== blockStart) {
        pendingRows.delete(rowNumber);
        continue;
      }
    }

    for (let r = blockStart; r <= blockEnd; r += 1) {
      pendingRows.delete(r);
    }

    const rawLicense = worksheet.getRow(blockStart).getCell(section.licenseCol).value;
    const licenseType = getCellText(rawLicense).trim();
    if (!licenseType) continue;

    if (!blocksByLicense.has(licenseType)) {
      blocksByLicense.set(licenseType, []);
    }
    blocksByLicense.get(licenseType).push({
      startRow: blockStart,
      endRow: blockEnd,
    });
  }

  return blocksByLicense;
}

function copyRowRangeToOutput({
  sourceWorksheet,
  outputWorksheet,
  sourceRowStart,
  sourceRowEnd,
  sourceColStart,
  sourceColEnd,
  targetRowStart,
  sourceMerges,
  outputMergeRefs,
}) {
  const rowOffset = targetRowStart - sourceRowStart;

  for (let col = sourceColStart; col <= sourceColEnd; col += 1) {
    const sourceColumn = sourceWorksheet.getColumn(col);
    const outputColumn = outputWorksheet.getColumn(col);
    outputColumn.width = sourceColumn.width;
    outputColumn.hidden = sourceColumn.hidden;
    outputColumn.outlineLevel = sourceColumn.outlineLevel;
    outputColumn.style = deepClone(sourceColumn.style || {});
  }

  for (let sourceRow = sourceRowStart; sourceRow <= sourceRowEnd; sourceRow += 1) {
    const targetRow = sourceRow + rowOffset;
    const sourceRowObj = sourceWorksheet.getRow(sourceRow);
    const outputRowObj = outputWorksheet.getRow(targetRow);

    outputRowObj.height = sourceRowObj.height;
    outputRowObj.hidden = sourceRowObj.hidden;
    outputRowObj.outlineLevel = sourceRowObj.outlineLevel;

    for (let col = sourceColStart; col <= sourceColEnd; col += 1) {
      const sourceCell = sourceRowObj.getCell(col);
      const outputCell = outputRowObj.getCell(col);
      outputCell.value = deepClone(sourceCell.value);
      outputCell.style = deepClone(sourceCell.style || {});
      outputCell.numFmt = sourceCell.numFmt;
      outputCell.dataValidation = deepClone(sourceCell.dataValidation || {});
    }
  }

  sourceMerges.forEach((merge) => {
    const intersectsRows = merge.endRow >= sourceRowStart && merge.startRow <= sourceRowEnd;
    const intersectsCols = merge.endCol >= sourceColStart && merge.startCol <= sourceColEnd;
    if (!intersectsRows || !intersectsCols) return;

    const startRow = merge.startRow + rowOffset;
    const endRow = merge.endRow + rowOffset;
    const ref = `${outputWorksheet.getCell(startRow, merge.startCol).address}:${outputWorksheet.getCell(endRow, merge.endCol).address}`;
    outputMergeRefs.add(ref);
  });
}

function buildLicenseWorkbook(sourceWorksheet, sections, sourceMerges, licenseType) {
  const outWorkbook = new ExcelJS.Workbook();
  const outputWorksheet = outWorkbook.addWorksheet(sourceWorksheet.name, {
    properties: deepClone(sourceWorksheet.properties || {}),
    views: deepClone(sourceWorksheet.views || []),
    pageSetup: deepClone(sourceWorksheet.pageSetup || {}),
    headerFooter: deepClone(sourceWorksheet.headerFooter || {}),
    state: sourceWorksheet.state,
  });

  const outputMergeRefs = new Set();
  let writeRow = 1;
  let wroteAnyRows = false;

  sections.forEach((section) => {
    const blocks = section.blocksByLicense.get(licenseType) || [];
    if (!blocks.length) return;

    copyRowRangeToOutput({
      sourceWorksheet,
      outputWorksheet,
      sourceRowStart: section.headerRow,
      sourceRowEnd: section.headerRow,
      sourceColStart: section.startCol,
      sourceColEnd: section.endCol,
      targetRowStart: writeRow,
      sourceMerges,
      outputMergeRefs,
    });
    wroteAnyRows = true;
    writeRow += 1;

    blocks.forEach((block) => {
      const blockRowCount = block.endRow - block.startRow + 1;
      copyRowRangeToOutput({
        sourceWorksheet,
        outputWorksheet,
        sourceRowStart: block.startRow,
        sourceRowEnd: block.endRow,
        sourceColStart: section.startCol,
        sourceColEnd: section.endCol,
        targetRowStart: writeRow,
        sourceMerges,
        outputMergeRefs,
      });
      wroteAnyRows = true;
      writeRow += blockRowCount;
    });

    writeRow += 1;
  });

  if (!wroteAnyRows) {
    return null;
  }

  outputMergeRefs.forEach((mergeRef) => {
    outputWorksheet.mergeCells(mergeRef);
  });

  return outWorkbook;
}

async function splitWorkbookByLicenseTypes(inputPath, targetFolder) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);

  const sourceWorksheet = workbook.worksheets[0];
  if (!sourceWorksheet) {
    console.log(`Skip (no sheets): ${inputPath}`);
    return;
  }

  const bounds = getWorksheetUsedRowBounds(sourceWorksheet);
  if (bounds.minRow === null || bounds.maxRow === null) {
    console.log(`Skip (empty first sheet): ${inputPath}`);
    return;
  }

  const sections = findSectionsWithLicenseType(sourceWorksheet, bounds.minRow, bounds.maxRow);
  if (!sections.length) {
    console.log(`Skip (no License Type sections): ${inputPath}`);
    return;
  }

  const sourceMerges = parseMerges(sourceWorksheet);
  const allLicenseTypes = new Set();

  sections.forEach((section) => {
    section.blocksByLicense = getLicenseBlocksForSection(sourceWorksheet, section, sourceMerges);
    section.blocksByLicense.forEach((_, licenseType) => {
      allLicenseTypes.add(licenseType);
    });
  });

  for (const licenseType of allLicenseTypes) {
    const outWorkbook = buildLicenseWorkbook(sourceWorksheet, sections, sourceMerges, licenseType);
    if (!outWorkbook) continue;

    const outputFileName = `${sanitizeFileName(licenseType)} Requirements.xlsx`;
    const outputPath = path.join(targetFolder, outputFileName);
    await outWorkbook.xlsx.writeFile(outputPath);
    console.log(`Wrote: ${outputPath}`);
  }
}

async function ensureWorkbookOutputFolders(sourceDir, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });

  const workbookFiles = getSourceXlsxFiles(sourceDir);
  if (!workbookFiles.length) {
    console.log(`No .xlsx files found in: ${sourceDir}`);
    return;
  }

  for (const fileName of workbookFiles) {
    const workbookName = path.parse(fileName).name;
    const folderName = sanitizeFileName(workbookName) || "untitled";
    const targetFolder = path.join(outputDir, folderName);
    const sourcePath = path.join(sourceDir, fileName);

    if (fs.existsSync(targetFolder)) {
      console.log(`Skip (exists): ${targetFolder}`);
      continue;
    }

    fs.mkdirSync(targetFolder, { recursive: true });
    console.log(`Created: ${targetFolder}`);
    await splitWorkbookByLicenseTypes(sourcePath, targetFolder);
  }
}

async function main() {
  const sourceArg = process.argv[2] || "source";
  const outputArg = process.argv[3] || "output";

  const sourceDir = path.resolve(process.cwd(), sourceArg);
  const outputDir = path.resolve(process.cwd(), outputArg);

  try {
    await ensureWorkbookOutputFolders(sourceDir, outputDir);
    console.log("Done.");
  } catch (error) {
    console.error(`Error: ${error.message}`);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}
