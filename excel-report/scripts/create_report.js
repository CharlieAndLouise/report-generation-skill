#!/usr/bin/env node
/**
 * Bundled Excel report builder for the excel-report skill.
 * Handles logo placement, header layout, and data formatting.
 *
 * Usage:
 *   node create_report.js \
 *     --title "Report Title" \
 *     --summary "Summary sentence" \
 *     --data /tmp/data.json \
 *     --logo resources/company_logo.png \
 *     --output output/report_20250115_120000.xlsx
 */

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const HEADER_BG = '1F4E79';
const HEADER_FG = 'FFFFFFFF';
const ALT_ROW_BG = 'FFD6E4F0';
const SUMMARY_BG = 'FFF2F2F2';
const SUMMARY_FG = 'FF595959';
const TITLE_COLOR = 'FF1F4E79';

const CURRENCY_KEYWORDS = ['salary', 'pay', 'wage', 'compensation', 'income', 'bonus', 'rate'];

function isCurrencyCol(key) {
  return CURRENCY_KEYWORDS.some(kw => key.toLowerCase().includes(kw));
}

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i += 2) {
    args[argv[i].replace('--', '')] = argv[i + 1];
  }
  return args;
}

async function createReport(title, summary, data, logoPath, outputPath) {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Report');

  // --- Logo + title area (rows 1-2) ---
  ws.getRow(1).height = 45;
  ws.getRow(2).height = 20;
  ws.getColumn('A').width = 4;
  ws.getColumn('B').width = 16;
  ws.getColumn('C').width = 5;

  if (logoPath && fs.existsSync(logoPath)) {
    const ext = path.extname(logoPath).slice(1).toLowerCase();
    const imageId = workbook.addImage({
      filename: logoPath,
      extension: ext === 'jpg' ? 'jpeg' : ext,
    });
    ws.addImage(imageId, {
      tl: { col: 0, row: 0 },
      br: { col: 2, row: 2 },
      editAs: 'oneCell',
    });
  }

  const titleCell = ws.getCell('C1');
  titleCell.value = title;
  titleCell.font = { bold: true, size: 14, color: { argb: TITLE_COLOR } };
  titleCell.alignment = { vertical: 'middle' };

  // --- Summary row (row 3) ---
  ws.getRow(3).height = 18;
  const nCols = data.length > 0 ? Object.keys(data[0]).length : 4;
  const endCol = Math.max(nCols, 4);

  ws.mergeCells(3, 1, 3, endCol);
  const summaryCell = ws.getCell('A3');
  summaryCell.value = summary;
  summaryCell.font = { italic: true, size: 10, color: { argb: SUMMARY_FG } };
  summaryCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SUMMARY_BG } };
  summaryCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  if (data.length === 0) {
    const dir = path.dirname(outputPath);
    if (dir) fs.mkdirSync(dir, { recursive: true });
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Report saved: ${outputPath}`);
    return;
  }

  const headers = Object.keys(data[0]);

  // --- Column headers (row 4) ---
  ws.getRow(4).height = 22;
  headers.forEach((key, i) => {
    const cell = ws.getCell(4, i + 1);
    cell.value = key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + HEADER_BG } };
    cell.font = { bold: true, color: { argb: HEADER_FG }, size: 11 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // --- Data rows (row 5+) ---
  data.forEach((record, rowOffset) => {
    const rowNum = rowOffset + 5;
    ws.getRow(rowNum).height = 16;
    headers.forEach((key, colOffset) => {
      const cell = ws.getCell(rowNum, colOffset + 1);
      const raw = record[key];

      if (isCurrencyCol(key)) {
        const num = parseFloat(String(raw).replace(/[$,]/g, ''));
        if (!isNaN(num)) {
          cell.value = num;
          cell.numFmt = '$#,##0';
        } else {
          cell.value = raw;
        }
      } else {
        cell.value = raw;
      }

      if (rowNum % 2 === 0) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT_ROW_BG } };
      }
      cell.alignment = { vertical: 'middle' };
    });
  });

  // --- Auto-fit column widths ---
  headers.forEach((key, i) => {
    const col = ws.getColumn(i + 1);
    const headerLabel = key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
    let maxLen = headerLabel.length;
    data.forEach(record => {
      const val = record[key];
      if (val != null) maxLen = Math.max(maxLen, String(val).length);
    });
    col.width = Math.min(maxLen + 4, 40);
  });

  const dir = path.dirname(outputPath);
  if (dir) fs.mkdirSync(dir, { recursive: true });
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Report saved: ${outputPath}`);
}

(async () => {
  const args = parseArgs(process.argv);
  if (!args.title || !args.summary || !args.data || !args.output) {
    console.error('Usage: node create_report.js --title "..." --summary "..." --data data.json --logo logo.png --output out.xlsx');
    process.exit(1);
  }
  const data = JSON.parse(fs.readFileSync(args.data, 'utf8'));
  await createReport(args.title, args.summary, data, args.logo || 'resources/company_logo.png', args.output);
})();
