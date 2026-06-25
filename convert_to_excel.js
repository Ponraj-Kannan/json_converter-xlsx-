const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Support CLI argument for custom directory, fallback to current working directory
const inputArg = process.argv[2];
let dirPath = process.cwd();

if (inputArg) {
  const resolved = path.resolve(inputArg);
  try {
    if (fs.existsSync(resolved)) {
      if (fs.statSync(resolved).isDirectory()) {
        dirPath = resolved;
      } else {
        dirPath = path.dirname(resolved);
      }
    }
  } catch (e) {
    // Fallback to current working directory
  }
}

// Find all JSON files starting with 'match'
const jsonFiles = fs.readdirSync(dirPath)
  .filter(file => file.startsWith('match') && file.endsWith('.json'))
  .sort((a, b) => {
    if (a === 'match.json') return -1;
    if (b === 'match.json') return 1;
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
  });

if (jsonFiles.length === 0) {
  console.error(`❌ No JSON files starting with 'match' found in: ${dirPath}`);
  process.exit(1);
}

// Write the output to match.xlsx in the target directory
const outputPath = path.join(dirPath, 'match.xlsx');

// Create workbook
const wb = new ExcelJS.Workbook();

// Create Summary sheet as the first sheet
const summaryWs = wb.addWorksheet('Summary');

// Create Overall Data sheet as the second sheet (populated after all match processing)
const overallWs = wb.addWorksheet('Overall Data');

// Common borders and alignment
const border = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } }
};

const centerAlign = { horizontal: 'center', vertical: 'middle' };
const headerAlign = { horizontal: 'center', vertical: 'middle', wrapText: true };

const summaryRowsData = [];

// Collector for the Overall Data sheet
const overallDataRows = [];  // [{ sheetName, headers, lastPlayPerSet }]
let overallHeaders = null;   // column headers from the first match processed

// Process each match*.json file one by one
jsonFiles.forEach(file => {
  const filePath = path.join(dirPath, file);
  const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  
  if (!data || data.length === 0) {
    console.warn(`⚠️ Warning: ${file} is empty, skipping.`);
    return;
  }
  
  const headers = Object.keys(data[0]);
  const sheetName = path.basename(file, '.json');
  
  // Add worksheet
  const ws = wb.addWorksheet(sheetName);
  
  // Set column widths
  ws.columns = headers.map((header, i) => ({
    key: header,
    width: i === 0 ? 6 : 15
  }));
  
  // Add header row
  const headerRow = ws.addRow(headers);
  headerRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = headerAlign;
    cell.border = border;
  });
  headerRow.height = 40;
  
  // Add data rows
  data.forEach(record => {
    const rowValues = headers.map(h => {
      const val = record[h];
      
      // Convert numeric strings to actual numbers to avoid Excel "number stored as text" warnings
      if (typeof val === 'string' && val.trim() !== '' && !isNaN(val)) {
        return Number(val);
      }
      return val;
    });
    
    const row = ws.addRow(rowValues);
    row.eachCell({ includeEmpty: true }, cell => {
      cell.alignment = centerAlign;
      cell.border = border;
    });
    row.height = 25;
  });
  
  // ── Per-Set Final Play Summary ────────────────────────────────────────────
  // Determine the set key (column header named 'Set')
  const setKey = headers.find(h => h.toLowerCase() === 'set') || 'Set';
  
  // Collect the last record for each set (preserving set order)
  const setOrderMap = new Map(); // setName → last record seen
  data.forEach(record => {
    const setName = record[setKey];
    if (setName) {
      setOrderMap.set(setName, record); // overwrites with each later record → ends up as the last one
    }
  });
  
  // Convert to sorted array based on first-appearance order
  const setOrder = [];
  data.forEach(record => {
    const setName = record[setKey];
    if (setName && !setOrder.includes(setName)) {
      setOrder.push(setName);
    }
  });
  
  const lastPlayPerSet = setOrder.map(setName => setOrderMap.get(setName));
  
  // Capture headers from first match; collect rows for Overall Data sheet
  if (!overallHeaders) overallHeaders = headers;
  overallDataRows.push({ sheetName, headers, lastPlayPerSet });
  
  // Leave 3 blank rows after the last data row, then write the section
  const lastDataRowIdx = data.length + 1; // 1-based Excel row of last data row
  const setHeaderRowIdx = lastDataRowIdx + 3; // 3 blank rows gap
  
  // Section header label (merged-style: write only in first cell, style all)
  const setHeaderRow = ws.getRow(setHeaderRowIdx);
  const setHeaderCell = ws.getCell(`A${setHeaderRowIdx}`);
  setHeaderCell.value = 'Final Play of Each Set';
  setHeaderCell.font = { bold: true, size: 12 };
  setHeaderCell.alignment = centerAlign;
  setHeaderCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF404040' } // dark grey background for section header
  };
  setHeaderCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  
  // Style the remaining cells in the header label row to match
  for (let col = 2; col <= headers.length; col++) {
    const cell = ws.getCell(setHeaderRowIdx, col);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF404040' }
    };
    cell.border = border;
  }
  setHeaderCell.border = border;
  setHeaderRow.height = 28;
  
  // Write the column-header row for the set summary (same as main headers)
  const setColHeaderRowIdx = setHeaderRowIdx + 1;
  const setColHeaderRow = ws.addRow(headers); // addRow always appends, use getRow/getCell for precise placement
  // Since addRow appends at current last row, we need to insert at exact position.
  // Re-approach: use ws.getRow() for all set-summary rows.
  ws.spliceRows(setColHeaderRowIdx, 0); // ensure row exists without data
  
  const setColHeaderExcelRow = ws.getRow(setColHeaderRowIdx);
  headers.forEach((h, i) => {
    const cell = setColHeaderExcelRow.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true };
    cell.alignment = headerAlign;
    cell.border = border;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9D9D9' } // light grey for sub-header
    };
  });
  setColHeaderExcelRow.height = 30;
  
  // Write each set's last-play row highlighted with #FFC000
  const SET_SUMMARY_COLOR = 'FFFFC000'; // #FFC000 in ARGB
  
  lastPlayPerSet.forEach((record, idx) => {
    const rowIdx = setColHeaderRowIdx + 1 + idx;
    const excelRow = ws.getRow(rowIdx);
    
    headers.forEach((h, colIdx) => {
      const cell = excelRow.getCell(colIdx + 1);
      const val = record[h];
      cell.value = (typeof val === 'string' && val.trim() !== '' && !isNaN(val))
        ? Number(val)
        : val;
      cell.alignment = centerAlign;
      cell.border = border;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: SET_SUMMARY_COLOR }
      };
    });
    excelRow.height = 25;
  });
  
  // ── 10-6 Trigger Detection & Row Highlighting ─────────────────────────────
  // Colors reused from the rest of the sheet
  const GREEN_HL = 'FFA8F4D0';  // success green (mint)
  const RED_HL   = 'FFF6A6B1';  // fail red (pink)
  // Excel column numbers for F (P1 Pts) and G (P2 Pts) — 1-based
  const COL_F = 6;
  const COL_G = 7;

  const triggeredSets = new Set(); // prevent double-triggering the same set
  let anyTriggerFail    = false;
  let anyTriggerSuccess = false;

  let di = 0;
  while (di < data.length) {
    const rec   = data[di];
    const p1Pts = Number(rec[headers[5]]) || 0;
    const p2Pts = Number(rec[headers[6]]) || 0;
    const curSet = rec[setKey];

    // Only trigger once per set on the exact score 10-6 or 6-10
    const isTrigger =
      (p1Pts === 10 && p2Pts === 6) ||
      (p1Pts === 6  && p2Pts === 10);

    if (isTrigger && !triggeredSets.has(curSet)) {
      triggeredSets.add(curSet);
      const triggerPlayer = p1Pts === 10 ? 1 : 2; // which player reached 10

      // Scan forward within the same set to determine outcome
      let outcomeSuccess = false;
      let hlEnd = di; // last data-index to highlight (inclusive)

      for (let j = di + 1; j < data.length; j++) {
        if (data[j][setKey] !== curSet) break; // set changed — stop
        hlEnd = j;
        const np1 = Number(data[j][headers[5]]) || 0;
        const np2 = Number(data[j][headers[6]]) || 0;

        if (np1 >= 10 && np2 >= 10) {
          // ── Deuce territory: both players at 10+ ──────────────────────────
          // Winner must have a 2-point lead
          const diff = Math.abs(np1 - np2);
          if (diff >= 2) {
            // Determine winner: the player ahead by 2
            outcomeSuccess = (triggerPlayer === 1) ? (np1 > np2) : (np2 > np1);
            break;
          }
          // diff < 2 → still contested, keep scanning
        } else {
          // ── Normal play: first player to reach 11 wins ────────────────────
          if (triggerPlayer === 1 && np1 >= 11) { outcomeSuccess = true;  break; }
          if (triggerPlayer === 2 && np2 >= 11) { outcomeSuccess = true;  break; }
          // Opponent wins without deuce
          if (triggerPlayer === 1 && np2 >= 11) { outcomeSuccess = false; break; }
          if (triggerPlayer === 2 && np1 >= 11) { outcomeSuccess = false; break; }
        }
      }

      if (outcomeSuccess) anyTriggerSuccess = true;
      else               anyTriggerFail    = true;

      const hlColor = outcomeSuccess ? GREEN_HL : RED_HL;

      // Apply fill to columns F & G from trigger row through end of triggered range
      for (let k = di; k <= hlEnd; k++) {
        const excelRowNum = k + 2; // +1 header row, +1 for 1-based
        [COL_F, COL_G].forEach(col => {
          const cell = ws.getCell(excelRowNum, col);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hlColor } };
        });
      }

      di = hlEnd + 1; // jump past the highlighted range
    } else {
      di++;
    }
  }

  // Overall status: FAIL if any triggered set failed; SUCCESS if all succeeded;
  // 'N/A' if no 10-6 trigger was found in any set
  const statusResult = anyTriggerFail
    ? 'Fail'
    : (anyTriggerSuccess ? 'Success' : 'N/A');
  const statusColor  = anyTriggerFail
    ? 'FFF6A6B1'
    : (anyTriggerSuccess ? 'FFA8F4D0' : 'FFDEDEDE');

  // ── MAX Sets + Status summary cells ───────────────────────────────────────
  const lastDataRow = data.length + 1;
  const computedSetHeaderRowIdx    = lastDataRow + 3;
  const computedSetColHeaderRowIdx = computedSetHeaderRowIdx + 1;
  const summaryRow = computedSetColHeaderRowIdx + lastPlayPerSet.length + 2;

  const cellD = ws.getCell(`D${summaryRow}`);
  const cellE = ws.getCell(`E${summaryRow}`);
  const cellK = ws.getCell(`K${summaryRow}`);

  const maxD = Math.max(...data.map(row => Number(row[headers[3]]) || 0));
  const maxE = Math.max(...data.map(row => Number(row[headers[4]]) || 0));

  cellD.value = { formula: `MAX(D2:D${lastDataRow})`, result: maxD };
  cellE.value = { formula: `MAX(E2:E${lastDataRow})`, result: maxE };
  cellK.value = statusResult;  // plain value — no longer driven by J/K columns

  [cellD, cellE, cellK].forEach(cell => {
    cell.alignment = centerAlign;
    cell.border    = border;
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: statusColor } };
  });

  ws.getRow(summaryRow).height = 25;
  
  console.log(`  📄 Added sheet: "${sheetName}" with ${data.length} rows and ${lastPlayPerSet.length} set(s) in summary`);
  
  // Collect details for the Summary sheet
  summaryRowsData.push({
    sheetName,
    player1: headers[3].replace(' Sets', ''),
    player2: headers[4].replace(' Sets', ''),
    maxDFormula: `'${sheetName}'!D${summaryRow}`,
    maxEFormula: `'${sheetName}'!E${summaryRow}`,
    statusFormula: `'${sheetName}'!K${summaryRow}`,
    maxDResult: maxD,
    maxEResult: maxE,
    statusResult: statusResult,
    statusColor: statusColor
  });
});

// --- Populate Overall Data Worksheet ---
if (overallDataRows.length > 0 && overallHeaders) {
  const SET_SUMMARY_COLOR = 'FFFFC000';
  const odHeaders = ['Match', ...overallHeaders];

  // Column widths: 'Match' col gets 15, '#' col gets 6, rest get 15
  overallWs.columns = odHeaders.map((h, i) => ({
    key: String(i),
    width: i === 0 ? 15 : (i === 1 ? 6 : 15)
  }));

  // Single header row at the top
  const odHeaderRow = overallWs.addRow(odHeaders);
  odHeaderRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.alignment = headerAlign;
    cell.border = border;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9D9D9' }
    };
  });
  odHeaderRow.height = 40;

  // One block per match
  overallDataRows.forEach(({ sheetName, headers: matchHeaders, lastPlayPerSet }, matchIdx) => {
    // Write a match-name label row (dark grey banner)
    const labelRow = overallWs.addRow([sheetName]);
    const labelCell = labelRow.getCell(1);
    labelCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    labelCell.alignment = centerAlign;
    labelCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF404040' }
    };
    labelCell.border = border;
    // Fill the rest of the label row in dark grey too
    for (let col = 2; col <= odHeaders.length; col++) {
      const c = labelRow.getCell(col);
      c.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF404040' }
      };
      c.border = border;
    }
    labelRow.height = 22;

    // Write each set's final-play row
    lastPlayPerSet.forEach(record => {
      const rowValues = [
        sheetName,
        ...matchHeaders.map(h => {
          // Use this match's own header keys to read values correctly
          const val = record[h];
          if (val === undefined || val === null) return '';
          return (typeof val === 'string' && val.trim() !== '' && !isNaN(val))
            ? Number(val)
            : val;
        })
      ];

      const dataRow = overallWs.addRow(rowValues);
      dataRow.eachCell({ includeEmpty: true }, cell => {
        cell.alignment = centerAlign;
        cell.border = border;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: SET_SUMMARY_COLOR }
        };
      });
      dataRow.height = 25;
    });

    // Blank separator row between match blocks (except after the last one)
    if (matchIdx < overallDataRows.length - 1) {
      overallWs.addRow([]).height = 10;
    }
  });

  console.log(`  📋 Populated "Overall Data" sheet with ${overallDataRows.reduce((s, m) => s + m.lastPlayPerSet.length, 0)} set-summary rows across ${overallDataRows.length} match(es).`);
} else {
  // No data — leave a placeholder message
  overallWs.addRow(['No match data available']);
}

// --- Populate Summary Worksheet ---
summaryWs.columns = [
  { key: 'match', width: 15 },
  { key: 'p1', width: 20 },
  { key: 'p2', width: 20 },
  { key: 'maxP1', width: 22 },
  { key: 'maxP2', width: 22 },
  { key: 'status', width: 15 }
];

const summaryHeaders = ['Match', 'Player 1', 'Player 2', 'Max Player 1 Sets', 'Max Player 2 Sets', 'Status'];
const summaryHeaderRow = summaryWs.addRow(summaryHeaders);
summaryHeaderRow.eachCell(cell => {
  cell.font = { bold: true };
  cell.alignment = headerAlign;
  cell.border = border;
});
summaryHeaderRow.height = 40;

summaryRowsData.forEach(item => {
  const row = summaryWs.addRow([
    item.sheetName,
    item.player1,
    item.player2,
    { formula: item.maxDFormula, result: item.maxDResult },
    { formula: item.maxEFormula, result: item.maxEResult },
    { formula: item.statusFormula, result: item.statusResult }
  ]);
  
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.alignment = centerAlign;
    cell.border = border;
    
    // Highlight Max Player 1 Sets, Max Player 2 Sets, and Status using same color as sheet
    if (colNumber >= 4) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: item.statusColor }
      };
    }
  });
  row.height = 25;
});

console.log(`  📊 Generated Summary sheet with dashboard overview.`);

// Write workbook to match.xlsx
wb.xlsx.writeFile(outputPath).then(() => {
  console.log(`\n✅ Successfully merged all JSON files into a single Excel file with a Summary dashboard: ${outputPath}`);
}).catch(err => {
  console.error('\n❌ Error writing file:', err.message);
});
