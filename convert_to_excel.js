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
  
  // Add summary cells (2 rows after the last data row)
  const lastDataRow = data.length + 1;
  const summaryRow = lastDataRow + 2;
  
  const cellD = ws.getCell(`D${summaryRow}`);
  const cellE = ws.getCell(`E${summaryRow}`);
  const cellK = ws.getCell(`K${summaryRow}`);
  
  // Evaluate Success / Fails color programmatically
  const lastDataRecord = data[data.length - 1];
  const valJ = Number(lastDataRecord[headers[9]]) || 0;
  const valK = Number(lastDataRecord[headers[10]]) || 0;
  
  const maxD = Math.max(...data.map(row => Number(row[headers[3]]) || 0));
  const maxE = Math.max(...data.map(row => Number(row[headers[4]]) || 0));
  
  const isSuccess = valJ >= valK ? (maxD > 0) : (maxE > 0);
  const statusResult = isSuccess ? 'Success' : 'Fails';
  const statusColor = isSuccess ? 'FFA8F4D0' : 'FFF6A6B1';
  
  cellD.value = { formula: `MAX(D2:D${lastDataRow})`, result: maxD };
  cellE.value = { formula: `MAX(E2:E${lastDataRow})`, result: maxE };
  cellK.value = { formula: `IF(J${lastDataRow}>=K${lastDataRow}, IF(D${summaryRow}>0, "Success", "Fails"), IF(E${summaryRow}>0, "Success", "Fails"))`, result: statusResult };
  
  // Style summary cells (center alignment, border, and conditional background color)
  const summaryCells = [cellD, cellE, cellK];
  summaryCells.forEach(cell => {
    cell.alignment = centerAlign;
    cell.border = border;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: statusColor }
    };
  });
  
  ws.getRow(summaryRow).height = 25;
  
  console.log(`  📄 Added sheet: "${sheetName}" with ${data.length} rows`);
  
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
