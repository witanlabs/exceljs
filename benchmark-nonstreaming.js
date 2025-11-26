/* eslint-disable no-console */
/**
 * Non-Streaming Benchmark for ExcelJS
 *
 * This benchmark exercises the non-streaming API which is our primary use case.
 * Run with: node --expose-gc benchmark-nonstreaming.js
 *
 * Operations tested:
 * 1. Read small XLSX file
 * 2. Read large XLSX file
 * 3. Write small workbook
 * 4. Write medium workbook (1000 rows)
 * 5. Write large workbook (10000 rows with styles)
 * 6. Cell operations at scale
 * 7. Merged cells operations
 * 8. Conditional formatting
 * 9. Tables
 * 10. Round-trip (read + modify + write)
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('./lib/exceljs.nodejs.js');

const RUNS = 3;
const OUTPUT_DIR = './spec/out';
const RESULTS_FILE = './benchmark-results.json';

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

const results = {
  timestamp: new Date().toISOString(),
  nodeVersion: process.version,
  platform: process.platform,
  arch: process.arch,
  benchmarks: {}
};

async function main() {
  console.log('='.repeat(70));
  console.log('ExcelJS Non-Streaming Benchmark');
  console.log('='.repeat(70));
  console.log(`Node ${process.version} | ${process.platform} ${process.arch}`);
  console.log(`Date: ${results.timestamp}`);
  console.log('='.repeat(70));

  try {
    // 1. Read small file
    await runBenchmark('read_small_xlsx', 'Read small XLSX (gold.xlsx ~9KB)', async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile('./spec/integration/data/gold.xlsx');
      return { worksheets: wb.worksheets.length };
    });

    // 2. Read medium file
    await runBenchmark('read_medium_xlsx', 'Read medium XLSX (images.xlsx ~23KB)', async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile('./spec/integration/data/images.xlsx');
      return { worksheets: wb.worksheets.length };
    });

    // 3. Read large file (if available)
    if (fs.existsSync('./spec/integration/data/huge.xlsx')) {
      await runBenchmark('read_large_xlsx', 'Read large XLSX (huge.xlsx ~14MB)', async () => {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile('./spec/integration/data/huge.xlsx');
        let rowCount = 0;
        wb.eachSheet(ws => {
          ws.eachRow(() => { rowCount++; });
        });
        return { worksheets: wb.worksheets.length, rows: rowCount };
      });
    }

    // 4. Write small workbook
    await runBenchmark('write_small_xlsx', 'Write small XLSX (100 rows, basic)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Data');

      // Add headers
      ws.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 30 },
        { header: 'Value', key: 'value', width: 15 },
        { header: 'Date', key: 'date', width: 15 },
      ];

      // Add 100 rows
      for (let i = 1; i <= 100; i++) {
        ws.addRow({ id: i, name: `Item ${i}`, value: Math.random() * 1000, date: new Date() });
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-small.xlsx`);
      return { rows: 100 };
    });

    // 5. Write medium workbook with styles
    await runBenchmark('write_medium_styled_xlsx', 'Write medium XLSX (1000 rows with styles)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('StyledData');

      ws.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 30 },
        { header: 'Amount', key: 'amount', width: 15 },
        { header: 'Status', key: 'status', width: 15 },
        { header: 'Date', key: 'date', width: 15 },
      ];

      // Style header row
      ws.getRow(1).font = { bold: true };
      ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };

      for (let i = 1; i <= 1000; i++) {
        const row = ws.addRow({
          id: i,
          name: `Item ${i}`,
          amount: Math.random() * 10000,
          status: i % 3 === 0 ? 'Active' : 'Inactive',
          date: new Date()
        });

        // Apply alternating row colors
        if (i % 2 === 0) {
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        }

        // Number format for amount
        row.getCell('amount').numFmt = '$#,##0.00';
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-medium-styled.xlsx`);
      return { rows: 1000 };
    });

    // 6. Write large workbook
    await runBenchmark('write_large_xlsx', 'Write large XLSX (10000 rows, 10 columns)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('LargeData');

      const columns = [];
      for (let c = 1; c <= 10; c++) {
        columns.push({ header: `Column${c}`, key: `col${c}`, width: 15 });
      }
      ws.columns = columns;

      for (let r = 1; r <= 10000; r++) {
        const rowData = {};
        for (let c = 1; c <= 10; c++) {
          rowData[`col${c}`] = `R${r}C${c}`;
        }
        ws.addRow(rowData);
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-large.xlsx`);
      return { rows: 10000, cols: 10 };
    });

    // 7. Cell operations benchmark
    await runBenchmark('cell_operations', 'Cell operations (5000 individual cell writes)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Cells');

      for (let r = 1; r <= 100; r++) {
        for (let c = 1; c <= 50; c++) {
          const cell = ws.getCell(r, c);
          cell.value = `${r}-${c}`;
          cell.font = { name: 'Arial', size: 10 };
        }
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-cells.xlsx`);
      return { cells: 5000 };
    });

    // 8. Merged cells benchmark
    await runBenchmark('merged_cells', 'Merged cells (500 merge operations)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Merged');

      // Fill some data first
      for (let r = 1; r <= 1000; r++) {
        for (let c = 1; c <= 10; c++) {
          ws.getCell(r, c).value = `${r}-${c}`;
        }
      }

      // Merge cells in groups
      for (let i = 0; i < 500; i++) {
        const startRow = i * 2 + 1;
        ws.mergeCells(startRow, 1, startRow + 1, 1);
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-merged.xlsx`);
      return { merges: 500 };
    });

    // 9. Conditional formatting benchmark
    await runBenchmark('conditional_formatting', 'Conditional formatting (100 rules)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('CF');

      // Add data
      for (let r = 1; r <= 1000; r++) {
        ws.getCell(r, 1).value = Math.random() * 100;
      }

      // Add conditional formatting rules
      for (let i = 0; i < 100; i++) {
        ws.addConditionalFormatting({
          ref: `A${i * 10 + 1}:A${i * 10 + 10}`,
          rules: [
            {
              type: 'cellIs',
              operator: 'greaterThan',
              formulae: [50],
              style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } },
            },
          ],
        });
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-cf.xlsx`);
      return { rules: 100 };
    });

    // 10. Tables benchmark
    await runBenchmark('tables', 'Tables (create table with 500 rows)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Tables');

      // Add table data
      const rows = [['Name', 'Amount', 'Date', 'Status']];
      for (let i = 1; i <= 500; i++) {
        rows.push([`Item ${i}`, Math.random() * 1000, new Date(), i % 2 === 0 ? 'Active' : 'Inactive']);
      }

      ws.addTable({
        name: 'SalesTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: true,
        style: {
          theme: 'TableStyleMedium2',
          showRowStripes: true,
        },
        columns: [
          { name: 'Name', totalsRowLabel: 'Total:', filterButton: true },
          { name: 'Amount', totalsRowFunction: 'sum', filterButton: true },
          { name: 'Date', filterButton: true },
          { name: 'Status', filterButton: true },
        ],
        rows: rows.slice(1),
      });

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-tables.xlsx`);
      return { tableRows: 500 };
    });

    // 11. Round-trip benchmark (read, modify, write)
    await runBenchmark('round_trip', 'Round-trip (read + modify + write)', async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile('./spec/integration/data/gold.xlsx');

      // Modify
      const ws = wb.getWorksheet(1);
      for (let r = 1; r <= 100; r++) {
        ws.getCell(r, 10).value = `Added ${r}`;
      }

      // Add new worksheet
      const ws2 = wb.addWorksheet('New Sheet');
      for (let r = 1; r <= 100; r++) {
        ws2.getCell(r, 1).value = r;
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-roundtrip.xlsx`);
      return { modified: true };
    });

    // 12. Formulas benchmark
    await runBenchmark('formulas', 'Formulas (1000 cells with formulas)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Formulas');

      // Add source data
      for (let r = 1; r <= 1000; r++) {
        ws.getCell(r, 1).value = Math.random() * 100;
        ws.getCell(r, 2).value = Math.random() * 100;
      }

      // Add formulas
      for (let r = 1; r <= 1000; r++) {
        ws.getCell(r, 3).value = { formula: `A${r}+B${r}` };
        ws.getCell(r, 4).value = { formula: `A${r}*B${r}` };
      }

      await wb.xlsx.writeFile(`${OUTPUT_DIR}/bench-formulas.xlsx`);
      return { formulas: 2000 };
    });

    // Print summary
    console.log('\n' + '='.repeat(70));
    console.log('SUMMARY');
    console.log('='.repeat(70));

    const summaryTable = [];
    for (const [name, data] of Object.entries(results.benchmarks)) {
      summaryTable.push({
        Benchmark: name,
        'Avg Time (ms)': data.avgMs.toFixed(2),
        'Min (ms)': data.minMs.toFixed(2),
        'Max (ms)': data.maxMs.toFixed(2),
        'Avg Mem (MB)': data.avgMemMB.toFixed(2),
      });
    }
    console.table(summaryTable);

    // Save results to file
    const existingResults = loadExistingResults();
    existingResults.push(results);
    // Keep last 10 runs
    while (existingResults.length > 10) {
      existingResults.shift();
    }
    fs.writeFileSync(RESULTS_FILE, JSON.stringify(existingResults, null, 2));
    console.log(`\nResults saved to ${RESULTS_FILE}`);

    // Compare with baseline if available
    if (existingResults.length > 1) {
      console.log('\n' + '='.repeat(70));
      console.log('COMPARISON WITH PREVIOUS RUN');
      console.log('='.repeat(70));

      const previous = existingResults[existingResults.length - 2];
      const comparison = [];

      for (const [name, current] of Object.entries(results.benchmarks)) {
        const prev = previous.benchmarks[name];
        if (prev) {
          const timeDiff = ((current.avgMs - prev.avgMs) / prev.avgMs * 100).toFixed(1);
          const memDiff = ((current.avgMemMB - prev.avgMemMB) / prev.avgMemMB * 100).toFixed(1);
          comparison.push({
            Benchmark: name,
            'Prev (ms)': prev.avgMs.toFixed(2),
            'Curr (ms)': current.avgMs.toFixed(2),
            'Time %': `${timeDiff > 0 ? '+' : ''}${timeDiff}%`,
            'Mem %': `${memDiff > 0 ? '+' : ''}${memDiff}%`,
          });
        }
      }
      console.table(comparison);
    }

  } catch (err) {
    console.error('Benchmark failed:', err);
    process.exit(1);
  }
}

function loadExistingResults() {
  try {
    if (fs.existsSync(RESULTS_FILE)) {
      return JSON.parse(fs.readFileSync(RESULTS_FILE, 'utf8'));
    }
  } catch (e) {
    console.warn('Could not load existing results:', e.message);
  }
  return [];
}

async function runBenchmark(name, description, fn) {
  console.log('\n' + '-'.repeat(70));
  console.log(`BENCHMARK: ${description}`);
  console.log('-'.repeat(70));

  // Warmup
  if (global.gc) global.gc();
  console.log('Warmup run...');
  const warmupStart = Date.now();
  await fn();
  console.log(`Warmup: ${Date.now() - warmupStart}ms`);

  const times = [];
  const memories = [];

  for (let i = 1; i <= RUNS; i++) {
    if (global.gc) global.gc();
    const memBefore = process.memoryUsage().heapUsed;

    const start = Date.now();
    const result = await fn();
    const elapsed = Date.now() - start;

    const memAfter = process.memoryUsage().heapUsed;
    const memUsed = (memAfter - memBefore) / 1024 / 1024;

    times.push(elapsed);
    memories.push(memUsed);

    console.log(`Run ${i}: ${elapsed}ms | Memory: ${memUsed.toFixed(2)}MB | ${JSON.stringify(result)}`);
  }

  const avgMs = times.reduce((a, b) => a + b, 0) / times.length;
  const minMs = Math.min(...times);
  const maxMs = Math.max(...times);
  const avgMemMB = memories.reduce((a, b) => a + b, 0) / memories.length;

  results.benchmarks[name] = {
    description,
    avgMs,
    minMs,
    maxMs,
    avgMemMB,
    runs: RUNS,
  };

  console.log(`Average: ${avgMs.toFixed(2)}ms | Memory: ${avgMemMB.toFixed(2)}MB`);
}

main();
