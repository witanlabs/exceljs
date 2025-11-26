const ExcelJS = verquire('exceljs');
const TableXform = verquire('xlsx/xform/table/table-xform');

describe('github issues', () => {
  describe('issue 1062 - table rows not loaded from xlsx', () => {
    it('should populate table rows when loading from xlsx', async () => {
      // Create a workbook with a table
      const wb1 = new ExcelJS.Workbook();
      const ws1 = wb1.addWorksheet('Sheet1');

      ws1.addTable({
        name: 'TestTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        columns: [
          {name: 'Name', filterButton: true},
          {name: 'Value', filterButton: true},
        ],
        rows: [
          ['Alice', 100],
          ['Bob', 200],
          ['Charlie', 300],
        ],
      });

      // Write to buffer
      const buffer = await wb1.xlsx.writeBuffer();

      // Read back
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet('Sheet1');
      const table = ws2.getTable('TestTable');

      // Verify table rows are populated
      expect(table.model.rows).to.have.lengthOf(3);
      expect(table.model.rows[0]).to.deep.equal(['Alice', 100]);
      expect(table.model.rows[1]).to.deep.equal(['Bob', 200]);
      expect(table.model.rows[2]).to.deep.equal(['Charlie', 300]);
    });

    it('should allow adding rows to a loaded table', async () => {
      // Create a workbook with a table
      const wb1 = new ExcelJS.Workbook();
      const ws1 = wb1.addWorksheet('Sheet1');

      ws1.addTable({
        name: 'TestTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        columns: [
          {name: 'Name', filterButton: true},
          {name: 'Value', filterButton: true},
        ],
        rows: [
          ['Alice', 100],
          ['Bob', 200],
        ],
      });

      // Write to buffer
      const buffer = await wb1.xlsx.writeBuffer();

      // Read back
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet('Sheet1');
      const table = ws2.getTable('TestTable');

      // Add a new row
      table.addRow(['Charlie', 300]);

      // Verify the row was added
      expect(table.model.rows).to.have.lengthOf(3);
      expect(table.model.rows[2]).to.deep.equal(['Charlie', 300]);

      // Verify worksheet cells are correct
      expect(ws2.getCell('A4').value).to.equal('Charlie');
      expect(ws2.getCell('B4').value).to.equal(300);
    });

    it('should preserve original data when loading table', async () => {
      // Create a workbook with a table
      const wb1 = new ExcelJS.Workbook();
      const ws1 = wb1.addWorksheet('Sheet1');

      ws1.addTable({
        name: 'TestTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        columns: [
          {name: 'Name', filterButton: true},
          {name: 'Value', filterButton: true},
        ],
        rows: [
          ['Alice', 100],
          ['Bob', 200],
        ],
      });

      // Write to buffer
      const buffer = await wb1.xlsx.writeBuffer();

      // Read back
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet('Sheet1');

      // Verify worksheet cells have correct data
      expect(ws2.getCell('A1').value).to.equal('Name');
      expect(ws2.getCell('B1').value).to.equal('Value');
      expect(ws2.getCell('A2').value).to.equal('Alice');
      expect(ws2.getCell('B2').value).to.equal(100);
      expect(ws2.getCell('A3').value).to.equal('Bob');
      expect(ws2.getCell('B3').value).to.equal(200);
    });

    it('should handle table with totals row', async () => {
      // Create a workbook with a table that has totals
      const wb1 = new ExcelJS.Workbook();
      const ws1 = wb1.addWorksheet('Sheet1');

      ws1.addTable({
        name: 'TestTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: true,
        columns: [
          {name: 'Name', totalsRowLabel: 'Total', filterButton: true},
          {name: 'Value', totalsRowFunction: 'sum', filterButton: true},
        ],
        rows: [
          ['Alice', 100],
          ['Bob', 200],
        ],
      });

      // Write to buffer
      const buffer = await wb1.xlsx.writeBuffer();

      // Read back
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet('Sheet1');
      const table = ws2.getTable('TestTable');

      // Verify only data rows are in table.rows (not header or totals)
      expect(table.model.rows).to.have.lengthOf(2);
      expect(table.model.rows[0]).to.deep.equal(['Alice', 100]);
      expect(table.model.rows[1]).to.deep.equal(['Bob', 200]);
    });

    it('should default headerRow to true when headerRowCount is omitted (OOXML spec)', () => {
      // This tests the fix for Excel-generated files that omit headerRowCount
      // OOXML spec says headerRowCount defaults to 1, meaning headerRow should be true
      const xform = new TableXform();

      // Simulate parsing table XML without headerRowCount attribute (as Excel generates)
      const mockNode = {
        name: 'table',
        attributes: {
          name: 'Table1',
          displayName: 'Table1',
          ref: 'A1:B4',
          // Note: no headerRowCount attribute - this is what Excel typically generates
        },
      };

      xform.parseOpen(mockNode);
      const {model} = xform;

      // headerRow should be true per OOXML default
      expect(model.headerRow).to.be.true();
    });

    it('should set headerRow to false when headerRowCount is explicitly 0', () => {
      const xform = new TableXform();

      const mockNode = {
        name: 'table',
        attributes: {
          name: 'Table1',
          displayName: 'Table1',
          ref: 'A1:B4',
          headerRowCount: '0',
        },
      };

      xform.parseOpen(mockNode);
      const {model} = xform;

      // headerRow should be false when explicitly set to 0
      expect(model.headerRow).to.be.false();
    });
  });
});
