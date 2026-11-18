const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const zlib = require('node:zlib');
const {jsPDF} = require('../vendor/jspdf.umd.min.js');
const autoTable = require('../vendor/jspdf.plugin.autotable.min.js');
const ExcelJS = require('../vendor/exceljs.min.js');

autoTable.applyPlugin(jsPDF);
const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'calculadora_costo_estructura.html'), 'utf8');
const levels = vm.runInNewContext(html.match(/const levelsData = (\[[\s\S]*?\]);/)[1]);
global.window = globalThis;
global.jspdf = {jsPDF};
global.ExcelJS = ExcelJS;
const context = {window: globalThis};
for (const file of ['report-data.js', 'report-pdf.js', 'report-excel.js']) {
  vm.runInThisContext(fs.readFileSync(path.join(root, file), 'utf8'), {filename: file});
}
for (const match of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) {
  if (match[1].trim()) new vm.Script(match[1]);
}
new vm.Script(fs.readFileSync(path.join(root, 'report-downloads.js'), 'utf8'));

function zipText(buffer, name) {
  let end = buffer.length - 22;
  while (end >= 0 && buffer.readUInt32LE(end) !== 0x06054b50) end--;
  assert(end >= 0, 'ZIP end record');
  let offset = buffer.readUInt32LE(end + 16);
  while (buffer.readUInt32LE(offset) === 0x02014b50) {
    const method = buffer.readUInt16LE(offset + 10);
    const size = buffer.readUInt32LE(offset + 20);
    const length = buffer.readUInt16LE(offset + 28);
    const local = buffer.readUInt32LE(offset + 42);
    const filename = buffer.subarray(offset + 46, offset + 46 + length).toString();
    if (filename === name) {
      const start = local + 30 + buffer.readUInt16LE(local + 26) + buffer.readUInt16LE(local + 28);
      const bytes = buffer.subarray(start, start + size);
      return (method === 8 ? zlib.inflateRawSync(bytes) : bytes).toString('utf8');
    }
    offset += 46 + length + buffer.readUInt16LE(offset + 30) + buffer.readUInt16LE(offset + 32);
  }
  throw new Error('Missing ZIP member: ' + name);
}

async function checkReport(count, stress = false) {
  const quantities = levels.map((_, index) => index < count ? (stress ? 999999 : index + 1) : 0);
  const report = context.window.StructureReport.createReport({
    levels, quantities,
    organization: 'DIRECCIÓN GENERAL DE ADMINISTRACIÓN Y FINANZAS DE LA SECRETARÍA DE DESARROLLO URBANO Y VIVIENDA. '.repeat(stress ? 5 : 1),
    validatedBy: 'María de los Ángeles López Hernández'.repeat(stress ? 5 : 1),
    generatedAt: new Date('2026-11-18T18:30:00Z')
  });
  const expectedExcelPages = count <= 8 ? 1 : 2;
  assert.equal(report.rows.length, count);
  assert.equal(report.pageCount, expectedExcelPages);
  assert.equal(report.totals.positions, quantities.reduce((a, b) => a + b, 0));
  const annualCents = levels.reduce((sum, item, index) => sum + Math.round(item.annual * 100) * quantities[index], 0);
  assert.equal(report.totals.annual, annualCents / 100);
  assert.match(report.generatedAtLabel, /12:30/);

  const doc = context.window.StructurePDF.createDocument(report);
  assert.equal(doc.getNumberOfPages(), 1, 'PDF always has one page: ' + count);
  assert.equal(doc.structureLayout.rowCount, count);
  assert(doc.structureLayout.signatureBottom <= 198, 'Signature and date stay above the page footer');
  assert.equal(doc.lastAutoTable.body.length, Math.max(count, 1), 'All levels remain in the PDF');
  if (count) assert.deepEqual(
    Array.from(doc.lastAutoTable.body, row => row.cells[0].text.join('')),
    Array.from(report.rows, row => row.level),
    'Each selected level appears once in the table'
  );
  assert(doc.lastAutoTable.finalY < 175, 'Room for signature');
  assert(Math.abs(doc.internal.pageSize.getWidth() - 297) < .1);
  assert(Math.abs(doc.internal.pageSize.getHeight() - 210) < .1);
  const pdfText = doc.internal.pages.flat().join('\n');
  assert(!pdfText.includes('C?lculo'), 'UTF-8 title preserved');

  const workbook = context.window.StructureExcel.createWorkbook(report);
  assert.equal(workbook.worksheets.length, 1);
  const sheet = workbook.worksheets[0];
  assert.equal(sheet.pageSetup.fitToWidth, 1);
  assert.equal(sheet.pageSetup.fitToHeight, expectedExcelPages);
  assert.equal(sheet.pageSetup.orientation, 'landscape');
  assert.equal(sheet.pageSetup.paperSize, 9);
  assert.equal(sheet.pageSetup.printTitlesRow, '9:9');
  assert.equal(sheet.getCell('A3').value, 'Organismo: ' + report.organization);
  const totalRow = 10 + Math.max(count, 1);
  assert.equal(sheet.getCell('G' + totalRow).alignment.shrinkToFit, true);
  assert.equal(sheet.getCell('G' + totalRow).result ?? sheet.getCell('G' + totalRow).value, report.totals.annual);
  for (let index = 0; index < count; index++) {
    assert.equal(sheet.getCell('F' + (10 + index)).value, quantities[index]);
    assert.equal(sheet.getCell('G' + (10 + index)).formula, 'C' + (10 + index) + '*F' + (10 + index));
    assert.equal(sheet.getCell('G' + (10 + index)).alignment.shrinkToFit, true);
  }
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  const saved = new ExcelJS.Workbook();
  await saved.xlsx.load(buffer);
  assert.equal(saved.worksheets[0].getCell('A3').value, 'Organismo: ' + report.organization);
  const sheetXML = zipText(buffer, 'xl/worksheets/sheet1.xml');
  assert.match(sheetXML, /fitToWidth="1"/);
  assert(sheetXML.includes('fitToHeight="' + expectedExcelPages + '"'));
  assert.match(sheetXML, /orientation="landscape"/);
  if (expectedExcelPages === 2) assert.match(sheetXML, /<rowBreaks/);
  const workbookXML = zipText(buffer, 'xl/workbook.xml');
  assert(workbookXML.includes('_xlnm.Print_Area'));
  assert(workbookXML.includes('_xlnm.Print_Titles'));

  if (process.argv.includes('--samples')) {
    const directory = path.join(root, 'tmp', 'exports');
    fs.mkdirSync(directory, {recursive: true});
    const name = 'report-' + count + (stress ? '-stress' : '');
    fs.writeFileSync(path.join(directory, name + '.pdf'), Buffer.from(doc.output('arraybuffer')));
    fs.writeFileSync(path.join(directory, name + '.xlsx'), buffer);
  }
  console.log('OK: ' + count + ' levels' + (stress ? ' (long names and large amounts)' : '') + ', PDF one page with every level, Excel formulas and print settings');
}

async function checkDownloadButtons() {
  const calls = [];
  const elements = {
    downloadPDFButton: {innerHTML: 'Descargar PDF', setAttribute() {}, removeAttribute() {}},
    downloadExcelButton: {innerHTML: 'Descargar Excel', setAttribute() {}, removeAttribute() {}},
    exportStatus: {hidden: true}
  };
  const browser = {
    console: {error() {}},
    document: {getElementById: id => elements[id]},
    window: {
      jspdf: {jsPDF}, ExcelJS,
      StructurePDF: {download: report => { calls.push(['pdf', report]); }},
      StructureExcel: {download: report => { calls.push(['excel', report]); }}
    }
  };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'report-downloads.js'), 'utf8'), browser);
  const report = {rows: []};
  await browser.window.StructureDownloads.download('pdf', report);
  await browser.window.StructureDownloads.download('excel', report);
  assert.deepEqual(calls.map(call => call[0]), ['pdf', 'excel']);
  assert.equal(elements.downloadPDFButton.disabled, false);
  assert.equal(elements.downloadExcelButton.disabled, false);
  assert.equal(elements.exportStatus.hidden, false);
  browser.window.StructurePDF.download = () => { throw new Error('Simulated export failure'); };
  await browser.window.StructureDownloads.download('pdf', report);
  assert.equal(elements.downloadPDFButton.disabled, false);
  assert.equal(elements.downloadPDFButton.innerHTML, 'Descargar PDF');
  assert.match(elements.exportStatus.textContent, /No se pudo generar/);
  console.log('OK: download buttons route formats, restore state and report errors');
}

(async () => {
  for (const count of [0, 1, 8, 9, 19]) await checkReport(count);
  await checkReport(8, true);
  await checkReport(19, true);
  await checkDownloadButtons();
})().catch(error => { console.error(error); process.exitCode = 1; });
