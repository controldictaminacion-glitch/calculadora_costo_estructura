(function (global) {
  'use strict';

  const COLORS = {
    guinda: 'FF9F2241', guindaDark: 'FF691C32', guindaHover: 'FF821B35',
    dorado: 'FFBC955C', doradoLight: 'FFD4C19C', doradoDark: 'FF9A7740',
    fondo: 'FFF4F1EA', blanco: 'FFFFFFFF', texto: 'FF0F172A', secundario: 'FF475569',
    borde: 'FFE8DAC6', suave: 'FFF1F5F9', alterno: 'FFF8FAFC'
  };
  const MONEY_FORMAT = '"$"#,##0.00';
  const FIRST_DATA_ROW = 10;

  function fill(argb) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
  }

  function font(color, size, bold) {
    return { name: 'Arial', size: size || 10, color: { argb: color || COLORS.texto }, bold: !!bold };
  }

  function mergedText(sheet, range, value, options) {
    sheet.mergeCells(range);
    const cell = sheet.getCell(range.split(':')[0]);
    cell.value = value;
    cell.font = font(COLORS.texto, 10);
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    Object.assign(cell, options || {});
    return cell;
  }

  function createWorkbook(report) {
    if (!global.ExcelJS) throw new Error('No se pudo cargar la herramienta de Excel. Recargue la página e intente nuevamente.');

    const workbook = new global.ExcelJS.Workbook();
    workbook.creator = 'DGDYPO';
    workbook.title = report.title;
    workbook.subject = 'Costo de estructura';
    workbook.created = report.generatedAt instanceof Date ? report.generatedAt : new Date();
    workbook.modified = workbook.created;
    workbook.calcProperties.fullCalcOnLoad = true;

    const rows = report.rows;
    const pageCount = report.pageCount === 2 || (report.pageCount !== 1 && rows.length > 8) ? 2 : 1;
    const sheet = workbook.addWorksheet('Costo de estructura', {
      properties: { tabColor: { argb: COLORS.guinda }, defaultRowHeight: 18 },
      views: [{ state: 'frozen', xSplit: 2, ySplit: 9, showGridLines: false }],
      pageSetup: {
        paperSize: 9,
        orientation: 'landscape',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: pageCount,
        horizontalCentered: true,
        verticalCentered: false,
        showGridLines: false,
        showRowColHeaders: false,
        margins: { left: 0.3, right: 0.3, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 }
      },
      headerFooter: { oddFooter: '&L&"Arial,Regular"&8&K475569DGDYPO&CPágina &P de &N' }
    });
    sheet.columns = [6, 30, 17, 17, 17, 10, 18, 18, 18].map(width => ({ width }));

    mergedText(sheet, 'A1:I1', report.title, { fill: fill(COLORS.guindaDark), font: font(COLORS.blanco, 18, true) });
    sheet.getRow(1).height = 28;
    mergedText(sheet, 'A2:I2', 'Gobierno de la Ciudad de México · DGDYPO', { font: font(COLORS.secundario, 10) });
    sheet.getRow(2).height = 18;

    const organization = 'Organismo: ' + (report.organization || 'No especificado');
    mergedText(sheet, 'A3:I3', organization, { fill: fill(COLORS.suave), font: font(COLORS.secundario, 10, true) });
    sheet.getRow(3).height = Math.max(30, Math.ceil(organization.length / 120) * 13 + 8);
    mergedText(sheet, 'A4:I4', 'Fecha de elaboración: ' + report.generatedAtLabel, { font: font(COLORS.secundario, 9) });
    sheet.getRow(4).height = 18;
    sheet.getRow(5).height = 8;

    const summaryGroups = [
      ['A', 'B', 'Total de plazas', 'F', report.totals.positions],
      ['C', 'D', 'Costo anual total', 'G', report.totals.annual],
      ['E', 'F', 'Mensual bruto total', 'H', report.totals.gross],
      ['G', 'I', 'Mensual neto aproximado', 'I', report.totals.net]
    ];
    const lastDataRow = FIRST_DATA_ROW + Math.max(1, rows.length) - 1;
    const totalRowNumber = lastDataRow + 1;
    summaryGroups.forEach(([start, end, label, totalColumn, result], index) => {
      const featured = index === 3;
      mergedText(sheet, `${start}6:${end}6`, label, { fill: fill(featured ? COLORS.guindaDark : COLORS.alterno), font: font(featured ? COLORS.blanco : COLORS.secundario, 9, true) });
      const value = mergedText(sheet, `${start}7:${end}7`, { formula: `${totalColumn}${totalRowNumber}`, result }, { fill: fill(featured ? COLORS.guindaDark : COLORS.alterno), font: font(featured ? COLORS.blanco : COLORS.guinda, 13, true) });
      value.numFmt = index === 0 ? '#,##0' : MONEY_FORMAT;
    });
    sheet.getRow(6).height = 22;
    sheet.getRow(7).height = 24;
    sheet.getRow(8).height = 8;

    const header = sheet.getRow(9);
    header.values = ['Nivel', 'Denominación general', 'Costo anual', 'Costo mensual\nbruto', 'Costo mensual\nneto aproximado', 'Cantidad de\nplazas', 'Subtotal anual', 'Subtotal mensual\nbruto', 'Subtotal mensual\nneto'];
    header.height = 42;
    header.eachCell(cell => {
      cell.fill = fill(COLORS.guindaDark);
      cell.font = font(COLORS.blanco, 9, true);
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { bottom: { style: 'medium', color: { argb: COLORS.dorado } } };
    });

    rows.forEach((item, index) => {
      const rowNumber = FIRST_DATA_ROW + index;
      const row = sheet.getRow(rowNumber);
      row.values = [
        String(item.level), item.designation, item.annual, item.gross, item.net, item.quantity,
        { formula: `C${rowNumber}*F${rowNumber}`, result: item.annualSubtotal },
        { formula: `D${rowNumber}*F${rowNumber}`, result: item.grossSubtotal },
        { formula: `E${rowNumber}*F${rowNumber}`, result: item.netSubtotal }
      ];
      row.height = Math.max(36, Math.ceil(item.designation.length / 30) * 12 + 8);
      row.eachCell((cell, column) => {
        cell.font = font(column >= 7 ? COLORS.guinda : COLORS.texto, 10, column === 1 || column === 9);
        cell.fill = fill(index % 2 ? COLORS.alterno : COLORS.blanco);
        cell.alignment = { vertical: 'middle', horizontal: column === 2 ? 'left' : (column === 1 || column === 6 ? 'center' : 'right'), wrapText: column === 2, shrinkToFit: column > 2 };
        cell.border = { bottom: { style: 'thin', color: { argb: COLORS.borde } } };
        if (column === 6 || column === 7) cell.border.left = { style: 'thin', color: { argb: COLORS.dorado } };
        cell.numFmt = column === 1 ? '@' : (column === 6 ? '#,##0' : (column > 2 ? MONEY_FORMAT : 'General'));
      });
    });

    if (!rows.length) {
      mergedText(sheet, `A${FIRST_DATA_ROW}:I${FIRST_DATA_ROW}`, 'Sin plazas capturadas.', { font: font(COLORS.texto, 10) });
      sheet.getRow(FIRST_DATA_ROW).height = 36;
    }

    const totalRow = sheet.getRow(totalRowNumber);
    mergedText(sheet, `A${totalRowNumber}:E${totalRowNumber}`, 'TOTALES', { font: font(COLORS.guinda, 10, true) });
    ['F', 'G', 'H', 'I'].forEach((column, index) => {
      const cached = [report.totals.positions, report.totals.annual, report.totals.gross, report.totals.net][index];
      const cell = sheet.getCell(`${column}${totalRowNumber}`);
      cell.value = rows.length ? { formula: `SUM(${column}${FIRST_DATA_ROW}:${column}${lastDataRow})`, result: cached } : 0;
      cell.numFmt = index === 0 ? '#,##0' : MONEY_FORMAT;
      cell.alignment = { horizontal: index === 0 ? 'center' : 'right', vertical: 'middle', wrapText: false, shrinkToFit: true };
    });
    totalRow.height = 28;
    totalRow.eachCell({ includeEmpty: true }, cell => {
      cell.fill = fill(COLORS.suave);
      cell.font = font(COLORS.guinda, 10, true);
      cell.border = { top: { style: 'medium', color: { argb: COLORS.dorado } }, bottom: { style: 'thin', color: { argb: COLORS.borde } } };
    });

    const validationRow = totalRowNumber + 2;
    sheet.getRow(totalRowNumber + 1).height = 10;
    mergedText(sheet, `A${validationRow}:E${validationRow}`, 'Validó por', { font: font(COLORS.guinda, 10, true) });
    mergedText(sheet, `F${validationRow}:I${validationRow}`, 'Firma', { font: font(COLORS.guinda, 10, true) });
    sheet.getRow(validationRow).height = 18;
    const nameRow = validationRow + 1;
    const name = report.validatedBy || '________________________________';
    const nameCell = mergedText(sheet, `A${nameRow}:E${nameRow}`, name);
    const signature = mergedText(sheet, `F${nameRow}:I${nameRow}`, ' ');
    [nameCell, signature].forEach(cell => { cell.border = { bottom: { style: 'thin', color: { argb: COLORS.borde } } }; });
    sheet.getRow(nameRow).height = Math.max(32, Math.ceil(name.length / 75) * 12 + 8);
    const lastRow = nameRow + 1;
    mergedText(sheet, `A${lastRow}:I${lastRow}`, 'Fecha y hora de elaboración del reporte: ' + report.generatedAtLabel, { font: font(COLORS.secundario, 9) });
    sheet.getRow(lastRow).height = 20;

    sheet.pageSetup.printArea = `A1:I${lastRow}`;
    sheet.pageSetup.printTitlesRow = '9:9';
    if (pageCount === 2 && rows.length > 1) {
      const requestedSplit = Number.isFinite(report.rowsPerPage) ? Math.floor(report.rowsPerPage) : Math.ceil(rows.length / 2);
      const split = Math.max(1, Math.min(rows.length - 1, requestedSplit));
      sheet.getRow(FIRST_DATA_ROW + split - 1).addPageBreak();
    }
    return workbook;
  }

  async function download(report) {
    const workbook = createWorkbook(report);
    const buffer = await workbook.xlsx.writeBuffer();
    const url = global.URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Costo_Estructura.xlsx';
    document.body.appendChild(link);
    link.click();
    link.remove();
    global.setTimeout(() => global.URL.revokeObjectURL(url), 1000);
  }

  global.StructureExcel = { createWorkbook, download };
})(typeof window !== 'undefined' ? window : globalThis);
