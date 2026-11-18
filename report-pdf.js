(function (global) {
  'use strict';

  const COLORS = {
    guinda: '#9F2241',
    guindaDark: '#691C32',
    guindaHover: '#821B35',
    dorado: '#BC955C',
    doradoLight: '#D4C19C',
    doradoDark: '#9A7740',
    fondo: '#F4F1EA',
    blanco: '#FFFFFF',
    texto: '#0F172A',
    secundario: '#475569',
    borde: '#E8DAC6',
    suave: '#F1F5F9',
    alterno: '#F8FAFC'
  };
  const MARGIN = 9;
  const TABLE_WIDTH = 279;
  const COLUMN_WIDTHS = [11, 56, 28, 29, 31, 16, 36, 36, 36];
  const HEADERS = [
    'Nivel',
    'Denominación general',
    'Costo anual',
    'Costo mensual bruto',
    'Costo mensual neto aproximado',
    'Cantidad de plazas',
    'Subtotal anual',
    'Subtotal mensual bruto',
    'Subtotal mensual neto'
  ];
  const money = new Intl.NumberFormat('es-MX', {
    style: 'currency', currency: 'MXN', minimumFractionDigits: 2, maximumFractionDigits: 2
  });

  function text(value) {
    return String(value == null ? '' : value)
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, '')
      .replace(/\u00a0/g, ' ');
  }

  function currency(value) {
    return text(money.format(Number(value) || 0));
  }

  function drawHeader(doc, report, layout) {
    const compact = layout.compact;
    const barHeight = compact ? 6 : 7;
    doc.setFillColor(COLORS.guindaDark);
    doc.rect(MARGIN, MARGIN, TABLE_WIDTH, barHeight, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(COLORS.blanco);
    doc.text('DGDYPO', MARGIN + 3, MARGIN + (compact ? 4.1 : 4.7));
    doc.setFont('helvetica', 'normal');
    doc.text('Gobierno de la Ciudad de México', MARGIN + TABLE_WIDTH - 3, MARGIN + (compact ? 4.1 : 4.7), { align: 'right' });
    doc.setDrawColor(COLORS.dorado);
    doc.setLineWidth(0.6);
    doc.line(MARGIN, MARGIN + barHeight + 0.5, MARGIN + TABLE_WIDTH, MARGIN + barHeight + 0.5);

    doc.setTextColor(COLORS.guinda);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(compact ? 12 : 14);
    doc.text(text(report.title || 'Cálculo de Costo de Estructura'), MARGIN, compact ? 21 : 24);

    doc.setTextColor(COLORS.secundario);
    const organizationFontSize = compact ? 7.5 : 8.5;
    doc.setFontSize(organizationFontSize);
    doc.text('Organismo', MARGIN, compact ? 26 : 30);
    doc.setFont('helvetica', 'normal');
    const organization = doc.splitTextToSize(text(report.organization || 'Sin organismo seleccionado'), TABLE_WIDTH);
    const organizationY = compact ? 29.5 : 34;
    doc.text(organization, MARGIN, organizationY, { lineHeightFactor: 1.15 });
    return organizationY + organization.length * organizationFontSize / doc.internal.scaleFactor * 1.15 + (compact ? 2 : 3);
  }

  function drawSummary(doc, report, startY, layout) {
    const compact = layout.compact;
    const gap = 3;
    const width = (TABLE_WIDTH - gap * 3) / 4;
    const entries = [
      ['Total de plazas', String(report.totals.positions || 0)],
      ['Costo anual total', currency(report.totals.annual)],
      ['Mensual bruto total', currency(report.totals.gross)],
      ['Mensual neto aproximado', currency(report.totals.net)]
    ];
    entries.forEach(function (entry, index) {
      const x = MARGIN + index * (width + gap);
      const featured = index === 3;
      doc.setFillColor(featured ? COLORS.guindaDark : COLORS.alterno);
      doc.setDrawColor(COLORS.borde);
      doc.setLineWidth(0.25);
      doc.roundedRect(x, startY, width, compact ? 13 : 17, 1.5, 1.5, 'FD');
      doc.setTextColor(featured ? COLORS.blanco : COLORS.secundario);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(compact ? 7 : 7.7);
      doc.text(entry[0], x + 3, startY + (compact ? 4 : 5));
      doc.setTextColor(featured ? COLORS.blanco : COLORS.guinda);
      doc.setFont('helvetica', 'bold');
      const valueFontSize = compact ? 10 : 11;
      doc.setFontSize(valueFontSize);
      const size = Math.min(valueFontSize, valueFontSize * (width - 6) / Math.max(1, doc.getTextWidth(entry[1])));
      doc.setFontSize(size);
      doc.text(entry[1], x + 3, startY + (compact ? 10 : 12.5));
    });
    return startY + (compact ? 17 : 22);
  }

  function tableRows(rows) {
    return rows.map(function (row) {
      return [
        text(row.level), text(row.designation), currency(row.annual),
        currency(row.gross), currency(row.net), String(row.quantity),
        currency(row.annualSubtotal), currency(row.grossSubtotal), currency(row.netSubtotal)
      ];
    });
  }

  function totalRow(report) {
    return [[
      { content: 'TOTALES', colSpan: 5, styles: { halign: 'right' } },
      String(report.totals.positions || 0), currency(report.totals.annual),
      currency(report.totals.gross), currency(report.totals.net)
    ]];
  }

  function signatureLayout(doc, report, layout) {
    const fontSize = layout.compact ? 7.5 : 8.5;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(fontSize);
    const lines = doc.splitTextToSize(text(report.validatedBy || 'Nombre de la persona que valida: ______________________________'), 180);
    const nameY = layout.compact ? 8 : 9.5;
    const timestampY = Math.max(layout.compact ? 17 : 18.5, nameY + (lines.length - 1) * fontSize / doc.internal.scaleFactor * 1.1 + 5);
    return { fontSize: fontSize, lines: lines, nameY: nameY, timestampY: timestampY, height: timestampY + 2 };
  }

  function drawSignature(doc, report, startY, signature) {
    doc.setDrawColor(COLORS.dorado);
    doc.setLineWidth(0.25);
    doc.line(MARGIN, startY, MARGIN + TABLE_WIDTH, startY);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(COLORS.guinda);
    doc.setFontSize(signature.fontSize);
    doc.text('Validó por', MARGIN, startY + 5);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(COLORS.texto);
    doc.text(signature.lines, MARGIN, startY + signature.nameY, { lineHeightFactor: 1.1 });
    doc.setDrawColor(COLORS.borde);
    doc.line(222, startY + 11, MARGIN + TABLE_WIDTH, startY + 11);
    doc.setFontSize(7.5);
    doc.setTextColor(COLORS.secundario);
    doc.text('Firma', 255, startY + 15, { align: 'center' });
    const timestamp = text(report.generatedAtLabel || new Intl.DateTimeFormat('es-MX', {
      dateStyle: 'long', timeStyle: 'short', timeZone: 'America/Mexico_City'
    }).format(report.generatedAt || new Date()));
    doc.text('Fecha y hora de elaboración: ' + timestamp, MARGIN, startY + signature.timestampY);
  }

  function drawFooter(doc) {
    const count = doc.getNumberOfPages();
    for (let page = 1; page <= count; page += 1) {
      doc.setPage(page);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      doc.setTextColor(COLORS.secundario);
      doc.text('DGDYPO - Cálculo de Costo de Estructura', MARGIN, 204);
      doc.text('Página ' + page + ' de ' + count, MARGIN + TABLE_WIDTH, 204, { align: 'right' });
    }
  }

  function buildDocument(report, layout) {
    const doc = new global.jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
    if (typeof doc.autoTable !== 'function') {
      throw new Error('No se pudo cargar el formato de la tabla PDF. Recargue la página e intente de nuevo.');
    }
    const rows = report.rows || [];
    doc.setProperties({
      title: text(report.title || 'Cálculo de Costo de Estructura'),
      subject: 'Tabulador de niveles, plazas y costos de estructura',
      author: 'DGDYPO', creator: 'Calculadora de Costo de Estructura'
    });

    const signature = signatureLayout(doc, report, layout);
    let y = drawHeader(doc, report, layout);
    y = drawSummary(doc, report, y, layout);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(layout.compact ? 8 : 9);
    doc.setTextColor(COLORS.guinda);
    doc.text('Tabulador de niveles y costos', MARGIN, y);
    y += 3;
    const body = rows.length ? tableRows(rows) : [[{
      content: 'Sin plazas capturadas.', colSpan: 9, styles: { halign: 'center', minCellHeight: 10 }
    }]];
    const columnStyles = {};
    COLUMN_WIDTHS.forEach(function (width, index) {
      columnStyles[index] = {
        cellWidth: width,
        halign: index === 0 || index === 5 ? 'center' : index === 1 ? 'left' : 'right'
      };
      if (index >= 6) columnStyles[index].textColor = COLORS.guinda;
    });
    doc.autoTable({
      startY: y,
      tableWidth: TABLE_WIDTH,
      margin: { left: MARGIN, right: MARGIN, top: MARGIN, bottom: 12 + signature.height + 3 },
      head: [HEADERS], body: body, foot: totalRow(report),
      theme: 'grid', pageBreak: 'avoid', rowPageBreak: 'avoid', showHead: 'firstPage', showFoot: 'lastPage',
      styles: {
        font: 'helvetica', fontSize: layout.fontSize, textColor: COLORS.texto,
        lineColor: COLORS.borde, lineWidth: 0.15,
        cellPadding: { top: layout.padding, right: 1.4, bottom: layout.padding, left: 1.4 },
        valign: 'middle', overflow: 'linebreak'
      },
      headStyles: { fillColor: COLORS.guindaDark, textColor: COLORS.blanco, fontStyle: 'bold', fontSize: Math.min(7.8, layout.fontSize), halign: 'center' },
      alternateRowStyles: { fillColor: COLORS.alterno },
      footStyles: { fillColor: COLORS.suave, textColor: COLORS.guinda, fontStyle: 'bold', halign: 'right' },
      columnStyles: columnStyles
    });
    const signatureStart = doc.lastAutoTable.finalY + 3;
    doc.structureLayout = {
      pageCount: doc.getNumberOfPages(), rowCount: rows.length,
      fontSize: layout.fontSize, tableTop: y, tableBottom: doc.lastAutoTable.finalY,
      signatureBottom: signatureStart + signature.height
    };
    if (doc.structureLayout.pageCount === 1 && doc.structureLayout.signatureBottom <= 198) {
      drawSignature(doc, report, signatureStart, signature);
      drawFooter(doc);
    }
    return doc;
  }

  function createDocument(report) {
    if (!global.jspdf || !global.jspdf.jsPDF) {
      throw new Error('No se pudo cargar la herramienta de PDF. Recargue la página e intente de nuevo.');
    }
    const count = (report.rows || []).length;
    const initial = count > 14
      ? { compact: true, fontSize: 7.2, padding: 0.75 }
      : count > 8
        ? { compact: true, fontSize: 7.7, padding: 0.95 }
        : { compact: false, fontSize: 8.2, padding: 1.5 };
    const layouts = [
      initial,
      { compact: true, fontSize: 7.2, padding: 0.7 },
      { compact: true, fontSize: 6.9, padding: 0.55 },
      { compact: true, fontSize: 6.7, padding: 0.4 }
    ];
    for (const layout of layouts) {
      const doc = buildDocument(report, layout);
      if (doc.structureLayout.pageCount === 1 && doc.structureLayout.signatureBottom <= 198) return doc;
    }
    throw new Error('El contenido excede el espacio disponible para una sola página. Revise la extensión del organismo y del nombre de validación.');
  }

  function download(report) {
    const doc = createDocument(report);
    doc.save('Costo_Estructura.pdf');
    return doc;
  }

  global.StructurePDF = { createDocument: createDocument, download: download };
})(window);
