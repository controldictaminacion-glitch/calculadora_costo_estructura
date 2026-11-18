(function (root) {
  'use strict';

  const SINGLE_PAGE_ROWS = 8;
  const roundMoney = value => Math.round((value + Number.EPSILON) * 100) / 100;

  function createReport(options) {
    const generatedAt = options.generatedAt || new Date();
    const rows = options.levels.flatMap((item, index) => {
      const quantity = Math.max(0, Math.trunc(Number(options.quantities[index]) || 0));
      if (!Number.isFinite(quantity) || quantity === 0) return [];
      return [{
        ...item,
        quantity,
        annualSubtotal: roundMoney(item.annual * quantity),
        grossSubtotal: roundMoney(item.gross * quantity),
        netSubtotal: roundMoney(item.net * quantity)
      }];
    });
    const totals = rows.reduce((result, row) => ({
      positions: result.positions + row.quantity,
      annual: roundMoney(result.annual + row.annualSubtotal),
      gross: roundMoney(result.gross + row.grossSubtotal),
      net: roundMoney(result.net + row.netSubtotal)
    }), {positions: 0, annual: 0, gross: 0, net: 0});
    const pageCount = rows.length <= SINGLE_PAGE_ROWS ? 1 : 2;

    return {
      title: 'Cálculo de Costo de Estructura',
      organization: String(options.organization || '').trim(),
      validatedBy: String(options.validatedBy || '').trim(),
      generatedAt,
      generatedAtLabel: new Intl.DateTimeFormat('es-MX', {
        dateStyle: 'long', timeStyle: 'short', timeZone: 'America/Mexico_City'
      }).format(generatedAt),
      rows,
      totals,
      pageCount,
      rowsPerPage: Math.max(1, Math.ceil(rows.length / pageCount))
    };
  }

  root.StructureReport = {createReport};
})(typeof window === 'undefined' ? globalThis : window);
