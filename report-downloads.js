(function (root) {
  'use strict';

  const libraries = {
    pdf: [
      ['vendor/jspdf.umd.min.js', () => Boolean(root.jspdf && root.jspdf.jsPDF)],
      ['vendor/jspdf.plugin.autotable.min.js', () => Boolean(root.jspdf && root.jspdf.jsPDF.API.autoTable)]
    ],
    excel: [['vendor/exceljs.min.js', () => Boolean(root.ExcelJS)]]
  };
  const pendingLibraries = new Map();

  function loadLibrary(path, isReady) {
    if (isReady()) return Promise.resolve();
    if (pendingLibraries.has(path)) return pendingLibraries.get(path);
    const pending = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = path;
      script.onload = () => {
        if (isReady()) resolve();
        else reject(new Error('La biblioteca no se inició: ' + path));
      };
      script.onerror = () => { script.remove(); reject(new Error('No se pudo cargar ' + path)); };
      document.head.appendChild(script);
    }).catch(error => { pendingLibraries.delete(path); throw error; });
    pendingLibraries.set(path, pending);
    return pending;
  }

  async function download(format, report) {
    const isPDF = format === 'pdf';
    const label = isPDF ? 'PDF' : 'Excel';
    const button = document.getElementById(isPDF ? 'downloadPDFButton' : 'downloadExcelButton');
    const status = document.getElementById('exportStatus');
    const originalLabel = button.innerHTML;
    button.disabled = true;
    button.setAttribute('aria-busy', 'true');
    button.textContent = 'Preparando ' + label + '…';
    status.hidden = false;
    status.textContent = 'Preparando el reporte ' + label + '…';
    try {
      for (const [path, isReady] of libraries[format]) await loadLibrary(path, isReady);
      await (isPDF ? root.StructurePDF : root.StructureExcel).download(report);
      status.textContent = 'El reporte ' + label + ' se generó correctamente.';
    } catch (error) {
      console.error('Error al generar el reporte ' + label, error);
      status.textContent = 'No se pudo generar el reporte ' + label + '. Recargue la página e intente de nuevo.';
    } finally {
      button.disabled = false;
      button.removeAttribute('aria-busy');
      button.innerHTML = originalLabel;
    }
  }

  root.StructureDownloads = {download};
})(typeof window === 'undefined' ? globalThis : window);
