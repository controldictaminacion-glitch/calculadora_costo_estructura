# Calculadora de costo de estructura

Seleccione un organismo, muestre los niveles en cero y capture las plazas por
nivel. El resumen calcula los costos anuales y mensuales de la estructura.

## Descargas

- **Descargar PDF:** genera una sola página A4 horizontal con todos los niveles
  capturados (hasta diecinueve), los totales y la firma. Ajusta el espacio y el
  tamaño de la tabla según la cantidad de datos.
- **Descargar Excel:** genera un `.xlsx` con números editables, fórmulas,
  encabezados repetidos y área de impresión definida. Conserva A4 horizontal y
  ajuste a una página de ancho y hasta una o dos de alto según los datos.

Los reportes incluyen únicamente niveles con plazas, el organismo, el nombre de
quien valida y la fecha de elaboración en horario de Ciudad de México. Mostrar
los niveles en cero en pantalla no agrega filas vacías a las descargas.

Las bibliotecas de exportación se encuentran en `vendor/` y se cargan solo al
descargar. Los archivos se generan en el navegador. La consulta de organismos
continúa usando la hoja de Google configurada en la página.

## Verificación

Ejecute `node tests/export-reports.cjs` para comprobar que el PDF conserva todos
los niveles en una página, nombres largos, importes grandes, fórmulas y
configuración de impresión de Excel. Agregue `--samples` para guardar muestras
PDF y XLSX en `tmp/exports/`.
