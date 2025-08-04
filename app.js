// app.js

// Referencias a elementos del DOM
const inputFile   = document.getElementById('input-file');
const btnProcesar = document.getElementById('procesar');

// Función para descargar el workbook como .xlsx
function downloadWorkbook(binaryStr, filename) {
  const buf  = new ArrayBuffer(binaryStr.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < binaryStr.length; ++i) {
    view[i] = binaryStr.charCodeAt(i) & 0xFF;
  }
  const blob = new Blob([buf], { type: 'application/octet-stream' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

btnProcesar.addEventListener('click', () => {
  const file = inputFile.files[0];
  if (!file) {
    alert('Por favor, selecciona un archivo .xlsx primero.');
    return;
  }

  const reader = new FileReader();
  reader.onload = evt => {
    // 1) Leer el archivo con cellDates para obtener objetos Date
    const data = evt.target.result;
    const wb   = XLSX.read(data, {
      type: 'binary',
      cellDates: true,
      raw: false
    });
    const sheet = wb.Sheets[wb.SheetNames[0]];

    // 2) Convertir toda la hoja a un array de filas
    const allRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: null,
      raw: false,
      cellDates: true
    });

    // 3) Encontrar y saltar la fila de encabezados reales
    const headerIdx = allRows.findIndex(r =>
      r[0] === 'Detalle' &&
      r[2] === 'Debe' &&
      r[3] === 'Haber' &&
      r[4] === 'Saldo'
    );
    if (headerIdx < 0) {
      alert('No se encontró la fila de encabezados (Detalle/Debe/Haber/Saldo).');
      return;
    }
    const rows = allRows.slice(headerIdx + 1);

    // 4) Preparar estructuras para acumular netos por cuenta y mes
    const dataMap  = {};       // { "codigo|desc": { "YYYY-MM": neto, ... } }
    const monthSet = new Set();
    let currentKey = null;

    // 5) Recorrer cada fila de datos
    rows.forEach(r => {
      const [colA, colB, colC, colD, colE] = r;

      // a) Detección de encabezado de cuenta
      const isHeader = colA != null && colB != null
                    && colC === null && colD === null && colE === null;

      // b) Detección de fila de totales
      const isTotal  = colA == null && (colC != null || colD != null || colE != null);

      // c) Parseo de fecha: Date, serial o string "DD/MM/YYYY"
      let fechaVal = null;
      if (colA instanceof Date) {
        fechaVal = colA;
      } else if (typeof colA === 'number') {
        const d = XLSX.SSF.parse_date_code(colA);
        fechaVal = new Date(d.y, d.m - 1, d.d);
      } else if (typeof colA === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(colA)) {
        const [d, m, y] = colA.split('/');
        fechaVal = new Date(+y, +m - 1, +d);
      }

      // 6) Lógica según tipo de fila
      if (isHeader) {
        currentKey = `${colA}|${colB}`;
        if (!dataMap[currentKey]) {
          dataMap[currentKey] = {};
        }
      }
      else if (!isTotal && fechaVal instanceof Date) {
        const yyyy     = fechaVal.getFullYear();
        const mm       = String(fechaVal.getMonth() + 1).padStart(2, '0');
        const monthKey = `${yyyy}-${mm}`;

        monthSet.add(monthKey);
        if (!dataMap[currentKey][monthKey]) {
          dataMap[currentKey][monthKey] = 0;
        }

        const debe  = colC != null
          ? parseFloat(colC.toString().replace(/\./g, '').replace(',', '.'))
          : 0;
        const haber = colD != null
          ? parseFloat(colD.toString().replace(/\./g, '').replace(',', '.'))
          : 0;

        dataMap[currentKey][monthKey] += (debe - haber);
      }
      // filas de totales o sin fecha se ignoran
    });

    // 7) Ordenar meses cronológicamente
    const meses = Array.from(monthSet).sort();

    // 8) Construir la matriz de salida
    const output = [];
    output.push(['Código', 'Descripción', ...meses]);
    Object.entries(dataMap).forEach(([key, vals]) => {
      const [codigo, desc] = key.split('|');
      const row = [codigo, desc, ...meses.map(m => vals[m] || 0)];
      output.push(row);
    });

    // 9) Generar y descargar el Excel resultante
    const ws    = XLSX.utils.aoa_to_sheet(output);
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, ws, 'SaldoMensual');
    const wbout = XLSX.write(newWb, { bookType: 'xlsx', type: 'binary' });
    downloadWorkbook(wbout, 'saldo_mensual.xlsx');
  };

  reader.readAsBinaryString(file);
});
