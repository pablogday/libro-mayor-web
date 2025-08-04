// app.js

// Referencias a elementos del DOM
const inputFile   = document.getElementById('input-file');
const startMonth  = document.getElementById('start-month');
const endMonth    = document.getElementById('end-month');
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

// Evento al pulsar “Procesar y descargar”
btnProcesar.addEventListener('click', () => {
  const file = inputFile.files[0];
  if (!file) {
    alert('Por favor, selecciona un archivo .xlsx primero.');
    return;
  }

  const reader = new FileReader();
  reader.onload = evt => {
    // Leemos el libro y la primera hoja
    const data  = evt.target.result;
    const wb    = XLSX.read(data, { type:'binary' });
    const sheet = wb.Sheets[wb.SheetNames[0]];

    // Convertimos a un array de filas; defval:null para celdas en blanco
    const allRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw:   false,
      defval: null
    });

    // Encontrar la fila de encabezados ("Detalle", "Descripcion", ...)
    const headerIdx = allRows.findIndex(r =>
      r[0] === 'Detalle' && r[1] === 'Descripcion'
    );
    // Tomar sólo las filas de datos (después de encabezados)
    const rows = headerIdx >= 0
      ? allRows.slice(headerIdx + 1)
      : allRows.slice(6); // fallback

    // Preparar el mapa de acumulación
    const dataMap = {}; // { 'codigo|desc': { 'YYYY-MM': neto, ... } }
    let currentKey = null;

    // Construir la lista de meses entre start y end
    const meses = [];
    let cursor = new Date(startMonth.value + '-01');
    const endDate = new Date(endMonth.value + '-01');
    while (cursor <= endDate) {
      const y = cursor.getFullYear();
      const m = String(cursor.getMonth() + 1).padStart(2,'0');
      meses.push(`${y}-${m}`);
      cursor.setMonth(cursor.getMonth() + 1);
    }

    // Procesar cada fila
    rows.forEach(r => {
      const [colA, colB, colC, colD, colE] = r;

      // Detectar fila de cabecera de cuenta
      const isHeader = colA && colB && colC === null && colD === null && colE === null;
      // Detectar fila de totales (colA nulo, pero hay valores en C/D/E)
      const isTotal  = !colA && (colC !== null || colD !== null || colE !== null);

      // Parsear fecha DD/MM/YYYY → Date
      let fechaVal = null;
      if (typeof colA === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(colA)) {
        const [d,m,y] = colA.split('/');
        fechaVal = new Date(+y, +m - 1, +d);
      }

      if (isHeader) {
        // Iniciar nueva cuenta
        currentKey = `${colA}|${colB}`;
        if (!dataMap[currentKey]) {
          dataMap[currentKey] = {};
          meses.forEach(mm => dataMap[currentKey][mm] = 0);
        }
      }
      else if (!isTotal && fechaVal instanceof Date) {
        // Es una transacción válida
        const yyyy     = fechaVal.getFullYear();
        const mm       = String(fechaVal.getMonth() + 1).padStart(2,'0');
        const monthKey = `${yyyy}-${mm}`;

        if (meses.includes(monthKey)) {
          // Convertimos colC/D a número (si vienen como texto)
          const debe  = colC !== null ? parseFloat(colC.toString().replace(/\./g,'').replace(',', '.')) : 0;
          const haber = colD !== null ? parseFloat(colD.toString().replace(/\./g,'').replace(',', '.')) : 0;
          dataMap[currentKey][monthKey] += (debe - haber);
        }
      }
      // totales y filas sin fecha quedan ignoradas
    });

    // Construir matriz de salida para SheetJS
    const output = [];
    output.push(['Código','Descripción', ...meses]);
    Object.entries(dataMap).forEach(([key, vals]) => {
      const [codigo, desc] = key.split('|');
      const row = [codigo, desc, ...meses.map(m => vals[m])];
      output.push(row);
    });

    // Generar y descargar el Excel
    const ws    = XLSX.utils.aoa_to_sheet(output);
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, ws, 'SaldoMensual');
    const wbout = XLSX.write(newWb, { bookType:'xlsx', type:'binary' });
    downloadWorkbook(wbout, 'saldo_mensual.xlsx');
  };

  reader.readAsBinaryString(file);
});
