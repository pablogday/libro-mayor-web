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
      : allRows.slice(6);

    // Construir el rango fijo de meses: Abril 2024 a Marzo 2025
    const meses = [
      '2024-04','2024-05','2024-06','2024-07','2024-08','2024-09',
      '2024-10','2024-11','2024-12',
      '2025-01','2025-02','2025-03'
    ];

    // Preparar el mapa de acumulación
    const dataMap = {}; // { 'codigo|desc': { 'YYYY-MM': neto, ... } }
    let currentKey = null;

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
          // Convertimos colC/D a número (maneja formato "1.234,56")
          const debe  = colC !== null ? parseFloat(colC.toString().replace(/\./g,'').replace(',', '.')) : 0;
          const haber = colD !== null ? parseFloat(colD.toString().replace(/\./g,'').replace(',', '.')) : 0;
          dataMap[currentKey][monthKey] += (debe - haber);
        }
      }
      // Ignoramos totales y filas sin fecha
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
