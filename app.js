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

// Al hacer clic en “Procesar y descargar”
btnProcesar.addEventListener('click', () => {
  const file = inputFile.files[0];
  if (!file) {
    alert('Por favor, selecciona un archivo .xlsx primero.');
    return;
  }

  const reader = new FileReader();
  reader.onload = evt => {
    const data  = evt.target.result;
    const wb    = XLSX.read(data, { type: 'binary' });
    const sheet = wb.Sheets[wb.SheetNames[0]];

    // Convertir hoja a array de filas
    const allRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: null
    });

    // Preparar estructuras
    const dataMap  = {}; // { "codigo|desc": { "YYYY-MM": neto, ... } }
    const monthSet = new Set();
    let currentKey = null;

    // Recorrer todas las filas
    allRows.forEach(r => {
      const [colA, colB, colC, colD, colE] = r;

      // Detectar inicio de cuenta contable
      const isHeader = colA && colB && colC === null && colD === null && colE === null;
      // Detectar fila de totales
      const isTotal  = !colA && (colC !== null || colD !== null || colE !== null);

      // Intentar parsear fecha DD/MM/YYYY
      let fechaVal = null;
      if (typeof colA === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(colA)) {
        const [d, m, y] = colA.split('/');
        fechaVal = new Date(+y, +m - 1, +d);
      }

      if (isHeader) {
        // Nueva cuenta: código|descripción
        currentKey = `${colA}|${colB}`;
        if (!dataMap[currentKey]) {
          dataMap[currentKey] = {};
        }
      }
      else if (!isTotal && fechaVal instanceof Date) {
        // Transacción válida
        const yyyy     = fechaVal.getFullYear();
        const mm       = String(fechaVal.getMonth() + 1).padStart(2, '0');
        const monthKey = `${yyyy}-${mm}`;
        monthSet.add(monthKey);

        // Inicializar si no existe
        if (!dataMap[currentKey][monthKey]) {
          dataMap[currentKey][monthKey] = 0;
        }
        // Convertir importes a número
        const debe  = colC !== null
          ? parseFloat(colC.toString().replace(/\./g, '').replace(',', '.'))
          : 0;
        const haber = colD !== null
          ? parseFloat(colD.toString().replace(/\./g, '').replace(',', '.'))
          : 0;

        dataMap[currentKey][monthKey] += (debe - haber);
      }
      // Totales y filas sin fecha se ignoran
    });

    // Crear array ordenado de meses
    const meses = Array.from(monthSet).sort();

    // Construir matriz de salida
    const output = [];
    output.push(['Código', 'Descripción', ...meses]);
    Object.entries(dataMap).forEach(([key, vals]) => {
      const [codigo, desc] = key.split('|');
      const row = [codigo, desc];
      meses.forEach(m => {
        row.push(vals[m] || 0);
      });
      output.push(row);
    });

    // Generar y descargar el Excel
    const ws    = XLSX.utils.aoa_to_sheet(output);
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, ws, 'SaldoMensual');
    const wbout = XLSX.write(newWb, { bookType: 'xlsx', type: 'binary' });
    downloadWorkbook(wbout, 'saldo_mensual.xlsx');
  };

  reader.readAsBinaryString(file);
});
