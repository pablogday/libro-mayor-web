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

// Al hacer clic en “Procesar y descargar”
btnProcesar.addEventListener('click', () => {
  const file = inputFile.files[0];
  if (!file) {
    alert('Por favor, selecciona un archivo .xlsx primero.');
    return;
  }

  const reader = new FileReader();
  reader.onload = evt => {
    const data    = evt.target.result;
    const wb      = XLSX.read(data, { type: 'binary' });
    const sheet   = wb.Sheets[wb.SheetNames[0]];
    const rows    = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

    // Variables auxiliares
    const dataMap = {}; // { 'codigo|desc': { 'YYYY-MM': neto, … } }
    let currentKey = null;

    // Inicializar rango de meses
    const meses = [];
    let cursor = new Date(startMonth.value + '-01');
    const endDate = new Date(endMonth.value + '-01');
    while (cursor <= endDate) {
      const y  = cursor.getFullYear();
      const m  = String(cursor.getMonth() + 1).padStart(2, '0');
      meses.push(`${y}-${m}`);
      cursor.setMonth(cursor.getMonth() + 1);
    }

    // Procesar cada fila
    rows.forEach(r => {
      const [colA, colB, colC, colD, colE] = r;
      const isHeader = colA && colB && !colC && !colD && !colE;
      const isTotal  = !colA && (colC || colD || colE);
      const fecha    = Date.parse(colA);
      if (isHeader) {
        currentKey = `${colA}|${colB}`;
        if (!dataMap[currentKey]) {
          dataMap[currentKey] = {};
          meses.forEach(m => dataMap[currentKey][m] = 0);
        }
      } else if (!isTotal && !isNaN(fecha)) {
        // Es una transacción
        const date     = new Date(colA.split('/').reverse().join('-')); // DD/MM/YYYY
        const yyyy     = date.getFullYear();
        const mm       = String(date.getMonth() + 1).padStart(2, '0');
        const monthKey = `${yyyy}-${mm}`;
        if (meses.includes(monthKey)) {
          const debe  = parseFloat(colC) || 0;
          const haber = parseFloat(colD) || 0;
          dataMap[currentKey][monthKey] += (debe - haber);
        }
      }
    });

    // Construir matriz para SheetJS
    const output = [];
    output.push(['Código', 'Descripción', ...meses]);
    Object.entries(dataMap).forEach(([key, vals]) => {
      const [codigo, desc] = key.split('|');
      const row = [codigo, desc, ...meses.map(m => vals[m])];
      output.push(row);
    });

    // Generar y descargar el Excel
    const ws   = XLSX.utils.aoa_to_sheet(output);
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, ws, 'SaldoMensual');
    const wbout = XLSX.write(newWb, { bookType: 'xlsx', type: 'binary' });
    downloadWorkbook(wbout, 'saldo_mensual.xlsx');
  };

  reader.readAsBinaryString(file);
});
