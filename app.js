// app.js

const inputFile   = document.getElementById('input-file');
const btnProcesar = document.getElementById('procesar');

function downloadWorkbook(bin, filename) {
  const buf  = new ArrayBuffer(bin.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < bin.length; ++i) view[i] = bin.charCodeAt(i) & 0xFF;
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
  if (!file) return alert('Selecciona primero un .xlsx');

  const reader = new FileReader();
  reader.onload = e => {
    const data  = e.target.result;
    const wb    = XLSX.read(data, { type: 'binary', cellDates: true, raw: false });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    // Trae fechas como Date si cellDates:true
    const allRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1, defval: null, raw: false, cellDates: true
    });

    // Encontrar encabezado real para saltarlo
    const headerIdx = allRows.findIndex(r =>
      r[0] === 'Detalle' && r[2] === 'Debe' && r[3] === 'Haber' && r[4] === 'Saldo'
    );
    if (headerIdx < 0) return alert('Formato inesperado: no hallo fila Detalle/Debe/Haber/Saldo');

    const rows = allRows.slice(headerIdx + 1);

    const dataMap  = {};
    const monthSet = new Set();
    let currentKey = null;

    rows.forEach(r => {
      const [colA, colB, colC, colD, colE] = r;

      // Cabecera de cuenta: A y B llenos, C–E null
      const isHeader = colA != null && colB != null
                    && colC  === null && colD  === null && colE  === null;
      // Totales: A null pero C/D/E con algo
      const isTotal  = colA == null && (colC != null || colD != null || colE != null);

      // 3 tipos de fecha posibles
      let fechaVal = null;
      if (colA instanceof Date) {
        fechaVal = colA;
      } else if (typeof colA === 'number') {
        // Serial Excel → Date
        const d = XLSX.SSF.parse_date_code(colA);
        fechaVal = new Date(d.y, d.m - 1, d.d);
      } else if (typeof colA === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(colA)) {
        const [d, m, y] = colA.split('/');
        fechaVal = new Date(+y, +m - 1, +d);
      }

      if (isHeader) {
        currentKey = `${colA}|${colB}`;
        if (!dataMap[currentKey]) dataMap[currentKey] = {};
      }
      else if (!isTotal && fechaVal instanceof Date) {
        const y = fechaVal.getFullYear();
        const m = String(fechaVal.getMonth() + 1).padStart(2, '0');
        const key = `${y}-${m}`;
        monthSet.add(key);

        if (!dataMap[currentKey][key]) dataMap[currentKey][key] = 0;
        const debe  = colC != null ? parseFloat(colC.toString().replace(/\./g,'').replace(',', '.')) : 0;
        const haber = colD != null ? parseFloat(colD.toString().replace(/\./g,'').replace(',', '.')) : 0;
        dataMap[currentKey][key] += (debe - haber);
      }
    });

    // Armar lista de meses ordenada
    const meses = Array.from(monthSet).sort();

    // Preparar matrix de salida
    const out = [];
    out.push(['Código','Descripción', ...meses]);
    Object.entries(dataMap).forEach(([k, vals]) => {
      const [cod, desc] = k.split('|');
      const row = [cod, desc, ...meses.map(m => vals[m] || 0)];
      out.push(row);
    });

    // Exportar
    const ws    = XLSX.utils.aoa_to_sheet(out);
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, ws, 'SaldoMensual');
    const wbout = XLSX.write(newWb, { bookType:'xlsx', type:'binary' });
    downloadWorkbook(wbout, 'saldo_mensual.xlsx');
  };

  reader.readAsBinaryString(file);
});
