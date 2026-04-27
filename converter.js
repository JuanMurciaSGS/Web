// Nombre del archivo: converter.js
// Dependencia: Se requiere la librería SheetJS (xlsx.full.min.js)

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('textFileInput');
    const convertButton = document.getElementById('convertButton');
    const statusMessage = document.getElementById('statusMessage');

    // 1. DETECCIÓN DEL TIPO DE REPORTE
    const toolSection = document.querySelector('.tool-section');
    const reportType = toolSection ? toolSection.getAttribute('data-report-type') : null; 

    // Asignar la función de procesamiento según el tipo de reporte
    let processFile;
    if (reportType === 'otc') {
        processFile = processOtcFile;
    } else if (reportType === 'unidentify') {
        processFile = processUnidentifyFile; 
    } else {
        processFile = processAgeingFile; 
    }

    // 2. EVENT LISTENER PRINCIPAL
    convertButton.addEventListener('click', () => {
        statusMessage.textContent = ''; 
        const file = fileInput.files[0];
        
        if (!file) {
            statusMessage.textContent = 'Por favor, selecciona un archivo.';
            statusMessage.style.color = 'red';
            return;
        }

        const reader = new FileReader();
        
        reader.onload = function(e) {
            processFile(e.target.result, file.name);
        };

        if (reportType === 'unidentify') {
             reader.readAsArrayBuffer(file);
        } else {
             reader.readAsText(file);
        }
    });

// ----------------------------------------------------------------------
// --- LÓGICA REPORTE UNIDENTIFY ---
// ----------------------------------------------------------------------
function processUnidentifyFile(dataArrayBuffer, fileName) {
    try {
        const workbook = XLSX.read(dataArrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        if (allRows.length === 0) {
            statusMessage.textContent = 'El archivo está vacío.';
            statusMessage.style.color = 'red';
            return;
        }

        const headers = allRows[0];
        const dataRows = allRows.slice(1);
        
        const CLASS_COLUMN_INDEX = 15; 
        const DATE_COLS = ['Receipt Date', 'Deposit Date', 'GL Date'];
        const NUMBER_COLS = ['Receipt Amount', 'Net Amount', 'Unapplied Amount', 'Unidentified Amount'];
        const TEXT_COLS = ['Receipt Number'];

        const colIndices = {};
        headers.forEach((header, index) => { colIndices[header.trim()] = index; });

        const dateIndices = DATE_COLS.map(name => colIndices[name]).filter(i => i !== undefined);
        const numberIndices = NUMBER_COLS.map(name => colIndices[name]).filter(i => i !== undefined);
        const textIndices = TEXT_COLS.map(name => colIndices[name]).filter(i => i !== undefined);
        
        const sheetsData = {}; 

        dataRows.forEach(row => {
            const classValue = String(row[CLASS_COLUMN_INDEX] || "SIN BANCO"); 
            dateIndices.forEach(idx => {
                if (row[idx]) {
                    let parsedDate = new Date(row[idx]);
                    if (!isNaN(parsedDate.getTime())) row[idx] = parsedDate;
                }
            });

            if (!sheetsData[classValue]) { sheetsData[classValue] = [headers]; }
            sheetsData[classValue].push(row);
        });

        const outputWorkbook = XLSX.utils.book_new();
        for (const classValue in sheetsData) {
            const data = sheetsData[classValue];
            const ws = XLSX.utils.aoa_to_sheet(data, { cellDates: true }); 
            
            for (let R = 1; R < data.length; ++R) {
                const row = data[R];
                for (let C = 0; C < row.length; ++C) {
                    const cell = ws[XLSX.utils.encode_cell({ c: C, r: R })];
                    if (!cell) continue;
                    if (dateIndices.includes(C)) {
                        cell.z = 'dd/mm/yyyy';
                    } else if (numberIndices.includes(C)) {
                        const val = parseFloat(String(cell.v).replace(/,/g, ''));
                        if (!isNaN(val)) { cell.v = val; cell.t = 'n'; cell.z = '#,##0.00'; }
                    } else if (textIndices.includes(C)) {
                        cell.t = 's'; cell.z = '@';
                    }
                }
            }
            XLSX.utils.book_append_sheet(outputWorkbook, ws, classValue.substring(0, 31));
        }
        XLSX.writeFile(outputWorkbook, fileName.replace(/\.[^/.]+$/, "") + "_UnidentifyReport.xlsx");
        statusMessage.textContent = `¡Conversión exitosa!`;
        statusMessage.style.color = 'green';
    } catch (error) {
        statusMessage.textContent = `Error: ${error.message}`;
    }
}

// ----------------------------------------------------------------------
// --- LÓGICA REPORTE AGEING (ACTUALIZADA: 88%, 12% Y CONCATENACIÓN) ---
// ----------------------------------------------------------------------
function processAgeingFile(fileContent, fileName) {
    try {
        const allRows = fileContent.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.split('\t'));

        if (allRows.length === 0) {
            statusMessage.textContent = 'El archivo está vacío.';
            statusMessage.style.color = 'red';
            return;
        }

        // AGREGAR CABECERAS NUEVAS
        let headers = [...allRows[0]];
        headers.push('BASE (88%)', 'IGV/RET (12%)', 'Descripcion_Cuenta');
        
        const dataRows = allRows.slice(1);

        const idxTrxDate = headers.indexOf('TRX_DATE');
        const idxTrxNumber = headers.indexOf('TRX_NUMBER');
        const idxInvAmt = headers.indexOf('INVOICE_AMT');
        const idxBalance = headers.indexOf('BALANCE');
        const idxBalFunct = headers.indexOf('BALANCE_FUNCT');
        const CLASS_COLUMN_INDEX = 9;
        
        // Índices para la concatenación
        const concatCols = [
            'COMPANY', 'ACCOUNT', 'SECTOR', 'ACTIVITY', 'COST CENTER', 
            'COST LEVEL', 'LOCATION', 'INTERCOMPANY', 'PROJECT', 
            'STATUTORY', 'RESERVED1', 'RESERVED2'
        ];
        const concatIndices = concatCols.map(col => headers.indexOf(col));

        // Índices de las nuevas columnas calculadas
        const idxBase88 = headers.length - 3;
        const idxIgv12 = headers.length - 2;
        const idxDescCuenta = headers.length - 1;

        const monthMap = { 'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'MAY': 4, 'JUN': 5, 
                           'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11 };

        const sheetsData = {}; 

        dataRows.forEach(row => {
            let processedRow = [...row];

            // 1. TRATAMIENTO DE FECHA
            if (idxTrxDate !== -1 && processedRow[idxTrxDate]) {
                let rawDate = processedRow[idxTrxDate].toUpperCase();
                let parts = rawDate.split(/[-/]/);
                if (parts.length === 3) {
                    const day = parseInt(parts[0]);
                    const monthIndex = monthMap[parts[1]];
                    const year = parseInt(parts[2]);
                    if (monthIndex !== undefined) {
                        const fullYear = year < 100 ? 2000 + year : year;
                        processedRow[idxTrxDate] = new Date(fullYear, monthIndex, day);
                    }
                }
            }

            // 2. TRATAMIENTO DE NÚMEROS Y CÁLCULOS (88% y 12%)
            let invAmtValue = 0;
            [idxInvAmt, idxBalance, idxBalFunct].forEach(idx => {
                if (idx !== -1 && processedRow[idx]) {
                    const cleanNum = parseFloat(processedRow[idx].replace(/,/g, ''));
                    const finalNum = isNaN(cleanNum) ? 0 : cleanNum;
                    processedRow[idx] = finalNum;
                    if (idx === idxInvAmt) invAmtValue = finalNum;
                }
            });
            processedRow[idxBase88] = invAmtValue * 0.88;
            processedRow[idxIgv12] = invAmtValue * 0.12;

            // 3. CONCATENACIÓN Descripcion_Cuenta
            const accountValues = concatIndices.map(idx => {
                return (idx !== -1 && processedRow[idx]) ? String(processedRow[idx]).trim() : "";
            });
            processedRow[idxDescCuenta] = accountValues.join('.');

            // 4. TRATAMIENTO TRX_NUMBER
            if (idxTrxNumber !== -1) processedRow[idxTrxNumber] = String(processedRow[idxTrxNumber]);

            const classValue = processedRow[CLASS_COLUMN_INDEX] || "SIN CLASIFICAR";
            if (!sheetsData[classValue]) { sheetsData[classValue] = [headers]; }
            sheetsData[classValue].push(processedRow);
        });

        const workbook = XLSX.utils.book_new();
        for (const classValue in sheetsData) {
            const data = sheetsData[classValue];
            const ws = XLSX.utils.aoa_to_sheet(data, { cellDates: true }); 

            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r + 1; R <= range.e.r; ++R) {
                if (idxTrxDate !== -1) {
                    const cell = ws[XLSX.utils.encode_cell({r: R, c: idxTrxDate})];
                    if (cell) cell.z = 'dd/mm/yyyy'; 
                }
                // Formato Moneda
                [idxInvAmt, idxBalance, idxBalFunct, idxBase88, idxIgv12].forEach(idx => {
                    if (idx !== -1) {
                        const cell = ws[XLSX.utils.encode_cell({r: R, c: idx})];
                        if (cell) cell.z = '#,##0.00'; 
                    }
                });
                // Formato Texto para la nueva columna concatenada
                const cellDesc = ws[XLSX.utils.encode_cell({r: R, c: idxDescCuenta})];
                if (cellDesc) { cellDesc.t = 's'; cellDesc.z = '@'; }
            }
            XLSX.utils.book_append_sheet(workbook, ws, classValue.substring(0, 31).replace(/[\\?*:[\]/]/g, ""));
        }
        XLSX.writeFile(workbook, fileName.replace(/\.[^/.]+$/, "") + "_AgeingReport.xlsx");
        statusMessage.textContent = `¡Conversión exitosa! Columnas de impuestos y cuenta contable añadidas.`;
        statusMessage.style.color = 'green';
    } catch (error) {
        statusMessage.textContent = `Error: ${error.message}`;
    }
}

// ----------------------------------------------------------------------
// --- LÓGICA REPORTE OTC ---
// ----------------------------------------------------------------------
function processOtcFile(fileContent, fileName) {
      try {
        const allRows = fileContent.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.split('\t'));

        if (allRows.length === 0) {
            statusMessage.textContent = 'El archivo está vacío.';
            statusMessage.style.color = 'red';
            return;
        }

        const workbook = XLSX.utils.book_new();
        const wsOriginal = XLSX.utils.aoa_to_sheet(allRows); 
        XLSX.utils.book_append_sheet(workbook, wsOriginal, "OTC Original");

        const adjustedRows = allRows.slice(16); 
        const TARGET_ACCOUNT_COLUMN_INDEX = 30; 
        const OLD_ACCOUNT = '4000427';
        const NEW_ACCOUNT = '4000425';
        const ALERT_CODE = 'F391501';
        const ALERT_COLUMN_INDEX = 36; 
        const SUM_COLUMN_INDEX = 51; 
        let alertCount = 0;
        let totalSum = 0;

        adjustedRows.forEach(row => {
            if (row.length > TARGET_ACCOUNT_COLUMN_INDEX && row[TARGET_ACCOUNT_COLUMN_INDEX] === OLD_ACCOUNT) {
                row[TARGET_ACCOUNT_COLUMN_INDEX] = NEW_ACCOUNT;
            }
            if (row.length > ALERT_COLUMN_INDEX && row[ALERT_COLUMN_INDEX] === ALERT_CODE) {
                alertCount++;
            }
            if (row.length > SUM_COLUMN_INDEX) {
                const value = parseFloat(row[SUM_COLUMN_INDEX]);
                if (!isNaN(value)) totalSum += value;
            }
        });

        const wsAdjusted = XLSX.utils.aoa_to_sheet(adjustedRows);
        XLSX.utils.book_append_sheet(workbook, wsAdjusted, "OTC Reclasificado");

        const outputFileName = fileName.replace(/\.[^/.]+$/, "") + "_OTC_Reporte.xlsx";
        XLSX.writeFile(workbook, outputFileName);
        
        const formattedSum = totalSum.toLocaleString('en-US', { minimumFractionDigits: 2 });
        let finalMessage = `¡Reporte OTC completado! Suma: $${formattedSum}`;
        statusMessage.textContent = finalMessage + (alertCount > 0 ? ` | ⚠️ Alerta: ${alertCount} hallazgos.` : "");
        statusMessage.style.color = alertCount > 0 ? 'orange' : 'green';
    } catch (error) {
        statusMessage.textContent = `Error: ${error.message}`;
    }
}

});

// --- FUNCIÓN UTILIDAD FECHAS ---
function datenum(v, date1904) {
    if(date1904) v+=1462;
    var epoch = v.getTime();
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}