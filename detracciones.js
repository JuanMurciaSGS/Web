// detracciones.js

document.addEventListener('DOMContentLoaded', () => {
    const sunatFileInput = document.getElementById('sunatFileInput');
    const planoFileInput = document.getElementById('planoFileInput');
    const processButton = document.getElementById('processButton');
    const statusMessage = document.getElementById('statusMessage');
    const downloadLink = document.getElementById('downloadLink');

    // --- 1. DEFINICIÓN DE COLUMNAS Y CONSTANTES ---
    const KEY_COLUMN = "Numero Constancia";
    const CRUCE_COLUMN = "Cruce"; // Nombre de la nueva columna
    const PLANO_SHEET_NAME = "SGS DEL PERU S.A.C.";
    
    // ** SOLO LAS COLUMNAS QUE NECESITAN FORMATO NUMÉRICO O FECHA **
    const CELL_FORMATS = {
        // Formato para Cantidades o Montos (2 decimales, separador de miles)
        "Monto de deposito": { type: 'n', format: '#,##0.00' }, 
        
        // Formato para Fechas (dd/mm/yyyy)
        "Fecha Pago": { type: 'n', format: 'dd/mm/yyyy' },
        "Fecha de Descarga": { type: 'n', format: 'yyyy-mm-dd' }
        
        // >> NOTA: Se ha ELIMINADO la columna "Numero de Documento Adquiriente" de esta lista.
        // Se manejará como texto por defecto.
    };
    
    const EXPECTED_PLANO_HEADERS = [
        "Fecha de Descarga",
        "Semana",
        "Nº",
        "Tipo de Cuenta",
        "Numero de Cuenta",
        KEY_COLUMN,
        "OPERACIÓN ORACLE",
        "Periodo Tributario",
        "RUC Proveedor",
        "Nombre Proveedor",
        "Tipo de Documento Adquiriente",
        "Numero de Documento Adquiriente", // Columna I
        "Nombre/Razon Social del Adquiriente",
        "Fecha Pago",
        "Monto de deposito",
        "Tipo Bien",
        "Tipo Operacion",
        "Tipo de Comprobante",
        "Serie de Comprobante",
        "Facturas",
        "Estado"
    ];

    // --- Lógica Principal: LECTURA, VALIDACIÓN Y CRUCE (Se mantiene igual) ---
    processButton.addEventListener('click', async () => {
        statusMessage.textContent = '';
        downloadLink.style.display = 'none';

        if (!sunatFileInput.files.length || !planoFileInput.files.length) {
            statusMessage.textContent = '🛑 Por favor, sube ambos archivos (SUNAT y Plano/Maestro).';
            statusMessage.style.color = 'red';
            return;
        }

        const sunatFile = sunatFileInput.files[0];
        const planoFile = planoFileInput.files[0];

        try {
            statusMessage.textContent = '⏳ Leyendo y validando Archivos...';
            statusMessage.style.color = 'blue';

            const sunatData = await readExcelFile(sunatFile);
            if (!sunatData.length || !Object.keys(sunatData[0]).includes(KEY_COLUMN)) {
                throw new Error(`El Archivo SUNAT no tiene datos válidos o le falta la columna clave '${KEY_COLUMN}'.`);
            }
            console.log(`✅ Archivo 1 (SUNAT) leído. Registros: ${sunatData.length}`);

            const planoData = await readExcelFile(planoFile, PLANO_SHEET_NAME);
            
            const planoHeaders = Object.keys(planoData[0]);
            validateFileStructure(planoHeaders, EXPECTED_PLANO_HEADERS, 'Archivo Plano/Maestro');
            console.log(`✅ Archivo 2 (Plano/Maestro) leído y validado. Registros: ${planoData.length} (Hoja: ${PLANO_SHEET_NAME})`);

            statusMessage.textContent = '🔄 Realizando Cruce de Registros...';

            const newRecords = findNewRecords(sunatData, planoData, KEY_COLUMN, CRUCE_COLUMN);

            if (newRecords.length === 0) {
                statusMessage.textContent = `🎉 Cruce completo. No se encontraron nuevos registros (0 nuevos) en el Archivo SUNAT.`;
                statusMessage.style.color = 'green';
                return;
            }

            const fileName = `Nuevos_Registros_Detracciones_${new Date().toISOString().slice(0, 10)}.xlsx`;
            exportToExcel(newRecords, fileName, sunatData, CELL_FORMATS); 
            
            statusMessage.textContent = `✅ Cruce finalizado. Se encontraron ${newRecords.length} nuevos registros. ¡Listo para descargar!`;
            statusMessage.style.color = 'green';
            downloadLink.style.display = 'block';

        } catch (error) {
            console.error('Error durante el proceso:', error);
            statusMessage.textContent = `🛑 Error fatal: ${error.message || 'Ocurrió un error inesperado.'}`;
            statusMessage.style.color = 'red';
        }
    });

    // --- FUNCIONES DE CRUCE Y EXPORTACIÓN ---

    /**
     * Realiza el cruce y prepara los datos para la exportación.
     */
    function findNewRecords(file1Data, file2Data, keyColumn, cruceColumn) {
        const existingKeys = new Set(
            file2Data
                .map(record => String(record[keyColumn]).trim())
                .filter(key => key.length > 0)
        );

        const newRecords = file1Data.filter(record => {
            const key = String(record[keyColumn]).trim();
            return key.length > 0 && !existingKeys.has(key);
        });
        
        const processedRecords = newRecords.map(record => {
             const newRecord = { ...record };
             
             // Conversión del Monto (Asegurar que sea número para el formato)
             const monto = newRecord['Monto de deposito'];
             if (monto) {
                 let cleanMonto = String(monto).replace(/[$,]/g, '').trim();
                 newRecord['Monto de deposito'] = parseFloat(cleanMonto) || 0;
             }
             
             // ** IMPORTANTE: Aseguramos que "Numero de Documento Adquiriente" sea un string **
             const docAdquiriente = newRecord['Numero de Documento Adquiriente'];
             if (docAdquiriente) {
                 newRecord['Numero de Documento Adquiriente'] = String(docAdquiriente).trim();
             }
             

             newRecord[cruceColumn] = "Nuevo";
             return newRecord;
        });

        return processedRecords;
    }

    /**
     * Genera y descarga un archivo Excel aplicando formatos numéricos y de fecha.
     */
    function exportToExcel(data, fileName, sourceData, formats) {
        if (data.length === 0) return;

        // 1. Obtener y ordenar los encabezados
        const originalHeaders = Object.keys(sourceData[0]);
        const constanciaIndex = originalHeaders.indexOf(KEY_COLUMN);
        
        let orderedHeaders = [...originalHeaders];
        if (constanciaIndex !== -1) {
            orderedHeaders.splice(constanciaIndex + 1, 0, CRUCE_COLUMN);
        } else {
            orderedHeaders.push(CRUCE_COLUMN);
        }
        
        // 2. Mapear los datos
        const dataForSheet = data.map(record => {
            let orderedRecord = {};
            orderedHeaders.forEach(header => {
                orderedRecord[header] = record[header] !== undefined ? record[header] : "";
            });
            return orderedRecord;
        });

        // 3. Crear la hoja de cálculo
        const worksheet = XLSX.utils.json_to_sheet(dataForSheet, { 
            header: orderedHeaders,
            skipHeader: false 
        });

        // 4. APLICAR FORMATO DE FECHAS Y NÚMEROS A LAS CELDAS
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        for(let C = range.s.c; C <= range.e.c; ++C) {
            const headerCellRef = XLSX.utils.encode_cell({ r: range.s.r, c: C });
            const header = worksheet[headerCellRef]?.v; 
            
            const formatInfo = formats[header];

            // Iterar sobre las filas de datos (a partir de la fila 1)
            for(let R = range.s.r + 1; R <= range.e.r; ++R) { 
                const cellAddress = { r: R, c: C };
                const cellRef = XLSX.utils.encode_cell(cellAddress);
                const cell = worksheet[cellRef];

                if (cell) {
                    
                    // A) Aplicar formato de FECHA/NÚMERO (si está en la lista)
                    if (formatInfo) {
                        cell.t = formatInfo.type; // Forzar a 'n' (number)
                        cell.z = formatInfo.format; // Aplicar formato de Excel
                    } 
                    
                    // B) Forzar TIPO TEXTO para la columna específica
                    else if (header === 'Numero de Documento Adquiriente') {
                         cell.t = 's'; // Forzar a 's' (string/text)
                         cell.z = undefined; // Quitar cualquier formato numérico
                    }
                    
                    // C) Si no tiene formato especial, SheetJS usa el tipo inferido o 's'
                }
            }
        }

        // 5. Crear y escribir el libro de trabajo
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Nuevos Registros");
        
        XLSX.writeFile(workbook, fileName);

        downloadLink.href = "#";
    }


    // --- FUNCIONES DE LECTURA Y VALIDACIÓN (Se mantienen las optimizadas) ---

    function validateFileStructure(readHeaders, expectedHeaders, fileName) {
        if (readHeaders.length !== expectedHeaders.length) {
            throw new Error(`El ${fileName} tiene una cantidad de columnas incorrecta. Se esperaban ${expectedHeaders.length} pero se encontraron ${readHeaders.length}.`);
        }

        const structureIsValid = expectedHeaders.every((expectedHeader, index) => {
            return expectedHeader === readHeaders[index];
        });

        if (!structureIsValid) {
            const mismatchedIndex = expectedHeaders.findIndex((expected, index) => expected !== readHeaders[index]);
            const expected = expectedHeaders[mismatchedIndex];
            const read = readHeaders[mismatchedIndex] || '[COLUMNA AUSENTE]';
            throw new Error(`Estructura incorrecta en el ${fileName} (Columna ${mismatchedIndex + 1}). Se esperaba: '${expected}', se encontró: '${read}'. Verifique el orden.`);
        }
    }

    /**
     * Función robusta para leer Archivos Excel (XLSX/XLS)
     * Se usa raw: true para conservar tipos (números, fechas seriales).
     */
    function readExcelFile(file, sheetName = null) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { 
                        type: 'array',
                        cellDates: false // Mantiene las fechas como números seriales
                    }); 

                    const targetSheetName = sheetName || workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[targetSheetName];

                    if (!worksheet) {
                        return reject(new Error(`La hoja '${targetSheetName}' no existe en el archivo.`));
                    }

                    // USAR raw: true para conservar los tipos
                    const json = XLSX.utils.sheet_to_json(worksheet, { 
                        raw: true, 
                        defval: "", 
                    });
                    
                    if (json.length === 0) {
                        return reject(new Error(`La hoja '${targetSheetName}' no tiene datos.`));
                    }
                    
                    // Filtrar filas completamente vacías
                    const finalData = json.filter(obj => {
                        const values = Object.values(obj);
                        return values.length > 0 && values.some(v => String(v).trim() !== '');
                    });

                    resolve(finalData);

                } catch (error) {
                    reject(new Error(`Fallo al procesar el archivo Excel: ${error.message}`));
                }
            };
            
            reader.onerror = (e) => reject(new Error(`Fallo al leer el archivo: ${file.name}`));
            reader.readAsArrayBuffer(file);
        });
    }

    // Funciones que no se usan en este paso:
    function parseCsvData() { return []; }
});