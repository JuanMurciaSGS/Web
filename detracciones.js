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
    
    // Encabezados esperados para el Archivo 2 (Plano/Maestro)
    // Se mantiene la estructura original del Archivo 2, sin la columna "Cruce" para la validación.
    // Sin embargo, si la columna "Cruce" debe ser parte de la estructura para la validación, la agregamos aquí:
    const EXPECTED_PLANO_HEADERS = [
        "Fecha de Descarga",
        "Semana",
        "Nº",
        "Tipo de Cuenta",
        "Numero de Cuenta",
        KEY_COLUMN,
        // ** NOTA IMPORTANTE: La columna "Cruce" no debe estar aquí si NO existe en el Archivo Plano **
        // Si el Archivo Plano/Maestro tiene la columna 'Cruce', déjala aquí:
        // CRUCE_COLUMN,
        "OPERACIÓN ORACLE",
        "Periodo Tributario",
        "RUC Proveedor",
        "Nombre Proveedor",
        "Tipo de Documento Adquiriente",
        "Numero de Documento Adquiriente",
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

    // Se asume que el Archivo 1 (SUNAT) tiene todos los datos que necesitamos, incluyendo las columnas del Archivo 2 (Plano)
    // excepto la columna "Cruce" que se va a inyectar.

    // --- Lógica Principal: LECTURA, VALIDACIÓN Y CRUCE ---
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

            // 1. Procesar Archivo 1 (SUNAT)
            const sunatData = await readExcelFile(sunatFile);
            if (!sunatData.length || !Object.keys(sunatData[0]).includes(KEY_COLUMN)) {
                throw new Error(`El Archivo SUNAT no tiene datos válidos o le falta la columna clave '${KEY_COLUMN}'.`);
            }
            console.log(`✅ Archivo 1 (SUNAT) leído. Registros: ${sunatData.length}`);

            // 2. Procesar Archivo 2 (Plano/Maestro)
            const planoData = await readExcelFile(planoFile, PLANO_SHEET_NAME);
            
            // Validar la estructura del Archivo 2
            const planoHeaders = Object.keys(planoData[0]);
            validateFileStructure(planoHeaders, EXPECTED_PLANO_HEADERS, 'Archivo Plano/Maestro');
            console.log(`✅ Archivo 2 (Plano/Maestro) leído y validado. Registros: ${planoData.length} (Hoja: ${PLANO_SHEET_NAME})`);

            statusMessage.textContent = '🔄 Realizando Cruce de Registros...';

            // 3. Realizar el Cruce y añadir la columna "Cruce"
            const newRecords = findNewRecords(sunatData, planoData, KEY_COLUMN, CRUCE_COLUMN);

            if (newRecords.length === 0) {
                statusMessage.textContent = `🎉 Cruce completo. No se encontraron nuevos registros (0 nuevos) en el Archivo SUNAT.`;
                statusMessage.style.color = 'green';
                return;
            }

            // 4. Exportar los Nuevos Registros
            const fileName = `Nuevos_Registros_Detracciones_${new Date().toISOString().slice(0, 10)}.xlsx`;
            exportToExcel(newRecords, fileName, sunatData); // Pasamos sunatData para obtener el orden de las columnas.
            
            // 5. Mostrar éxito y enlace de descarga
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
     * Realiza el cruce y prepara los datos para la exportación inyectando la columna 'Cruce'.
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
        
        // 1. Inyectar la columna "Cruce" con valor "Nuevo" en cada registro
        const processedRecords = newRecords.map(record => {
             // Clonar el registro original
            const newRecord = { ...record };
            // Establecer el valor de la columna Cruce
            newRecord[cruceColumn] = "Nuevo";
            return newRecord;
        });

        return processedRecords;
    }

    /**
     * Genera y descarga un archivo Excel manteniendo el orden de las columnas del Archivo 1,
     * e insertando la columna "Cruce" en la posición solicitada.
     */
    function exportToExcel(data, fileName, sourceData) {
        if (data.length === 0) return;

        // 1. Obtener el orden de las columnas del Archivo 1 (sunatData)
        // Se asume que el primer registro contiene todos los encabezados originales.
        const originalHeaders = Object.keys(sourceData[0]);

        // 2. Determinar la posición de inserción de la columna "Cruce"
        const constanciaIndex = originalHeaders.indexOf(KEY_COLUMN);
        
        // El nuevo orden de encabezados:
        let orderedHeaders = [...originalHeaders];
        
        if (constanciaIndex !== -1) {
            // Insertar "Cruce" justo después de "Numero Constancia"
            orderedHeaders.splice(constanciaIndex + 1, 0, CRUCE_COLUMN);
        } else {
            // Si no encuentra la columna clave, añade "Cruce" al final (fallback)
            orderedHeaders.push(CRUCE_COLUMN);
        }
        
        // 3. Mapear los datos para asegurar el orden de las columnas en el Excel
        const dataForSheet = data.map(record => {
            let orderedRecord = {};
            orderedHeaders.forEach(header => {
                // Si la columna es 'Cruce', usa el valor inyectado ('Nuevo').
                // Para las demás, usa el valor del registro (o cadena vacía si no existe).
                orderedRecord[header] = record[header] !== undefined ? record[header] : "";
            });
            return orderedRecord;
        });

        // 4. Crear la hoja de cálculo
        const worksheet = XLSX.utils.json_to_sheet(dataForSheet, { 
            header: orderedHeaders 
        });

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Nuevos Registros");
        
        // Escribir el archivo y forzar la descarga
        XLSX.writeFile(workbook, fileName);

        downloadLink.href = "#";
    }


    // --- FUNCIONES DE LECTURA Y VALIDACIÓN (Se mantienen las funciones robustas) ---

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
     * @param {File} file - El objeto File a leer.
     * @param {string} [sheetName] - Nombre de la hoja a leer. Si es null/undefined, lee la primera hoja.
     */
    function readExcelFile(file, sheetName = null) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { 
                        type: 'array',
                        cellDates: false 
                    }); 

                    const targetSheetName = sheetName || workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[targetSheetName];

                    if (!worksheet) {
                        return reject(new Error(`La hoja '${targetSheetName}' no existe en el archivo.`));
                    }

                    const json = XLSX.utils.sheet_to_json(worksheet, { 
                        raw: false, 
                        defval: "", 
                        header: 1, 
                        range: "A1:ZZ10000"
                    });
                    
                    if (json.length < 2) {
                        return reject(new Error(`La hoja '${targetSheetName}' está vacía o no tiene encabezados válidos (Fila 1).`));
                    }

                    const headers = json[0].map(h => 
                        String(h)
                            .trim()
                            .replace(/\s+/g, ' ')
                            .trim()
                    ).filter(h => h.length > 0); 

                    const dataRows = json.slice(1);
                    
                    const finalData = dataRows.map(row => {
                        let obj = {};
                        headers.forEach((header, index) => {
                            obj[header] = row[index] !== undefined ? String(row[index]).trim() : '';
                        });
                        return obj;
                    }).filter(obj => {
                        const values = Object.values(obj);
                        return values.length > 0 && values.some(v => v !== '');
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