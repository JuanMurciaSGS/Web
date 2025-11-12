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
        processFile = processUnidentifyFile; // Lógica para XLS/XLSX
    } else {
        processFile = processAgeingFile; // Lógica para TXT/TSV
    }

    // 2. EVENT LISTENER PRINCIPAL (Adaptado para leer TXT o XLS/XLSX)
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
            // La función de procesamiento recibe el contenido (texto o ArrayBuffer)
            processFile(e.target.result, file.name);
        };

        // Si es Unidentify, leemos el archivo como ArrayBuffer (necesario para XLS/XLSX).
        if (reportType === 'unidentify') {
             reader.readAsArrayBuffer(file);
        } else {
             // Si es OTC o Ageing, leemos el archivo como texto (necesario para TXT/TSV).
             reader.readAsText(file);
        }
    });

// ----------------------------------------------------------------------
// --- LÓGICA ESPECÍFICA PARA EL REPORTE UNIDENTIFY (Archivos XLS/XLSX) ---
// ----------------------------------------------------------------------
function processUnidentifyFile(dataArrayBuffer, fileName) {
    try {
        // Cargar el libro de trabajo (workbook) desde el ArrayBuffer
        const workbook = XLSX.read(dataArrayBuffer, { type: 'array' });
        
        // Asumimos que queremos la primera hoja
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convertir la hoja a formato Array of Arrays (AoA)
        // Opciones para manejar fechas correctamente
        const allRows = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1, 
            raw: true,
            // Convertir celdas de fecha a objetos Date de JavaScript
            cellDates: true, 
            // Formato de fecha para la conversión, si es necesario
            dateNF: 'dd/mm/yyyy' 
        });

        if (allRows.length === 0) {
            statusMessage.textContent = 'El archivo está vacío o la hoja no contiene datos.';
            statusMessage.style.color = 'red';
            return;
        }

        const headers = allRows[0];
        const dataRows = allRows.slice(1);
        
        // === CONFIGURACIÓN DE CLASIFICACIÓN ===
        // Columna 16 es el índice 15
        const CLASS_COLUMN_INDEX = 15; 
        const CLASS_COLUMN_NAME = headers[CLASS_COLUMN_INDEX];
        
        if (!CLASS_COLUMN_NAME) {
            statusMessage.textContent = `Error: La columna de clasificación (índice ${CLASS_COLUMN_INDEX}) no existe.`;
            statusMessage.style.color = 'red';
            return;
        }

        // 1. Análisis y Clasificación de Datos (Similar a Ageing)
        const sheetsData = {}; 

        // Índices de las columnas que contienen fechas (6, 14, 15 -> 5, 13, 14)
        const DATE_INDICES = [5, 13, 14];

        dataRows.forEach(row => {
            
            // Si el valor es un objeto Date (debido a cellDates: true), lo convertimos al formato DD/MM/AAAA
            DATE_INDICES.forEach(index => {
                if (row[index] instanceof Date) {
                    const day = row[index].getDate().toString().padStart(2, '0');
                    const month = (row[index].getMonth() + 1).toString().padStart(2, '0');
                    const year = row[index].getFullYear();
                    row[index] = `${day}/${month}/${year}`;
                }
            });

            // Usamos String() para asegurar que el valor se pueda usar como clave (nombre de la hoja)
            const classValue = String(row[CLASS_COLUMN_INDEX] || "SIN BANCO"); 
            
            if (!sheetsData[classValue]) {
                sheetsData[classValue] = [headers];
            }
            sheetsData[classValue].push(row);
        });

        // 2. Creación del Nuevo Archivo XLSX con Múltiples Hojas
        const outputWorkbook = XLSX.utils.book_new();

        for (const classValue in sheetsData) {
            if (sheetsData.hasOwnProperty(classValue)) {
                const data = sheetsData[classValue];
                const ws = XLSX.utils.aoa_to_sheet(data); 
                // Limitar el nombre a 31 caracteres
                const sheetName = classValue.substring(0, 31); 
                XLSX.utils.book_append_sheet(outputWorkbook, ws, sheetName);
            }
        }

        // 3. Descarga del Archivo Excel (.xlsx)
        const outputFileName = fileName.replace(/\.[^/.]+$/, "") + "_UnidentifyReport.xlsx";
        
        XLSX.writeFile(outputWorkbook, outputFileName);
        
        statusMessage.textContent = `¡Conversión y clasificación exitosa! Archivo Excel con ${Object.keys(sheetsData).length} hojas (por Banco) descargado.`;
        statusMessage.style.color = 'green';

    } catch (error) {
        console.error("Error durante el procesamiento del archivo Unidentify:", error);
        statusMessage.textContent = `Error al procesar el archivo. Asegúrate de que sea un archivo XLS/XLSX válido. Detalle: ${error.message}`;
        statusMessage.style.color = 'red';
    }
}

// ----------------------------------------------------------------------
// --- LÓGICA ESPECÍFICA PARA EL REPORTE AGEING (Múltiples Hojas) ---
// ----------------------------------------------------------------------
function processAgeingFile(fileContent, fileName) {
    try {
        const allRows = fileContent.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.split('\t')); // Asumimos separador TAB (\t)

        if (allRows.length === 0) {
            statusMessage.textContent = 'El archivo está vacío.';
            statusMessage.style.color = 'red';
            return;
        }

        const headers = allRows[0];
        const dataRows = allRows.slice(1);
        
        // Columna 10 (índice 9) para clasificar
        const CLASS_COLUMN_INDEX = 9; 
        const CLASS_COLUMN_NAME = headers[CLASS_COLUMN_INDEX];
        
        if (!CLASS_COLUMN_NAME) {
            statusMessage.textContent = `Error: La columna de clasificación (índice ${CLASS_COLUMN_INDEX}) no existe.`;
            statusMessage.style.color = 'red';
            return;
        }

        // Análisis y Clasificación de Datos en sheetsData
        const sheetsData = {}; 
        dataRows.forEach(row => {
            const classValue = row[CLASS_COLUMN_INDEX] || "SIN CLASIFICAR";
            
            if (!sheetsData[classValue]) {
                sheetsData[classValue] = [headers];
            }
            sheetsData[classValue].push(row);
        });

        // Creación del Archivo XLSX con Múltiples Hojas
        const workbook = XLSX.utils.book_new();

        for (const classValue in sheetsData) {
            if (sheetsData.hasOwnProperty(classValue)) {
                const data = sheetsData[classValue];
                const ws = XLSX.utils.aoa_to_sheet(data); 
                const sheetName = classValue.substring(0, 31);
                XLSX.utils.book_append_sheet(workbook, ws, sheetName);
            }
        }

        // Descarga
        const outputFileName = fileName.replace(/\.[^/.]+$/, "") + "_AgeingReport.xlsx";
        XLSX.writeFile(workbook, outputFileName);
        
        statusMessage.textContent = `¡Conversión y análisis exitoso! Archivo Excel con ${Object.keys(sheetsData).length} hojas descargado.`;
        statusMessage.style.color = 'green';

    } catch (error) {
        console.error("Error durante el procesamiento del archivo AGEING:", error);
        statusMessage.textContent = `Error al procesar el archivo AGEING. Asegúrate del formato. Detalle: ${error.message}`;
        statusMessage.style.color = 'red';
    }
}

// ----------------------------------------------------------------------
// --- LÓGICA ESPECÍFICA PARA EL REPORTE OTC (Original + Reclasificado con Alerta y Suma) ---
// ----------------------------------------------------------------------
function processOtcFile(fileContent, fileName) {
     try {
        // 1. Parsear el contenido (datos originales)
        const allRows = fileContent.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.split('\t')); // Asumimos separador TAB (\t)

        if (allRows.length === 0) {
            statusMessage.textContent = 'El archivo está vacío.';
            statusMessage.style.color = 'red';
            return;
        }

        const workbook = XLSX.utils.book_new();

        // 2a. Hoja 1: "OTC Original" (Datos completos sin modificar)
        const wsOriginal = XLSX.utils.aoa_to_sheet(allRows); 
        XLSX.utils.book_append_sheet(workbook, wsOriginal, "OTC Original");

        // 2b. Preparación de datos para la Hoja 2: "OTC Reclasificado"
        
        // Eliminar las primeras 16 filas (índice 0 hasta 15)
        const adjustedRows = allRows.slice(16); 

        // Configuración de la Reclasificación
        const TARGET_ACCOUNT_COLUMN_INDEX = 30; // Columna 31 es el índice 30
        const OLD_ACCOUNT = '4000427';
        const NEW_ACCOUNT = '4000425';

        // Configuración y contador de Alerta
        const ALERT_CODE = 'F391501';
        const ALERT_COLUMN_INDEX = 36; // Columna 37 es el índice 36
        let alertCount = 0;

        // Configuración y Acumulador de Suma
        const SUM_COLUMN_INDEX = 51; // Columna 52 es el índice 51
        let totalSum = 0;

        // Iterar, reclasificar, contar la alerta Y REALIZAR LA SUMA
        adjustedRows.forEach(row => {
            // Reclasificación (Columna 31 / Índice 30)
            if (row.length > TARGET_ACCOUNT_COLUMN_INDEX) {
                if (row[TARGET_ACCOUNT_COLUMN_INDEX] === OLD_ACCOUNT) {
                    row[TARGET_ACCOUNT_COLUMN_INDEX] = NEW_ACCOUNT;
                }
            }
            
            // Conteo de Alerta (Columna 37 / Índice 36)
            if (row.length > ALERT_COLUMN_INDEX) {
                if (row[ALERT_COLUMN_INDEX] === ALERT_CODE) {
                    alertCount++;
                }
            }
            
            // CÁLCULO DE LA SUMA (Columna 52 / Índice 51)
            if (row.length > SUM_COLUMN_INDEX) {
                // Intentamos convertir el valor a un número y lo sumamos
                const value = parseFloat(row[SUM_COLUMN_INDEX]);
                if (!isNaN(value)) {
                    totalSum += value;
                }
            }
        });

        // Hoja 2: Crear la hoja de cálculo con los datos ajustados
        const wsAdjusted = XLSX.utils.aoa_to_sheet(adjustedRows);
        XLSX.utils.book_append_sheet(workbook, wsAdjusted, "OTC Reclasificado");

        // 3. Descarga del Archivo Excel (.xlsx)
        const outputFileName = fileName.replace(/\.[^/.]+$/, "") + "_OTC_Reporte.xlsx";
        
        XLSX.writeFile(workbook, outputFileName);
        
        // 4. Mensaje de Estado y Alerta (INCLUYENDO LA SUMA)

        // Formatear la suma para mostrarla mejor, con 2 decimales y separador de miles
        const formattedSum = totalSum.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        
        let finalMessage = `¡Reporte OTC completado! Archivo Excel con las hojas 'OTC Original' y 'OTC Reclasificado' descargado.`;
        let messageColor = 'green';
        
        // Agregar la suma al mensaje principal
        finalMessage += ` | **Suma Columna 52: $${formattedSum}**`;
        
        // Agregar la alerta si aplica
        if (alertCount > 0) {
            finalMessage += ` | ¡⚠️ ALERTA! Código ${ALERT_CODE} encontrado ${alertCount} veces.`;
            messageColor = 'orange'; 
        }

        statusMessage.textContent = finalMessage;
        statusMessage.style.color = messageColor;
        
    } catch (error) {
        console.error("Error durante el procesamiento del archivo OTC:", error);
        statusMessage.textContent = `Error al procesar el archivo OTC. Asegúrate del formato. Detalle: ${error.message}`;
        statusMessage.style.color = 'red';
    }
}
});