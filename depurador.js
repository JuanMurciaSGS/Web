const fileInput = document.getElementById('fileInput');
const messageDiv = document.getElementById('message');
const downloadButton = document.getElementById('downloadButton'); 

// Contenedor para los nuevos botones de descarga
const downloadContainer = document.createElement('div'); 
// Asume que downloadButton está en el DOM, si no lo está, busca el contenedor de descarga
const parentElement = downloadButton ? downloadButton.parentNode : document.body;
if (downloadButton) {
    downloadButton.parentNode.insertBefore(downloadContainer, downloadButton.nextSibling);
    downloadButton.style.display = 'none'; // Ocultamos el botón original
} else {
    // Si el botón original no existe, insertamos el contenedor en el cuerpo o un div existente
    document.body.appendChild(downloadContainer); 
}

// Definición de las columnas que deben permanecer
const COLUMNAS_A_MANTENER = [
    'RECEIVABLE_ACCOUNT',
    'SECTOR',
    'COST CENTER',
    'LOCATION',
    'Glosa',
    'Castigo moneda original',
    'Moneda' 
];

let processedData = {}; 
let originalFileName = '';

// Escucha cuando se selecciona un archivo
fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        originalFileName = file.name.replace(/\.[^/.]+$/, "");
        messageDiv.textContent = `Archivo cargado: ${file.name}. Filtrando ceros, ordenando y separando por moneda...`;
        messageDiv.style.backgroundColor = '#d1ecf1';
        messageDiv.style.color = '#0c5460';
        downloadContainer.innerHTML = ''; // Limpiamos botones de descargas anteriores
        readExcel(file);
    }
});

/**
 * Lee el archivo de Excel, extrae los datos, los filtra (quita ceros) y los ordena.
 * @param {File} file - El archivo seleccionado por el usuario.
 */
function readExcel(file) {
    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            let jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonSheet.length < 2) {
                messageDiv.textContent = '❌ Error: El archivo parece estar vacío o solo tiene encabezados.';
                messageDiv.style.backgroundColor = '#f8d7da';
                messageDiv.style.color = '#721c24';
                return;
            }
            
            const headers = jsonSheet[0];
            let dataRows = jsonSheet.slice(1);

            // Obtener el índice de la columna a filtrar y ordenar
            const castigoIndex = headers.map(h => String(h).trim()).indexOf('Castigo moneda original');

            if (castigoIndex === -1) {
                messageDiv.textContent = '❌ Error: La columna "Castigo moneda original" no fue encontrada en el archivo.';
                messageDiv.style.backgroundColor = '#f8d7da';
                messageDiv.style.color = '#721c24';
                return;
            }
            
            // --- 1. FILTRAR: Quitar filas donde el monto sea 0 ---
            const initialLength = dataRows.length;
            dataRows = dataRows.filter(row => {
                // Limpia el valor de la celda de caracteres no numéricos y lo convierte a número
                const val = Number(String(row[castigoIndex]).replace(/[^0-9.-]/g, '')) || 0;
                return val !== 0; // Solo mantiene las filas donde el valor NO es cero
            });
            const filteredLength = dataRows.length;
            
            // --- 2. ORDENAR LOS DATOS por "Castigo moneda original" (de menor a mayor) ---
            dataRows.sort((a, b) => {
                // Se asegura que los valores se interpreten como números
                const valA = Number(String(a[castigoIndex]).replace(/[^0-9.-]/g, '')) || 0;
                const valB = Number(String(b[castigoIndex]).replace(/[^0-9.-]/g, '')) || 0;
                return valA - valB;
            });
            
            // 3. Procesar, separar y ordenar las columnas
            const separatedData = processAndSeparateData(headers, dataRows);
            
            processedData = {};
            let totalRecordsOutput = 0;
            
            downloadContainer.innerHTML = ''; 

            // 4. Crear un libro de trabajo por cada moneda encontrada
            for (const currency in separatedData) {
                const rows = separatedData[currency];
                if (rows.length > 0) {
                    const newWorksheet = XLSX.utils.json_to_sheet(rows);
                    const newWorkbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, `Reporte ${currency}`);
                    
                    processedData[currency] = newWorkbook;
                    totalRecordsOutput += rows.length;
                    
                    // Crear un botón de descarga por moneda
                    const button = document.createElement('button');
                    button.textContent = `⬇️ Descargar ${currency} Ordenado (${rows.length} registros)`;
                    button.className = 'download-currency-button';
                    button.style.cssText = 'padding: 10px 15px; background-color: #007bff; color: white; border: none; border-radius: 5px; cursor: pointer; margin-right: 10px; margin-top: 10px; font-weight: bold;';
                    
                    button.addEventListener('click', () => {
                        const timestamp = new Date().getTime();
                        const newFileName = `${originalFileName}_${currency}_ORDENADO_${timestamp}.xlsx`;
                        XLSX.writeFile(processedData[currency], newFileName);
                    });
                    downloadContainer.appendChild(button);
                }
            }
            
            messageDiv.textContent = `✅ Archivo procesado con éxito. Filas eliminadas (Castigo = 0): ${initialLength - filteredLength}. Total de registros a descargar: ${totalRecordsOutput}.`;
            messageDiv.style.backgroundColor = '#d4edda';
            messageDiv.style.color = '#155724';

        } catch (error) {
            console.error("Error al procesar:", error);
            messageDiv.textContent = `❌ Error al procesar el archivo: ${error.message}. Asegúrate de que es un archivo Excel válido.`;
            messageDiv.style.backgroundColor = '#f8d7da';
            messageDiv.style.color = '#721c24';
        }
    };

    reader.onerror = (error) => {
        messageDiv.textContent = `❌ Error de lectura de archivo: ${error.message}`;
        messageDiv.style.backgroundColor = '#f8d7da';
        messageDiv.style.color = '#721c24';
    };

    reader.readAsArrayBuffer(file);
}

/**
 * Aplica el formato de columnas, mueve montos a la columna DEBE 
 * con valor absoluto (positivo), y separa los datos por la columna 'Moneda'.
 */
function processAndSeparateData(headers, dataRows) {
    const separatedOutput = { 'PEN': [], 'USD': [] }; 

    const headerMap = new Map();
    headers.forEach((header, index) => {
        const cleanHeader = header ? String(header).trim() : null; 
        if (cleanHeader) {
            headerMap.set(cleanHeader, index);
        }
    });

    // Función auxiliar para obtener el valor de una columna original
    const getOriginalValue = (colName, row) => {
        const index = headerMap.get(colName);
        return (index !== undefined && row[index] !== undefined) ? row[index] : '';
    };


    dataRows.forEach(row => {
        // Obtener la moneda
        let currency = String(getOriginalValue('Moneda', row)).toUpperCase().trim();
        
        if (currency !== 'PEN' && currency !== 'USD') {
            return; 
        }

        // --- LÓGICA DE MONTO DEBE ---
        const castigoValueRaw = String(getOriginalValue('Castigo moneda original', row));
        // Limpia el valor, lo convierte a número
        const castigoValue = Number(castigoValueRaw.replace(/[^0-9.-]/g, '')) || 0;
        
        // El monto siempre será el valor absoluto (positivo)
        const debeValue = Math.abs(castigoValue); 
        const haberValue = 0; // Se mantiene en 0.
        // ------------------------------------


        // Aplicar el orden de columnas deseado:
        const newRow = {};
        
        // ************************************************
        // ORDEN DE COLUMNAS FINAL
        // ************************************************
        
        newRow['FCODE'] = 'F391501'; 
        newRow['RECEIVABLE_ACCOUNT'] = getOriginalValue('RECEIVABLE_ACCOUNT', row);
        newRow['SECTOR'] = getOriginalValue('SECTOR', row);
        newRow['ACT'] = '000000'; 
        newRow['COST CENTER'] = getOriginalValue('COST CENTER', row);
        newRow['CL'] = '00';
        newRow['LOCATION'] = getOriginalValue('LOCATION', row); 
        newRow['INTERCOMPANY'] = '0000000';
        newRow['PROJECT'] = '00000000';
        newRow['STATUTORY'] = 'B';
        newRow['RESERVED 1'] = '000000';
        newRow['RESERVED 2'] = '000000';
        
        // --- COLUMNAS DE MONTO ---
        newRow['DEBE'] = debeValue.toFixed(2); 
        newRow['HABER'] = haberValue.toFixed(2);
        
        newRow['Glosa'] = getOriginalValue('Glosa', row);
        newRow['Castigo moneda original'] = castigoValueRaw;
        
        // ************************************************
        
        // Separar la fila en el array correspondiente (PEN o USD)
        separatedOutput[currency].push(newRow);
    });

    return separatedOutput;
}