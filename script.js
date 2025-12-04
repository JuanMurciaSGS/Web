document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('asientoForm');
    const asientoBody = document.getElementById('asientoBody');

    // Función para generar una única fila (<tr>) (Mantenemos esta función igual)
    function generarFila(data, rol, tipoAsiento) {
        let cta, valorDebe, valorHaber;

        if (rol === 'ingreso') { 
            cta = '4000411';
            if (tipoAsiento === 'diferido') {
                valorDebe = data.valor; valorHaber = '0';
            } else { // 'reconocimiento'
                valorDebe = '0'; valorHaber = data.valor;
            }
        } else { // rol === 'pasivo'
            cta = '3500600';
            if (tipoAsiento === 'diferido') {
                valorDebe = '0'; valorHaber = data.valor;
            } else { // 'reconocimiento'
                valorDebe = data.valor; valorHaber = '0';
            }
        }

        const esCtaIngreso = cta === '4000411';
        const sector = esCtaIngreso ? data.sector : "'000";
        const codeActivity = esCtaIngreso ? data.codeActivity : "'000000";
        const cecos = esCtaIngreso ? data.cecos : "'0000";
        const localidad = esCtaIngreso ? data.localidad : "'000";
        const statutory = 'B'; 

        return {
            html: `
                <tr>
                    <td>F391501</td>
                    <td>${cta}</td>
                    <td>${sector}</td>
                    <td>${codeActivity}</td>
                    <td>${cecos}</td>
                    <td>'00</td>
                    <td>${localidad}</td>
                    <td>'0000000</td>
                    <td>'00000000</td>
                    <td>${statutory}</td>
                    <td>'000000</td>
                    <td>'000000</td>
                    <td>${valorDebe}</td>
                    <td>${valorHaber}</td>
                    <td>${data.glosa}</td>
                </tr>
            `,
            esDebe: valorDebe !== '0'
        };
    }

    // Manejador del evento de envío del formulario
    form.addEventListener('submit', function(e) {
        e.preventDefault();

        const valorInput = document.getElementById('valor');
        const glosaInput = document.getElementById('glosa'); // Añadimos la referencia a Glosa para limpiar
        const valorNumerico = parseFloat(valorInput.value);

        // 🚨 VALIDACIÓN: Aseguramos que el valor sea positivo y válido
        if (isNaN(valorNumerico) || valorNumerico <= 0) {
            alert('🚨 Error: El "Valor del Movimiento" debe ser un número positivo mayor a cero.');
            valorInput.focus();
            return;
        }

        // 1. Obtener valores del formulario
        const tipoAsiento = document.getElementById('tipoAsiento').value;
        const data = {
            sector: document.getElementById('sector').value.trim(),
            codeActivity: document.getElementById('codeActivity').value.trim(),
            cecos: document.getElementById('cecos').value.trim(),
            localidad: document.getElementById('localidad').value.trim(),
            valor: Math.abs(valorNumerico).toFixed(2), 
            glosa: document.getElementById('glosa').value.trim()
        };

        // 2. *** ELIMINAMOS: asientoBody.innerHTML = ''; *** (Ya no limpiamos la tabla)

        // 3. Generar ambas filas
        const fila4000 = generarFila(data, 'ingreso', tipoAsiento);
        const fila3500 = generarFila(data, 'pasivo', tipoAsiento);
        
        let filasHTML = '';

        // 4. Determinar el orden de inserción: La línea que va al DEBE debe ir primero
        if (fila4000.esDebe) {
            filasHTML = fila4000.html + fila3500.html;
        } else {
            filasHTML = fila3500.html + fila4000.html;
        }

        // 5. *** MODIFICAMOS: Usamos insertAdjacentHTML para AGREGAR las filas ***
        asientoBody.insertAdjacentHTML('beforeend', filasHTML);
        
        // 6. Limpiamos solo los campos variables para el siguiente ingreso
        valorInput.value = ''; // Limpia el valor
        glosaInput.value = ''; // Limpia la glosa
        valorInput.focus(); // Enfoca para el siguiente ingreso
    });
});

document.addEventListener('DOMContentLoaded', function() {
    // ... (Mantener las declaraciones de form, asientoBody y la función generarFila) ...
    const form = document.getElementById('asientoForm');
    const asientoBody = document.getElementById('asientoBody');
    const exportButton = document.getElementById('exportarExcel'); // Nueva referencia

    // ... (Mantener la función generarFila igual) ...

    // ===========================================
    // NUEVA FUNCIÓN: Exportar a Excel (CSV)
    // ===========================================
    function exportarTablaAExcel() {
        const table = document.getElementById('tablaGenerada');
        let csv = [];
        
        // Iterar sobre todas las filas de la tabla (incluyendo thead y tbody)
        const rows = table.querySelectorAll('tr');

        rows.forEach(function(row) {
            let rowData = [];
            // Iterar sobre las celdas (<th> y <td>) de la fila
            const cols = row.querySelectorAll('th, td');
            
            cols.forEach(function(col) {
                // Limpiar el texto: quitar espacios extra y reemplazar comas por punto y coma
                // Esto ayuda a evitar problemas si los campos de glosa tienen comas.
                let data = col.innerText.replace(/(\r\n|\n|\r)/gm, '').replace(/"/g, '""');
                // Envolvemos los datos con comillas dobles para manejar datos con espacios o comas
                rowData.push('"' + data + '"');
            });
            
            // Unir las celdas de la fila con comas y añadir la fila al CSV
            csv.push(rowData.join(','));
        });

        // 1. Unir todas las filas CSV con saltos de línea
        let csvFile = csv.join('\n');

        // 2. Crear un Blob (paquete de datos) con el contenido CSV
        // El prefijo UTF-8 BOM ayuda a Excel a mostrar caracteres especiales (ñ, tildes) correctamente.
        const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvFile], { type: 'text/csv;charset=utf-8;' });
        
        // 3. Crear un enlace de descarga
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'Asientos_Contables_' + new Date().toISOString().slice(0, 10) + '.csv'; // Nombre del archivo

        // 4. Simular un clic para forzar la descarga
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }
    // ===========================================
    // FIN: Exportar a Excel
    // ===========================================
    
    // ... (Manejador del evento de envío del formulario - submit) ...

    form.addEventListener('submit', function(e) {
        e.preventDefault();
        // ... (todo el código de generación y validación aquí) ...
        
        // *** Código de exportación de tabla (Paso 5) ***
        // (Tu código de generación de filas debe estar aquí. Solo muestro el final del manejador)
        // 5. Insertar las filas
        asientoBody.insertAdjacentHTML('beforeend', filasHTML);
        
        // 6. Limpiamos solo los campos variables
        valorInput.value = ''; 
        glosaInput.value = ''; 
        valorInput.focus();
    });
    
    // ===========================================
    // CONECTAR EL NUEVO BOTÓN
    // ===========================================
    exportButton.addEventListener('click', exportarTablaAExcel);
});