document.addEventListener('DOMContentLoaded', function() {
    // 1. Referencias al DOM
    const form = document.getElementById('asientoForm');
    const asientoBody = document.getElementById('asientoBody');
    const exportButton = document.getElementById('exportarExcel');
    
    // Función para generar una única fila (<tr>)
    function generarFila(data, rol, tipoAsiento) {
        let cta, valorDebe, valorHaber;

        // **Ajuste de CTAs y lógica DEBE/HABER**
        if (rol === 'ingreso') { // Asiento/Extorno que afecta a la cuenta de INGRESO (4000)
            cta = '4000415';
            if (tipoAsiento === 'asiento') { // ASENTAR INGRESO: Ingreso va al DEBE
                valorDebe = '0'; valorHaber = data.valor;
            } else { // EXTORNAR INGRESO: Ingreso va al HABER
                valorDebe = data.valor; valorHaber = '0';
            }
        } else { // rol === 'pasivo' (Afecta a la cuenta de PASIVO/DIFERIDO 1651)
            cta = '1651111';
            if (tipoAsiento === 'asiento') { // ASENTAR INGRESO: Pasivo va al HABER
                valorDebe = data.valor; valorHaber = '0';
            } else { // EXTORNAR INGRESO: Pasivo va al DEBE
                valorDebe = '0'; valorHaber = data.valor;
            }
        }

        // Definición de campos
        const esCtaIngreso = cta === '4000415';
        // Usamos valores por defecto si los campos del formulario están vacíos
        const sector = esCtaIngreso ? data.sector || '90' : "'000";
        const codeActivity = esCtaIngreso ? data.codeActivity || '9053' : "'000000";
        const cecos = esCtaIngreso ? data.cecos || '9115' : "'0000";
        const localidad = esCtaIngreso ? data.localidad || '024' : "'000";
        const statutory = 'B';
        const cl = '00'; 
        
        // Formato para el HTML
        const html = `
            <tr>
                <td>F391501</td>
                <td>${cta}</td>
                <td>${sector}</td>
                <td>${codeActivity}</td>
                <td>${cecos}</td>
                <td>'${cl}</td>
                <td>${localidad}</td>
                <td>0000000</td>
                <td>00000000</td>
                <td>${statutory}</td>
                <td>000000</td>
                <td>000000</td>
                <td style="text-align: right; font-weight: bold;">${valorDebe}</td>
                <td style="text-align: right; font-weight: bold;">${valorHaber}</td>
                <td>${data.glosa}</td>
            </tr>
        `;
        
        return {
            html: html,
            esDebe: valorDebe !== '0'
        };
    }

    // ===========================================
    // Manejador del evento de envío del formulario (Generación)
    // ===========================================
    form.addEventListener('submit', function(e) {
        e.preventDefault();

        const valorInput = document.getElementById('valor');
        const glosaInput = document.getElementById('glosa');
        const valorNumerico = parseFloat(valorInput.value);

        // VALIDACIÓN
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

        // 2. Generar ambas filas
        const fila4000 = generarFila(data, 'ingreso', tipoAsiento);
        const fila1651 = generarFila(data, 'pasivo', tipoAsiento);
        
        let filasHTML = '';

        // 3. Determinar el orden de inserción: La línea que va al DEBE debe ir primero
        // Esto asegura el formato contable (DEBE arriba, HABER abajo)
        if (fila4000.esDebe) {
            filasHTML = fila4000.html + fila1651.html;
        } else {
            filasHTML = fila1651.html + fila4000.html;
        }

        // 4. Agregar las filas al final de la tabla (insertAdjacentHTML)
        asientoBody.insertAdjacentHTML('beforeend', filasHTML);
        
        // 5. Limpiamos solo los campos variables
        valorInput.value = ''; 
        glosaInput.value = ''; 
        valorInput.focus();
    });

    // ===========================================
    // Función: Exportar a Excel (CSV)
    // ===========================================
    function exportarTablaAExcel() {
        const table = document.getElementById('tablaGenerada');
        // Verificar si hay datos en el cuerpo de la tabla
        if (asientoBody.rows.length === 0) {
            alert('No hay asientos generados para exportar.');
            return;
        }

        let csv = [];
        
        // 1. Obtener encabezados y filas
        const rows = table.querySelectorAll('tr');

        rows.forEach(function(row) {
            let rowData = [];
            const cols = row.querySelectorAll('th, td');
            
            cols.forEach(function(col) {
                let data = col.innerText.replace(/(\r\n|\n|\r)/gm, '').trim();
                // Si el campo es numérico (DEBE/HABER), asegurarse de que el separador decimal sea punto (.)
                if (col.cellIndex >= 12 && col.cellIndex <= 13) {
                     data = data.replace(',', '.');
                }
                // Envolvemos con comillas dobles
                rowData.push('"' + data.replace(/"/g, '""') + '"');
            });
            csv.push(rowData.join(','));
        });

        let csvFile = csv.join('\n');

        // 2. Crear Blob con BOM (Byte Order Mark) para UTF-8 y compatibilidad con tildes en Excel
        const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvFile], { type: 'text/csv;charset=utf-8;' });
        
        // 3. Crear enlace y simular clic
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'Asientos_Contables_' + new Date().toISOString().slice(0, 10) + '.csv';
        
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }
    
    // 4. Conectar el evento del botón de Exportar
    exportButton.addEventListener('click', exportarTablaAExcel);
});