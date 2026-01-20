// Conexión Socket.io
const socket = io();

// Elementos del DOM
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileSelected = document.getElementById('fileSelected');
const fileName = document.getElementById('fileName');
const companiesList = document.getElementById('companiesList');
const companiesCount = document.getElementById('companiesCount');
const btnStart = document.getElementById('btnStart');
const btnStop = document.getElementById('btnStop');
const btnExport = document.getElementById('btnExport');
const btnClearConsole = document.getElementById('btnClearConsole');
const consoleEl = document.getElementById('console');
const resultsCard = document.getElementById('resultsCard');
const resultsTitle = document.getElementById('resultsTitle');
const resultsContent = document.getElementById('resultsContent');
const btnCloseResults = document.getElementById('btnCloseResults');

let empresas = [];

// Drag and Drop
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
});

async function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        addLog('error', '❌ Solo se permiten archivos Excel (.xlsx, .xls)');
        return;
    }

    const formData = new FormData();
    formData.append('excelFile', file);

    try {
        addLog('info', '📤 Subiendo archivo...');
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (data.success) {
            empresas = data.empresas;
            displayEmpresas();

            // Mostrar archivo seleccionado
            dropZone.querySelector('.drop-zone-content').style.display = 'none';
            fileSelected.style.display = 'flex';
            fileName.textContent = file.name;

            addLog('success', `✅ Archivo cargado correctamente`);
            addLog('info', `📊 Se encontraron ${data.totalEmpresas} empresas`);

            btnStart.disabled = false;
        } else {
            addLog('error', `❌ ${data.error}`);
        }
    } catch (error) {
        addLog('error', `❌ Error al subir archivo: ${error.message}`);
    }
}

function displayEmpresas() {
    companiesCount.textContent = `${empresas.length} empresas`;

    if (empresas.length === 0) {
        companiesList.innerHTML = `
            <div class="empty-state">
                <span>👥</span>
                <p>No hay empresas cargadas</p>
            </div>
        `;
        return;
    }

    companiesList.innerHTML = empresas.map(emp => `
        <div class="company-item ${emp.estado}" data-id="${emp.id}">
            <div>
                <div class="company-name">${emp.nombre}</div>
                <div class="company-rut">${emp.rut}</div>
            </div>
            <span class="status-badge ${emp.estado}">${getStatusText(emp.estado)}</span>
        </div>
    `).join('');

    // Agregar click handlers para ver resultados
    document.querySelectorAll('.company-item').forEach(item => {
        item.addEventListener('click', () => viewResults(item.dataset.id));
    });
}

function getStatusText(estado) {
    const textos = {
        'pendiente': '⏳ Pendiente',
        'procesando': '🔄 Procesando',
        'completado': '✅ Listo',
        'error': '❌ Error'
    };
    return textos[estado] || estado;
}

async function viewResults(empresaId) {
    try {
        const response = await fetch(`/resultados/${empresaId}`);
        const data = await response.json();

        if (data.success && data.data) {
            const empresa = empresas.find(e => e.id === parseInt(empresaId));
            resultsTitle.textContent = `Resultados: ${empresa?.nombre || 'Empresa'}`;
            resultsContent.innerHTML = renderResultados(data.data);
            resultsCard.style.display = 'block';
        } else {
            addLog('warning', '⚠️ No hay datos disponibles para esta empresa');
        }
    } catch (error) {
        addLog('error', `❌ Error al cargar resultados: ${error.message}`);
    }
}

function renderResultados(data) {
    let html = '';

    // Boletas Emitidas
    if (data.boletasEmitidas) {
        html += `<h3 style="margin-bottom: 0.5rem; color: #3182ce;">📤 Boletas de Honorarios Emitidas 2025</h3>`;
        html += renderTabla(data.boletasEmitidas, true);
        if (data.boletasEmitidas.totales) {
            html += `<div class="summary-box">💰 Honorarios Brutos Emitidos: $${(data.boletasEmitidas.totales.honorarioBruto || 0).toLocaleString('es-CL')}</div>`;
        }
    }

    // Boletas Recibidas
    if (data.boletasRecibidas) {
        html += `<h3 style="margin: 1rem 0 0.5rem; color: #3182ce;">📥 Boletas de Honorarios Recibidas 2025</h3>`;
        html += renderTabla(data.boletasRecibidas, false);
        if (data.boletasRecibidas.totales) {
            html += `<div class="summary-box">💰 Honorarios Brutos Recibidos: $${(data.boletasRecibidas.totales.honorarioBruto || 0).toLocaleString('es-CL')}</div>`;
        }
    }

    // BTE Recibidas
    if (data.bteRecibidas) {
        html += `<h3 style="margin: 1rem 0 0.5rem; color: #38a169;">📥 BTE Recibidas 2025 (Prestación de Servicios de Terceros)</h3>`;
        html += renderTablaBTE(data.bteRecibidas);
        if (data.bteRecibidas.totales) {
            html += `<div class="summary-box" style="background: rgba(56, 161, 105, 0.1); border-color: rgba(56, 161, 105, 0.3);">💰 Monto Total BTE Recibidas: $${(data.bteRecibidas.totales.montoTotal || 0).toLocaleString('es-CL')}</div>`;
        }
    }

    // BTE Emitidas
    if (data.bteEmitidas) {
        html += `<h3 style="margin: 1rem 0 0.5rem; color: #38a169;">📤 BTE Emitidas 2025 (Prestación de Servicios de Terceros)</h3>`;
        html += renderTablaBTE(data.bteEmitidas);
        if (data.bteEmitidas.totales) {
            html += `<div class="summary-box" style="background: rgba(56, 161, 105, 0.1); border-color: rgba(56, 161, 105, 0.3);">💰 Monto Total BTE Emitidas: $${(data.bteEmitidas.totales.montoTotal || 0).toLocaleString('es-CL')}</div>`;
        }
    }

    return html || '<p>No hay datos disponibles</p>';
}

function renderTabla(data, incluirFolios) {
    if (!data.meses || data.meses.length === 0) {
        return '<p style="color: #94a3b8;">Sin datos para este período</p>';
    }

    let headers = ['Período'];
    if (incluirFolios) headers.push('F.Inicial', 'F.Final');
    headers.push('Vigentes', 'Anuladas', 'Hon. Bruto', 'Ret. Terc.', 'Ret. Cont.', 'Líquido');

    let html = '<table class="results-table"><thead><tr>';
    headers.forEach(h => html += `<th>${h}</th>`);
    html += '</tr></thead><tbody>';

    data.meses.forEach(m => {
        html += '<tr>';
        html += `<td>${m.periodo}</td>`;
        if (incluirFolios) {
            html += `<td>${m.folioInicial || '-'}</td>`;
            html += `<td>${m.folioFinal || '-'}</td>`;
        }
        html += `<td>${m.vigentes || 0}</td>`;
        html += `<td>${m.anuladas || 0}</td>`;
        html += `<td>$${(m.honorarioBruto || 0).toLocaleString('es-CL')}</td>`;
        html += `<td>$${(m.retencionTerceros || 0).toLocaleString('es-CL')}</td>`;
        html += `<td>$${(m.retencionContribuyente || 0).toLocaleString('es-CL')}</td>`;
        html += `<td>$${(m.totalLiquido || 0).toLocaleString('es-CL')}</td>`;
        html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
}

function renderTablaBTE(data) {
    if (!data.meses || data.meses.length === 0) {
        return '<p style="color: #94a3b8;">Sin datos para este período</p>';
    }

    const headers = ['Período', 'Cantidad', 'Monto Neto', 'Monto Exento', 'IVA', 'Monto Total'];

    let html = '<table class="results-table"><thead><tr>';
    headers.forEach(h => html += `<th>${h}</th>`);
    html += '</tr></thead><tbody>';

    data.meses.forEach(m => {
        html += '<tr>';
        html += `<td>${m.periodo}</td>`;
        html += `<td>${m.cantidad || 0}</td>`;
        html += `<td>$${(m.montoNeto || 0).toLocaleString('es-CL')}</td>`;
        html += `<td>$${(m.montoExento || 0).toLocaleString('es-CL')}</td>`;
        html += `<td>$${(m.montoIva || 0).toLocaleString('es-CL')}</td>`;
        html += `<td>$${(m.montoTotal || 0).toLocaleString('es-CL')}</td>`;
        html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
}

// Botones de acción
btnStart.addEventListener('click', () => {
    socket.emit('startProcess');
    btnStart.disabled = true;
    btnStop.disabled = false;
});

btnStop.addEventListener('click', async () => {
    await fetch('/stop', { method: 'POST' });
    btnStop.disabled = true;
    btnStart.disabled = false;
});

btnClearConsole.addEventListener('click', () => {
    consoleEl.innerHTML = '';
    addLog('info', '🧹 Consola limpiada');
});

btnCloseResults.addEventListener('click', () => {
    resultsCard.style.display = 'none';
});

btnExport.addEventListener('click', () => {
    window.location.href = '/exportar';
});

// Socket events
socket.on('log', (data) => {
    addLog(data.type, data.message);
});

socket.on('empresaUpdate', (data) => {
    const empresa = empresas.find(e => e.id === data.id);
    if (empresa) {
        empresa.estado = data.estado;
        displayEmpresas();
    }
});

socket.on('processComplete', () => {
    btnStart.disabled = false;
    btnStop.disabled = true;
    btnExport.disabled = false;
    addLog('success', '🎉 ¡Proceso completado! Puedes exportar los resultados.');
});

function addLog(type, message) {
    const time = new Date().toLocaleTimeString();
    const line = document.createElement('div');
    line.className = `console-line ${type}`;
    line.innerHTML = `<span class="timestamp">[${time}]</span><span class="message">${message}</span>`;
    consoleEl.appendChild(line);
    consoleEl.scrollTop = consoleEl.scrollHeight;
}
