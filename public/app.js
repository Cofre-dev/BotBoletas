// ─── Socket ────────────────────────────────────────────────────
const socket = io();

// ─── DOM refs ─────────────────────────────────────────────────
const dropZone          = document.getElementById('dropZone');
const fileInput         = document.getElementById('fileInput');
const fileSelected      = document.getElementById('fileSelected');
const fileName          = document.getElementById('fileName');
const companiesList     = document.getElementById('companiesList');
const companiesCount    = document.getElementById('companiesCount');
const btnStart          = document.getElementById('btnStart');
const btnStop           = document.getElementById('btnStop');
const btnClearConsole   = document.getElementById('btnClearConsole');
const consoleEl         = document.getElementById('console');
const resultsCard       = document.getElementById('resultsCard');
const resultsTitle      = document.getElementById('resultsTitle');
const resultsContent    = document.getElementById('resultsContent');
const btnCloseResults   = document.getElementById('btnCloseResults');
const statsRow          = document.getElementById('statsRow');
const statTotal         = document.getElementById('statTotal');
const statCompleted     = document.getElementById('statCompleted');
const statErrors        = document.getElementById('statErrors');
const statPending       = document.getElementById('statPending');
const progressBarContainer = document.getElementById('progressBarContainer');
const progressBarFill      = document.getElementById('progressBarFill');
const currentCard       = document.getElementById('currentCard');
const currentCompanyEl  = document.getElementById('currentCompany');
const currentProgressEl = document.getElementById('currentProgress');
const statusIndicator   = document.getElementById('statusIndicator');
const statusText        = document.getElementById('statusText');

// ─── State ────────────────────────────────────────────────────
let empresas = [];

// ─── Init ─────────────────────────────────────────────────────
addLog('info', 'Bienvenido al Bot de Boletas SII');
setupRipples();
setupCardTilt();
setupMouseTrail();

// ─── ════════════════════════════════════════════════════════ ──
//     ANIMATION SYSTEMS
// ─── ════════════════════════════════════════════════════════ ──

// Count-up animation for stat numbers
function animateCount(el, to, duration = 550) {
    const from = parseInt(el.textContent) || 0;
    if (from === to) return;

    const start = performance.now();

    function easeOutCubic(t) {
        return 1 - Math.pow(1 - t, 3);
    }

    function step(now) {
        const elapsed = now - start;
        const t = Math.min(elapsed / duration, 1);
        el.textContent = Math.round(from + (to - from) * easeOutCubic(t));
        if (t < 1) requestAnimationFrame(step);
        else el.textContent = to;
    }

    requestAnimationFrame(step);
}

// Bump a stat element (spring scale animation)
function bumpStat(el) {
    el.classList.remove('stat-bump');
    void el.offsetWidth; // force reflow
    el.classList.add('stat-bump');
}

// Button ripple wave on click
function setupRipples() {
    document.querySelectorAll('.btn-primary, .btn-danger').forEach(btn => {
        btn.addEventListener('click', function (e) {
            const rect = this.getBoundingClientRect();
            const wave = document.createElement('span');
            wave.className = 'ripple-wave';
            wave.style.left = (e.clientX - rect.left) + 'px';
            wave.style.top  = (e.clientY - rect.top)  + 'px';
            this.appendChild(wave);
            wave.addEventListener('animationend', () => wave.remove());
        });
    });
}

// Subtle 3D tilt on stat & upload cards
function setupCardTilt() {
    const tiltTargets = document.querySelectorAll('.stat-card');

    tiltTargets.forEach(card => {
        card.addEventListener('mousemove', (e) => {
            const rect = card.getBoundingClientRect();
            const x = (e.clientX - rect.left) / rect.width  - 0.5;
            const y = (e.clientY - rect.top)  / rect.height - 0.5;
            card.style.transform = `perspective(480px) rotateX(${-y * 7}deg) rotateY(${x * 7}deg) scale(1.03) translateY(-2px)`;
        });

        card.addEventListener('mouseleave', () => {
            card.style.transform = '';
        });
    });
}

// Mouse trail — subtle glowing dots following cursor
function setupMouseTrail() {
    let lastTrail = 0;

    document.addEventListener('mousemove', (e) => {
        const now = Date.now();
        if (now - lastTrail < 45) return; // ~22fps cap
        lastTrail = now;

        // Only trail inside main content
        const main = document.querySelector('.main-content');
        if (!main) return;
        const rect = main.getBoundingClientRect();
        if (e.clientX < rect.left || e.clientX > rect.right ||
            e.clientY < rect.top  || e.clientY > rect.bottom) return;

        const dot = document.createElement('div');
        dot.className = 'mouse-trail-dot';
        dot.style.left = e.clientX + 'px';
        dot.style.top  = e.clientY + 'px';
        document.body.appendChild(dot);
        dot.addEventListener('animationend', () => dot.remove());
    });
}

// Canvas particle burst on process completion
function triggerParticles() {
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:9999;';
    canvas.width  = window.innerWidth;
    canvas.height = window.innerHeight;
    document.body.appendChild(canvas);

    const ctx    = canvas.getContext('2d');
    const colors = ['#3b82f6', '#10b981', '#f59e0b', '#a78bfa', '#34d399', '#60a5fa', '#f472b6'];
    const cx     = canvas.width  * 0.5;
    const cy     = canvas.height * 0.38;
    const particles = [];

    for (let i = 0; i < 110; i++) {
        const angle = (Math.PI * 2 * i) / 110;
        const speed = Math.random() * 9 + 2;
        particles.push({
            x:        cx,
            y:        cy,
            vx:       Math.cos(angle) * speed * (0.4 + Math.random() * 0.9),
            vy:       Math.sin(angle) * speed * (0.4 + Math.random() * 0.9) - 2,
            r:        Math.random() * 4.5 + 1.5,
            color:    colors[Math.floor(Math.random() * colors.length)],
            life:     1,
            decay:    Math.random() * 0.016 + 0.009,
            gravity:  0.18,
            rotation: Math.random() * Math.PI * 2,
            rotSpeed: (Math.random() - 0.5) * 0.18,
            shape:    Math.random() > 0.45 ? 'circle' : 'rect',
        });
    }

    function draw() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);

        let alive = 0;
        particles.forEach(p => {
            if (p.life <= 0) return;
            p.x  += p.vx;
            p.y  += p.vy;
            p.vy += p.gravity;
            p.vx *= 0.99;
            p.life -= p.decay;
            p.rotation += p.rotSpeed;
            alive++;

            ctx.save();
            ctx.globalAlpha = Math.max(0, p.life);
            ctx.fillStyle   = p.color;
            ctx.translate(p.x, p.y);
            ctx.rotate(p.rotation);

            if (p.shape === 'circle') {
                ctx.beginPath();
                ctx.arc(0, 0, p.r, 0, Math.PI * 2);
                ctx.fill();
            } else {
                ctx.fillRect(-p.r, -p.r * 0.5, p.r * 2, p.r);
            }

            ctx.restore();
        });

        if (alive > 0) requestAnimationFrame(draw);
        else canvas.remove();
    }

    draw();
}

// ─── ════════════════════════════════════════════════════════ ──
//     DRAG & DROP
// ─── ════════════════════════════════════════════════════════ ──

dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

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

// ─── Upload ───────────────────────────────────────────────────
async function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        addLog('error', 'Solo se permiten archivos Excel (.xlsx, .xls)');
        return;
    }

    const formData = new FormData();
    formData.append('excelFile', file);

    try {
        addLog('info', 'Subiendo archivo...');
        const response = await fetch('/upload', { method: 'POST', body: formData });
        const data = await response.json();

        if (data.success) {
            empresas = data.empresas;
            displayEmpresas();

            dropZone.querySelector('.drop-zone-content').style.display = 'none';
            fileSelected.style.display = 'flex';
            fileName.textContent = file.name;

            addLog('success', `Archivo cargado — ${data.totalEmpresas} empresas encontradas`);
            btnStart.disabled = false;
        } else {
            addLog('error', data.error);
        }
    } catch (error) {
        addLog('error', `Error al subir archivo: ${error.message}`);
    }
}

// ─── Display companies ────────────────────────────────────────
function displayEmpresas() {
    companiesCount.textContent = empresas.length;

    if (empresas.length === 0) {
        companiesList.innerHTML = `
            <div class="empty-state">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none">
                    <path d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
                </svg>
                <p>Sin empresas cargadas</p>
            </div>`;
        return;
    }

    companiesList.innerHTML = empresas.map((emp, idx) => `
        <div class="company-item ${emp.estado}" data-id="${emp.id}" style="--i:${idx}">
            <div>
                <div class="company-name">${emp.nombre}</div>
                <div class="company-rut">${emp.rut}</div>
            </div>
            <span class="status-badge ${emp.estado}">${getStatusText(emp.estado)}</span>
        </div>
    `).join('');

    document.querySelectorAll('.company-item').forEach(item => {
        item.addEventListener('click', () => viewResults(item.dataset.id));
    });
}

function getStatusText(estado) {
    const textos = { pendiente: 'Pendiente', procesando: 'Procesando', completado: 'Listo', error: 'Error' };
    return textos[estado] || estado;
}

// ─── Stats & progress ─────────────────────────────────────────
const prevStatValues = { total: 0, completed: 0, errors: 0, pending: 0 };

function updateStats() {
    const total      = empresas.length;
    const completed  = empresas.filter(e => e.estado === 'completado').length;
    const errors     = empresas.filter(e => e.estado === 'error').length;
    const processing = empresas.filter(e => e.estado === 'procesando').length;
    const pending    = empresas.filter(e => e.estado === 'pendiente').length + processing;
    const done       = completed + errors;

    // Count-up + bump for changed values
    if (total !== prevStatValues.total)         { animateCount(statTotal,     total);     bumpStat(statTotal); }
    if (completed !== prevStatValues.completed) { animateCount(statCompleted, completed); bumpStat(statCompleted); }
    if (errors !== prevStatValues.errors)       { animateCount(statErrors,    errors);    bumpStat(statErrors); }
    if (pending !== prevStatValues.pending)     { animateCount(statPending,   pending);   bumpStat(statPending); }

    prevStatValues.total     = total;
    prevStatValues.completed = completed;
    prevStatValues.errors    = errors;
    prevStatValues.pending   = pending;

    // Progress bar
    if (total > 0) {
        progressBarFill.style.width = Math.round((done / total) * 100) + '%';
    }

    // Current company indicator
    const procesando = empresas.find(e => e.estado === 'procesando');
    if (procesando) {
        currentCard.style.display = 'block';
        currentCompanyEl.textContent  = procesando.nombre;
        currentProgressEl.textContent = `${done + 1} / ${total}`;
    }
}

// ─── View results ─────────────────────────────────────────────
async function viewResults(empresaId) {
    try {
        const response = await fetch(`/resultados/${empresaId}`);
        const data = await response.json();

        if (data.success && data.data) {
            const empresa = empresas.find(e => e.id === parseInt(empresaId));
            resultsTitle.textContent = empresa?.nombre || 'Empresa';
            resultsContent.innerHTML = renderResultados(data.data);
            resultsCard.style.display = 'block';
        } else {
            addLog('warning', 'No hay datos disponibles para esta empresa');
        }
    } catch (error) {
        addLog('error', `Error al cargar resultados: ${error.message}`);
    }
}

function renderResultados(data) {
    let html = '';

    if (data.boletasEmitidas) {
        html += `<h3 style="margin-bottom:0.5rem;color:#93c5fd;font-size:0.85rem;">Boletas de Honorarios Emitidas 2025</h3>`;
        html += renderTabla(data.boletasEmitidas, true);
        if (data.boletasEmitidas.totales)
            html += `<div class="summary-box">Honorarios Brutos Emitidos: $${(data.boletasEmitidas.totales.honorarioBruto || 0).toLocaleString('es-CL')}</div>`;
    }

    if (data.boletasRecibidas) {
        html += `<h3 style="margin:1rem 0 0.5rem;color:#93c5fd;font-size:0.85rem;">Boletas de Honorarios Recibidas 2025</h3>`;
        html += renderTabla(data.boletasRecibidas, false);
        if (data.boletasRecibidas.totales)
            html += `<div class="summary-box">Honorarios Brutos Recibidos: $${(data.boletasRecibidas.totales.honorarioBruto || 0).toLocaleString('es-CL')}</div>`;
    }

    if (data.bteRecibidas) {
        html += `<h3 style="margin:1rem 0 0.5rem;color:#6ee7b7;font-size:0.85rem;">BTE Recibidas 2025</h3>`;
        html += renderTablaBTE(data.bteRecibidas);
        if (data.bteRecibidas.totales)
            html += `<div class="summary-box" style="background:rgba(16,185,129,0.08);border-color:rgba(16,185,129,0.2);color:#6ee7b7;">Monto Total BTE Recibidas: $${(data.bteRecibidas.totales.montoTotal || 0).toLocaleString('es-CL')}</div>`;
    }

    if (data.bteEmitidas) {
        html += `<h3 style="margin:1rem 0 0.5rem;color:#6ee7b7;font-size:0.85rem;">BTE Emitidas 2025</h3>`;
        html += renderTablaBTE(data.bteEmitidas);
        if (data.bteEmitidas.totales)
            html += `<div class="summary-box" style="background:rgba(16,185,129,0.08);border-color:rgba(16,185,129,0.2);color:#6ee7b7;">Monto Total BTE Emitidas: $${(data.bteEmitidas.totales.montoTotal || 0).toLocaleString('es-CL')}</div>`;
    }

    return html || '<p style="color:#4b5e78;font-size:0.85rem;">No hay datos disponibles</p>';
}

function renderTabla(data, incluirFolios) {
    if (!data.meses || data.meses.length === 0)
        return '<p style="color:#4b5e78;font-size:0.82rem;padding:0.5rem 0;">Sin datos para este período</p>';

    let headers = ['Período'];
    if (incluirFolios) headers.push('F.Inicial', 'F.Final');
    headers.push('Vigentes', 'Anuladas', 'Hon. Bruto', 'Ret. Terc.', 'Ret. Cont.', 'Líquido');

    let html = '<table class="results-table"><thead><tr>';
    headers.forEach(h => html += `<th>${h}</th>`);
    html += '</tr></thead><tbody>';

    data.meses.forEach(m => {
        html += '<tr>';
        html += `<td>${m.periodo}</td>`;
        if (incluirFolios) { html += `<td>${m.folioInicial || '-'}</td><td>${m.folioFinal || '-'}</td>`; }
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
    if (!data.meses || data.meses.length === 0)
        return '<p style="color:#4b5e78;font-size:0.82rem;padding:0.5rem 0;">Sin datos para este período</p>';

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

// ─── Button handlers ──────────────────────────────────────────
btnStart.addEventListener('click', () => {
    socket.emit('startProcess');
    btnStart.disabled = true;
    btnStop.disabled  = false;

    statsRow.style.display = 'grid';
    progressBarContainer.classList.add('visible');

    statusIndicator.className = 'status-indicator running';
    statusText.textContent    = 'Ejecutando';
    addLog('info', 'Iniciando proceso de extracción...');
});

btnStop.addEventListener('click', async () => {
    await fetch('/stop', { method: 'POST' });
    btnStop.disabled  = true;
    btnStart.disabled = false;
    currentCard.style.display = 'none';
    statusIndicator.className = 'status-indicator stopped';
    statusText.textContent    = 'Detenido';
    addLog('warning', 'Proceso detenido manualmente');
});

btnClearConsole.addEventListener('click', () => {
    consoleEl.innerHTML = '';
    addLog('info', 'Consola limpiada');
});

btnCloseResults.addEventListener('click', () => {
    resultsCard.style.display = 'none';
});

// ─── Socket events ────────────────────────────────────────────
socket.on('log', (data) => addLog(data.type, data.message));

socket.on('empresaUpdate', (data) => {
    const empresa = empresas.find(e => e.id === data.id);
    if (!empresa) return;

    const prev = empresa.estado;
    empresa.estado = data.estado;
    displayEmpresas();
    updateStats();

    // Flash green on completion
    if (data.estado === 'completado' && prev === 'procesando') {
        setTimeout(() => {
            const el = companiesList.querySelector(`[data-id="${data.id}"]`);
            if (el) el.classList.add('just-completed');
        }, 50);
    }

    // Auto-scroll to processing item
    if (data.estado === 'procesando') {
        setTimeout(() => {
            const el = companiesList.querySelector(`[data-id="${data.id}"]`);
            if (el) el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 80);
    }
});

socket.on('processComplete', () => {
    btnStart.disabled = false;
    btnStop.disabled  = true;
    currentCard.style.display = 'none';

    statusIndicator.className = 'status-indicator done';
    statusText.textContent    = 'Completado';

    progressBarFill.style.width = '100%';
    updateStats();

    // 🎉 Particle burst
    triggerParticles();

    addLog('success', 'Proceso completado — descargando Excel...');

    // Small delay so particles are visible before navigation
    setTimeout(() => { window.location.href = '/exportar'; }, 1200);
});

// ─── Logger ───────────────────────────────────────────────────
function addLog(type, message) {
    const time   = new Date().toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const badges = { info: 'INF', success: 'OK', warning: 'WARN', error: 'ERR' };
    const line   = document.createElement('div');
    line.className   = `console-line ${type}`;
    line.innerHTML   = `<span class="console-time">${time}</span><span class="console-badge ${type}">${badges[type] || 'LOG'}</span><span class="message">${message}</span>`;
    consoleEl.appendChild(line);
    consoleEl.scrollTop = consoleEl.scrollHeight;
}
