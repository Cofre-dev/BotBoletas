const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const multer = require('multer');
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs');
const SIIBot = require('./bot/siiBot');

const app = express();
const server = http.createServer(app);
const io = new Server(server);

// Configuración de multer para subir archivos
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = path.join(__dirname, 'uploads');
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir, { recursive: true });
        }
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        cb(null, 'empresas_' + Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (ext !== '.xlsx' && ext !== '.xls') {
            return cb(new Error('Solo se permiten archivos Excel (.xlsx, .xls)'));
        }
        cb(null, true);
    }
});

// Servir archivos estáticos
app.use(express.static('public'));
app.use(express.json());

// Variable global para almacenar datos de empresas y resultados
let empresasData = [];
let resultadosEmpresas = {};
let botInstance = null;
let isProcessing = false;

// Subir archivo Excel
app.post('/upload', upload.single('excelFile'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No se subió ningún archivo' });
        }

        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Procesar datos (Columna A: empresa, B: RUT, C: clave)
        empresasData = [];
        for (let i = 1; i < data.length; i++) { // Empezar desde 1 para saltar encabezados
            const row = data[i];
            if (row[0] && row[1] && row[2]) {
                empresasData.push({
                    id: i,
                    nombre: String(row[0]).trim(),
                    rut: String(row[1]).trim(),
                    clave: String(row[2]).trim(),
                    estado: 'pendiente'
                });
            }
        }

        // Limpiar resultados anteriores
        resultadosEmpresas = {};

        res.json({
            success: true,
            totalEmpresas: empresasData.length,
            empresas: empresasData.map(e => ({ id: e.id, nombre: e.nombre, rut: e.rut, estado: e.estado }))
        });

    } catch (error) {
        console.error('Error al procesar archivo:', error);
        res.status(500).json({ error: 'Error al procesar el archivo Excel' });
    }
});

// Obtener lista de empresas
app.get('/empresas', (req, res) => {
    res.json({
        empresas: empresasData.map(e => ({
            id: e.id,
            nombre: e.nombre,
            rut: e.rut,
            estado: e.estado
        })),
        totalEmpresas: empresasData.length
    });
});

// Obtener resultados de una empresa específica
app.get('/resultados/:empresaId', (req, res) => {
    const empresaId = parseInt(req.params.empresaId);
    const resultado = resultadosEmpresas[empresaId];

    if (resultado) {
        res.json({ success: true, data: resultado });
    } else {
        res.json({ success: false, message: 'No hay datos disponibles para esta empresa' });
    }
});

// Obtener estado del proceso
app.get('/status', (req, res) => {
    res.json({
        isProcessing,
        empresas: empresasData.map(e => ({ id: e.id, nombre: e.nombre, estado: e.estado }))
    });
});

// Detener el proceso
app.post('/stop', async (req, res) => {
    if (botInstance) {
        await botInstance.stop();
        isProcessing = false;
        io.emit('log', { type: 'warning', message: '⚠️ Proceso detenido por el usuario' });
    }
    res.json({ success: true });
});

// Exportar a Excel
app.get('/exportar', async (req, res) => {
    try {
        const ExcelJS = require('exceljs');
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Bot Boletas SII';
        workbook.created = new Date();

        // ─── Paleta de colores ────────────────────────────────────
        const C = {
            // Encabezado empresa
            coHeader:     'FF0D1B2A',
            coHeaderSub:  'FF1A2F44',
            coHeaderText: 'FFFFFFFF',
            coHeaderMuted:'FF6B9AB8',

            // Boletas (azul)
            blueSection:  'FF1D3461',
            blueHeader:   'FF2E5EAA',
            blueEvenRow:  'FFEEF4FB',
            blueLight:    'FFD6E4F5',

            // BTE (verde)
            greenSection: 'FF1A4731',
            greenHeader:  'FF2D7D55',
            greenEvenRow: 'FFE8F5EE',
            greenLight:   'FFBBDECE',

            // Total (ámbar)
            totalBg:      'FFCA8A04',
            totalBorder:  'FF9A6403',

            // Resumen mensaje
            msgBg:        'FFFFF8E8',
            msgBorder:    'FFCA8A04',
            msgText:      'FF7B4000',

            // Resumen hoja
            resumeHdr:    'FF0D1B2A',
            resumeColHdr: 'FF1D3461',
            resumeEven:   'FFF0F5FF',
            resumeTotal:  'FF1D3461',
        };

        const MONEY_FMT = '#,##0';
        const INT_FMT   = '#,##0';

        // ─── Helper: estilizar celda ─────────────────────────────
        function sc(cell, {
            bg, fg = 'FF1A1A1A', bold = false, italic = false,
            size = 9.5, center = false, right = false,
            indent = 0, numFmt = null, wrap = false, border = null
        }) {
            if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.font = { name: 'Calibri', bold, italic, size, color: { argb: fg } };
            cell.alignment = {
                vertical: 'middle',
                horizontal: center ? 'center' : right ? 'right' : 'left',
                indent,
                wrapText: wrap,
            };
            if (numFmt) cell.numFmt = numFmt;
            if (border) cell.border = border;
        }

        const THIN = {
            top:    { style: 'thin',   color: { argb: 'FFD0DCE8' } },
            bottom: { style: 'thin',   color: { argb: 'FFD0DCE8' } },
            left:   { style: 'thin',   color: { argb: 'FFD0DCE8' } },
            right:  { style: 'thin',   color: { argb: 'FFD0DCE8' } },
        };

        const THIN_GREEN = {
            top:    { style: 'thin',   color: { argb: 'FFBBD9C8' } },
            bottom: { style: 'thin',   color: { argb: 'FFBBD9C8' } },
            left:   { style: 'thin',   color: { argb: 'FFBBD9C8' } },
            right:  { style: 'thin',   color: { argb: 'FFBBD9C8' } },
        };

        // ─── Helper: escribir fila de sección ────────────────────
        function writeSectionTitle(sheet, rowN, label, nCols, bgColor) {
            sheet.mergeCells(rowN, 1, rowN, nCols);
            const cell = sheet.getCell(rowN, 1);
            cell.value = label;
            sc(cell, { bg: bgColor, fg: C.coHeaderText, bold: true, size: 11, indent: 1 });
            sheet.getRow(rowN).height = 22;
        }

        // Helper: encabezados de columna
        function writeColHeaders(sheet, rowN, headers, bgColor, borderColor) {
            const r = sheet.getRow(rowN);
            r.height = 20;
            headers.forEach((h, i) => {
                const cell = r.getCell(i + 1);
                cell.value = h;
                sc(cell, { bg: bgColor, fg: 'FFFFFFFF', bold: true, size: 9, center: true });
                cell.border = {
                    top:    { style: 'thin',   color: { argb: borderColor } },
                    bottom: { style: 'medium', color: { argb: borderColor } },
                    left:   { style: 'thin',   color: { argb: 'FFFFFFFF' } },
                    right:  { style: 'thin',   color: { argb: 'FFFFFFFF' } },
                };
            });
        }

        // Helper: fila de datos con filas alternadas
        function writeDataRow(sheet, rowN, values, evenRowBg, monetaryCols, borderStyle) {
            const isEven = rowN % 2 === 0;
            const bg = isEven ? evenRowBg : 'FFFFFFFF';
            const r  = sheet.getRow(rowN);
            r.height = 17;
            values.forEach((val, i) => {
                const col  = i + 1;
                const cell = r.getCell(col);
                cell.value = val;
                const isMoney = monetaryCols.includes(col);
                sc(cell, {
                    bg,
                    size: 9,
                    center: col === 1,
                    right: isMoney,
                    numFmt: isMoney ? MONEY_FMT : null,
                });
                cell.border = borderStyle;
            });
        }

        // Helper: fila de totales
        function writeTotalRow(sheet, rowN, values, monetaryCols) {
            const r = sheet.getRow(rowN);
            r.height = 19;
            values.forEach((val, i) => {
                const col  = i + 1;
                const cell = r.getCell(col);
                cell.value = val;
                const isMoney = monetaryCols.includes(col);
                sc(cell, {
                    bg: C.totalBg, fg: 'FFFFFFFF', bold: true, size: 9,
                    center: col === 1,
                    right: isMoney,
                    numFmt: isMoney ? MONEY_FMT : null,
                });
                cell.border = {
                    top:    { style: 'medium', color: { argb: C.totalBorder } },
                    bottom: { style: 'medium', color: { argb: C.totalBorder } },
                    left:   { style: 'thin',   color: { argb: C.totalBg } },
                    right:  { style: 'thin',   color: { argb: C.totalBg } },
                };
            });
        }

        // Helper: caja resumen/mensaje
        function writeSummaryBox(sheet, rowN, label, nCols) {
            sheet.mergeCells(rowN, 1, rowN, nCols);
            const cell = sheet.getCell(rowN, 1);
            cell.value = label;
            sc(cell, {
                bg: C.msgBg, fg: C.msgText, bold: true, italic: true,
                size: 10, center: true,
            });
            cell.border = {
                top:    { style: 'medium', color: { argb: C.msgBorder } },
                bottom: { style: 'medium', color: { argb: C.msgBorder } },
                left:   { style: 'medium', color: { argb: C.msgBorder } },
                right:  { style: 'medium', color: { argb: C.msgBorder } },
            };
            sheet.getRow(rowN).height = 20;
        }

        // ─── Hoja RESUMEN ─────────────────────────────────────────
        const resSheet = workbook.addWorksheet('Resumen');
        resSheet.properties.tabColor = { argb: C.resumeHdr };

        resSheet.columns = [
            { width: 36 },  // A Empresa
            { width: 16 },  // B RUT
            { width: 22 },  // C Hon. Emit.
            { width: 22 },  // D Hon. Rec.
            { width: 22 },  // E BTE Rec.
            { width: 22 },  // F BTE Emit.
            { width: 11 },  // G Estado
        ];

        // Título
        resSheet.mergeCells('A1:G1');
        sc(resSheet.getCell('A1'), {
            bg: C.resumeHdr, fg: 'FFFFFFFF', bold: true, size: 12, center: true,
        });
        resSheet.getCell('A1').value = 'RESUMEN GENERAL — BOLETAS DE HONORARIOS Y BTE — 2025';
        resSheet.getRow(1).height = 27;

        // Fecha generación
        resSheet.mergeCells('A2:G2');
        sc(resSheet.getCell('A2'), { bg: C.coHeaderSub, fg: C.coHeaderMuted, size: 9, center: true });
        resSheet.getCell('A2').value =
            `Generado: ${new Date().toLocaleDateString('es-CL', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`;
        resSheet.getRow(2).height = 16;

        // Separador
        resSheet.getRow(3).height = 7;

        // Encabezados resumen
        const resHeaders = ['EMPRESA / CONTRIBUYENTE', 'RUT', 'HON. BRUTO EMITIDO', 'HON. BRUTO RECIBIDO', 'MONTO BTE RECIBIDO', 'MONTO BTE EMITIDO', 'ESTADO'];
        const resHdrRow  = resSheet.getRow(4);
        resHdrRow.height = 22;
        resHeaders.forEach((h, i) => {
            const cell = resHdrRow.getCell(i + 1);
            cell.value = h;
            sc(cell, { bg: C.resumeColHdr, fg: 'FFFFFFFF', bold: true, size: 9, center: i > 0 });
            cell.border = THIN;
        });
        resSheet.views = [{ state: 'frozen', ySplit: 4 }];

        let resRow = 5;
        const resTotals = { emit: 0, rec: 0, bteRec: 0, bteEmit: 0 };

        // ─── Hojas por empresa ────────────────────────────────────
        const TAB_COLORS = ['FF2558A6', 'FF1B7A4A', 'FF7057B8', 'FF8B3A3A', 'FF1A7070'];
        let sheetIdx = 0;

        for (const empresaId in resultadosEmpresas) {
            const resultado = resultadosEmpresas[empresaId];
            const empresa   = empresasData.find(e => e.id === parseInt(empresaId));
            if (!empresa || !resultado) continue;

            const sheetName = empresa.nombre.substring(0, 31).replace(/[*?:/\\[\]]/g, '').trim();
            const sheet     = workbook.addWorksheet(sheetName);
            sheet.properties.tabColor = { argb: TAB_COLORS[sheetIdx % TAB_COLORS.length] };
            sheetIdx++;

            // Anchos de columnas (máximo 9 cols para boletas emitidas)
            sheet.columns = [
                { width: 13 },  // A Período
                { width: 13 },  // B Folio Ini / Vigentes
                { width: 13 },  // C Folio Fin / Anuladas
                { width: 11 },  // D Vigentes / Hon. Bruto
                { width: 11 },  // E Anuladas / Ret. Terc.
                { width: 20 },  // F Hon. Bruto / Ret. Contrib / Monto Total
                { width: 18 },  // G Ret. Terc. / Liq.
                { width: 22 },  // H Ret. Contrib.
                { width: 20 },  // I Total Líquido
            ];

            // ── Bloque encabezado empresa (filas 1-3) ─────────────
            sheet.mergeCells('A1:I1');
            sc(sheet.getCell('A1'), { bg: C.coHeader, fg: C.coHeaderText, bold: true, size: 13, indent: 1 });
            sheet.getCell('A1').value = resultado.contribuyente || empresa.nombre;
            sheet.getRow(1).height = 26;

            sheet.mergeCells('A2:I2');
            sc(sheet.getCell('A2'), { bg: C.coHeader, fg: C.coHeaderMuted, size: 10, indent: 1 });
            sheet.getCell('A2').value = `RUT: ${resultado.rut || empresa.rut}`;
            sheet.getRow(2).height = 19;

            sheet.mergeCells('A3:I3');
            sc(sheet.getCell('A3'), { bg: C.coHeaderSub, fg: 'FF4D7A99', size: 8.5, indent: 1 });
            sheet.getCell('A3').value =
                `Generado: ${new Date().toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' })}  ·  Bot Boletas SII`;
            sheet.getRow(3).height = 15;

            sheet.getRow(4).height = 8; // separador

            let row = 5;
            sheet.views = [{ state: 'frozen', ySplit: 4 }];

            // ── BOLETAS EMITIDAS ──────────────────────────────────
            writeSectionTitle(sheet, row, '  BOLETAS DE HONORARIOS EMITIDAS — 2025', 9, C.blueSection);
            row++;

            if (resultado.boletasEmitidas?.meses?.length) {
                writeColHeaders(sheet, row,
                    ['PERIODO', 'FOLIO INICIAL', 'FOLIO FINAL', 'VIGENTES', 'ANULADAS', 'HONORARIO BRUTO', 'RET. TERCEROS', 'RET. CONTRIBUYENTE', 'TOTAL LÍQUIDO'],
                    C.blueHeader, 'FF1A3D7A');
                row++;

                resultado.boletasEmitidas.meses.forEach(mes => {
                    writeDataRow(sheet, row,
                        [mes.periodo, mes.folioInicial || '—', mes.folioFinal || '—',
                         mes.vigentes || 0, mes.anuladas || 0,
                         mes.honorarioBruto || 0, mes.retencionTerceros || 0,
                         mes.retencionContribuyente || 0, mes.totalLiquido || 0],
                        C.blueEvenRow, [6, 7, 8, 9], THIN);
                    row++;
                });

                const te = resultado.boletasEmitidas.totales || {};
                writeTotalRow(sheet, row,
                    ['TOTALES', '', '', te.vigentes || 0, te.anuladas || 0,
                     te.honorarioBruto || 0, te.retencionTerceros || 0,
                     te.retencionContribuyente || 0, te.totalLiquido || 0],
                    [6, 7, 8, 9]);
                row += 2;

                writeSummaryBox(sheet, row,
                    `HONORARIOS BRUTOS EMITIDOS: $${(te.honorarioBruto || 0).toLocaleString('es-CL')}`, 9);
                row += 2;

                resTotals.emit += te.honorarioBruto || 0;
            } else {
                sheet.mergeCells(row, 1, row, 9);
                sc(sheet.getCell(row, 1), { fg: 'FF999999', size: 9, indent: 1 });
                sheet.getCell(row, 1).value = 'Sin datos para este período';
                sheet.getRow(row).height = 16;
                row += 2;
            }

            // ── BOLETAS RECIBIDAS ─────────────────────────────────
            writeSectionTitle(sheet, row, '  BOLETAS DE HONORARIOS RECIBIDAS — 2025', 7, C.blueSection);
            row++;

            if (resultado.boletasRecibidas?.meses?.length) {
                writeColHeaders(sheet, row,
                    ['PERIODO', 'VIGENTES', 'ANULADAS', 'HONORARIO BRUTO', 'RET. TERCEROS', 'RET. CONTRIBUYENTE', 'TOTAL LÍQUIDO'],
                    C.blueHeader, 'FF1A3D7A');
                row++;

                resultado.boletasRecibidas.meses.forEach(mes => {
                    writeDataRow(sheet, row,
                        [mes.periodo, mes.vigentes || 0, mes.anuladas || 0,
                         mes.honorarioBruto || 0, mes.retencionTerceros || 0,
                         mes.retencionContribuyente || 0, mes.totalLiquido || 0],
                        C.blueEvenRow, [4, 5, 6, 7], THIN);
                    row++;
                });

                const tr = resultado.boletasRecibidas.totales || {};
                writeTotalRow(sheet, row,
                    ['TOTALES', tr.vigentes || 0, tr.anuladas || 0,
                     tr.honorarioBruto || 0, tr.retencionTerceros || 0,
                     tr.retencionContribuyente || 0, tr.totalLiquido || 0],
                    [4, 5, 6, 7]);
                row += 2;

                writeSummaryBox(sheet, row,
                    `HONORARIOS BRUTOS RECIBIDOS: $${(tr.honorarioBruto || 0).toLocaleString('es-CL')}`, 7);
                row += 2;

                resTotals.rec += tr.honorarioBruto || 0;
            } else {
                sheet.mergeCells(row, 1, row, 7);
                sc(sheet.getCell(row, 1), { fg: 'FF999999', size: 9, indent: 1 });
                sheet.getCell(row, 1).value = 'Sin datos para este período';
                sheet.getRow(row).height = 16;
                row += 2;
            }

            // ── BTE RECIBIDAS ─────────────────────────────────────
            writeSectionTitle(sheet, row, '  BTE RECIBIDAS (PRESTACIÓN DE SERVICIOS DE TERCEROS) — 2025', 6, C.greenSection);
            row++;

            if (resultado.bteRecibidas?.meses?.length) {
                writeColHeaders(sheet, row,
                    ['PERIODO', 'CANTIDAD', 'MONTO NETO', 'MONTO EXENTO', 'IVA', 'MONTO TOTAL'],
                    C.greenHeader, 'FF1A5233');
                row++;

                resultado.bteRecibidas.meses.forEach(mes => {
                    writeDataRow(sheet, row,
                        [mes.periodo, mes.cantidad || 0, mes.montoNeto || 0,
                         mes.montoExento || 0, mes.montoIva || 0, mes.montoTotal || 0],
                        C.greenEvenRow, [3, 4, 5, 6], THIN_GREEN);
                    row++;
                });

                const tbr = resultado.bteRecibidas.totales || {};
                writeTotalRow(sheet, row,
                    ['TOTALES', tbr.cantidad || 0, tbr.montoNeto || 0,
                     tbr.montoExento || 0, tbr.montoIva || 0, tbr.montoTotal || 0],
                    [3, 4, 5, 6]);
                row += 2;

                writeSummaryBox(sheet, row,
                    `MONTO TOTAL BTE RECIBIDAS: $${(tbr.montoTotal || 0).toLocaleString('es-CL')}`, 6);
                row += 2;

                resTotals.bteRec += tbr.montoTotal || 0;
            } else {
                sheet.mergeCells(row, 1, row, 6);
                sc(sheet.getCell(row, 1), { fg: 'FF999999', size: 9, indent: 1 });
                sheet.getCell(row, 1).value = 'Sin datos para este período';
                sheet.getRow(row).height = 16;
                row += 2;
            }

            // ── BTE EMITIDAS ──────────────────────────────────────
            writeSectionTitle(sheet, row, '  BTE EMITIDAS (PRESTACIÓN DE SERVICIOS DE TERCEROS) — 2025', 6, C.greenSection);
            row++;

            if (resultado.bteEmitidas?.meses?.length) {
                writeColHeaders(sheet, row,
                    ['PERIODO', 'CANTIDAD', 'MONTO NETO', 'MONTO EXENTO', 'IVA', 'MONTO TOTAL'],
                    C.greenHeader, 'FF1A5233');
                row++;

                resultado.bteEmitidas.meses.forEach(mes => {
                    writeDataRow(sheet, row,
                        [mes.periodo, mes.cantidad || 0, mes.montoNeto || 0,
                         mes.montoExento || 0, mes.montoIva || 0, mes.montoTotal || 0],
                        C.greenEvenRow, [3, 4, 5, 6], THIN_GREEN);
                    row++;
                });

                const tbe = resultado.bteEmitidas.totales || {};
                writeTotalRow(sheet, row,
                    ['TOTALES', tbe.cantidad || 0, tbe.montoNeto || 0,
                     tbe.montoExento || 0, tbe.montoIva || 0, tbe.montoTotal || 0],
                    [3, 4, 5, 6]);
                row += 2;

                writeSummaryBox(sheet, row,
                    `MONTO TOTAL BTE EMITIDAS: $${(tbe.montoTotal || 0).toLocaleString('es-CL')}`, 6);

                resTotals.bteEmit += tbe.montoTotal || 0;
            } else {
                sheet.mergeCells(row, 1, row, 6);
                sc(sheet.getCell(row, 1), { fg: 'FF999999', size: 9, indent: 1 });
                sheet.getCell(row, 1).value = 'Sin datos para este período';
                sheet.getRow(row).height = 16;
            }

            // ── Fila en hoja Resumen ──────────────────────────────
            const isResEven = resRow % 2 === 0;
            const resBg     = isResEven ? C.resumeEven : 'FFFFFFFF';
            const resDataRow = resSheet.getRow(resRow);
            resDataRow.height = 18;

            const resVals = [
                resultado.contribuyente || empresa.nombre,
                resultado.rut || empresa.rut,
                resultado.boletasEmitidas?.totales?.honorarioBruto  || 0,
                resultado.boletasRecibidas?.totales?.honorarioBruto || 0,
                resultado.bteRecibidas?.totales?.montoTotal         || 0,
                resultado.bteEmitidas?.totales?.montoTotal          || 0,
                empresa.estado === 'completado' ? 'Completado' : 'Error',
            ];

            resVals.forEach((val, i) => {
                const cell = resDataRow.getCell(i + 1);
                cell.value = val;
                const isMoney = i >= 2 && i <= 5;
                sc(cell, {
                    bg: resBg, size: 9,
                    center: i === 6,
                    right: isMoney,
                    numFmt: isMoney ? MONEY_FMT : null,
                });
                cell.border = THIN;
                // Color estado
                if (i === 6) {
                    cell.font = {
                        name: 'Calibri', bold: true, size: 9,
                        color: { argb: val === 'Completado' ? 'FF1A7A4A' : 'FF9B2222' },
                    };
                }
            });
            resRow++;
        }

        // ── Fila totales hoja Resumen ─────────────────────────────
        if (resRow > 5) {
            resSheet.getRow(resRow).height = 7; // separador
            resRow++;

            const resTotRow = resSheet.getRow(resRow);
            resTotRow.height = 20;
            const resTotVals = [
                'TOTALES', '',
                resTotals.emit, resTotals.rec, resTotals.bteRec, resTotals.bteEmit, '',
            ];
            resTotVals.forEach((val, i) => {
                const cell = resTotRow.getCell(i + 1);
                cell.value = val;
                const isMoney = i >= 2 && i <= 5;
                sc(cell, {
                    bg: C.resumeTotal, fg: 'FFFFFFFF', bold: true, size: 9.5,
                    center: i === 0 || i === 6,
                    right: isMoney,
                    numFmt: isMoney ? MONEY_FMT : null,
                });
                cell.border = {
                    top:    { style: 'medium', color: { argb: C.totalBorder } },
                    bottom: { style: 'medium', color: { argb: C.totalBorder } },
                    left:   { style: 'thin',   color: { argb: C.resumeTotal } },
                    right:  { style: 'thin',   color: { argb: C.resumeTotal } },
                };
            });
        }

        // ── Generar y enviar ──────────────────────────────────────
        const exportDir = path.join(__dirname, 'exports');
        if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir, { recursive: true });

        const { injectVBA } = require('./bot/excelVBA');

        // Generar buffer XLSX desde ExcelJS
        const xlsxBuffer = await workbook.xlsx.writeBuffer();

        // Intentar inyectar macros VBA → convierte a XLSM
        const xlsmBuffer = await injectVBA(xlsxBuffer);

        const ext      = xlsmBuffer ? 'xlsm' : 'xlsx';
        const fileName = `Boletas_SII_${new Date().toISOString().slice(0, 10)}.${ext}`;
        const filePath = path.join(exportDir, fileName);

        fs.writeFileSync(filePath, xlsmBuffer || xlsxBuffer);
        res.download(filePath, fileName);

    } catch (error) {
        console.error('Error al exportar:', error);
        res.status(500).json({ error: 'Error al generar el archivo Excel' });
    }
});

// Socket.io para comunicación en tiempo real
io.on('connection', (socket) => {
    console.log('Cliente conectado');

    socket.on('startProcess', async () => {
        if (isProcessing) {
            socket.emit('log', { type: 'warning', message: '⚠️ Ya hay un proceso en ejecución' });
            return;
        }

        if (empresasData.length === 0) {
            socket.emit('log', { type: 'error', message: '❌ No hay empresas cargadas. Sube un archivo Excel primero.' });
            return;
        }

        isProcessing = true;
        socket.emit('log', { type: 'info', message: '🚀 Iniciando proceso de extracción...' });
        socket.emit('log', { type: 'info', message: `📊 Total de empresas a procesar: ${empresasData.length}` });

        try {
            botInstance = new SIIBot(io);

            for (let i = 0; i < empresasData.length; i++) {
                if (!isProcessing) break;

                const empresa = empresasData[i];
                socket.emit('log', { type: 'info', message: `\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━` });
                socket.emit('log', { type: 'info', message: `📌 Procesando empresa ${i + 1}/${empresasData.length}: ${empresa.nombre}` });

                empresa.estado = 'procesando';
                socket.emit('empresaUpdate', { id: empresa.id, estado: 'procesando' });

                try {
                    const resultado = await botInstance.procesarEmpresa(empresa);
                    resultadosEmpresas[empresa.id] = resultado;
                    empresa.estado = 'completado';
                    socket.emit('empresaUpdate', { id: empresa.id, estado: 'completado' });
                    socket.emit('log', { type: 'success', message: `✅ Empresa ${empresa.nombre} procesada correctamente` });
                } catch (error) {
                    console.error(`Error procesando ${empresa.nombre}:`, error);
                    empresa.estado = 'error';
                    socket.emit('empresaUpdate', { id: empresa.id, estado: 'error' });
                    socket.emit('log', { type: 'error', message: `❌ Error en ${empresa.nombre}: ${error.message}` });
                }
            }

            socket.emit('log', { type: 'success', message: '\n🎉 Proceso completado!' });
            socket.emit('processComplete');
        } catch (error) {
            socket.emit('log', { type: 'error', message: `❌ Error general: ${error.message}` });
        } finally {
            isProcessing = false;
            if (botInstance) {
                await botInstance.close();
            }
        }
    });

    socket.on('disconnect', () => {
        console.log('Cliente desconectado');
    });
});

const PORT = process.env.PORT || 3001;
server.listen(PORT, () => {
    console.log(`🚀 Servidor ejecutándose en http://localhost:${PORT}`);
});
