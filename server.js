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

        for (const empresaId in resultadosEmpresas) {
            const resultado = resultadosEmpresas[empresaId];
            const empresa = empresasData.find(e => e.id === parseInt(empresaId));

            if (!empresa || !resultado) continue;

            // Crear hoja para esta empresa
            const sheetName = empresa.nombre.substring(0, 31).replace(/[*?:/\\[\]]/g, '');
            const sheet = workbook.addWorksheet(sheetName);

            // Estilo para encabezados
            const headerStyle = {
                font: { bold: true, color: { argb: 'FFFFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1a365d' } },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };

            // Información del contribuyente
            sheet.mergeCells('A1:I1');
            sheet.getCell('A1').value = `Contribuyente: ${resultado.contribuyente || empresa.nombre}`;
            sheet.getCell('A1').font = { bold: true, size: 14 };

            sheet.mergeCells('A2:I2');
            sheet.getCell('A2').value = `RUT: ${resultado.rut || empresa.rut}`;
            sheet.getCell('A2').font = { bold: true, size: 12 };

            // === BOLETAS EMITIDAS ===
            sheet.mergeCells('A4:I4');
            sheet.getCell('A4').value = 'BOLETAS EMITIDAS - AÑO 2025';
            sheet.getCell('A4').font = { bold: true, size: 14, color: { argb: 'FF1a365d' } };
            sheet.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFe2e8f0' } };

            // Encabezados boletas emitidas
            const headersEmitidas = ['PERIODO', 'FOLIO INICIAL', 'FOLIO FINAL', 'VIGENTES', 'ANULADAS', 'HONORARIO BRUTO', 'RET. TERCEROS', 'RET. CONTRIBUYENTE', 'TOTAL LÍQUIDO'];
            const headerRowEmitidas = sheet.getRow(5);
            headersEmitidas.forEach((header, idx) => {
                const cell = headerRowEmitidas.getCell(idx + 1);
                cell.value = header;
                cell.font = headerStyle.font;
                cell.fill = headerStyle.fill;
                cell.alignment = headerStyle.alignment;
            });

            // Datos de boletas emitidas
            if (resultado.boletasEmitidas && resultado.boletasEmitidas.meses) {
                let rowNum = 6;
                resultado.boletasEmitidas.meses.forEach(mes => {
                    const row = sheet.getRow(rowNum);
                    row.getCell(1).value = mes.periodo;
                    row.getCell(2).value = mes.folioInicial || '';
                    row.getCell(3).value = mes.folioFinal || '';
                    row.getCell(4).value = mes.vigentes || 0;
                    row.getCell(5).value = mes.anuladas || 0;
                    row.getCell(6).value = mes.honorarioBruto || 0;
                    row.getCell(7).value = mes.retencionTerceros || 0;
                    row.getCell(8).value = mes.retencionContribuyente || 0;
                    row.getCell(9).value = mes.totalLiquido || 0;
                    rowNum++;
                });

                // Fila de totales
                const totalRowEmitidas = sheet.getRow(rowNum);
                totalRowEmitidas.getCell(1).value = 'TOTALES';
                totalRowEmitidas.getCell(1).font = { bold: true };
                totalRowEmitidas.getCell(4).value = resultado.boletasEmitidas.totales?.vigentes || 0;
                totalRowEmitidas.getCell(5).value = resultado.boletasEmitidas.totales?.anuladas || 0;
                totalRowEmitidas.getCell(6).value = resultado.boletasEmitidas.totales?.honorarioBruto || 0;
                totalRowEmitidas.getCell(7).value = resultado.boletasEmitidas.totales?.retencionTerceros || 0;
                totalRowEmitidas.getCell(8).value = resultado.boletasEmitidas.totales?.retencionContribuyente || 0;
                totalRowEmitidas.getCell(9).value = resultado.boletasEmitidas.totales?.totalLiquido || 0;
                totalRowEmitidas.eachCell(cell => { cell.font = { bold: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFe2e8f0' } }; });
                rowNum += 2;

                // Mensaje resumen de honorarios brutos emitidos
                sheet.mergeCells(`A${rowNum}:I${rowNum}`);
                const mensajeEmitidas = sheet.getCell(`A${rowNum}`);
                mensajeEmitidas.value = `El contribuyente ${resultado.contribuyente || empresa.nombre} tiene como HONORARIOS BRUTOS EMITIDOS: $${(resultado.boletasEmitidas.totales?.honorarioBruto || 0).toLocaleString('es-CL')}`;
                mensajeEmitidas.font = { bold: true, size: 12, color: { argb: 'FF1a365d' } };
                mensajeEmitidas.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } };
                rowNum += 3;

                // === BOLETAS RECIBIDAS ===
                sheet.mergeCells(`A${rowNum}:I${rowNum}`);
                sheet.getCell(`A${rowNum}`).value = 'BOLETAS RECIBIDAS - AÑO 2025';
                sheet.getCell(`A${rowNum}`).font = { bold: true, size: 14, color: { argb: 'FF1a365d' } };
                sheet.getCell(`A${rowNum}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFe2e8f0' } };
                rowNum++;

                // Encabezados boletas recibidas
                const headersRecibidas = ['PERIODO', 'VIGENTES', 'ANULADAS', 'HONORARIO BRUTO', 'RET. TERCEROS', 'RET. CONTRIBUYENTE', 'TOTAL LÍQUIDO'];
                const headerRowRecibidas = sheet.getRow(rowNum);
                headersRecibidas.forEach((header, idx) => {
                    const cell = headerRowRecibidas.getCell(idx + 1);
                    cell.value = header;
                    cell.font = headerStyle.font;
                    cell.fill = headerStyle.fill;
                    cell.alignment = headerStyle.alignment;
                });
                rowNum++;

                if (resultado.boletasRecibidas && resultado.boletasRecibidas.meses) {
                    resultado.boletasRecibidas.meses.forEach(mes => {
                        const row = sheet.getRow(rowNum);
                        row.getCell(1).value = mes.periodo;
                        row.getCell(2).value = mes.vigentes || 0;
                        row.getCell(3).value = mes.anuladas || 0;
                        row.getCell(4).value = mes.honorarioBruto || 0;
                        row.getCell(5).value = mes.retencionTerceros || 0;
                        row.getCell(6).value = mes.retencionContribuyente || 0;
                        row.getCell(7).value = mes.totalLiquido || 0;
                        rowNum++;
                    });

                    // Fila de totales recibidas
                    const totalRowRecibidas = sheet.getRow(rowNum);
                    totalRowRecibidas.getCell(1).value = 'TOTALES';
                    totalRowRecibidas.getCell(1).font = { bold: true };
                    totalRowRecibidas.getCell(2).value = resultado.boletasRecibidas.totales?.vigentes || 0;
                    totalRowRecibidas.getCell(3).value = resultado.boletasRecibidas.totales?.anuladas || 0;
                    totalRowRecibidas.getCell(4).value = resultado.boletasRecibidas.totales?.honorarioBruto || 0;
                    totalRowRecibidas.getCell(5).value = resultado.boletasRecibidas.totales?.retencionTerceros || 0;
                    totalRowRecibidas.getCell(6).value = resultado.boletasRecibidas.totales?.retencionContribuyente || 0;
                    totalRowRecibidas.getCell(7).value = resultado.boletasRecibidas.totales?.totalLiquido || 0;
                    totalRowRecibidas.eachCell(cell => { cell.font = { bold: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFe2e8f0' } }; });
                    rowNum += 2;

                    // Mensaje resumen de honorarios brutos recibidos
                    sheet.mergeCells(`A${rowNum}:G${rowNum}`);
                    const mensajeRecibidas = sheet.getCell(`A${rowNum}`);
                    mensajeRecibidas.value = `El contribuyente ${resultado.contribuyente || empresa.nombre} tiene como HONORARIOS BRUTOS RECIBIDOS: $${(resultado.boletasRecibidas.totales?.honorarioBruto || 0).toLocaleString('es-CL')}`;
                    mensajeRecibidas.font = { bold: true, size: 12, color: { argb: 'FF1a365d' } };
                    mensajeRecibidas.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } };
                    rowNum += 3;
                }

                // === BTE RECIBIDAS ===
                if (resultado.bteRecibidas && resultado.bteRecibidas.meses) {
                    sheet.mergeCells(`A${rowNum}:F${rowNum}`);
                    sheet.getCell(`A${rowNum}`).value = 'BTE RECIBIDAS (PRESTACIÓN DE SERVICIOS DE TERCEROS) - AÑO 2025';
                    sheet.getCell(`A${rowNum}`).font = { bold: true, size: 14, color: { argb: 'FF276749' } };
                    sheet.getCell(`A${rowNum}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFc6f6d5' } };
                    rowNum++;

                    const headersBTE = ['PERIODO', 'CANTIDAD', 'MONTO NETO', 'MONTO EXENTO', 'IVA', 'MONTO TOTAL'];
                    const headerRowBTERec = sheet.getRow(rowNum);
                    headersBTE.forEach((header, idx) => {
                        const cell = headerRowBTERec.getCell(idx + 1);
                        cell.value = header;
                        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF276749' } };
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    });
                    rowNum++;

                    resultado.bteRecibidas.meses.forEach(mes => {
                        const row = sheet.getRow(rowNum);
                        row.getCell(1).value = mes.periodo;
                        row.getCell(2).value = mes.cantidad || 0;
                        row.getCell(3).value = mes.montoNeto || 0;
                        row.getCell(4).value = mes.montoExento || 0;
                        row.getCell(5).value = mes.montoIva || 0;
                        row.getCell(6).value = mes.montoTotal || 0;
                        rowNum++;
                    });

                    const totalRowBTERec = sheet.getRow(rowNum);
                    totalRowBTERec.getCell(1).value = 'TOTALES';
                    totalRowBTERec.getCell(2).value = resultado.bteRecibidas.totales?.cantidad || 0;
                    totalRowBTERec.getCell(3).value = resultado.bteRecibidas.totales?.montoNeto || 0;
                    totalRowBTERec.getCell(4).value = resultado.bteRecibidas.totales?.montoExento || 0;
                    totalRowBTERec.getCell(5).value = resultado.bteRecibidas.totales?.montoIva || 0;
                    totalRowBTERec.getCell(6).value = resultado.bteRecibidas.totales?.montoTotal || 0;
                    totalRowBTERec.eachCell(cell => { cell.font = { bold: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFc6f6d5' } }; });
                    rowNum += 2;

                    sheet.mergeCells(`A${rowNum}:F${rowNum}`);
                    const mensajeBTERec = sheet.getCell(`A${rowNum}`);
                    mensajeBTERec.value = `El contribuyente ${resultado.contribuyente || empresa.nombre} tiene como MONTO TOTAL BTE RECIBIDAS: $${(resultado.bteRecibidas.totales?.montoTotal || 0).toLocaleString('es-CL')}`;
                    mensajeBTERec.font = { bold: true, size: 12, color: { argb: 'FF276749' } };
                    mensajeBTERec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFc6f6d5' } };
                    rowNum += 3;
                }

                // === BTE EMITIDAS ===
                if (resultado.bteEmitidas && resultado.bteEmitidas.meses) {
                    sheet.mergeCells(`A${rowNum}:F${rowNum}`);
                    sheet.getCell(`A${rowNum}`).value = 'BTE EMITIDAS (PRESTACIÓN DE SERVICIOS DE TERCEROS) - AÑO 2025';
                    sheet.getCell(`A${rowNum}`).font = { bold: true, size: 14, color: { argb: 'FF276749' } };
                    sheet.getCell(`A${rowNum}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFc6f6d5' } };
                    rowNum++;

                    const headersBTEEm = ['PERIODO', 'CANTIDAD', 'MONTO NETO', 'MONTO EXENTO', 'IVA', 'MONTO TOTAL'];
                    const headerRowBTEEm = sheet.getRow(rowNum);
                    headersBTEEm.forEach((header, idx) => {
                        const cell = headerRowBTEEm.getCell(idx + 1);
                        cell.value = header;
                        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF276749' } };
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    });
                    rowNum++;

                    resultado.bteEmitidas.meses.forEach(mes => {
                        const row = sheet.getRow(rowNum);
                        row.getCell(1).value = mes.periodo;
                        row.getCell(2).value = mes.cantidad || 0;
                        row.getCell(3).value = mes.montoNeto || 0;
                        row.getCell(4).value = mes.montoExento || 0;
                        row.getCell(5).value = mes.montoIva || 0;
                        row.getCell(6).value = mes.montoTotal || 0;
                        rowNum++;
                    });

                    const totalRowBTEEm = sheet.getRow(rowNum);
                    totalRowBTEEm.getCell(1).value = 'TOTALES';
                    totalRowBTEEm.getCell(2).value = resultado.bteEmitidas.totales?.cantidad || 0;
                    totalRowBTEEm.getCell(3).value = resultado.bteEmitidas.totales?.montoNeto || 0;
                    totalRowBTEEm.getCell(4).value = resultado.bteEmitidas.totales?.montoExento || 0;
                    totalRowBTEEm.getCell(5).value = resultado.bteEmitidas.totales?.montoIva || 0;
                    totalRowBTEEm.getCell(6).value = resultado.bteEmitidas.totales?.montoTotal || 0;
                    totalRowBTEEm.eachCell(cell => { cell.font = { bold: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFc6f6d5' } }; });
                    rowNum += 2;

                    sheet.mergeCells(`A${rowNum}:F${rowNum}`);
                    const mensajeBTEEm = sheet.getCell(`A${rowNum}`);
                    mensajeBTEEm.value = `El contribuyente ${resultado.contribuyente || empresa.nombre} tiene como MONTO TOTAL BTE EMITIDAS: $${(resultado.bteEmitidas.totales?.montoTotal || 0).toLocaleString('es-CL')}`;
                    mensajeBTEEm.font = { bold: true, size: 12, color: { argb: 'FF276749' } };
                    mensajeBTEEm.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFc6f6d5' } };
                }
            }

            // Ajustar ancho de columnas
            sheet.columns.forEach(column => {
                column.width = 18;
            });
        }

        // Generar archivo
        const exportDir = path.join(__dirname, 'exports');
        if (!fs.existsSync(exportDir)) {
            fs.mkdirSync(exportDir, { recursive: true });
        }

        const fileName = `Boletas_SII_${new Date().toISOString().slice(0, 10)}.xlsx`;
        const filePath = path.join(exportDir, fileName);

        await workbook.xlsx.writeFile(filePath);

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

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`🚀 Servidor ejecutándose en http://localhost:${PORT}`);
});
