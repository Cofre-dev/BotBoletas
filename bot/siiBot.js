const puppeteer = require('puppeteer');

class SIIBot {
    constructor(io) {
        this.io = io;
        this.browser = null;
        this.page = null;
        this.isRunning = true;
    }

    log(type, message) {
        this.io.emit('log', { type, message });
    }

    async initialize() {
        this.log('info', '🌐 Iniciando navegador...');
        this.browser = await puppeteer.launch({
            headless: false,
            defaultViewport: { width: 1280, height: 800 },
            args: ['--start-maximized', '--disable-blink-features=AutomationControlled']
        });
        this.page = await this.browser.newPage();
        await this.page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    }

    async stop() {
        this.isRunning = false;
    }

    async close() {
        if (this.browser) {
            await this.browser.close();
            this.browser = null;
            this.page = null;
        }
    }

    async wait(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async procesarEmpresa(empresa) {
        if (!this.browser) {
            await this.initialize();
        }

        const resultado = {
            contribuyente: empresa.nombre,
            rut: empresa.rut,
            boletasEmitidas: null,
            boletasRecibidas: null,
            bteRecibidas: null,
            bteEmitidas: null
        };

        try {
            // 1. Ir a la página de login del SII
            this.log('info', '📍 Navegando al portal SII...');
            await this.page.goto('https://zeusr.sii.cl//AUT2000/InicioAutenticacion/IngresoRutClave.html?https://misiir.sii.cl/cgi_misii/siihome.cgi', {
                waitUntil: 'networkidle2',
                timeout: 60000
            });
            await this.wait(2000);

            // 2. Ingresar RUT
            this.log('info', '🔑 Ingresando credenciales...');
            await this.page.waitForSelector('#rutcntr', { timeout: 10000 });
            await this.page.click('#rutcntr');
            await this.page.type('#rutcntr', empresa.rut, { delay: 50 });
            await this.wait(500);

            // 3. Ingresar clave
            await this.page.waitForSelector('#clave', { timeout: 10000 });
            await this.page.click('#clave');
            await this.page.type('#clave', empresa.clave, { delay: 50 });
            await this.wait(500);

            // 4. Click en botón de ingreso
            await this.page.click('button[type="submit"], input[type="submit"], #bt_ingresar');
            await this.wait(5000);

            // 5. Manejar modal "Antes de continuar" si aparece
            await this.handleModalActualizarDatos();

            // 6. Navegar a Boletas de Honorarios Emitidas
            this.log('info', '📄 Navegando a Boletas de Honorarios Emitidas...');
            resultado.boletasEmitidas = await this.obtenerBoletasEmitidas();

            // 7. Volver al inicio y navegar a Boletas Recibidas
            this.log('info', '📄 Navegando a Boletas de Honorarios Recibidas...');
            resultado.boletasRecibidas = await this.obtenerBoletasRecibidas();

            // 8. Obtener BTE's Recibidas (Boletas de prestación de servicios de terceros)
            this.log('info', '📄 Navegando a BTE Recibidas...');
            resultado.bteRecibidas = await this.obtenerBTERecibidas();

            // 9. Obtener BTE's Emitidas
            this.log('info', '📄 Navegando a BTE Emitidas...');
            resultado.bteEmitidas = await this.obtenerBTEEmitidas();

            // 10. Cerrar sesión
            this.log('info', '🚪 Cerrando sesión...');
            await this.cerrarSesion();

            return resultado;

        } catch (error) {
            this.log('error', `Error en empresa ${empresa.nombre}: ${error.message}`);
            // Intentar cerrar sesión y continuar
            try {
                await this.cerrarSesion();
            } catch (e) {
                // Ignorar error al cerrar sesión
            }
            throw error;
        }
    }

    async handleModalActualizarDatos() {
        try {
            await this.wait(2000);
            // Buscar y clickear "Actualizar más tarde"
            const btnActualizarMasTarde = await this.page.$('#btnActualizarMasTarde');
            if (btnActualizarMasTarde) {
                this.log('info', '📌 Cerrando modal de actualización de datos...');
                await btnActualizarMasTarde.click();
                await this.wait(1000);
            }
        } catch (e) {
            // Modal no apareció, continuar
        }
    }

    async handleModalImportante() {
        try {
            await this.wait(1500);
            // Buscar botón "Cerrar" del modal informativo
            const btnCerrar = await this.page.$('button[data-dismiss="modal"][onclick*="modalInforma"]');
            if (btnCerrar) {
                this.log('info', '📌 Cerrando modal informativo...');
                await btnCerrar.click();
                await this.wait(1000);
            } else {
                // Intentar otro selector
                const allButtons = await this.page.$$('button.btn-default');
                for (const btn of allButtons) {
                    const text = await this.page.evaluate(el => el.textContent, btn);
                    if (text.includes('Cerrar')) {
                        await btn.click();
                        await this.wait(1000);
                        break;
                    }
                }
            }
        } catch (e) {
            // Modal no apareció, continuar
        }
    }

    async obtenerBoletasEmitidas() {
        try {
            // Navegar a "Trámites en línea"
            await this.wait(2000);

            // Click en "Trámites en línea"
            this.log('info', '📍 Buscando Trámites en línea...');
            await this.page.evaluate(() => {
                const items = document.querySelectorAll('li span, li div');
                for (const item of items) {
                    if (item.textContent.includes('Trámites en línea')) {
                        item.closest('li').click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // Click en "Boletas de honorarios electrónicas"
            this.log('info', '📍 Buscando Boletas de honorarios electrónicas...');
            await this.page.evaluate(() => {
                const headers = document.querySelectorAll('h4 span, h4');
                for (const h of headers) {
                    if (h.textContent.includes('Boletas de honorarios electrónicas')) {
                        h.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // Manejar modal informativo si aparece
            await this.handleModalImportante();

            // Click en "Emisor de boleta de honorarios"
            this.log('info', '📍 Accediendo a Emisor de boleta de honorarios...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes('Emisor de boleta de honorarios')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // Click en "Consultas sobre boletas de honorarios electrónicas"
            this.log('info', '📍 Accediendo a Consultas sobre boletas...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a, h4');
                for (const link of links) {
                    if (link.textContent.includes('Consultas sobre boletas de honorarios electrónicas')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // Click en "Consultar boletas emitidas"
            this.log('info', '📍 Accediendo a Consultar boletas emitidas...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes('Consultar boletas emitidas')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // Seleccionar año 2025
            this.log('info', '📅 Seleccionando año 2025...');
            await this.page.select('select[name="cbanoinformeanual"]', '2025');
            await this.wait(1000);

            // Click en consultar
            await this.page.click('input[name="cmdconsultar12"], #cmdconsultar124');
            await this.wait(4000);

            // Extraer datos de la tabla
            this.log('info', '📊 Extrayendo datos de boletas emitidas...');
            const datosEmitidas = await this.extraerDatosTablaEmitidas();

            return datosEmitidas;

        } catch (error) {
            this.log('error', `Error al obtener boletas emitidas: ${error.message}`);
            return null;
        }
    }

    async obtenerBoletasRecibidas() {
        try {
            // Volver al home del SII
            this.log('info', '📍 Volviendo al inicio del SII...');
            await this.page.goto('https://misiir.sii.cl/cgi_misii/siihome.cgi', {
                waitUntil: 'networkidle2',
                timeout: 60000
            });
            await this.wait(2000);

            // Manejar modal de actualización de datos
            await this.handleModalActualizarDatos();

            // Navegar a Trámites en línea
            this.log('info', '📍 Navegando a Trámites en línea...');
            await this.page.evaluate(() => {
                const items = document.querySelectorAll('li span, li div');
                for (const item of items) {
                    if (item.textContent.includes('Trámites en línea')) {
                        item.closest('li').click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // Click en "Boletas de honorarios electrónicas"
            await this.page.evaluate(() => {
                const headers = document.querySelectorAll('h4 span, h4');
                for (const h of headers) {
                    if (h.textContent.includes('Boletas de honorarios electrónicas')) {
                        h.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // Manejar modal informativo
            await this.handleModalImportante();

            // Click en "Emisor de boleta de honorarios"
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes('Emisor de boleta de honorarios')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // Click en "Consultas sobre boletas de honorarios electrónicas"
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a, h4');
                for (const link of links) {
                    if (link.textContent.includes('Consultas sobre boletas de honorarios electrónicas')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // Click en "Consultar boletas recibidas"
            this.log('info', '📍 Accediendo a Consultar boletas recibidas...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes('Consultar boletas recibidas')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // Seleccionar año 2025
            this.log('info', '📅 Seleccionando año 2025...');
            await this.page.select('select[name="cbanoinformeanual"]', '2025');
            await this.wait(1000);

            // Click en consultar
            await this.page.click('input[name="cmdconsultar12"], #cmdconsultar124');
            await this.wait(4000);

            // Extraer datos de la tabla
            this.log('info', '📊 Extrayendo datos de boletas recibidas...');
            const datosRecibidas = await this.extraerDatosTablaRecibidas();

            return datosRecibidas;

        } catch (error) {
            this.log('error', `Error al obtener boletas recibidas: ${error.message}`);
            return null;
        }
    }

    async extraerDatosTablaEmitidas() {
        // Esperar a que los scripts se ejecuten completamente
        this.log('info', '⏳ Esperando que la página cargue completamente...');
        await this.wait(3000);

        // Intentar obtener datos de xml_values directamente (más confiable)
        const datosXml = await this.page.evaluate(() => {
            if (typeof xml_values !== 'undefined') {
                return {
                    contribuyente: xml_values['nombre_contribuyente'] || '',
                    rut: (xml_values['rut_arrastre'] || '') + '-' + (xml_values['dv_arrastre'] || ''),
                    año: xml_values['anio_consulta'] || '',
                    ene: { bruto: xml_values['ene1'], terceros: xml_values['ene2'], contrib: xml_values['ene3'], folioIni: xml_values['ene4'], folioFin: xml_values['ene5'], vigentes: xml_values['ene6'], anuladas: xml_values['ene7'], liquido: xml_values['sumene'] },
                    feb: { bruto: xml_values['feb1'], terceros: xml_values['feb2'], contrib: xml_values['feb3'], folioIni: xml_values['feb4'], folioFin: xml_values['feb5'], vigentes: xml_values['feb6'], anuladas: xml_values['feb7'], liquido: xml_values['sumfeb'] },
                    mar: { bruto: xml_values['mar1'], terceros: xml_values['mar2'], contrib: xml_values['mar3'], folioIni: xml_values['mar4'], folioFin: xml_values['mar5'], vigentes: xml_values['mar6'], anuladas: xml_values['mar7'], liquido: xml_values['summar'] },
                    abr: { bruto: xml_values['abr1'], terceros: xml_values['abr2'], contrib: xml_values['abr3'], folioIni: xml_values['abr4'], folioFin: xml_values['abr5'], vigentes: xml_values['abr6'], anuladas: xml_values['abr7'], liquido: xml_values['sumabr'] },
                    may: { bruto: xml_values['may1'], terceros: xml_values['may2'], contrib: xml_values['may3'], folioIni: xml_values['may4'], folioFin: xml_values['may5'], vigentes: xml_values['may6'], anuladas: xml_values['may7'], liquido: xml_values['summay'] },
                    jun: { bruto: xml_values['jun1'], terceros: xml_values['jun2'], contrib: xml_values['jun3'], folioIni: xml_values['jun4'], folioFin: xml_values['jun5'], vigentes: xml_values['jun6'], anuladas: xml_values['jun7'], liquido: xml_values['sumjun'] },
                    jul: { bruto: xml_values['jul1'], terceros: xml_values['jul2'], contrib: xml_values['jul3'], folioIni: xml_values['jul4'], folioFin: xml_values['jul5'], vigentes: xml_values['jul6'], anuladas: xml_values['jul7'], liquido: xml_values['sumjul'] },
                    ago: { bruto: xml_values['ago1'], terceros: xml_values['ago2'], contrib: xml_values['ago3'], folioIni: xml_values['ago4'], folioFin: xml_values['ago5'], vigentes: xml_values['ago6'], anuladas: xml_values['ago7'], liquido: xml_values['sumago'] },
                    sep: { bruto: xml_values['sep1'], terceros: xml_values['sep2'], contrib: xml_values['sep3'], folioIni: xml_values['sep4'], folioFin: xml_values['sep5'], vigentes: xml_values['sep6'], anuladas: xml_values['sep7'], liquido: xml_values['sumsep'] },
                    oct: { bruto: xml_values['oct1'], terceros: xml_values['oct2'], contrib: xml_values['oct3'], folioIni: xml_values['oct4'], folioFin: xml_values['oct5'], vigentes: xml_values['oct6'], anuladas: xml_values['oct7'], liquido: xml_values['sumoct'] },
                    nov: { bruto: xml_values['nov1'], terceros: xml_values['nov2'], contrib: xml_values['nov3'], folioIni: xml_values['nov4'], folioFin: xml_values['nov5'], vigentes: xml_values['nov6'], anuladas: xml_values['nov7'], liquido: xml_values['sumnov'] },
                    dic: { bruto: xml_values['dic1'], terceros: xml_values['dic2'], contrib: xml_values['dic3'], folioIni: xml_values['dic4'], folioFin: xml_values['dic5'], vigentes: xml_values['dic6'], anuladas: xml_values['dic7'], liquido: xml_values['sumdic'] },
                    totales: { vigentes: xml_values['tot6'], anuladas: xml_values['tot7'], bruto: xml_values['tot1'], terceros: xml_values['tot2'], contrib: xml_values['tot3'], liquido: xml_values['sumtot'] }
                };
            }
            return null;
        });

        if (datosXml) {
            this.log('success', '✅ Datos obtenidos de xml_values');
            const mesesNombres = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
            const mesesKeys = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

            const parseNum = (val) => parseInt(val) || 0;

            const resultado = {
                contribuyente: datosXml.contribuyente,
                rut: datosXml.rut,
                meses: mesesKeys.map((key, idx) => ({
                    periodo: mesesNombres[idx],
                    folioInicial: parseNum(datosXml[key].folioIni),
                    folioFinal: parseNum(datosXml[key].folioFin),
                    vigentes: parseNum(datosXml[key].vigentes),
                    anuladas: parseNum(datosXml[key].anuladas),
                    honorarioBruto: parseNum(datosXml[key].bruto),
                    retencionTerceros: parseNum(datosXml[key].terceros),
                    retencionContribuyente: parseNum(datosXml[key].contrib),
                    totalLiquido: parseNum(datosXml[key].liquido)
                })),
                totales: {
                    vigentes: parseNum(datosXml.totales.vigentes),
                    anuladas: parseNum(datosXml.totales.anuladas),
                    honorarioBruto: parseNum(datosXml.totales.bruto),
                    retencionTerceros: parseNum(datosXml.totales.terceros),
                    retencionContribuyente: parseNum(datosXml.totales.contrib),
                    totalLiquido: parseNum(datosXml.totales.liquido)
                }
            };
            return resultado;
        }

        // Fallback: extraer desde la tabla (usando innerText para obtener valores renderizados)
        this.log('info', '📊 Extrayendo desde tabla HTML...');
        const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

        const datos = await this.page.evaluate((meses) => {
            const resultado = {
                contribuyente: '',
                rut: '',
                meses: [],
                totales: { vigentes: 0, anuladas: 0, honorarioBruto: 0, retencionTerceros: 0, retencionContribuyente: 0, totalLiquido: 0 }
            };

            // Buscar contribuyente y RUT en las filas de información
            const allTds = document.querySelectorAll('td');
            for (let i = 0; i < allTds.length; i++) {
                const td = allTds[i];
                if (td.innerText && td.innerText.includes('Contribuyente:')) {
                    const nextTd = allTds[i + 1];
                    if (nextTd) resultado.contribuyente = nextTd.innerText.trim();
                }
                if (td.innerText && td.innerText.includes('RUT:')) {
                    const nextTd = allTds[i + 1];
                    if (nextTd) resultado.rut = nextTd.innerText.trim();
                }
            }

            // Buscar la tabla principal con datos
            const tables = document.querySelectorAll('table');
            let dataTable = null;
            for (const table of tables) {
                if (table.innerText.includes('PERIODOS') && table.innerText.includes('HONORARIO BRUTO')) {
                    dataTable = table;
                    break;
                }
            }

            if (!dataTable) return resultado;

            const rows = dataTable.querySelectorAll('tr');
            const parseNumber = (text) => {
                if (!text) return 0;
                const cleaned = text.replace(/\./g, '').replace(/,/g, '').replace(/\s/g, '').replace(/\u00A0/g, '').trim();
                return parseInt(cleaned) || 0;
            };

            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 9) {
                    const periodoText = (cells[0].innerText || '').trim().toUpperCase();
                    const mesEncontrado = meses.find(m => periodoText.includes(m));

                    if (mesEncontrado) {
                        resultado.meses.push({
                            periodo: mesEncontrado,
                            folioInicial: parseNumber(cells[1].innerText),
                            folioFinal: parseNumber(cells[2].innerText),
                            vigentes: parseNumber(cells[3].innerText),
                            anuladas: parseNumber(cells[4].innerText),
                            honorarioBruto: parseNumber(cells[5].innerText),
                            retencionTerceros: parseNumber(cells[6].innerText),
                            retencionContribuyente: parseNumber(cells[7].innerText),
                            totalLiquido: parseNumber(cells[8].innerText)
                        });
                    }

                    if (periodoText.includes('TOTAL')) {
                        resultado.totales.vigentes = parseNumber(cells[3]?.innerText);
                        resultado.totales.anuladas = parseNumber(cells[4]?.innerText);
                        resultado.totales.honorarioBruto = parseNumber(cells[5]?.innerText);
                        resultado.totales.retencionTerceros = parseNumber(cells[6]?.innerText);
                        resultado.totales.retencionContribuyente = parseNumber(cells[7]?.innerText);
                        resultado.totales.totalLiquido = parseNumber(cells[8]?.innerText);
                    }
                }
            }

            // Calcular totales si no se encontraron
            if (resultado.totales.honorarioBruto === 0 && resultado.meses.length > 0) {
                resultado.totales.honorarioBruto = resultado.meses.reduce((sum, m) => sum + m.honorarioBruto, 0);
                resultado.totales.retencionTerceros = resultado.meses.reduce((sum, m) => sum + m.retencionTerceros, 0);
                resultado.totales.retencionContribuyente = resultado.meses.reduce((sum, m) => sum + m.retencionContribuyente, 0);
                resultado.totales.totalLiquido = resultado.meses.reduce((sum, m) => sum + m.totalLiquido, 0);
                resultado.totales.vigentes = resultado.meses.reduce((sum, m) => sum + m.vigentes, 0);
                resultado.totales.anuladas = resultado.meses.reduce((sum, m) => sum + m.anuladas, 0);
            }

            return resultado;
        }, meses);

        return datos;
    }

    async extraerDatosTablaRecibidas() {
        // Esperar a que los scripts se ejecuten completamente
        this.log('info', '⏳ Esperando que la página cargue completamente...');
        await this.wait(3000);

        // Intentar obtener datos de xml_values directamente (más confiable)
        const datosXml = await this.page.evaluate(() => {
            if (typeof xml_values !== 'undefined') {
                return {
                    ene: { bruto: xml_values['ene1'], terceros: xml_values['ene2'], contrib: xml_values['ene3'], vigentes: xml_values['ene6'], anuladas: xml_values['ene7'], liquido: xml_values['sumene'] },
                    feb: { bruto: xml_values['feb1'], terceros: xml_values['feb2'], contrib: xml_values['feb3'], vigentes: xml_values['feb6'], anuladas: xml_values['feb7'], liquido: xml_values['sumfeb'] },
                    mar: { bruto: xml_values['mar1'], terceros: xml_values['mar2'], contrib: xml_values['mar3'], vigentes: xml_values['mar6'], anuladas: xml_values['mar7'], liquido: xml_values['summar'] },
                    abr: { bruto: xml_values['abr1'], terceros: xml_values['abr2'], contrib: xml_values['abr3'], vigentes: xml_values['abr6'], anuladas: xml_values['abr7'], liquido: xml_values['sumabr'] },
                    may: { bruto: xml_values['may1'], terceros: xml_values['may2'], contrib: xml_values['may3'], vigentes: xml_values['may6'], anuladas: xml_values['may7'], liquido: xml_values['summay'] },
                    jun: { bruto: xml_values['jun1'], terceros: xml_values['jun2'], contrib: xml_values['jun3'], vigentes: xml_values['jun6'], anuladas: xml_values['jun7'], liquido: xml_values['sumjun'] },
                    jul: { bruto: xml_values['jul1'], terceros: xml_values['jul2'], contrib: xml_values['jul3'], vigentes: xml_values['jul6'], anuladas: xml_values['jul7'], liquido: xml_values['sumjul'] },
                    ago: { bruto: xml_values['ago1'], terceros: xml_values['ago2'], contrib: xml_values['ago3'], vigentes: xml_values['ago6'], anuladas: xml_values['ago7'], liquido: xml_values['sumago'] },
                    sep: { bruto: xml_values['sep1'], terceros: xml_values['sep2'], contrib: xml_values['sep3'], vigentes: xml_values['sep6'], anuladas: xml_values['sep7'], liquido: xml_values['sumsep'] },
                    oct: { bruto: xml_values['oct1'], terceros: xml_values['oct2'], contrib: xml_values['oct3'], vigentes: xml_values['oct6'], anuladas: xml_values['oct7'], liquido: xml_values['sumoct'] },
                    nov: { bruto: xml_values['nov1'], terceros: xml_values['nov2'], contrib: xml_values['nov3'], vigentes: xml_values['nov6'], anuladas: xml_values['nov7'], liquido: xml_values['sumnov'] },
                    dic: { bruto: xml_values['dic1'], terceros: xml_values['dic2'], contrib: xml_values['dic3'], vigentes: xml_values['dic6'], anuladas: xml_values['dic7'], liquido: xml_values['sumdic'] },
                    totales: { vigentes: xml_values['tot6'], anuladas: xml_values['tot7'], bruto: xml_values['tot1'], terceros: xml_values['tot2'], contrib: xml_values['tot3'], liquido: xml_values['sumtot'] }
                };
            }
            return null;
        });

        if (datosXml) {
            this.log('success', '✅ Datos obtenidos de xml_values');
            const mesesNombres = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
            const mesesKeys = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

            const parseNum = (val) => parseInt(val) || 0;

            const resultado = {
                meses: mesesKeys.map((key, idx) => ({
                    periodo: mesesNombres[idx],
                    vigentes: parseNum(datosXml[key].vigentes),
                    anuladas: parseNum(datosXml[key].anuladas),
                    honorarioBruto: parseNum(datosXml[key].bruto),
                    retencionTerceros: parseNum(datosXml[key].terceros),
                    retencionContribuyente: parseNum(datosXml[key].contrib),
                    totalLiquido: parseNum(datosXml[key].liquido)
                })),
                totales: {
                    vigentes: parseNum(datosXml.totales.vigentes),
                    anuladas: parseNum(datosXml.totales.anuladas),
                    honorarioBruto: parseNum(datosXml.totales.bruto),
                    retencionTerceros: parseNum(datosXml.totales.terceros),
                    retencionContribuyente: parseNum(datosXml.totales.contrib),
                    totalLiquido: parseNum(datosXml.totales.liquido)
                }
            };
            return resultado;
        }

        // Fallback: extraer desde la tabla (usando innerText para obtener valores renderizados)
        this.log('info', '📊 Extrayendo desde tabla HTML...');
        const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

        const datos = await this.page.evaluate((meses) => {
            const resultado = {
                meses: [],
                totales: { vigentes: 0, anuladas: 0, honorarioBruto: 0, retencionTerceros: 0, retencionContribuyente: 0, totalLiquido: 0 }
            };

            // Buscar la tabla principal con datos
            const tables = document.querySelectorAll('table');
            let dataTable = null;
            for (const table of tables) {
                if (table.innerText.includes('PERIODOS') && table.innerText.includes('HONORARIO BRUTO')) {
                    dataTable = table;
                    break;
                }
            }

            if (!dataTable) return resultado;

            const rows = dataTable.querySelectorAll('tr');
            const parseNumber = (text) => {
                if (!text) return 0;
                const cleaned = text.replace(/\./g, '').replace(/,/g, '').replace(/\s/g, '').replace(/\u00A0/g, '').trim();
                return parseInt(cleaned) || 0;
            };

            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 7) {
                    const periodoText = (cells[0].innerText || '').trim().toUpperCase();
                    const mesEncontrado = meses.find(m => periodoText.includes(m));

                    if (mesEncontrado) {
                        resultado.meses.push({
                            periodo: mesEncontrado,
                            vigentes: parseNumber(cells[1].innerText),
                            anuladas: parseNumber(cells[2].innerText),
                            honorarioBruto: parseNumber(cells[3].innerText),
                            retencionTerceros: parseNumber(cells[4].innerText),
                            retencionContribuyente: parseNumber(cells[5].innerText),
                            totalLiquido: parseNumber(cells[6].innerText)
                        });
                    }

                    if (periodoText.includes('TOTAL')) {
                        resultado.totales.vigentes = parseNumber(cells[1]?.innerText);
                        resultado.totales.anuladas = parseNumber(cells[2]?.innerText);
                        resultado.totales.honorarioBruto = parseNumber(cells[3]?.innerText);
                        resultado.totales.retencionTerceros = parseNumber(cells[4]?.innerText);
                        resultado.totales.retencionContribuyente = parseNumber(cells[5]?.innerText);
                        resultado.totales.totalLiquido = parseNumber(cells[6]?.innerText);
                    }
                }
            }

            // Calcular totales si no se encontraron
            if (resultado.totales.honorarioBruto === 0 && resultado.meses.length > 0) {
                resultado.totales.honorarioBruto = resultado.meses.reduce((sum, m) => sum + m.honorarioBruto, 0);
                resultado.totales.retencionTerceros = resultado.meses.reduce((sum, m) => sum + m.retencionTerceros, 0);
                resultado.totales.retencionContribuyente = resultado.meses.reduce((sum, m) => sum + m.retencionContribuyente, 0);
                resultado.totales.totalLiquido = resultado.meses.reduce((sum, m) => sum + m.totalLiquido, 0);
                resultado.totales.vigentes = resultado.meses.reduce((sum, m) => sum + m.vigentes, 0);
                resultado.totales.anuladas = resultado.meses.reduce((sum, m) => sum + m.anuladas, 0);
            }

            return resultado;
        }, meses);

        return datos;
    }

    async obtenerBTERecibidas() {
        try {
            // Volver al home del SII
            this.log('info', '📍 Volviendo al inicio del SII para BTE Recibidas...');
            await this.page.goto('https://misiir.sii.cl/cgi_misii/siihome.cgi', {
                waitUntil: 'networkidle2',
                timeout: 60000
            });
            await this.wait(2000);

            // Manejar modal de actualización de datos
            await this.handleModalActualizarDatos();

            // 1. Click en "Trámites en línea"
            this.log('info', '📍 Navegando a Trámites en línea...');
            await this.page.evaluate(() => {
                const items = document.querySelectorAll('li span, li div');
                for (const item of items) {
                    if (item.textContent.includes('Trámites en línea')) {
                        item.closest('li').click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // 2. Click en "Boletas de honorarios electrónicas"
            await this.page.evaluate(() => {
                const headers = document.querySelectorAll('h4 span, h4');
                for (const h of headers) {
                    if (h.textContent.includes('Boletas de honorarios electrónicas')) {
                        h.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // 3. Manejar modal informativo
            await this.handleModalImportante();

            // 4. Click en "Boleta de prestación de servicios de terceros electrónica"
            this.log('info', '📍 Accediendo a Boleta de prestación de servicios de terceros...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes('Boleta de prestación de servicios de terceros')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // 5. Click en "Consulta de BTE's recibidas"
            this.log('info', '📍 Accediendo a Consulta de BTE recibidas...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes("Consulta de BTE") && link.textContent.includes("recibidas")) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // 6. Seleccionar año 2025
            this.log('info', '📅 Seleccionando año 2025...');
            await this.page.select('select[name="ANOA"], #ANOA', '2025');
            await this.wait(1000);

            // 7. Click en Consultar
            await this.page.click('input[name="consultaA"]');
            await this.wait(4000);

            // 8. Extraer datos de la tabla
            this.log('info', '📊 Extrayendo datos de BTE recibidas...');
            const datos = await this.extraerDatosBTE();

            return datos;

        } catch (error) {
            this.log('error', `Error al obtener BTE recibidas: ${error.message}`);
            return null;
        }
    }

    async obtenerBTEEmitidas() {
        try {
            // Volver al home del SII
            this.log('info', '📍 Volviendo al inicio del SII para BTE Emitidas...');
            await this.page.goto('https://misiir.sii.cl/cgi_misii/siihome.cgi', {
                waitUntil: 'networkidle2',
                timeout: 60000
            });
            await this.wait(2000);

            // Manejar modal de actualización de datos
            await this.handleModalActualizarDatos();

            // 1. Click en "Trámites en línea"
            this.log('info', '📍 Navegando a Trámites en línea...');
            await this.page.evaluate(() => {
                const items = document.querySelectorAll('li span, li div');
                for (const item of items) {
                    if (item.textContent.includes('Trámites en línea')) {
                        item.closest('li').click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // 2. Click en "Boletas de honorarios electrónicas"
            await this.page.evaluate(() => {
                const headers = document.querySelectorAll('h4 span, h4');
                for (const h of headers) {
                    if (h.textContent.includes('Boletas de honorarios electrónicas')) {
                        h.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);

            // 3. Manejar modal informativo
            await this.handleModalImportante();

            // 4. Click en "Boleta de prestación de servicios de terceros electrónica"
            this.log('info', '📍 Accediendo a Boleta de prestación de servicios de terceros...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes('Boleta de prestación de servicios de terceros')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // 5. Click en "Consulta de BTE's emitidas"
            this.log('info', '📍 Accediendo a Consulta de BTE emitidas...');
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.includes("Consulta de BTE") && link.textContent.includes("emitidas")) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(3000);

            // 6. Seleccionar año 2025
            this.log('info', '📅 Seleccionando año 2025...');
            await this.page.select('select[name="ANOA"], #ANOA', '2025');
            await this.wait(1000);

            // 7. Click en Consultar
            await this.page.click('input[name="consultaA"]');
            await this.wait(4000);

            // 8. Extraer datos de la tabla
            this.log('info', '📊 Extrayendo datos de BTE emitidas...');
            const datos = await this.extraerDatosBTE();

            return datos;

        } catch (error) {
            this.log('error', `Error al obtener BTE emitidas: ${error.message}`);
            return null;
        }
    }

    async extraerDatosBTE() {
        // Esperar a que la página cargue
        await this.wait(3000);

        const datos = await this.page.evaluate(() => {
            const resultado = {
                meses: [],
                totales: {
                    cantidad: 0,
                    montoNeto: 0,
                    montoExento: 0,
                    montoIva: 0,
                    montoTotal: 0
                }
            };

            const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
                'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

            // Buscar la tabla de datos
            const tables = document.querySelectorAll('table');
            let dataTable = null;

            for (const table of tables) {
                const text = table.innerText || table.textContent;
                if (text.includes('PERIODO') || text.includes('Periodo') ||
                    text.includes('ENERO') || text.includes('Enero') ||
                    text.includes('Monto Neto') || text.includes('MONTO')) {
                    dataTable = table;
                    break;
                }
            }

            if (!dataTable) return resultado;

            const rows = dataTable.querySelectorAll('tr');
            const parseNumber = (text) => {
                if (!text) return 0;
                const cleaned = text.replace(/\./g, '').replace(/,/g, '').replace(/\s/g, '').replace(/\u00A0/g, '').replace(/\$/g, '').trim();
                return parseInt(cleaned) || 0;
            };

            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 2) {
                    const periodoText = (cells[0].innerText || '').trim().toUpperCase();
                    const mesEncontrado = meses.find(m => periodoText.includes(m));

                    if (mesEncontrado) {
                        const rowData = {
                            periodo: mesEncontrado,
                            cantidad: parseNumber(cells[1]?.innerText),
                            montoNeto: parseNumber(cells[2]?.innerText),
                            montoExento: parseNumber(cells[3]?.innerText),
                            montoIva: parseNumber(cells[4]?.innerText),
                            montoTotal: parseNumber(cells[5]?.innerText)
                        };
                        resultado.meses.push(rowData);
                    }

                    if (periodoText.includes('TOTAL')) {
                        resultado.totales.cantidad = parseNumber(cells[1]?.innerText);
                        resultado.totales.montoNeto = parseNumber(cells[2]?.innerText);
                        resultado.totales.montoExento = parseNumber(cells[3]?.innerText);
                        resultado.totales.montoIva = parseNumber(cells[4]?.innerText);
                        resultado.totales.montoTotal = parseNumber(cells[5]?.innerText);
                    }
                }
            }

            // Calcular totales si no se encontraron
            if (resultado.totales.montoTotal === 0 && resultado.meses.length > 0) {
                resultado.totales.cantidad = resultado.meses.reduce((sum, m) => sum + (m.cantidad || 0), 0);
                resultado.totales.montoNeto = resultado.meses.reduce((sum, m) => sum + (m.montoNeto || 0), 0);
                resultado.totales.montoExento = resultado.meses.reduce((sum, m) => sum + (m.montoExento || 0), 0);
                resultado.totales.montoIva = resultado.meses.reduce((sum, m) => sum + (m.montoIva || 0), 0);
                resultado.totales.montoTotal = resultado.meses.reduce((sum, m) => sum + (m.montoTotal || 0), 0);
            }

            return resultado;
        });

        return datos;
    }

    async cerrarSesion() {
        try {
            // Buscar y clickear cerrar sesión
            await this.page.evaluate(() => {
                const links = document.querySelectorAll('a');
                for (const link of links) {
                    if (link.textContent.toLowerCase().includes('cerrar sesión') ||
                        link.textContent.toLowerCase().includes('salir')) {
                        link.click();
                        return true;
                    }
                }
                return false;
            });
            await this.wait(2000);
        } catch (e) {
            // Ignorar errores al cerrar sesión
        }
    }
}

module.exports = SIIBot;
