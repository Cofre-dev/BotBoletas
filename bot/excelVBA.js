/**
 * excelVBA.js
 * Inyecta un vbaProject.bin en un buffer XLSX convirtiéndolo en XLSM.
 * Requiere que exista templates/vba_template.xlsm (generado por scripts/crear-template-vba.ps1).
 */

const fs   = require('fs');
const path = require('path');

const TEMPLATE_PATH = path.join(__dirname, '..', 'templates', 'vba_template.xlsm');

/**
 * @param {Buffer} xlsxBuffer
 * @returns {Promise<Buffer|null>}
 */
async function injectVBA(xlsxBuffer) {
    if (!fs.existsSync(TEMPLATE_PATH)) {
        console.log('[VBA] Template no encontrado, exportando sin macros. Ejecuta scripts/crear-template-vba.bat para habilitar macros.');
        return null;
    }

    try {
        const JSZip = require('jszip');

        // ── 1. Extraer vbaProject.bin del template ────────────────
        const templateZip = await JSZip.loadAsync(fs.readFileSync(TEMPLATE_PATH));
        const vbaEntry    = templateZip.file('xl/vbaProject.bin');

        if (!vbaEntry) {
            console.warn('[VBA] El template no contiene xl/vbaProject.bin.');
            return null;
        }

        const vbaBin = await vbaEntry.async('nodebuffer');

        // ── 2. Abrir el XLSX de destino ───────────────────────────
        const zip = await JSZip.loadAsync(xlsxBuffer);

        // ── 3. Inyectar el binario VBA ────────────────────────────
        zip.file('xl/vbaProject.bin', vbaBin);

        // ── 4. Parchear [Content_Types].xml ──────────────────────
        const ctFile = zip.file('[Content_Types].xml');
        let ctXml = await ctFile.async('text');

        if (!ctXml.includes('vbaProject')) {
            ctXml = ctXml.replace(
                '</Types>',
                '<Override PartName="/xl/vbaProject.bin"' +
                ' ContentType="application/vnd.ms-office.vbaProject"/></Types>'
            );
            zip.file('[Content_Types].xml', ctXml);
        }

        // ── 5. Parchear xl/_rels/workbook.xml.rels ────────────────
        const relsFile = zip.file('xl/_rels/workbook.xml.rels');
        let relsXml = await relsFile.async('text');

        if (!relsXml.includes('vbaProject')) {
            relsXml = relsXml.replace(
                '</Relationships>',
                '<Relationship Id="rIdVBA"' +
                ' Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"' +
                ' Target="vbaProject.bin"/></Relationships>'
            );
            zip.file('xl/_rels/workbook.xml.rels', relsXml);
        }

        // ── 6. Generar buffer XLSM ────────────────────────────────
        const result = await zip.generateAsync({
            type:        'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
        });

        return result;

    } catch (err) {
        console.error('[VBA] Error al inyectar macros:', err.message);
        return null;
    }
}

module.exports = { injectVBA, TEMPLATE_PATH };
