/**
 * Test Script - Conciliazione Mastrino - Solo Step 1 (1:1)
 * Match esatto senza raggruppamento commissioni
 * Output: unico foglio con movimenti residui (formato semplificato)
 */

const XLSX = require('xlsx');
const path = require('path');

// === CONFIGURAZIONE ===
const FILE_PATH = path.join(__dirname, 'ANTICIPI A FORNITORI.xlsx');
const OUTPUT_PATH = path.join(__dirname, 'residui-step1.xlsx');
const COD_ANTICIPO = 513;
const COD_SALDO = 27;

// Indici colonne originali
const IDX = {
    DataReg: 9,
    NumDoc: 10,
    CodCausale: 13,
    DesCausale: 14,
    Des1Causale: 15,
    Dare: 18,
    Avere: 19
};

// === UTILITY ===
function excelDateToJS(serial) {
    if (!serial) return null;
    const utc_days = Math.floor(serial - 25569);
    return new Date(utc_days * 86400 * 1000);
}

function formatDate(date) {
    if (!date) return '';
    return date.toLocaleDateString('it-IT');
}

function formatImporto(n) {
    if (n === null || n === undefined) return '0,00';
    return n.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// === CARICAMENTO DATI ===
function caricaDati() {
    const wb = XLSX.readFile(FILE_PATH);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(ws, { header: 1 });

    // Mappa tutti i movimenti rilevanti (513 e 27) con i dati semplificati
    const movimenti = [];

    for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        const cod = row[IDX.CodCausale];

        // Solo anticipi (513) e saldi (27)
        if (cod === COD_ANTICIPO || cod === COD_SALDO) {
            movimenti.push({
                rigaOriginale: i + 1,
                dataSerial: row[IDX.DataReg],
                data: excelDateToJS(row[IDX.DataReg]),
                numDoc: row[IDX.NumDoc] || '',
                codCausale: row[IDX.CodCausale],
                desCausale: row[IDX.DesCausale] || '',
                des1Causale: row[IDX.Des1Causale] || '',
                dare: row[IDX.Dare] || null,
                avere: row[IDX.Avere] || null,
                conciliato: false,
                matchRiga: null
            });
        }
    }

    return movimenti;
}

// === STEP 1: Conciliazione 1-a-1 esatta ===
function step1_UnoAUno(movimenti) {
    console.log('\n' + '='.repeat(60));
    console.log('STEP 1: Conciliazione 1-a-1 (importo esatto)');
    console.log('='.repeat(60));

    const anticipi = movimenti.filter(m => m.codCausale === COD_ANTICIPO && m.dare);
    const saldi = movimenti.filter(m => m.codCausale === COD_SALDO && m.avere);

    let matches = 0;

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        // Cerca saldo con stesso importo esatto e data successiva
        const saldo = saldi.find(s =>
            !s.conciliato &&
            Math.abs(s.avere - ant.dare) < 0.01 &&
            s.dataSerial > ant.dataSerial
        );

        if (saldo) {
            ant.conciliato = true;
            ant.matchRiga = saldo.rigaOriginale;
            saldo.conciliato = true;
            saldo.matchRiga = ant.rigaOriginale;
            matches++;

            console.log(`✓ Riga ${ant.rigaOriginale} → Riga ${saldo.rigaOriginale}: ${ant.des1Causale || 'N/D'} €${formatImporto(ant.dare)}`);
        }
    }

    console.log(`\nTotale conciliazioni 1:1: ${matches}`);
    return matches;
}

// === GENERA EXCEL RESIDUI ===
function generaExcelResidui(movimenti) {
    const wb = XLSX.utils.book_new();

    // Filtra solo i non conciliati e ordina per data
    const residui = movimenti
        .filter(m => !m.conciliato)
        .sort((a, b) => (a.dataSerial || 0) - (b.dataSerial || 0));

    // Crea array di dati con header
    const data = [
        ['DataReg', 'NumDoc', 'CodCausale', 'DesCausale', 'Des1Causale', 'Dare', 'Avere']
    ];

    for (const m of residui) {
        data.push([
            formatDate(m.data),
            m.numDoc,
            m.codCausale,
            m.desCausale,
            m.des1Causale,
            m.dare || '',
            m.avere || ''
        ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Residui');

    XLSX.writeFile(wb, OUTPUT_PATH);
    console.log(`\n✅ Excel generato: ${OUTPUT_PATH}`);
    console.log(`   ${residui.length} movimenti residui`);
}

// === MAIN ===
function main() {
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║      TEST CONCILIAZIONE MASTRINO - SOLO STEP 1 (1:1)       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');

    const movimenti = caricaDati();

    const anticipi = movimenti.filter(m => m.codCausale === COD_ANTICIPO);
    const saldi = movimenti.filter(m => m.codCausale === COD_SALDO);

    console.log(`\nDati caricati:`);
    console.log(`  - Anticipi a Fornitori (513): ${anticipi.length}`);
    console.log(`  - Saldi Fattura (27): ${saldi.length}`);
    console.log(`  - Totale Dare: €${formatImporto(anticipi.reduce((s, a) => s + (a.dare || 0), 0))}`);
    console.log(`  - Totale Avere: €${formatImporto(saldi.reduce((s, a) => s + (a.avere || 0), 0))}`);

    // Esegui Step 1
    const matches = step1_UnoAUno(movimenti);

    // Riepilogo
    console.log('\n' + '='.repeat(60));
    console.log('RIEPILOGO');
    console.log('='.repeat(60));

    const antConc = anticipi.filter(a => a.conciliato);
    const antNC = anticipi.filter(a => !a.conciliato);
    const salConc = saldi.filter(s => s.conciliato);
    const salNC = saldi.filter(s => !s.conciliato);

    console.log(`\nANTICIPI (Dare):`);
    console.log(`  Conciliati: ${antConc.length} (€${formatImporto(antConc.reduce((s, a) => s + (a.dare || 0), 0))})`);
    console.log(`  Residui: ${antNC.length} (€${formatImporto(antNC.reduce((s, a) => s + (a.dare || 0), 0))})`);

    console.log(`\nSALDI (Avere):`);
    console.log(`  Conciliati: ${salConc.length} (€${formatImporto(salConc.reduce((s, a) => s + (a.avere || 0), 0))})`);
    console.log(`  Residui: ${salNC.length} (€${formatImporto(salNC.reduce((s, a) => s + (a.avere || 0), 0))})`);

    // Genera Excel
    generaExcelResidui(movimenti);

    console.log('\n' + '='.repeat(60));
    console.log('TEST COMPLETATO');
    console.log('='.repeat(60));
}

main();
