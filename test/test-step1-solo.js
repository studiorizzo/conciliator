/**
 * Test Script - Conciliazione Mastrino - Solo Step 1 (1:1)
 * Match esatto senza raggruppamento commissioni
 */

const XLSX = require('xlsx');
const path = require('path');

// === CONFIGURAZIONE ===
const FILE_PATH = path.join(__dirname, 'ANTICIPI A FORNITORI.xlsx');
const OUTPUT_PATH = path.join(__dirname, 'residui-step1.xlsx');
const COD_ANTICIPO = 513;
const COD_SALDO = 27;

// Indici colonne
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
    if (!date) return 'N/A';
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
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const anticipi = [];
    const saldi = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const cod = row[IDX.CodCausale];

        if (cod === COD_ANTICIPO && row[IDX.Dare]) {
            anticipi.push({
                riga: i + 1,
                data: excelDateToJS(row[IDX.DataReg]),
                dataSerial: row[IDX.DataReg],
                numDoc: row[IDX.NumDoc] || '',
                fornitore: row[IDX.Des1Causale] || 'N/D',
                importo: row[IDX.Dare],
                conciliato: false,
                matchRiga: null
            });
        } else if (cod === COD_SALDO && row[IDX.Avere]) {
            saldi.push({
                riga: i + 1,
                data: excelDateToJS(row[IDX.DataReg]),
                dataSerial: row[IDX.DataReg],
                numDoc: (row[IDX.NumDoc] || '').toString().trim(),
                importo: row[IDX.Avere],
                conciliato: false,
                matchRiga: null
            });
        }
    }

    return { anticipi, saldi };
}

// === STEP 1: Conciliazione 1-a-1 esatta ===
function step1_UnoAUno(anticipi, saldi) {
    console.log('\n' + '='.repeat(60));
    console.log('STEP 1: Conciliazione 1-a-1 (importo esatto)');
    console.log('='.repeat(60));

    let matches = 0;
    const conciliazioni = [];

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        // Cerca saldo con stesso importo esatto e data successiva
        const saldo = saldi.find(s =>
            !s.conciliato &&
            Math.abs(s.importo - ant.importo) < 0.01 &&
            s.dataSerial > ant.dataSerial
        );

        if (saldo) {
            ant.conciliato = true;
            ant.matchRiga = saldo.riga;
            saldo.conciliato = true;
            saldo.matchRiga = ant.riga;
            matches++;

            conciliazioni.push({
                antRiga: ant.riga,
                antData: formatDate(ant.data),
                antFornitore: ant.fornitore,
                antImporto: ant.importo,
                salRiga: saldo.riga,
                salData: formatDate(saldo.data),
                salNumDoc: saldo.numDoc,
                salImporto: saldo.importo
            });

            console.log(`✓ Riga ${ant.riga} → Riga ${saldo.riga}: ${ant.fornitore} €${formatImporto(ant.importo)}`);
        }
    }

    console.log(`\nTotale conciliazioni 1:1: ${matches}`);
    return { matches, conciliazioni };
}

// === GENERA EXCEL RESIDUI ===
function generaExcelResidui(anticipi, saldi, conciliazioni) {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Anticipi non conciliati
    const anticipiNC = anticipi.filter(a => !a.conciliato).map(a => ({
        'Riga': a.riga,
        'Data': formatDate(a.data),
        'Fornitore': a.fornitore,
        'NumDoc': a.numDoc,
        'Importo': a.importo
    }));

    if (anticipiNC.length > 0) {
        const wsAnticpi = XLSX.utils.json_to_sheet(anticipiNC);
        XLSX.utils.book_append_sheet(wb, wsAnticpi, 'Anticipi Non Conciliati');
    }

    // Sheet 2: Saldi non conciliati
    const saldiNC = saldi.filter(s => !s.conciliato).map(s => ({
        'Riga': s.riga,
        'Data': formatDate(s.data),
        'NumDoc': s.numDoc,
        'Importo': s.importo
    }));

    if (saldiNC.length > 0) {
        const wsSaldi = XLSX.utils.json_to_sheet(saldiNC);
        XLSX.utils.book_append_sheet(wb, wsSaldi, 'Saldi Non Conciliati');
    }

    // Sheet 3: Conciliazioni effettuate
    const conc = conciliazioni.map(c => ({
        'Anticipo Riga': c.antRiga,
        'Anticipo Data': c.antData,
        'Fornitore': c.antFornitore,
        'Importo': c.antImporto,
        'Saldo Riga': c.salRiga,
        'Saldo Data': c.salData,
        'Saldo NumDoc': c.salNumDoc
    }));

    if (conc.length > 0) {
        const wsConc = XLSX.utils.json_to_sheet(conc);
        XLSX.utils.book_append_sheet(wb, wsConc, 'Conciliazioni 1-1');
    }

    // Sheet 4: Riepilogo
    const riepilogo = [
        ['RIEPILOGO CONCILIAZIONE STEP 1'],
        [],
        ['Descrizione', 'Quantità', 'Importo'],
        ['Anticipi totali', anticipi.length, anticipi.reduce((s, a) => s + a.importo, 0)],
        ['Anticipi conciliati', anticipi.filter(a => a.conciliato).length, anticipi.filter(a => a.conciliato).reduce((s, a) => s + a.importo, 0)],
        ['Anticipi residui', anticipiNC.length, anticipiNC.reduce((s, a) => s + a.Importo, 0)],
        [],
        ['Saldi totali', saldi.length, saldi.reduce((s, a) => s + a.importo, 0)],
        ['Saldi conciliati', saldi.filter(s => s.conciliato).length, saldi.filter(s => s.conciliato).reduce((s, a) => s + a.importo, 0)],
        ['Saldi residui', saldiNC.length, saldiNC.reduce((s, a) => s + a.Importo, 0)]
    ];
    const wsRiep = XLSX.utils.aoa_to_sheet(riepilogo);
    XLSX.utils.book_append_sheet(wb, wsRiep, 'Riepilogo');

    XLSX.writeFile(wb, OUTPUT_PATH);
    console.log(`\n✅ Excel generato: ${OUTPUT_PATH}`);
}

// === MAIN ===
function main() {
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║      TEST CONCILIAZIONE MASTRINO - SOLO STEP 1 (1:1)       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');

    const { anticipi, saldi } = caricaDati();

    console.log(`\nDati caricati:`);
    console.log(`  - Anticipi a Fornitori (513): ${anticipi.length}`);
    console.log(`  - Saldi Fattura (27): ${saldi.length}`);
    console.log(`  - Totale Dare: €${formatImporto(anticipi.reduce((s, a) => s + a.importo, 0))}`);
    console.log(`  - Totale Avere: €${formatImporto(saldi.reduce((s, a) => s + a.importo, 0))}`);

    // Esegui Step 1
    const { matches, conciliazioni } = step1_UnoAUno(anticipi, saldi);

    // Riepilogo
    console.log('\n' + '='.repeat(60));
    console.log('RIEPILOGO');
    console.log('='.repeat(60));

    const antNC = anticipi.filter(a => !a.conciliato);
    const salNC = saldi.filter(s => !s.conciliato);

    console.log(`\nANTICIPI:`);
    console.log(`  Conciliati: ${matches} (€${formatImporto(anticipi.filter(a => a.conciliato).reduce((s, a) => s + a.importo, 0))})`);
    console.log(`  Residui: ${antNC.length} (€${formatImporto(antNC.reduce((s, a) => s + a.importo, 0))})`);

    console.log(`\nSALDI:`);
    console.log(`  Conciliati: ${matches} (€${formatImporto(saldi.filter(s => s.conciliato).reduce((s, a) => s + a.importo, 0))})`);
    console.log(`  Residui: ${salNC.length} (€${formatImporto(salNC.reduce((s, a) => s + a.importo, 0))})`);

    // Genera Excel
    generaExcelResidui(anticipi, saldi, conciliazioni);

    console.log('\n' + '='.repeat(60));
    console.log('TEST COMPLETATO');
    console.log('='.repeat(60));
}

main();
