/**
 * Test Script - Conciliazione Mastrino
 * Logica di riconciliazione tra Anticipi a Fornitori (Dare) e Saldi Fattura (Avere)
 */

const XLSX = require('xlsx');

// === CONFIGURAZIONE ===
const FILE_PATH = 'ANTICIPI A FORNITORI.xlsx';
const COD_ANTICIPO = 513;  // Anticipo a Fornitori (Dare)
const COD_SALDO = 27;      // Saldo Fattura (Avere)

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

    const anticipi = [];  // Dare (513)
    const saldi = [];     // Avere (27)

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const cod = row[IDX.CodCausale];

        if (cod === COD_ANTICIPO && row[IDX.Dare]) {
            anticipi.push({
                id: i,
                riga: i + 1,
                data: excelDateToJS(row[IDX.DataReg]),
                dataSerial: row[IDX.DataReg],
                numDoc: row[IDX.NumDoc],
                fornitore: row[IDX.Des1Causale] || 'N/D',
                importo: row[IDX.Dare],
                conciliato: false,
                matchedWith: []
            });
        } else if (cod === COD_SALDO && row[IDX.Avere]) {
            saldi.push({
                id: i,
                riga: i + 1,
                data: excelDateToJS(row[IDX.DataReg]),
                dataSerial: row[IDX.DataReg],
                numDoc: row[IDX.NumDoc],
                importo: row[IDX.Avere],
                conciliato: false,
                matchedWith: []
            });
        }
    }

    return { anticipi, saldi };
}

// === STEP 1: Conciliazione 1-a-1 ===
function step1_UnoAUno(anticipi, saldi) {
    console.log('\n' + '='.repeat(60));
    console.log('STEP 1: Conciliazione 1-a-1 (importo esatto)');
    console.log('='.repeat(60));

    let matches = 0;

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        // Cerca saldo con stesso importo e data successiva
        const saldo = saldi.find(s =>
            !s.conciliato &&
            Math.abs(s.importo - ant.importo) < 0.01 &&
            s.dataSerial > ant.dataSerial
        );

        if (saldo) {
            ant.conciliato = true;
            ant.matchedWith = [saldo.id];
            ant.tipoMatch = '1:1';
            saldo.conciliato = true;
            saldo.matchedWith = [ant.id];
            matches++;

            console.log(`✓ Match: Anticipo ${formatDate(ant.data)} ${ant.fornitore} €${formatImporto(ant.importo)}`);
            console.log(`         → Saldo ${formatDate(saldo.data)} Doc:${saldo.numDoc} €${formatImporto(saldo.importo)}`);
        }
    }

    console.log(`\nRisultato Step 1: ${matches} conciliazioni 1-a-1`);
    return matches;
}

// === STEP 2: Conciliazione 1-a-molti ===
function* combinazioni(arr, sommaTarget, dataMinima, maxElements = 5) {
    // Genera combinazioni di elementi che sommano al target
    const n = arr.length;

    function* genera(start, current, somma) {
        if (Math.abs(somma - sommaTarget) < 0.01) {
            yield [...current];
            return;
        }
        if (somma > sommaTarget + 0.01 || current.length >= maxElements) return;

        for (let i = start; i < n; i++) {
            const el = arr[i];
            if (el.dataSerial > dataMinima) {
                current.push(el);
                yield* genera(i + 1, current, somma + el.importo);
                current.pop();
            }
        }
    }

    yield* genera(0, [], 0);
}

function step2_UnoAMolti(anticipi, saldi) {
    console.log('\n' + '='.repeat(60));
    console.log('STEP 2: Conciliazione 1-a-molti');
    console.log('='.repeat(60));

    let matches = 0;
    const saldiDisponibili = saldi.filter(s => !s.conciliato);

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        // Cerca combinazione di saldi che sommano all'anticipo
        for (const combo of combinazioni(saldiDisponibili, ant.importo, ant.dataSerial)) {
            // Verifica che tutti i saldi della combo siano ancora disponibili
            if (combo.every(s => !s.conciliato)) {
                ant.conciliato = true;
                ant.matchedWith = combo.map(s => s.id);
                ant.tipoMatch = `1:${combo.length}`;

                combo.forEach(s => {
                    s.conciliato = true;
                    s.matchedWith.push(ant.id);
                });

                matches++;
                console.log(`✓ Match: Anticipo ${formatDate(ant.data)} ${ant.fornitore} €${formatImporto(ant.importo)}`);
                combo.forEach(s => {
                    console.log(`         → Saldo ${formatDate(s.data)} Doc:${s.numDoc} €${formatImporto(s.importo)}`);
                });
                break;
            }
        }
    }

    console.log(`\nRisultato Step 2: ${matches} conciliazioni 1-a-molti`);
    return matches;
}

// === STEP 3: Conciliazione molti-a-1 ===
function step3_MoltiAUno(anticipi, saldi) {
    console.log('\n' + '='.repeat(60));
    console.log('STEP 3: Conciliazione molti-a-1 (stesso fornitore)');
    console.log('='.repeat(60));

    let matches = 0;
    const anticipiNonConciliati = anticipi.filter(a => !a.conciliato);
    const saldiNonConciliati = saldi.filter(s => !s.conciliato);

    // Raggruppa anticipi per fornitore
    const perFornitore = {};
    for (const ant of anticipiNonConciliati) {
        if (!perFornitore[ant.fornitore]) {
            perFornitore[ant.fornitore] = [];
        }
        perFornitore[ant.fornitore].push(ant);
    }

    for (const saldo of saldiNonConciliati) {
        if (saldo.conciliato) continue;

        // Per ogni fornitore, cerca combinazione di anticipi che sommano al saldo
        for (const [fornitore, antList] of Object.entries(perFornitore)) {
            const antDisponibili = antList.filter(a => !a.conciliato && a.dataSerial < saldo.dataSerial);

            for (const combo of combinazioni(antDisponibili, saldo.importo, 0)) {
                // Tutti gli anticipi devono avere data < saldo
                if (combo.every(a => !a.conciliato && a.dataSerial < saldo.dataSerial)) {
                    saldo.conciliato = true;
                    saldo.matchedWith = combo.map(a => a.id);
                    saldo.tipoMatch = `${combo.length}:1`;

                    combo.forEach(a => {
                        a.conciliato = true;
                        a.matchedWith.push(saldo.id);
                        a.tipoMatch = `${combo.length}:1`;
                    });

                    matches++;
                    console.log(`✓ Match: Saldo ${formatDate(saldo.data)} Doc:${saldo.numDoc} €${formatImporto(saldo.importo)}`);
                    console.log(`         Fornitore: ${fornitore}`);
                    combo.forEach(a => {
                        console.log(`         ← Anticipo ${formatDate(a.data)} €${formatImporto(a.importo)}`);
                    });
                    break;
                }
            }
            if (saldo.conciliato) break;
        }
    }

    console.log(`\nRisultato Step 3: ${matches} conciliazioni molti-a-1`);
    return matches;
}

// === RIEPILOGO ===
function stampaRiepilogo(anticipi, saldi) {
    console.log('\n' + '='.repeat(60));
    console.log('RIEPILOGO FINALE');
    console.log('='.repeat(60));

    const antConciliati = anticipi.filter(a => a.conciliato);
    const antNonConciliati = anticipi.filter(a => !a.conciliato);
    const salConciliati = saldi.filter(s => s.conciliato);
    const salNonConciliati = saldi.filter(s => !s.conciliato);

    console.log('\nANTICIPI A FORNITORI (Dare):');
    console.log(`  Totali: ${anticipi.length}`);
    console.log(`  Conciliati: ${antConciliati.length} (€${formatImporto(antConciliati.reduce((s,a) => s + a.importo, 0))})`);
    console.log(`  Non conciliati: ${antNonConciliati.length} (€${formatImporto(antNonConciliati.reduce((s,a) => s + a.importo, 0))})`);

    console.log('\nSALDI FATTURA (Avere):');
    console.log(`  Totali: ${saldi.length}`);
    console.log(`  Conciliati: ${salConciliati.length} (€${formatImporto(salConciliati.reduce((s,a) => s + a.importo, 0))})`);
    console.log(`  Non conciliati: ${salNonConciliati.length} (€${formatImporto(salNonConciliati.reduce((s,a) => s + a.importo, 0))})`);

    // Dettaglio tipi di match
    const tipiMatch = {};
    antConciliati.forEach(a => {
        tipiMatch[a.tipoMatch] = (tipiMatch[a.tipoMatch] || 0) + 1;
    });
    console.log('\nTipi di conciliazione:');
    Object.entries(tipiMatch).sort().forEach(([tipo, count]) => {
        console.log(`  ${tipo}: ${count}`);
    });

    if (antNonConciliati.length > 0) {
        console.log('\n--- ANTICIPI NON CONCILIATI ---');
        antNonConciliati.forEach(a => {
            console.log(`  Riga ${a.riga}: ${formatDate(a.data)} ${a.fornitore} €${formatImporto(a.importo)}`);
        });
    }

    if (salNonConciliati.length > 0) {
        console.log('\n--- SALDI NON CONCILIATI ---');
        salNonConciliati.forEach(s => {
            console.log(`  Riga ${s.riga}: ${formatDate(s.data)} Doc:${s.numDoc} €${formatImporto(s.importo)}`);
        });
    }
}

// === MAIN ===
function main() {
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║     TEST CONCILIAZIONE MASTRINO - ANTICIPI A FORNITORI     ║');
    console.log('╚════════════════════════════════════════════════════════════╝');

    const { anticipi, saldi } = caricaDati();

    console.log(`\nDati caricati:`);
    console.log(`  - Anticipi a Fornitori (513): ${anticipi.length} movimenti`);
    console.log(`  - Saldi Fattura (27): ${saldi.length} movimenti`);
    console.log(`  - Totale Dare: €${formatImporto(anticipi.reduce((s,a) => s + a.importo, 0))}`);
    console.log(`  - Totale Avere: €${formatImporto(saldi.reduce((s,a) => s + a.importo, 0))}`);

    // Esegui i 3 step
    const match1 = step1_UnoAUno(anticipi, saldi);
    const match2 = step2_UnoAMolti(anticipi, saldi);
    const match3 = step3_MoltiAUno(anticipi, saldi);

    stampaRiepilogo(anticipi, saldi);

    console.log('\n' + '='.repeat(60));
    console.log('TEST COMPLETATO');
    console.log('='.repeat(60));
}

main();
