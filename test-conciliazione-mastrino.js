/**
 * Test Script - Conciliazione Mastrino v2
 * Con raggruppamento commissioni e tolleranza
 */

const XLSX = require('xlsx');

// === CONFIGURAZIONE ===
const FILE_PATH = 'ANTICIPI A FORNITORI.xlsx';
const COD_ANTICIPO = 513;
const COD_SALDO = 27;
const TOLLERANZA = 1.00; // Tolleranza in euro per i match

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

function matchConTolleranza(importo1, importo2, tolleranza = TOLLERANZA) {
    return Math.abs(importo1 - importo2) <= tolleranza;
}

// === CARICAMENTO DATI ===
function caricaDati() {
    const wb = XLSX.readFile(FILE_PATH);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const anticipi = [];
    const saldiRaw = [];

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
                matchedWith: [],
                tipoMatch: null
            });
        } else if (cod === COD_SALDO && row[IDX.Avere]) {
            saldiRaw.push({
                id: i,
                riga: i + 1,
                data: excelDateToJS(row[IDX.DataReg]),
                dataSerial: row[IDX.DataReg],
                numDoc: (row[IDX.NumDoc] || '').toString().trim(),
                importo: row[IDX.Avere],
            });
        }
    }

    return { anticipi, saldiRaw };
}

// === STEP 0: Raggruppa saldi per NumDoc e Data ===
function step0_RaggruppaSaldi(saldiRaw) {
    console.log('\n' + '='.repeat(60));
    console.log('STEP 0: Raggruppamento saldi per NumDoc/Data');
    console.log('='.repeat(60));

    const gruppi = {};

    for (const saldo of saldiRaw) {
        // Chiave: data + numDoc
        const key = `${saldo.dataSerial}_${saldo.numDoc}`;

        if (!gruppi[key]) {
            gruppi[key] = {
                id: saldo.id,
                riga: saldo.riga,
                data: saldo.data,
                dataSerial: saldo.dataSerial,
                numDoc: saldo.numDoc,
                importo: 0,
                componenti: [],
                conciliato: false,
                matchedWith: []
            };
        }

        gruppi[key].importo += saldo.importo;
        gruppi[key].componenti.push({
            riga: saldo.riga,
            importo: saldo.importo
        });
    }

    const saldi = Object.values(gruppi);

    // Mostra raggruppamenti con più componenti
    const raggruppati = saldi.filter(s => s.componenti.length > 1);
    console.log(`\nSaldi raggruppati (con commissioni): ${raggruppati.length}`);
    raggruppati.forEach(s => {
        console.log(`  Doc ${s.numDoc} del ${formatDate(s.data)}: €${formatImporto(s.importo)}`);
        s.componenti.forEach(c => {
            console.log(`    - Riga ${c.riga}: €${formatImporto(c.importo)}`);
        });
    });

    console.log(`\nTotale saldi dopo raggruppamento: ${saldi.length} (da ${saldiRaw.length} righe)`);

    return saldi;
}

// === STEP 1: Conciliazione 1-a-1 con tolleranza ===
function step1_UnoAUno(anticipi, saldi) {
    console.log('\n' + '='.repeat(60));
    console.log(`STEP 1: Conciliazione 1-a-1 (tolleranza ±€${TOLLERANZA})`);
    console.log('='.repeat(60));

    let matches = 0;
    let matchesConDiff = 0;

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        // Cerca saldo con importo simile e data successiva
        const saldo = saldi.find(s =>
            !s.conciliato &&
            matchConTolleranza(s.importo, ant.importo) &&
            s.dataSerial > ant.dataSerial
        );

        if (saldo) {
            const diff = Math.abs(saldo.importo - ant.importo);

            ant.conciliato = true;
            ant.matchedWith = [saldo.id];
            ant.tipoMatch = '1:1';
            ant.differenza = diff;
            saldo.conciliato = true;
            saldo.matchedWith = [ant.id];
            matches++;

            if (diff > 0.01) {
                matchesConDiff++;
                console.log(`✓ Match: Anticipo ${formatDate(ant.data)} ${ant.fornitore} €${formatImporto(ant.importo)}`);
                console.log(`         → Saldo ${formatDate(saldo.data)} Doc:${saldo.numDoc} €${formatImporto(saldo.importo)} (diff: €${formatImporto(diff)})`);
            }
        }
    }

    console.log(`\nRisultato Step 1: ${matches} conciliazioni 1-a-1`);
    if (matchesConDiff > 0) {
        console.log(`  (di cui ${matchesConDiff} con differenza > €0.01)`);
    }
    return matches;
}

// === STEP 2: Conciliazione 1-a-molti con tolleranza ===
function* combinazioni(arr, sommaTarget, dataMinima, tolleranza = TOLLERANZA, maxElements = 5) {
    const n = arr.length;

    function* genera(start, current, somma) {
        if (matchConTolleranza(somma, sommaTarget, tolleranza) && current.length > 0) {
            yield { elementi: [...current], somma, diff: Math.abs(somma - sommaTarget) };
        }
        if (somma > sommaTarget + tolleranza || current.length >= maxElements) return;

        for (let i = start; i < n; i++) {
            const el = arr[i];
            if (el.dataSerial > dataMinima && !el.conciliato) {
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
    console.log(`STEP 2: Conciliazione 1-a-molti (tolleranza ±€${TOLLERANZA})`);
    console.log('='.repeat(60));

    let matches = 0;

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        const saldiDisponibili = saldi.filter(s => !s.conciliato && s.dataSerial > ant.dataSerial);

        // Cerca la migliore combinazione (quella con minore differenza)
        let migliorCombo = null;

        for (const combo of combinazioni(saldiDisponibili, ant.importo, ant.dataSerial)) {
            if (!migliorCombo || combo.diff < migliorCombo.diff) {
                // Verifica che tutti siano ancora disponibili
                if (combo.elementi.every(s => !s.conciliato)) {
                    migliorCombo = combo;
                    if (combo.diff < 0.01) break; // Match perfetto trovato
                }
            }
        }

        if (migliorCombo && migliorCombo.elementi.length > 1) {
            ant.conciliato = true;
            ant.matchedWith = migliorCombo.elementi.map(s => s.id);
            ant.tipoMatch = `1:${migliorCombo.elementi.length}`;
            ant.differenza = migliorCombo.diff;

            migliorCombo.elementi.forEach(s => {
                s.conciliato = true;
                s.matchedWith.push(ant.id);
            });

            matches++;
            console.log(`✓ Match: Anticipo ${formatDate(ant.data)} ${ant.fornitore} €${formatImporto(ant.importo)}`);
            migliorCombo.elementi.forEach(s => {
                console.log(`         → Saldo ${formatDate(s.data)} Doc:${s.numDoc} €${formatImporto(s.importo)}`);
            });
            if (migliorCombo.diff > 0.01) {
                console.log(`         (diff: €${formatImporto(migliorCombo.diff)})`);
            }
        }
    }

    console.log(`\nRisultato Step 2: ${matches} conciliazioni 1-a-molti`);
    return matches;
}

// === STEP 3: Conciliazione molti-a-1 ===
function step3_MoltiAUno(anticipi, saldi) {
    console.log('\n' + '='.repeat(60));
    console.log(`STEP 3: Conciliazione molti-a-1 (stesso fornitore, tolleranza ±€${TOLLERANZA})`);
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
            if (fornitore === 'N/D') continue; // Salta fornitori sconosciuti per molti:1

            const antDisponibili = antList.filter(a => !a.conciliato && a.dataSerial < saldo.dataSerial);

            for (const combo of combinazioni(antDisponibili, saldo.importo, 0)) {
                if (combo.elementi.every(a => !a.conciliato && a.dataSerial < saldo.dataSerial)) {
                    saldo.conciliato = true;
                    saldo.matchedWith = combo.elementi.map(a => a.id);

                    combo.elementi.forEach(a => {
                        a.conciliato = true;
                        a.matchedWith.push(saldo.id);
                        a.tipoMatch = `${combo.elementi.length}:1`;
                        a.differenza = combo.diff / combo.elementi.length;
                    });

                    matches++;
                    console.log(`✓ Match: Saldo ${formatDate(saldo.data)} Doc:${saldo.numDoc} €${formatImporto(saldo.importo)}`);
                    console.log(`         Fornitore: ${fornitore}`);
                    combo.elementi.forEach(a => {
                        console.log(`         ← Anticipo ${formatDate(a.data)} €${formatImporto(a.importo)}`);
                    });
                    if (combo.diff > 0.01) {
                        console.log(`         (diff: €${formatImporto(combo.diff)})`);
                    }
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

    // Differenze totali
    const diffTotale = antConciliati.reduce((s, a) => s + (a.differenza || 0), 0);
    if (diffTotale > 0) {
        console.log(`\nDifferenza totale per tolleranza: €${formatImporto(diffTotale)}`);
    }

    if (antNonConciliati.length > 0) {
        console.log('\n--- ANTICIPI NON CONCILIATI ---');
        antNonConciliati.forEach(a => {
            console.log(`  Riga ${a.riga}: ${formatDate(a.data)} ${a.fornitore} €${formatImporto(a.importo)}`);
        });
    }

    if (salNonConciliati.length > 0) {
        console.log('\n--- SALDI NON CONCILIATI ---');
        salNonConciliati.forEach(s => {
            const comp = s.componenti.length > 1 ? ` (${s.componenti.length} comp.)` : '';
            console.log(`  Riga ${s.riga}: ${formatDate(s.data)} Doc:${s.numDoc} €${formatImporto(s.importo)}${comp}`);
        });
    }
}

// === MAIN ===
function main() {
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║   TEST CONCILIAZIONE MASTRINO v2 - CON TOLLERANZA          ║');
    console.log('╚════════════════════════════════════════════════════════════╝');

    const { anticipi, saldiRaw } = caricaDati();

    console.log(`\nDati caricati:`);
    console.log(`  - Anticipi a Fornitori (513): ${anticipi.length} movimenti`);
    console.log(`  - Saldi Fattura raw (27): ${saldiRaw.length} righe`);
    console.log(`  - Totale Dare: €${formatImporto(anticipi.reduce((s,a) => s + a.importo, 0))}`);
    console.log(`  - Totale Avere: €${formatImporto(saldiRaw.reduce((s,a) => s + a.importo, 0))}`);

    // Step 0: Raggruppa saldi
    const saldi = step0_RaggruppaSaldi(saldiRaw);

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
