// === CONCILIATOR MASTRINO - JavaScript ===
// Riconciliazione interna delle voci Dare e Avere nel mastrino contabile

// === GESTIONE SIDEBAR UPLOAD ===
const uploadFab = document.getElementById('upload-fab');
const uploadSidebar = document.getElementById('upload-sidebar');
const downloadFab = document.getElementById('download-fab');
const welcomeScreen = document.getElementById('welcome-screen');
const loadingOverlay = document.getElementById('loading-overlay');

function mostraLoading(title = 'Conciliazione in corso...', subtitle = 'Analisi dei movimenti e ricerca corrispondenze') {
    const loadingTitle = document.getElementById('loading-title');
    const loadingSubtitle = document.getElementById('loading-subtitle');
    if (loadingTitle) loadingTitle.textContent = title;
    if (loadingSubtitle) loadingSubtitle.textContent = subtitle;
    loadingOverlay.classList.add('active');
}

function nascondiLoading() {
    loadingOverlay.classList.remove('active');
}

uploadFab.addEventListener('click', () => {
    const isOpen = uploadSidebar.classList.contains('open');
    if (isOpen) {
        closeUploadSidebar();
    } else {
        openUploadSidebar();
    }
});

function openUploadSidebar() {
    closeConfigSidebar();
    uploadSidebar.classList.add('open');
    overlay.classList.add('active');
    uploadFab.classList.add('active');
    fab.classList.add('hidden');
    if (downloadFab.classList.contains('visible')) {
        downloadFab.classList.add('hidden');
    }
}

function closeUploadSidebar() {
    uploadSidebar.classList.remove('open');
    overlay.classList.remove('active');
    uploadFab.classList.remove('active');
    fab.classList.remove('hidden');
    if (downloadFab.classList.contains('visible')) {
        downloadFab.classList.remove('hidden');
    }
}

downloadFab.addEventListener('click', () => {
    if (risultati) {
        generaExcel();
    }
});

function mostraDownloadFab() {
    downloadFab.classList.add('visible');
}

function nascondiDownloadFab() {
    downloadFab.classList.remove('visible');
}

function conciliaFromUpload() {
    if (!mastrino) {
        alert('Devi prima caricare il file mastrino!');
        return;
    }

    updateCommissioniConfig();
    closeUploadSidebar();

    if (welcomeScreen) {
        welcomeScreen.style.display = 'none';
    }

    avviaConciliazioneAutomatica();

    setTimeout(() => {
        const btnConcilia = document.getElementById('btn-concilia-upload');
        const btnReload = document.getElementById('btn-reload-upload');

        if (btnConcilia && hasConciliatoUnaVolta) {
            btnConcilia.disabled = true;
            btnConcilia.style.opacity = '0.4';
            btnConcilia.style.cursor = 'not-allowed';
            btnConcilia.title = 'Prima conciliazione completata. Usa il reload per ricominciare oppure vai in Configurazione per riconciliare con nuovi parametri.';

            if (btnReload) {
                btnReload.classList.add('highlighted');
                btnReload.title = 'Clicca qui per ricaricare e ricominciare da capo';
            }
        }
    }, 1000);
}

function mostraConfigFab() {
    fab.classList.add('visible');
}

function resetPagina() {
    if (confirm('Sei sicuro di voler ricaricare la pagina? Tutti i dati andranno persi.')) {
        location.reload();
    }
}

function confermaUscita() {
    // Se ci sono dati processati, chiedi conferma
    if (hasConciliatoUnaVolta) {
        return confirm('Sei sicuro di voler uscire? Tutti i dati andranno persi.');
    }
    return true;
}

// === GESTIONE SIDEBAR CONFIGURAZIONE ===
const fab = document.getElementById('fab');
const sidebar = document.getElementById('sidebar');
const overlay = document.getElementById('overlay');

let sliderValues = {
    importo: 0.01
};

fab.addEventListener('click', () => {
    const isOpen = sidebar.classList.contains('open');
    if (isOpen) {
        closeConfigSidebar();
    } else {
        openConfigSidebar();
    }
});

overlay.addEventListener('click', () => {
    closeConfigSidebar();
    closeUploadSidebar();
});

function openConfigSidebar() {
    closeUploadSidebar();
    sidebar.classList.add('open');
    overlay.classList.add('active');
    fab.classList.add('active');
    uploadFab.classList.add('hidden');
    if (downloadFab.classList.contains('visible')) {
        downloadFab.classList.add('hidden');
    }
}

function closeConfigSidebar() {
    sidebar.classList.remove('open');
    overlay.classList.remove('active');
    fab.classList.remove('active');
    uploadFab.classList.remove('hidden');
    if (downloadFab.classList.contains('visible')) {
        downloadFab.classList.remove('hidden');
    }
}

function closeSidebar() {
    closeConfigSidebar();
}

let isDragging = false;
let currentSlider = null;

function setSliderValue(event, id, min, max) {
    const track = event.currentTarget;
    const rect = track.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const percentage = Math.max(0, Math.min(100, (x / rect.width) * 100));
    updateSlider(id, percentage, min, max);
}

document.querySelectorAll('.slider-thumb').forEach(thumb => {
    thumb.addEventListener('mousedown', function(e) {
        e.preventDefault();
        e.stopPropagation();

        isDragging = true;
        const fill = this.parentElement;
        const track = fill.parentElement;
        const sliderId = fill.id.replace('fill-', '');

        let min, max;
        if (sliderId === 'importo') {
            min = 0.01; max = 1;
        }

        currentSlider = { id: sliderId, track: track, min: min, max: max };
        document.body.style.cursor = 'grabbing';
    });
});

document.addEventListener('mousemove', function(e) {
    if (!isDragging || !currentSlider) return;

    const rect = currentSlider.track.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const percentage = Math.max(0, Math.min(100, (x / rect.width) * 100));
    updateSlider(currentSlider.id, percentage, currentSlider.min, currentSlider.max);
});

document.addEventListener('mouseup', function() {
    if (isDragging) {
        isDragging = false;
        currentSlider = null;
        document.body.style.cursor = '';
    }
});

function updateSlider(id, percentage, min, max) {
    const fill = document.getElementById(`fill-${id}`);
    const valueEl = document.getElementById(`val-${id}`);

    fill.style.width = percentage + '%';

    const value = min + (max - min) * (percentage / 100);
    let displayValue = value.toFixed(2);
    sliderValues[id] = parseFloat(displayValue);

    if (id === 'importo') {
        valueEl.textContent = '\u20AC ' + displayValue;
    }
}

function addCommissionSidebar() {
    const input = document.getElementById('new-commission');
    const value = parseFloat(input.value);

    if (!value || value <= 0) {
        alert('Inserisci un valore valido maggiore di 0');
        return;
    }

    const list = document.getElementById('commission-list');
    const existingItems = Array.from(list.querySelectorAll('.commission-item'));

    for (let item of existingItems) {
        const existingValue = parseFloat(item.querySelector('label').textContent.replace('\u20AC ', ''));
        if (existingValue === value) {
            alert('Commissione gia presente');
            return;
        }
    }

    const item = document.createElement('div');
    item.className = 'commission-item';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = 'comm-' + value.toString().replace('.', '');
    checkbox.checked = true;
    checkbox.onchange = updateCommissioniConfig;

    const label = document.createElement('label');
    label.htmlFor = checkbox.id;
    label.textContent = '\u20AC ' + value.toFixed(2);

    item.appendChild(checkbox);
    item.appendChild(label);

    let inserted = false;
    for (let i = 0; i < existingItems.length; i++) {
        const existingValue = parseFloat(existingItems[i].querySelector('label').textContent.replace('\u20AC ', ''));
        if (value < existingValue) {
            list.insertBefore(item, existingItems[i]);
            inserted = true;
            break;
        }
    }

    if (!inserted) {
        list.appendChild(item);
    }

    input.value = '';
}

function updateCommissioniConfig() {
    COMMISSIONI_ERROR = Array.from(document.querySelectorAll('#commission-list input:checked')).map(c => {
        const text = c.nextElementSibling.textContent.replace('\u20AC ', '');
        return parseFloat(text);
    });
}

function resetConfig() {
    updateSlider('importo', 0, 0.01, 1);

    document.querySelectorAll('#commission-list input').forEach(c => {
        if (c.id === 'comm-075' || c.id === 'comm-175') {
            c.checked = true;
        } else {
            c.closest('.commission-item').remove();
        }
    });

    updateCommissioniConfig();
}

function applicaConfigESalva() {
    if (!mastrino) {
        alert('Devi prima caricare il file mastrino!');
        return;
    }

    updateCommissioniConfig();
    closeConfigSidebar();
    avviaConciliazioneAutomatica();
}

// === STATO GLOBALE ===
let mastrino = null;
let risultati = null;
let hasConciliatoUnaVolta = false;

// Commissioni che vanno segnalate come ERROR (movimenti bancari non pertinenti al mastrino)
let COMMISSIONI_ERROR = [0.75, 1.75];

// Indici colonne del mastrino
const IDX = {
    DataReg: 9,
    NumDoc: 10,
    CodCausale: 13,
    DesCausale: 14,
    Des1Causale: 15,
    Dare: 18,
    Avere: 19
};

// Codici causale
const COD_BILANCIO_APERTURA = 25;
const COD_ANTICIPO = 513;
const COD_SALDO = 27;

console.log('Conciliator Mastrino - Inizializzazione');

// === UTILITY FUNCTIONS ===

function excelDateToJS(serial) {
    if (!serial) return null;
    if (serial instanceof Date) return serial;
    if (typeof serial === 'number') {
        const utc_days = Math.floor(serial - 25569);
        return new Date(utc_days * 86400 * 1000);
    }
    return null;
}

function formatData(date) {
    if (!date) return '';
    const g = date.getDate().toString().padStart(2, '0');
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const a = date.getFullYear();
    return `${g}/${m}/${a}`;
}

function formatImportoItaliano(value) {
    if (value === null || value === undefined) return '';
    return value.toLocaleString('it-IT', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function isCommissioneError(importo, tolleranza = 0.01) {
    return COMMISSIONI_ERROR.some(comm => {
        const diff = Math.abs(importo - comm);
        return diff <= tolleranza;
    });
}

// Normalizzazione descrizione causale
function normalizzaDescrizione(desCausale, des1Causale) {
    const desc = (desCausale || '').toUpperCase().trim();
    const desc1 = (des1Causale || '').trim();

    // Mapping descrizioni
    if (desc.includes('BILANC') && desc.includes('APERTURA')) {
        return 'Bilancio di Apertura' + (desc1 ? ' ' + desc1 : '');
    }

    if (desc.includes('ANTICI') && (desc.includes('FORN') || desc.includes('FONR'))) {
        return 'Anticipo a Fornitori' + (desc1 ? ' - ' + desc1 : '');
    }

    if (desc.includes('S/DO') && desc.includes('FATTURA')) {
        return 'Saldo fattura' + (desc1 ? ' - ' + desc1 : '');
    }

    if (desc.includes('EROGAZ')) {
        return 'Erogazione' + (desc1 ? ' - ' + desc1 : '');
    }

    if (desc.toLowerCase().includes('pagamento')) {
        return 'Pagamento' + (desc1 ? ' - ' + desc1 : '');
    }

    if (desc.includes('CANONE DI AFFIT')) {
        return 'Canone di affitto' + (desc1 ? ' - ' + desc1 : '');
    }

    if (desc.includes('CANONE')) {
        return 'Canone' + (desc1 ? ' - ' + desc1 : '');
    }

    if (desc.includes('MOV.GENERICO') || desc.includes('MOV. GENERICO')) {
        return 'Movimento generico' + (desc1 ? ' - ' + desc1 : '');
    }

    // Default: ritorna originale
    return desc + (desc1 ? ' ' + desc1 : '');
}

// === FILE HANDLING ===

function setupDropzone(dropzoneId, inputId, infoId) {
    const dropzone = document.getElementById(dropzoneId);
    const input = document.getElementById(inputId);
    const info = document.getElementById(infoId);

    if (!dropzone || !input) return;

    dropzone.addEventListener('click', () => input.click());

    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('active');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('active');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('active');
        const file = e.dataTransfer.files[0];
        if (file) handleFile(file, dropzone, info);
    });

    input.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleFile(file, dropzone, info);
    });
}

function handleFile(file, dropzone, infoEl) {
    const inizioOverlay = Date.now();

    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        showAlert('Formato file non valido. Usa file Excel (.xlsx o .xls)', 'error');
        return;
    }

    mostraLoading('Lettura Mastrino in corso...', `Lettura del file ${file.name}`);

    setTimeout(() => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                mastrino = parseMastrino(workbook);
                infoEl.textContent = `\u2713 ${file.name} - ${mastrino.length} movimenti caricati`;

                infoEl.style.display = 'block';
                dropzone.classList.add('uploaded');

                const tempoTrascorso = Date.now() - inizioOverlay;
                const delayMinimo = 800;
                const ritardoNecessario = Math.max(0, delayMinimo - tempoTrascorso);

                setTimeout(() => {
                    nascondiLoading();
                    showAlert('File caricato con successo!', 'success');
                }, ritardoNecessario);

            } catch (error) {
                console.error('Errore parsing file:', error);
                nascondiLoading();
                showAlert('Errore nel parsing del file: ' + error.message, 'error');
            }
        };

        reader.readAsArrayBuffer(file);
    }, 20);
}

function parseMastrino(workbook) {
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    console.log('Parsing mastrino...');
    console.log('   Righe totali:', rawData.length);

    const movimenti = [];
    let progCounter = 1;

    for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];

        // Salta righe vuote
        if (!row || row.length === 0) continue;

        const dataReg = row[IDX.DataReg];
        const dare = row[IDX.Dare] || null;
        const avere = row[IDX.Avere] || null;

        // Deve avere almeno data e uno tra dare/avere
        if (!dataReg && !dare && !avere) continue;

        const codCausale = row[IDX.CodCausale] || null;
        const desCausale = row[IDX.DesCausale] || '';
        const des1Causale = row[IDX.Des1Causale] || '';

        movimenti.push({
            prog: `M${progCounter++}`,
            rigaOriginale: i + 1,
            dataSerial: dataReg,
            data: excelDateToJS(dataReg),
            numDoc: row[IDX.NumDoc] || '',
            codCausale: codCausale,
            desCausale: desCausale,
            des1Causale: des1Causale,
            descrizione: normalizzaDescrizione(desCausale, des1Causale),
            dare: dare,
            avere: avere,
            conciliato: false,
            isError: false,
            matchProgs: [],
            saldoRiconciliazione: null
        });
    }

    console.log('   Movimenti validi:', movimenti.length);
    return movimenti;
}

function showAlert(message, type = 'success') {
    const container = document.getElementById('alert-container');
    const alert = document.createElement('div');
    alert.className = `alert alert-${type}`;
    alert.textContent = message;
    container.appendChild(alert);

    setTimeout(() => alert.remove(), 5000);
}

// === FUNZIONE TROVA COMBINAZIONE (1-a-molti) ===

/**
 * Trova una combinazione di saldi che sommati uguagliano il target (entro tolleranza)
 * @param {Array} saldi - Array di oggetti saldo disponibili (non conciliati)
 * @param {number} target - Importo da raggiungere (il dare dell'anticipo)
 * @param {number} tolleranza - Tolleranza sull'importo (default 0.01)
 * @param {number} maxElementi - Numero massimo di elementi nella combinazione (default 5)
 * @returns {Array|null} - Array di saldi che formano la combinazione, o null se non trovata
 */
function trovaCombinazione(saldi, target, tolleranza = 0.01, maxElementi = 5) {
    // Ordina per importo decrescente per trovare prima le combinazioni più "dense"
    const saldiOrdinati = [...saldi].sort((a, b) => b.avere - a.avere);

    let risultato = null;

    // Funzione ricorsiva con backtracking
    function cerca(indice, combinazioneCorrente, sommaCorrente) {
        // Se abbiamo già trovato una combinazione, esci
        if (risultato) return;

        // Se la somma è nel range, abbiamo trovato!
        if (Math.abs(sommaCorrente - target) <= tolleranza && combinazioneCorrente.length >= 2) {
            risultato = [...combinazioneCorrente];
            return;
        }

        // Se abbiamo superato il target o raggiunto max elementi, backtrack
        if (sommaCorrente > target + tolleranza || combinazioneCorrente.length >= maxElementi) {
            return;
        }

        // Prova ad aggiungere altri saldi
        for (let i = indice; i < saldiOrdinati.length; i++) {
            const saldo = saldiOrdinati[i];

            // Salta se già nella combinazione o già conciliato
            if (saldo.conciliato) continue;

            combinazioneCorrente.push(saldo);
            cerca(i + 1, combinazioneCorrente, sommaCorrente + saldo.avere);
            combinazioneCorrente.pop();

            // Se trovato, esci dal loop
            if (risultato) return;
        }
    }

    cerca(0, [], 0);
    return risultato;
}

// === ALGORITMO CONCILIAZIONE ===

function concilia(mast, config) {
    console.log('==========================================');
    console.log('INIZIO CONCILIAZIONE MASTRINO');
    console.log('==========================================');
    console.log('Configurazione:', config);

    // Reset solo isError (dipende dalla config commissioni)
    // I record già conciliati vengono preservati
    mast.forEach(m => {
        m.isError = false;
        // Non tocchiamo: conciliato, matchProgs, saldoRiconciliazione
    });

    // Marca commissioni come ERROR
    console.log('\n--- Identificazione commissioni (ERROR) ---');
    let countError = 0;
    mast.forEach(m => {
        const importo = m.dare || m.avere || 0;
        if (isCommissioneError(importo, config.tolleranzaImporto)) {
            m.isError = true;
            countError++;
            console.log(`   ERROR: Riga ${m.rigaOriginale} - ${formatImportoItaliano(importo)} (commissione)`);
        }
    });
    console.log(`   Totale commissioni ERROR: ${countError}`);

    // Separa anticipi e saldi (escludi bilancio apertura e commissioni)
    const anticipi = mast.filter(m =>
        m.codCausale === COD_ANTICIPO &&
        m.dare &&
        !m.isError
    );

    const saldi = mast.filter(m =>
        m.codCausale === COD_SALDO &&
        m.avere &&
        !m.isError
    );

    console.log(`\nAnticip a Fornitori (513): ${anticipi.length}`);
    console.log(`Saldi Fattura (27): ${saldi.length}`);

    // STEP 1: Conciliazione automatica (tolleranza importo)
    console.log('\n========== STEP 1: CONCILIAZIONE AUTOMATICA ==========');
    console.log(`Tolleranza importo: ${config.tolleranzaImporto}`);

    // STEP 1a: Conciliazione 1-a-1
    console.log('\n--- STEP 1a: Conciliazione 1-a-1 ---');
    let matches = 0;

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        // Cerca saldo con stesso importo esatto e data successiva
        const saldo = saldi.find(s =>
            !s.conciliato &&
            Math.abs(s.avere - ant.dare) <= config.tolleranzaImporto &&
            s.dataSerial > ant.dataSerial
        );

        if (saldo) {
            ant.conciliato = true;
            ant.matchProgs = [saldo.prog];
            ant.saldoRiconciliazione = Math.round((ant.dare - saldo.avere) * 100) / 100;

            saldo.conciliato = true;
            saldo.matchProgs = [ant.prog];

            matches++;
            console.log(`   Match: ${ant.prog} (Riga ${ant.rigaOriginale}) -> ${saldo.prog} (Riga ${saldo.rigaOriginale}): ${formatImportoItaliano(ant.dare)}`);
        }
    }

    console.log(`   Totale conciliazioni 1-a-1: ${matches}`);

    // STEP 1b: Conciliazione 1-a-molti (1 Dare -> N Avere)
    console.log('\n--- STEP 1b: Conciliazione 1-a-molti ---');
    let matches1aN = 0;

    for (const ant of anticipi) {
        if (ant.conciliato) continue;

        // Saldi disponibili con data successiva all'anticipo
        const saldiDisponibili = saldi.filter(s =>
            !s.conciliato &&
            s.dataSerial > ant.dataSerial
        );

        if (saldiDisponibili.length < 2) continue; // Serve almeno 2 saldi per 1-a-molti

        // Cerca combinazione di saldi che sommano all'anticipo
        const combinazione = trovaCombinazione(saldiDisponibili, ant.dare, config.tolleranzaImporto);

        if (combinazione && combinazione.length >= 2) {
            // Concilia l'anticipo con i saldi trovati
            ant.conciliato = true;
            ant.matchProgs = combinazione.map(s => s.prog);

            const sommaAvere = combinazione.reduce((sum, s) => sum + s.avere, 0);
            ant.saldoRiconciliazione = Math.round((ant.dare - sommaAvere) * 100) / 100;

            combinazione.forEach(s => {
                s.conciliato = true;
                s.matchProgs = [ant.prog];
            });

            matches1aN++;
            console.log(`   Match 1-a-${combinazione.length}: ${ant.prog} (Riga ${ant.rigaOriginale}) -> [${combinazione.map(s => s.prog).join(', ')}]: ${formatImportoItaliano(ant.dare)}`);
        }
    }

    console.log(`\nTotale conciliazioni 1-a-molti: ${matches1aN}`);
    console.log(`Totale conciliazioni Step 1: ${matches + matches1aN}`);

    // Calcola statistiche
    const antConc = anticipi.filter(a => a.conciliato);
    const antNC = anticipi.filter(a => !a.conciliato);
    const salConc = saldi.filter(s => s.conciliato);
    const salNC = saldi.filter(s => !s.conciliato);

    const stats = {
        totaleMovimenti: mast.length,
        anticipiTotali: anticipi.length,
        anticipiConciliati: antConc.length,
        anticipiNonConciliati: antNC.length,
        importoAnticipiConciliati: antConc.reduce((s, a) => s + (a.dare || 0), 0),
        importoAnticipiNonConciliati: antNC.reduce((s, a) => s + (a.dare || 0), 0),
        saldiTotali: saldi.length,
        saldiConciliati: salConc.length,
        saldiNonConciliati: salNC.length,
        importoSaldiConciliati: salConc.reduce((s, a) => s + (a.avere || 0), 0),
        importoSaldiNonConciliati: salNC.reduce((s, a) => s + (a.avere || 0), 0),
        commissioni: countError
    };

    console.log('\n--- RIEPILOGO ---');
    console.log(`Anticipi conciliati: ${stats.anticipiConciliati}/${stats.anticipiTotali}`);
    console.log(`Saldi conciliati: ${stats.saldiConciliati}/${stats.saldiTotali}`);
    console.log(`Commissioni (ERROR): ${stats.commissioni}`);

    return { movimenti: mast, stats: stats };
}

// === AVVIO CONCILIAZIONE ===

function avviaConciliazioneAutomatica() {
    if (!mastrino) {
        showAlert('Nessun file caricato!', 'error');
        return;
    }

    mostraLoading('Conciliazione in corso...', 'Analisi dei movimenti e ricerca corrispondenze');

    const config = {
        tolleranzaImporto: sliderValues.importo
    };

    setTimeout(() => {
        try {
            risultati = concilia(mastrino, config);
            hasConciliatoUnaVolta = true;

            mostraRisultati(risultati);
            mostraConfigFab();
            mostraDownloadFab();

            nascondiLoading();
            showAlert('Conciliazione completata!', 'success');

        } catch (error) {
            console.error('Errore durante la conciliazione:', error);
            nascondiLoading();
            showAlert('Errore durante la conciliazione: ' + error.message, 'error');
        }
    }, 500);
}

// === VISUALIZZAZIONE RISULTATI ===

function mostraRisultati(ris) {
    const resultsSection = document.getElementById('results-section');
    resultsSection.style.display = 'block';

    // Popola summary
    const summaryGrid = document.getElementById('summaryGrid');
    summaryGrid.innerHTML = `
        <div class="summary-card">
            <h3>Anticipi Conciliati</h3>
            <div class="value">${ris.stats.anticipiConciliati} / ${ris.stats.anticipiTotali}</div>
            <span class="percentage">\u20AC ${formatImportoItaliano(ris.stats.importoAnticipiConciliati)}</span>
        </div>
        <div class="summary-card">
            <h3>Saldi Conciliati</h3>
            <div class="value">${ris.stats.saldiConciliati} / ${ris.stats.saldiTotali}</div>
            <span class="percentage">\u20AC ${formatImportoItaliano(ris.stats.importoSaldiConciliati)}</span>
        </div>
        <div class="summary-card">
            <h3>Anticipi Non Conciliati</h3>
            <div class="value">${ris.stats.anticipiNonConciliati}</div>
            <span class="percentage">\u20AC ${formatImportoItaliano(ris.stats.importoAnticipiNonConciliati)}</span>
        </div>
        <div class="summary-card">
            <h3>Saldi Non Conciliati</h3>
            <div class="value">${ris.stats.saldiNonConciliati}</div>
            <span class="percentage">\u20AC ${formatImportoItaliano(ris.stats.importoSaldiNonConciliati)}</span>
        </div>
        <div class="summary-card">
            <h3>Commissioni (Error)</h3>
            <div class="value">${ris.stats.commissioni}</div>
        </div>
    `;

    // Popola tabella
    popolaTabella(ris.movimenti);
}

function popolaTabella(movimenti) {
    const table = document.getElementById('table-mastrino');

    // Ordina movimenti secondo le regole specificate:
    // 1. Bilancio apertura prima
    // 2. Tutti i movimenti (dare, avere, error) ordinati per data
    //    MA gli avere conciliati "saltano la fila" e seguono il loro dare corrispondente

    const bilancio = movimenti.filter(m => m.codCausale === COD_BILANCIO_APERTURA);

    // Movimenti dare (inclusi error)
    const dareMovimenti = movimenti.filter(m =>
        m.dare &&
        m.codCausale !== COD_BILANCIO_APERTURA
    );

    // Movimenti avere NON conciliati (inclusi error) - quelli conciliati saltano la fila
    const avereNonConciliati = movimenti.filter(m =>
        m.avere &&
        m.codCausale !== COD_BILANCIO_APERTURA &&
        !m.conciliato
    );

    // Unisci dare e avere non conciliati e ordina per data
    const tuttiMovimentiOrdinati = [...dareMovimenti, ...avereNonConciliati]
        .sort((a, b) => (a.dataSerial || 0) - (b.dataSerial || 0));

    // Costruisci array ordinato
    const movimentiOrdinati = [];

    // 1. Bilancio apertura
    bilancio.forEach(m => movimentiOrdinati.push({ movimento: m, tipo: 'bilancio' }));

    // 2. Tutti i movimenti ordinati per data, con avere conciliati che seguono il dare
    tuttiMovimentiOrdinati.forEach(mov => {
        if (mov.dare) {
            // E' un movimento dare
            const tipo = mov.isError ? 'error' : 'dare';
            movimentiOrdinati.push({ movimento: mov, tipo: tipo });

            // Se conciliato, aggiungi avere corrispondenti subito dopo (saltano la fila)
            if (mov.conciliato && mov.matchProgs.length > 0) {
                mov.matchProgs.forEach(progAvere => {
                    const avereMatch = movimenti.find(m => m.prog === progAvere);
                    if (avereMatch) {
                        movimentiOrdinati.push({ movimento: avereMatch, tipo: 'avere-match', parentProg: mov.prog });
                    }
                });
            }
        } else {
            // E' un movimento avere non conciliato (segue l'ordine per data)
            const tipo = mov.isError ? 'error' : 'avere-non-conciliato';
            movimentiOrdinati.push({ movimento: mov, tipo: tipo });
        }
    });

    // Genera HTML tabella
    let html = `
        <thead>
            <tr>
                <th>Stato</th>
                <th>Data Reg.</th>
                <th>Num. Doc.</th>
                <th>Descrizione</th>
                <th>Dare</th>
                <th>Avere</th>
                <th>Saldo Riconciliazione</th>
            </tr>
        </thead>
        <tbody>
    `;

    movimentiOrdinati.forEach((item, index) => {
        const m = item.movimento;
        let rowClass = '';
        let statoHtml = '';
        let expandIcon = '';

        // Determina classe riga e stato
        if (item.tipo === 'bilancio') {
            rowClass = 'row-bilancio';
            statoHtml = '<span class="badge-stato" style="background:#fff3cd;color:#856404;">APERTURA</span>';
        } else if (item.tipo === 'error') {
            rowClass = 'row-error';
            statoHtml = '<span class="badge-stato badge-error">ERROR</span>';
        } else if (item.tipo === 'avere-match') {
            // Avere conciliato: sfondo bianco, badge con solo contorno
            rowClass = 'row-avere-match hidden'; // Inizia nascosto (collassato)
            statoHtml = '<span class="badge-stato badge-conciliato-outline">CONCILIATO</span>';
        } else if (m.conciliato) {
            rowClass = 'row-dare-conciliato';
            statoHtml = '<span class="badge-stato badge-conciliato">CONCILIATO</span>';
            if (m.matchProgs.length > 0) {
                // Icona freccia verso destra (collassato di default)
                expandIcon = `<span class="expand-icon" data-prog="${m.prog}">&#9654;</span>`;
            }
        } else if (m.avere) {
            // Avere non conciliato: badge outline arancione
            statoHtml = '<span class="badge-stato badge-non-conciliato-outline">NON CONCILIATO</span>';
        } else {
            // Dare non conciliato: badge pieno arancione
            statoHtml = '<span class="badge-stato badge-non-conciliato">NON CONCILIATO</span>';
        }

        // Data attributi per filtri
        let dataStato = 'non-conciliato';
        if (m.isError) dataStato = 'error';
        else if (m.conciliato) dataStato = 'conciliato';
        else if (m.codCausale === COD_BILANCIO_APERTURA) dataStato = 'apertura';

        // Saldo riconciliazione (solo per dare conciliati)
        let saldoHtml = '';
        if (item.tipo === 'dare' && m.conciliato && m.saldoRiconciliazione !== null) {
            if (Math.abs(m.saldoRiconciliazione) < 0.01) {
                saldoHtml = '<span class="saldo-zero">\u20AC 0,00</span>';
            } else {
                saldoHtml = `<span class="saldo-differenza">\u20AC ${formatImportoItaliano(m.saldoRiconciliazione)}</span>`;
            }
        }

        html += `
            <tr class="${rowClass}" data-stato="${dataStato}" data-prog="${m.prog}" ${item.parentProg ? `data-parent="${item.parentProg}"` : ''}>
                <td>${expandIcon}${statoHtml}</td>
                <td>${formatData(m.data)}</td>
                <td>${m.numDoc || ''}</td>
                <td>${m.descrizione}</td>
                <td class="importo-dare">${m.dare ? '\u20AC ' + formatImportoItaliano(m.dare) : ''}</td>
                <td class="importo-avere">${m.avere ? '\u20AC ' + formatImportoItaliano(m.avere) : ''}</td>
                <td>${saldoHtml}</td>
            </tr>
        `;
    });

    html += '</tbody>';
    table.innerHTML = html;

    // Aggiungi event listener per espandere/collassare
    document.querySelectorAll('.expand-icon').forEach(icon => {
        icon.addEventListener('click', (e) => {
            e.stopPropagation();
            const prog = icon.dataset.prog;
            const isExpanded = icon.classList.contains('expanded');

            // Toggle icona
            icon.classList.toggle('expanded');
            // &#9654; = freccia destra (collassato), &#9660; = freccia giu' (espanso)
            icon.innerHTML = isExpanded ? '&#9654;' : '&#9660;';

            // Toggle righe avere corrispondenti
            document.querySelectorAll(`tr[data-parent="${prog}"]`).forEach(row => {
                row.classList.toggle('hidden');
            });
        });
    });

    // Click su riga dare per toggle
    document.querySelectorAll('.row-dare-conciliato').forEach(row => {
        row.addEventListener('click', () => {
            const icon = row.querySelector('.expand-icon');
            if (icon) icon.click();
        });
    });
}

// === FILTRI ===

function toggleFilters() {
    const menu = document.getElementById('filters-menu');
    const toggle = document.getElementById('filters-toggle');
    menu.classList.toggle('show');
    toggle.classList.toggle('active');
}

function applyFilters() {
    const statoFilter = document.querySelector('input[name="stato-filter"]:checked').value;
    const table = document.getElementById('table-mastrino');
    const rows = table.querySelectorAll('tbody tr');
    const resetBtn = document.getElementById('btn-reset-filters');

    let hasActiveFilter = statoFilter !== 'all';

    rows.forEach(row => {
        const stato = row.dataset.stato;
        let visible = true;

        if (statoFilter !== 'all') {
            if (statoFilter === 'conciliati' && stato !== 'conciliato') visible = false;
            if (statoFilter === 'non-conciliati' && stato !== 'non-conciliato') visible = false;
            if (statoFilter === 'error' && stato !== 'error') visible = false;
        }

        row.classList.toggle('filter-hidden', !visible);
    });

    // Evidenzia pulsante reset se ci sono filtri attivi
    resetBtn.classList.toggle('active', hasActiveFilter);

    // Chiudi menu filtri
    document.getElementById('filters-menu').classList.remove('show');
    document.getElementById('filters-toggle').classList.remove('active');
}

function resetFilters() {
    document.querySelector('input[name="stato-filter"][value="all"]').checked = true;
    applyFilters();
}

// Chiudi menu filtri quando si clicca fuori
document.addEventListener('click', (e) => {
    const dropdown = document.getElementById('filters-dropdown');
    if (dropdown && !dropdown.contains(e.target)) {
        document.getElementById('filters-menu').classList.remove('show');
        document.getElementById('filters-toggle').classList.remove('active');
    }
});

// === GENERAZIONE EXCEL ===

function generaExcel() {
    if (!risultati) {
        showAlert('Nessun risultato da esportare!', 'error');
        return;
    }

    const wb = XLSX.utils.book_new();

    // Foglio principale con tutti i movimenti
    const data = [
        ['Stato', 'Data Reg.', 'Num. Doc.', 'Cod. Causale', 'Descrizione', 'Dare', 'Avere', 'Saldo Riconciliazione']
    ];

    risultati.movimenti.forEach(m => {
        let stato = 'Non Conciliato';
        if (m.isError) stato = 'ERROR';
        else if (m.conciliato) stato = 'Conciliato';
        else if (m.codCausale === COD_BILANCIO_APERTURA) stato = 'Apertura';

        data.push([
            stato,
            formatData(m.data),
            m.numDoc || '',
            m.codCausale || '',
            m.descrizione,
            m.dare || '',
            m.avere || '',
            m.saldoRiconciliazione !== null ? m.saldoRiconciliazione : ''
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Conciliazione');

    // Foglio riepilogo
    const riepilogo = [
        ['Riepilogo Conciliazione Mastrino'],
        [],
        ['Movimenti Totali', risultati.stats.totaleMovimenti],
        [],
        ['Anticipi Totali', risultati.stats.anticipiTotali],
        ['Anticipi Conciliati', risultati.stats.anticipiConciliati],
        ['Anticipi Non Conciliati', risultati.stats.anticipiNonConciliati],
        ['Importo Anticipi Conciliati', risultati.stats.importoAnticipiConciliati],
        ['Importo Anticipi Non Conciliati', risultati.stats.importoAnticipiNonConciliati],
        [],
        ['Saldi Totali', risultati.stats.saldiTotali],
        ['Saldi Conciliati', risultati.stats.saldiConciliati],
        ['Saldi Non Conciliati', risultati.stats.saldiNonConciliati],
        ['Importo Saldi Conciliati', risultati.stats.importoSaldiConciliati],
        ['Importo Saldi Non Conciliati', risultati.stats.importoSaldiNonConciliati],
        [],
        ['Commissioni (ERROR)', risultati.stats.commissioni]
    ];

    const wsRiepilogo = XLSX.utils.aoa_to_sheet(riepilogo);
    XLSX.utils.book_append_sheet(wb, wsRiepilogo, 'Riepilogo');

    // Download
    const now = new Date();
    const timestamp = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}_${now.getHours().toString().padStart(2,'0')}${now.getMinutes().toString().padStart(2,'0')}`;
    XLSX.writeFile(wb, `conciliazione_mastrino_${timestamp}.xlsx`);

    showAlert('Excel scaricato con successo!', 'success');
}

// === INIZIALIZZAZIONE ===

// Setup dropzone sidebar
setupDropzone('dropzone-mastrino-sidebar', 'file-mastrino-sidebar', 'info-mastrino-sidebar');

console.log('Conciliator Mastrino - Pronto');
