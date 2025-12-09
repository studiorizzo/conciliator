// === CONCILIATOR BANCA v4 - JavaScript ===
// Extracted from original HTML for modular structure

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
    if (!estrattoConto || !mastrino) {
        alert('⚠️ Devi prima caricare entrambi i file!');
        return;
    }

    const periodi = getPeriodi();
    if (periodi.length === 0) {
        alert('⚠️ Seleziona almeno un periodo!');
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
            btnConcilia.title = 'Prima conciliazione completata. Usa ↻ per ricominciare oppure vai in ⚙️ per riconciliare con nuovi parametri.';

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
    if (confirm('⚠️ Sei sicuro di voler ricaricare la pagina? Tutti i dati andranno persi.')) {
        location.reload();
    }
}

// === GESTIONE SIDEBAR CONFIGURAZIONE ===
const fab = document.getElementById('fab');
const sidebar = document.getElementById('sidebar');
const overlay = document.getElementById('overlay');

let sliderValues = {
    movimenti: 1,
    importo: 0.01,
    giorni: 0
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

function toggleTag(tag) {
    tag.classList.toggle('selected');
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
        if (sliderId === 'movimenti') {
            min = 1; max = 20;
        } else if (sliderId === 'importo') {
            min = 0.01; max = 1;
        } else if (sliderId === 'giorni') {
            min = 0; max = 15;
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
    let displayValue;

    if (id === 'movimenti' || id === 'giorni') {
        displayValue = Math.round(value);
        sliderValues[id] = displayValue;
    } else {
        displayValue = value.toFixed(2);
        sliderValues[id] = parseFloat(displayValue);
    }

    if (id === 'importo') {
        valueEl.textContent = '€ ' + displayValue;
    } else {
        valueEl.textContent = displayValue;
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
        const existingValue = parseFloat(item.querySelector('label').textContent.replace('€ ', ''));
        if (existingValue === value) {
            alert('Commissione già presente');
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
    label.textContent = '€ ' + value.toFixed(2);

    item.appendChild(checkbox);
    item.appendChild(label);

    let inserted = false;
    for (let i = 0; i < existingItems.length; i++) {
        const existingValue = parseFloat(existingItems[i].querySelector('label').textContent.replace('€ ', ''));
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
    COMMISSIONI_AMMESSE = Array.from(document.querySelectorAll('#commission-list input:checked')).map(c => {
        const text = c.nextElementSibling.textContent.replace('€ ', '');
        return parseFloat(text);
    });
}

function resetConfig() {
    updateSlider('movimenti', 0, 1, 20);
    updateSlider('importo', 0, 0.01, 1);
    updateSlider('giorni', 0, 0, 15);

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
    if (!estrattoConto || !mastrino) {
        alert('⚠️ Devi prima caricare entrambi i file!');
        return;
    }

    const periodi = getPeriodi();
    if (periodi.length === 0) {
        alert('⚠️ Seleziona almeno un periodo!');
        return;
    }

    updateCommissioniConfig();
    closeConfigSidebar();
    avviaConciliazioneAutomatica();
}

// === STATO GLOBALE ===
let estrattoConto = null;
let mastrino = null;
let risultati = null;
let annoRiferimento = null;

let COMMISSIONI_AMMESSE = [0.75, 1.75];

// === TABELLA COMMISSIONI DA REGISTRARE ===
let commissioniDaRegistrare = [];

function aggiungiCommissioneDaRegistrare(importo, dataBanca, progBanca, descrizioneBanca, importoBanca) {
    const esiste = commissioniDaRegistrare.find(c =>
        Math.abs(c.importo - importo) <= 0.01 &&
        c.prog_banca === progBanca
    );

    if (!esiste) {
        commissioniDaRegistrare.push({
            prog_banca: progBanca,
            data_banca: dataBanca,
            descrizione: descrizioneBanca || '',
            importo: importo,
            importo_banca: importoBanca
        });
        console.log(`   📝 Registrata commissione €${formatImportoItaliano(importo)} per ${progBanca}`);
    }
}

function isMatchSicuro(mov1, mov2, importo1, importo2, config, tolleranzaGiorni) {
    const dataEsatta = tolleranzaGiorni === 0 &&
                       mov1.data && mov2.data &&
                       mov1.data.getTime() === mov2.data.getTime();

    const diffImporto = Math.round(Math.abs(importo1 - importo2) * 100) / 100;
    const importoEsatto = diffImporto <= 0.01;

    return dataEsatta && importoEsatto;
}

console.log('🚀 Conciliator Banca v4 - Inizializzazione');

// === UTILITY FUNCTIONS ===

function parseImportoItaliano(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === 'string' && value.trim() === '') return null;
    if (typeof value === 'number' && !isNaN(value)) return value;

    let str = value.toString().trim();
    if (str === '') return null;

    str = str.replace(/\s/g, '');

    if (str.includes('.') && str.includes(',')) {
        str = str.replace(/\./g, '').replace(',', '.');
    }
    else if (str.includes(',') && !str.includes('.')) {
        str = str.replace(',', '.');
    }
    else if (str.includes('.')) {
        const parts = str.split('.');
        if (parts.length === 2 && parts[1].length === 3) {
            str = str.replace('.', '');
        }
    }

    const result = parseFloat(str);
    return isNaN(result) ? null : result;
}

function formatImportoItaliano(value) {
    if (value === null || value === undefined) return '';
    return value.toLocaleString('it-IT', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function parseDataEsplicita(dataStr, annoDefault) {
    if (!dataStr) return null;

    try {
        if (dataStr instanceof Date) {
            return isNaN(dataStr.getTime()) ? null : dataStr;
        }

        if (typeof dataStr === 'number') {
            const date = XLSX.SSF.parse_date_code(dataStr);
            return new Date(date.y, date.m - 1, date.d);
        }

        const str = dataStr.toString().trim();

        if (str.includes('/')) {
            const parts = str.split('/');
            if (parts.length === 2) {
                const giorno = parseInt(parts[0]);
                const mese = parseInt(parts[1]);
                return new Date(annoDefault || new Date().getFullYear(), mese - 1, giorno);
            } else if (parts.length === 3) {
                const giorno = parseInt(parts[0]);
                const mese = parseInt(parts[1]);
                const anno = parseInt(parts[2]);
                return new Date(anno, mese - 1, giorno);
            }
        }
        else if (str.includes('-')) {
            return new Date(str);
        }

        return null;
    } catch (e) {
        console.error('❌ Errore parsing data:', dataStr, e);
        return null;
    }
}

function estraiAnno(valutaStr) {
    if (!valutaStr) return null;

    try {
        if (valutaStr instanceof Date) {
            return valutaStr.getFullYear();
        }

        if (typeof valutaStr === 'number') {
            const date = XLSX.SSF.parse_date_code(valutaStr);
            return date.y;
        }

        const str = valutaStr.toString().trim();
        if (str.includes('/')) {
            const parts = str.split('/');
            if (parts.length === 3) {
                return parseInt(parts[2]);
            } else if (parts.length === 2) {
                return new Date().getFullYear();
            }
        }

        return null;
    } catch (e) {
        console.error('❌ Errore estrazione anno:', valutaStr, e);
        return null;
    }
}

function formatData(date) {
    if (!date) return '';
    const g = date.getDate().toString().padStart(2, '0');
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const a = date.getFullYear();
    return `${g}/${m}/${a}`;
}

function isInTrimestre(data, trimestre, anno) {
    if (!data || !anno) return false;

    const ranges = {
        'Q1': { start: new Date(anno, 0, 1), end: new Date(anno, 2, 31) },
        'Q2': { start: new Date(anno, 3, 1), end: new Date(anno, 5, 30) },
        'Q3': { start: new Date(anno, 6, 1), end: new Date(anno, 8, 30) },
        'Q4': { start: new Date(anno, 9, 1), end: new Date(anno, 11, 31) }
    };

    const range = ranges[trimestre];
    return data >= range.start && data <= range.end;
}

function isCommissioneAmmessa(importo, tolleranza = 0.01) {
    return COMMISSIONI_AMMESSE.some(comm => {
        const diff = Math.round(Math.abs(importo - comm) * 100) / 100;
        return diff <= tolleranza;
    });
}

// === FILE HANDLING ===

function setupDropzone(dropzoneId, inputId, infoId, tipo) {
    const dropzone = document.getElementById(dropzoneId);
    const input = document.getElementById(inputId);
    const info = document.getElementById(infoId);

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
        if (file) handleFile(file, tipo, dropzone, info);
    });

    input.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleFile(file, tipo, dropzone, info);
    });
}

function handleFile(file, tipo, dropzone, infoEl) {
    const inizioOverlay = Date.now();

    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        showAlert('❌ Formato file non valido. Usa file Excel (.xlsx o .xls)', 'error');
        return;
    }

    const tipoNome = tipo === 'estratto' ? 'Estratto Conto' : 'Mastrino Contabile';
    mostraLoading(`Lettura ${tipoNome} in corso...`, `Lettura del file ${file.name}`);

    setTimeout(() => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                if (tipo === 'estratto') {
                    estrattoConto = parseEstrattoConto(workbook);
                    commissioniDaRegistrare = [];
                    infoEl.textContent = `✓ ${file.name} - ${estrattoConto.length} movimenti caricati`;
                } else {
                    mastrino = parseMastrino(workbook);
                    commissioniDaRegistrare = [];
                    infoEl.textContent = `✓ ${file.name} - ${mastrino.length} movimenti caricati`;
                }

                infoEl.style.display = 'block';
                dropzone.classList.add('uploaded');

                const tempoTrascorso = Date.now() - inizioOverlay;
                const delayMinimo = 800;
                const ritardoNecessario = Math.max(0, delayMinimo - tempoTrascorso);

                setTimeout(() => {
                    nascondiLoading();
                    checkReady();
                    showAlert('✅ File caricato con successo!', 'success');
                }, ritardoNecessario);

            } catch (error) {
                console.error('❌ Errore parsing file:', error);
                nascondiLoading();
                showAlert('❌ Errore nel parsing del file: ' + error.message, 'error');
            }
        };

        reader.readAsArrayBuffer(file);
    }, 20);
}

function parseEstrattoConto(workbook) {
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet);

    console.log('📊 Parsing estratto conto...');
    console.log('   Righe totali:', json.length);

    if (!annoRiferimento) {
        for (let row of json) {
            if (row.Valuta) {
                annoRiferimento = estraiAnno(row.Valuta);
                console.log('📅 Anno riferimento estratto:', annoRiferimento);
                break;
            }
        }
    }

    const movimenti = json.filter(row => {
        const desc = (row['Descrizione operazione'] || '').toString().toLowerCase();
        const isSaldo = desc.includes('saldo iniziale') || desc.includes('saldo finale');
        const hasData = row.Data && row.Data !== '';
        return !isSaldo && hasData;
    });

    console.log('   Movimenti validi (esclusi saldi):', movimenti.length);

    let progCounter = 1;
    const parsed = movimenti.map((row) => {
        const prog = `E${progCounter++}`;
        const dataStr = row.Data ? row.Data.toString() : '';
        const dataValutaStr = row.Valuta ? row.Valuta.toString() : '';
        const data = parseDataEsplicita(dataStr, annoRiferimento);
        const entrate = parseImportoItaliano(row.Entrate);
        const uscite = parseImportoItaliano(row.Uscite);

        return {
            prog: prog,
            data_str_originale: dataStr,
            valuta_str_originale: dataValutaStr,
            data: data,
            descrizione: row['Descrizione operazione'] || '',
            entrate: entrate,
            uscite: uscite,
            conciliato: false,
            mastrino_progs: '',
            banca_progs: '',
            num_movimenti_mastrino: 0,
            num_movimenti_banca: 0,
            commissione_mastrino: 0,
            commissione_calcolata: 0
        };
    });

    console.log('   Entrate:', parsed.filter(p => p.entrate !== null).length);
    console.log('   Uscite:', parsed.filter(p => p.uscite !== null).length);
    console.log('   Con data valida:', parsed.filter(p => p.data !== null).length);

    return parsed;
}

function parseMastrino(workbook) {
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet);

    console.log('📊 Parsing mastrino...');
    console.log('   Righe totali:', json.length);

    const movimenti = json.filter(row => {
        return row.DataReg && (row.Dare || row.Avere);
    });

    console.log('   Movimenti validi:', movimenti.length);

    let progCounter = 1;
    const parsed = movimenti.map((row) => {
        const prog = `M${progCounter++}`;
        const dataReg = parseDataEsplicita(row.DataReg, annoRiferimento);
        const dataDoc = parseDataEsplicita(row.DataDoc, annoRiferimento);
        const dare = parseImportoItaliano(row.Dare);
        const avere = parseImportoItaliano(row.Avere);
        const saldo = parseImportoItaliano(row.Saldo);

        let descrizioneCompleta = [
            row.DesCausale,
            row.Des1Causale,
            row.Des2Causale,
            row.Des3Causale
        ].filter(x => x).join(' ');

        return {
            prog: prog,
            data_reg_str_originale: row.DataReg ? row.DataReg.toString() : '',
            data: dataReg,
            data_doc: dataDoc,
            num_doc: row.NumDoc,
            ndoc_orig: row.NdocOrig,
            cod_causale: row.CodCausale,
            descrizione: descrizioneCompleta,
            dare: dare,
            avere: avere,
            saldo: saldo,
            conciliato: false,
            banca_prog: '',
            tipo: dare ? 'dare' : 'avere'
        };
    });

    console.log('  DARE:', parsed.filter(p => p.dare !== null).length);
    console.log('  AVERE:', parsed.filter(p => p.avere !== null).length);
    console.log('  Con data valida:', parsed.filter(p => p.data !== null).length);

    return parsed;
}

function checkReady() {
    // Non più necessario - button concilia è nella sidebar
}

function showAlert(message, type = 'success') {
    const container = document.getElementById('alert-container');
    const alert = document.createElement('div');
    alert.className = `alert alert-${type}`;
    alert.textContent = message;
    container.appendChild(alert);

    setTimeout(() => alert.remove(), 5000);
}

// === ALGORITMO CONCILIAZIONE ===

function* combinations(arr, k) {
    if (k === 0) {
        yield [];
        return;
    }
    if (arr.length < k) return;

    for (let i = 0; i <= arr.length - k; i++) {
        const first = arr[i];
        const rest = arr.slice(i + 1);
        for (const combo of combinations(rest, k - 1)) {
            yield [first, ...combo];
        }
    }
}

function trovaCombinazioni(movimenti, target, dataTarget, campoImporto, config, tolleranzaGiorni = 0) {
    const validi = movimenti.filter(m => {
        if (m.conciliato || !m.data || !dataTarget) return false;

        const diffGiorni = Math.abs((m.data.getTime() - dataTarget.getTime()) / (1000 * 60 * 60 * 24));
        return diffGiorni <= tolleranzaGiorni;
    });

    if (validi.length === 0) return null;

    for (let n = 1; n <= Math.min(validi.length, config.maxMovimenti); n++) {
        for (const combo of combinations(validi, n)) {
            const somma = combo.reduce((sum, m) => sum + (m[campoImporto] || 0), 0);
            const diff = Math.round(Math.abs(somma - target) * 100) / 100;
            if (diff <= config.tolleranzaImporto) {
                return combo;
            }
        }
    }

    return null;
}

function trovaCombinazioniMoltiA1(movimenti, target, dataTarget, campoImporto, config, tolleranzaGiorni = 0) {
    const validi = movimenti.filter(m => {
        if (m.conciliato || !m.data || !dataTarget) return false;

        const diffGiorni = Math.abs((m.data.getTime() - dataTarget.getTime()) / (1000 * 60 * 60 * 24));
        return diffGiorni <= tolleranzaGiorni;
    });

    if (validi.length < 2) return null;

    for (let n = 2; n <= Math.min(validi.length, config.maxMovimenti); n++) {
        for (const combo of combinations(validi, n)) {
            const somma = combo.reduce((sum, m) => sum + (m[campoImporto] || 0), 0);
            const diff = Math.round(Math.abs(somma - target) * 100) / 100;
            if (diff <= config.tolleranzaImporto) {
                return combo;
            }
        }
    }

    return null;
}

function concilia(estratto, mast, config, periodi, onProgress, isIncremental = false) {
    console.log('═══════════════════════════════════════════');
    console.log(`🔄 INIZIO CONCILIAZIONE ${isIncremental ? '(INCREMENTALE)' : 'v3'}`);
    console.log('═══════════════════════════════════════════');
    console.log('📅 Anno riferimento:', annoRiferimento);
    console.log('📋 Periodi selezionati:', periodi);
    console.log('⚙️  Configurazione:', config);

    console.log('📊 ESTRATTO CONTO - movimenti totali:', estratto.length);
    console.log('   - Entrate:', estratto.filter(m => m.entrate).length);
    console.log('   - Uscite:', estratto.filter(m => m.uscite).length);
    console.log('   - Con data valida:', estratto.filter(m => m.data).length);

    console.log('📊 MASTRINO - movimenti totali:', mast.length);
    console.log('   - DARE:', mast.filter(m => m.dare).length);
    console.log('   - AVERE:', mast.filter(m => m.avere).length);
    console.log('   - Con data valida:', mast.filter(m => m.data).length);

    let estrattoFiltrato = estratto;
    let mastFiltrato = mast;

    if (periodi && periodi.length > 0 && annoRiferimento) {
        console.log(`📍 Filtro per periodo (anno ${annoRiferimento})...`);

        estrattoFiltrato = estratto.filter(m => {
            if (!m.data) return false;
            return periodi.some(q => isInTrimestre(m.data, q, annoRiferimento));
        });

        mastFiltrato = mast.filter(m => {
            if (!m.data) return false;
            return periodi.some(q => isInTrimestre(m.data, q, annoRiferimento));
        });

        console.log(`   ✓ Estratto filtrato: ${estrattoFiltrato.length} movimenti`);
        console.log(`   ✓ Mastrino filtrato: ${mastFiltrato.length} movimenti`);
    }

    const entrate = estrattoFiltrato.filter(m => {
        if (!m.entrate) return false;
        if (isIncremental && m.conciliato_sicuro) return false;
        return true;
    });
    const uscite = estrattoFiltrato.filter(m => {
        if (!m.uscite) return false;
        if (isIncremental && m.conciliato_sicuro) return false;
        return true;
    });
    const dare = mastFiltrato.filter(m => {
        if (!m.dare) return false;
        if (isIncremental && m.conciliato_sicuro) return false;
        return true;
    });
    const avere = mastFiltrato.filter(m => {
        if (!m.avere) return false;
        if (isIncremental && m.conciliato_sicuro) return false;
        return true;
    });

    console.log('═══════════════════════════════════════════');
    console.log('📊 DATI FILTRATI:');
    console.log('   Entrate banca:', entrate.length);
    console.log('   Uscite banca:', uscite.length);
    console.log('   DARE mastrino:', dare.length);
    console.log('   AVERE mastrino:', avere.length);
    if (isIncremental) {
        console.log('   (Esclusi movimenti conciliati sicuri dalla prima passata)');
    }
    console.log('═══════════════════════════════════════════');

    let progress = 0;
    const totalSteps = 16;

    console.log('');
    console.log('▶▶▶ FASE 1: Conciliazione con DATA ESATTA ◀◀◀');
    console.log('');

    // 1. Entrate 1-a-molti
    console.log('1️⃣  Entrate 1-a-molti (data esatta)...');
    progress++;
    onProgress(progress / totalSteps * 50);
    let contatore = 0;
    entrate.forEach(e => {
        if (e.conciliato) return;
        const combo = trovaCombinazioni(dare, e.entrate, e.data, 'dare', config, 0);
        if (combo) {
            const isSicuro = !isIncremental && combo.length === 1 &&
                             isMatchSicuro(e, combo[0], e.entrate, combo[0].dare, config, 0);

            e.conciliato = true;
            e.conciliato_sicuro = isSicuro;
            e.num_movimenti_mastrino = combo.length;
            e.num_movimenti_banca = 1;
            e.mastrino_progs = combo.map(m => m.prog).join(', ');
            combo.forEach(m => {
                m.conciliato = true;
                m.conciliato_sicuro = isSicuro;
                m.banca_prog = e.prog;
            });
            contatore++;
            console.log(`   ✓ ${e.prog} (€${formatImportoItaliano(e.entrate)}) → ${e.mastrino_progs}${isSicuro ? ' 🔒' : ''}`);
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    // 2. Entrate molti-a-1
    console.log('2️⃣  Entrate molti-a-1 (data esatta)...');
    progress++;
    onProgress(progress / totalSteps * 50);
    contatore = 0;
    dare.forEach(d => {
        if (d.conciliato) return;
        const combo = trovaCombinazioniMoltiA1(entrate, d.dare, d.data, 'entrate', config, 0);
        if (combo) {
            d.conciliato = true;
            d.banca_prog = combo.map(e => e.prog).join(', ');
            d.num_movimenti_banca = combo.length;
            combo.forEach(e => {
                e.conciliato = true;
                e.num_movimenti_mastrino = 1;
                e.num_movimenti_banca = combo.length;
                e.mastrino_progs = d.prog;
                e.banca_progs = combo.map(x => x.prog).join(', ');
            });
            contatore++;
            console.log(`   ✓ ${d.prog} (€${formatImportoItaliano(d.dare)}) ← ${d.banca_prog}`);
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    // 3. Uscite 1-a-molti (commissioni nel mastrino)
    console.log('3️⃣  Uscite 1-a-molti con commissioni NEL mastrino (data esatta)...');
    progress++;
    onProgress(progress / totalSteps * 50);
    contatore = 0;
    uscite.forEach(u => {
        if (u.conciliato) return;

        const combo = trovaCombinazioni(avere, u.uscite, u.data, 'avere', config, 0);

        if (combo) {
            let commissione = 0;
            const progs = [];
            const movimentiNonCommissione = combo.filter(a => !isCommissioneAmmessa(a.avere, config.tolleranzaImporto));

            const isSicuro = !isIncremental &&
                             movimentiNonCommissione.length === 1 &&
                             combo.length <= 2 &&
                             isMatchSicuro(u, movimentiNonCommissione[0], u.uscite - commissione, movimentiNonCommissione[0].avere, config, 0);

            combo.forEach(a => {
                if (isCommissioneAmmessa(a.avere, config.tolleranzaImporto)) {
                    commissione = a.avere;
                    a.tipo = 'commissione';
                } else {
                    progs.push(a.prog);
                }
                a.conciliato = true;
                a.conciliato_sicuro = isSicuro;
                a.banca_prog = u.prog;
                a.num_movimenti_banca = 1;
            });
            u.conciliato = true;
            u.conciliato_sicuro = isSicuro;
            u.num_movimenti_mastrino = combo.length;
            u.num_movimenti_banca = 1;
            u.commissione_mastrino = commissione;
            u.mastrino_progs = progs.join(', ');
            contatore++;
            if (commissione > 0) {
                console.log(`   ✓ ${u.prog} (€${formatImportoItaliano(u.uscite)}) → ${u.mastrino_progs} + COMM €${formatImportoItaliano(commissione)}${isSicuro ? ' 🔒' : ''}`);
            } else {
                console.log(`   ✓ ${u.prog} (€${formatImportoItaliano(u.uscite)}) → ${u.mastrino_progs}${isSicuro ? ' 🔒' : ''}`);
            }
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    // 4. Uscite con commissioni calcolate
    console.log('4️⃣  Uscite 1-a-molti con commissioni CALCOLATE (data esatta)...');
    progress++;
    onProgress(progress / totalSteps * 50);
    contatore = 0;
    uscite.forEach(u => {
        if (u.conciliato) return;
        for (const comm of COMMISSIONI_AMMESSE) {
            const combo = trovaCombinazioni(avere, u.uscite - comm, u.data, 'avere', config, 0);
            if (combo) {
                u.conciliato = true;
                u.num_movimenti_mastrino = combo.length;
                u.num_movimenti_banca = 1;
                u.commissione_calcolata = comm;
                u.mastrino_progs = combo.map(m => m.prog).join(', ');
                combo.forEach(a => {
                    a.conciliato = true;
                    a.banca_prog = u.prog;
                });

                aggiungiCommissioneDaRegistrare(comm, u.data, u.prog, u.descrizione, u.uscite);

                contatore++;
                console.log(`   ✓ ${u.prog} (€${formatImportoItaliano(u.uscite)}) → ${u.mastrino_progs} + CALC €${formatImportoItaliano(comm)}`);
                break;
            }
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    // 5. Uscite molti-a-1
    console.log('5️⃣  Uscite molti-a-1 (data esatta)...');
    progress++;
    onProgress(progress / totalSteps * 50);
    contatore = 0;
    avere.forEach(a => {
        if (a.conciliato || isCommissioneAmmessa(a.avere)) return;
        const combo = trovaCombinazioniMoltiA1(uscite, a.avere, a.data, 'uscite', config, 0);
        if (combo) {
            a.conciliato = true;
            a.banca_prog = combo.map(u => u.prog).join(', ');
            a.num_movimenti_banca = combo.length;
            combo.forEach(u => {
                u.conciliato = true;
                u.num_movimenti_mastrino = 1;
                u.num_movimenti_banca = combo.length;
                u.mastrino_progs = a.prog;
                u.banca_progs = combo.map(x => x.prog).join(', ');
            });
            contatore++;
            console.log(`   ✓ ${a.prog} (€${formatImportoItaliano(a.avere)}) ← ${a.banca_prog}`);
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    // 5B. Uscite molti-a-1 con commissioni NEL mastrino (data esatta)
    console.log('5️⃣B Uscite molti-a-1 con commissioni NEL mastrino (data esatta)...');
    progress++;
    onProgress(progress / totalSteps * 50);
    contatore = 0;

    avere.forEach(a => {
        if (a.conciliato || isCommissioneAmmessa(a.avere)) return;

        const commDisponibili = avere.filter(comm =>
            !comm.conciliato &&
            isCommissioneAmmessa(comm.avere) &&
            comm.data && a.data &&
            comm.data.getTime() === a.data.getTime()
        );

        let combo = trovaCombinazioniMoltiA1(uscite, a.avere, a.data, 'uscite', config, 0);

        if (!combo && commDisponibili.length > 0) {
            for (const comm of commDisponibili) {
                const targetConComm = Math.round((a.avere + comm.avere) * 100) / 100;
                combo = trovaCombinazioniMoltiA1(uscite, targetConComm, a.data, 'uscite', config, 0);

                if (combo) {
                    a.conciliato = true;
                    a.banca_prog = combo.map(u => u.prog).join(', ');
                    a.num_movimenti_banca = combo.length;

                    comm.conciliato = true;
                    comm.tipo = 'commissione';
                    comm.banca_prog = a.banca_prog;

                    combo.forEach(u => {
                        u.conciliato = true;
                        u.num_movimenti_mastrino = 2;
                        u.num_movimenti_banca = combo.length;
                        u.mastrino_progs = a.prog;
                        u.commissione_mastrino = comm.avere;
                        u.banca_progs = combo.map(x => x.prog).join(', ');
                    });

                    contatore++;
                    console.log(`   ✓ ${a.prog} (€${formatImportoItaliano(a.avere)}) ← ${a.banca_prog} + COMM ${comm.prog} (€${formatImportoItaliano(comm.avere)})`);
                    break;
                }
            }
        } else if (combo) {
            a.conciliato = true;
            a.banca_prog = combo.map(u => u.prog).join(', ');
            a.num_movimenti_banca = combo.length;
            combo.forEach(u => {
                u.conciliato = true;
                u.num_movimenti_mastrino = 1;
                u.num_movimenti_banca = combo.length;
                u.mastrino_progs = a.prog;
                u.banca_progs = combo.map(x => x.prog).join(', ');
            });
            contatore++;
            console.log(`   ✓ ${a.prog} (€${formatImportoItaliano(a.avere)}) ← ${a.banca_prog}`);
        }
    });

    console.log(`   Totale conciliate: ${contatore}\n`);

    // 5C. Uscite molti-a-1 con commissioni CALCOLATE (data esatta)
    console.log('5️⃣C Uscite molti-a-1 con commissioni CALCOLATE (data esatta)...');
    progress++;
    onProgress(progress / totalSteps * 50);
    contatore = 0;

    avere.forEach(a => {
        if (a.conciliato || isCommissioneAmmessa(a.avere)) return;

        for (const comm of COMMISSIONI_AMMESSE) {
            const targetConComm = Math.round((a.avere + comm) * 100) / 100;
            const combo = trovaCombinazioniMoltiA1(uscite, targetConComm, a.data, 'uscite', config, 0);

            if (combo) {
                a.conciliato = true;
                a.banca_prog = combo.map(u => u.prog).join(', ');
                a.num_movimenti_banca = combo.length;

                combo.forEach(u => {
                    u.conciliato = true;
                    u.num_movimenti_mastrino = 1;
                    u.num_movimenti_banca = combo.length;
                    u.mastrino_progs = a.prog;
                    u.commissione_calcolata = comm;
                    u.banca_progs = combo.map(x => x.prog).join(', ');

                    aggiungiCommissioneDaRegistrare(comm, u.data, u.prog, u.descrizione, u.uscite);
                });

                contatore++;
                console.log(`   ✓ ${a.prog} (€${formatImportoItaliano(a.avere)}) ← ${a.banca_prog} + CALC €${formatImportoItaliano(comm)}`);
                break;
            }
        }
    });

    console.log(`   Totale conciliate: ${contatore}\n`);

    // 11A. Riconciliazione commissioni calcolate (data esatta)
    console.log('');
    console.log('🔧 PASSAGGIO 11A: Riconciliazione commissioni calcolate (data esatta)');
    console.log('');
    progress++;
    onProgress(progress / totalSteps * 50);
    contatore = 0;

    for (let i = commissioniDaRegistrare.length - 1; i >= 0; i--) {
        const commCalc = commissioniDaRegistrare[i];

        const match = avere.find(a =>
            !a.conciliato &&
            isCommissioneAmmessa(a.avere) &&
            Math.abs(a.avere - commCalc.importo) <= 0.01 &&
            a.data && commCalc.data_banca &&
            a.data.getTime() === commCalc.data_banca.getTime()
        );

        if (match) {
            match.conciliato = true;
            match.tipo = 'commissione';
            match.banca_prog = commCalc.prog_banca;
            match.num_movimenti_banca = 1;

            const uscitaBanca = uscite.find(u => u.prog === commCalc.prog_banca);
            if (uscitaBanca) {
                if (uscitaBanca.mastrino_progs) {
                    uscitaBanca.mastrino_progs += ', ' + match.prog;
                } else {
                    uscitaBanca.mastrino_progs = match.prog;
                }

                if (uscitaBanca.commissione_calcolata) {
                    uscitaBanca.commissione_mastrino = uscitaBanca.commissione_calcolata;
                    uscitaBanca.commissione_calcolata = 0;
                }

                if (uscitaBanca.num_movimenti_mastrino) {
                    uscitaBanca.num_movimenti_mastrino++;
                } else {
                    uscitaBanca.num_movimenti_mastrino = 1;
                }
            }

            commissioniDaRegistrare.splice(i, 1);

            contatore++;
            console.log(`   ✓ ${commCalc.prog_banca}: Commissione €${formatImportoItaliano(commCalc.importo)} → ${match.prog}`);
        }
    }

    console.log(`   Totale commissioni riconciliate: ${contatore}`);
    console.log('');

    console.log('');
    console.log(`▶▶▶ FASE 2: Conciliazione con TOLLERANZA ±${config.tolleranzaGiorni} GIORNI ◀◀◀`);
    console.log('');

    const tolGiorni = parseInt(config.tolleranzaGiorni);

    // Ripete con tolleranza
    console.log('6️⃣  Entrate 1-a-molti (tolleranza)...');
    progress++;
    onProgress(progress / totalSteps * 50 + 50);
    contatore = 0;
    entrate.filter(e => !e.conciliato).forEach(e => {
        const combo = trovaCombinazioni(dare, e.entrate, e.data, 'dare', config, tolGiorni);
        if (combo) {
            e.conciliato = true;
            e.num_movimenti_mastrino = combo.length;
            e.num_movimenti_banca = 1;
            e.mastrino_progs = combo.map(m => m.prog).join(', ');
            combo.forEach(m => {
                m.conciliato = true;
                m.banca_prog = e.prog;
            });
            contatore++;
            console.log(`   ✓ ${e.prog} → ${e.mastrino_progs} (con tolleranza)`);
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    console.log('7️⃣  Entrate molti-a-1 (tolleranza)...');
    progress++;
    onProgress(progress / totalSteps * 50 + 50);
    contatore = 0;
    dare.filter(d => !d.conciliato).forEach(d => {
        const combo = trovaCombinazioniMoltiA1(entrate, d.dare, d.data, 'entrate', config, tolGiorni);
        if (combo) {
            d.conciliato = true;
            d.banca_prog = combo.map(e => e.prog).join(', ');
            d.num_movimenti_banca = combo.length;
            combo.forEach(e => {
                e.conciliato = true;
                e.num_movimenti_mastrino = 1;
                e.num_movimenti_banca = combo.length;
                e.mastrino_progs = d.prog;
                e.banca_progs = combo.map(x => x.prog).join(', ');
            });
            contatore++;
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    console.log('8️⃣  Uscite 1-a-molti con commissioni NEL mastrino (tolleranza)...');
    progress++;
    onProgress(progress / totalSteps * 50 + 50);
    contatore = 0;
    uscite.filter(u => !u.conciliato).forEach(u => {
        const combo = trovaCombinazioni(avere, u.uscite, u.data, 'avere', config, tolGiorni);
        if (combo) {
            let commissione = 0;
            const progs = [];
            combo.forEach(a => {
                if (isCommissioneAmmessa(a.avere)) {
                    commissione = a.avere;
                    a.tipo = 'commissione';
                } else {
                    progs.push(a.prog);
                }
                a.conciliato = true;
                a.banca_prog = u.prog;
            });
            u.conciliato = true;
            u.num_movimenti_mastrino = combo.length;
            u.num_movimenti_banca = 1;
            u.commissione_mastrino = commissione;
            u.mastrino_progs = progs.join(', ');
            contatore++;
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    console.log('9️⃣  Uscite 1-a-molti con commissioni CALCOLATE (tolleranza)...');
    progress++;
    onProgress(progress / totalSteps * 50 + 50);
    contatore = 0;
    uscite.filter(u => !u.conciliato).forEach(u => {
        for (const comm of COMMISSIONI_AMMESSE) {
            const combo = trovaCombinazioni(avere, u.uscite - comm, u.data, 'avere', config, tolGiorni);
            if (combo) {
                u.conciliato = true;
                u.num_movimenti_mastrino = combo.length;
                u.num_movimenti_banca = 1;
                u.commissione_calcolata = comm;
                u.mastrino_progs = combo.map(m => m.prog).join(', ');
                combo.forEach(a => {
                    a.conciliato = true;
                    a.banca_prog = u.prog;
                });

                aggiungiCommissioneDaRegistrare(comm, u.data, u.prog, u.descrizione, u.uscite);

                contatore++;
                break;
            }
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    console.log('🔟 Uscite molti-a-1 (tolleranza)...');
    progress++;
    onProgress(progress / totalSteps * 50 + 50);
    contatore = 0;
    avere.filter(a => !a.conciliato && !isCommissioneAmmessa(a.avere)).forEach(a => {
        const combo = trovaCombinazioniMoltiA1(uscite, a.avere, a.data, 'uscite', config, tolGiorni);
        if (combo) {
            a.conciliato = true;
            a.banca_prog = combo.map(u => u.prog).join(', ');
            a.num_movimenti_banca = combo.length;
            combo.forEach(u => {
                u.conciliato = true;
                u.num_movimenti_mastrino = 1;
                u.num_movimenti_banca = combo.length;
                u.mastrino_progs = a.prog;
                u.banca_progs = combo.map(x => x.prog).join(', ');
            });
            contatore++;
        }
    });
    console.log(`   Totale conciliate: ${contatore}\n`);

    // 10B. Uscite molti-a-1 con commissioni NEL mastrino (tolleranza)
    console.log('🔟B Uscite molti-a-1 con commissioni NEL mastrino (tolleranza)...');
    progress++;
    onProgress(progress / totalSteps * 50 + 50);
    contatore = 0;

    avere.filter(a => !a.conciliato && !isCommissioneAmmessa(a.avere)).forEach(a => {

        const commDisponibili = avere.filter(comm =>
            !comm.conciliato &&
            isCommissioneAmmessa(comm.avere) &&
            comm.data && a.data &&
            Math.abs((comm.data.getTime() - a.data.getTime()) / (1000 * 60 * 60 * 24)) <= tolGiorni
        );

        let combo = trovaCombinazioniMoltiA1(uscite, a.avere, a.data, 'uscite', config, tolGiorni);

        if (!combo && commDisponibili.length > 0) {
            for (const comm of commDisponibili) {
                const targetConComm = Math.round((a.avere + comm.avere) * 100) / 100;
                combo = trovaCombinazioniMoltiA1(uscite, targetConComm, a.data, 'uscite', config, tolGiorni);

                if (combo) {
                    a.conciliato = true;
                    a.banca_prog = combo.map(u => u.prog).join(', ');
                    a.num_movimenti_banca = combo.length;

                    comm.conciliato = true;
                    comm.tipo = 'commissione';
                    comm.banca_prog = a.banca_prog;

                    combo.forEach(u => {
                        u.conciliato = true;
                        u.num_movimenti_mastrino = 2;
                        u.num_movimenti_banca = combo.length;
                        u.mastrino_progs = a.prog;
                        u.commissione_mastrino = comm.avere;
                        u.banca_progs = combo.map(x => x.prog).join(', ');
                    });

                    contatore++;
                    break;
                }
            }
        } else if (combo) {
            a.conciliato = true;
            a.banca_prog = combo.map(u => u.prog).join(', ');
            a.num_movimenti_banca = combo.length;
            combo.forEach(u => {
                u.conciliato = true;
                u.num_movimenti_mastrino = 1;
                u.num_movimenti_banca = combo.length;
                u.mastrino_progs = a.prog;
                u.banca_progs = combo.map(x => x.prog).join(', ');
            });
            contatore++;
        }
    });

    console.log(`   Totale conciliate: ${contatore}\n`);

    // 10C. Uscite molti-a-1 con commissioni CALCOLATE (tolleranza)
    console.log('🔟C Uscite molti-a-1 con commissioni CALCOLATE (tolleranza)...');
    progress++;
    onProgress(progress / totalSteps * 50 + 50);
    contatore = 0;

    avere.filter(a => !a.conciliato && !isCommissioneAmmessa(a.avere)).forEach(a => {

        for (const comm of COMMISSIONI_AMMESSE) {
            const targetConComm = Math.round((a.avere + comm) * 100) / 100;
            const combo = trovaCombinazioniMoltiA1(uscite, targetConComm, a.data, 'uscite', config, tolGiorni);

            if (combo) {
                a.conciliato = true;
                a.banca_prog = combo.map(u => u.prog).join(', ');
                a.num_movimenti_banca = combo.length;

                combo.forEach(u => {
                    u.conciliato = true;
                    u.num_movimenti_mastrino = 1;
                    u.num_movimenti_banca = combo.length;
                    u.mastrino_progs = a.prog;
                    u.commissione_calcolata = comm;
                    u.banca_progs = combo.map(x => x.prog).join(', ');

                    aggiungiCommissioneDaRegistrare(comm, u.data, u.prog, u.descrizione, u.uscite);
                });

                contatore++;
                break;
            }
        }
    });

    console.log(`   Totale conciliate: ${contatore}\n`);

    // 11B. Riconciliazione commissioni calcolate (tolleranza giorni)
    console.log('');
    console.log('🔧 PASSAGGIO 11B: Riconciliazione commissioni calcolate (tolleranza giorni)');
    console.log('');
    progress++;
    onProgress(100);
    contatore = 0;

    for (let i = commissioniDaRegistrare.length - 1; i >= 0; i--) {
        const commCalc = commissioniDaRegistrare[i];

        let miglioreMatch = null;
        let minDiffGiorni = Infinity;

        avere.filter(a =>
            !a.conciliato &&
            isCommissioneAmmessa(a.avere) &&
            Math.abs(a.avere - commCalc.importo) <= 0.01 &&
            a.data && commCalc.data_banca
        ).forEach(a => {
            const diffGiorni = Math.abs((a.data.getTime() - commCalc.data_banca.getTime()) / (1000 * 60 * 60 * 24));

            if (diffGiorni <= tolGiorni && diffGiorni < minDiffGiorni) {
                minDiffGiorni = diffGiorni;
                miglioreMatch = a;
            }
        });

        if (miglioreMatch) {
            miglioreMatch.conciliato = true;
            miglioreMatch.tipo = 'commissione';
            miglioreMatch.banca_prog = commCalc.prog_banca;
            miglioreMatch.num_movimenti_banca = 1;

            const uscitaBanca = uscite.find(u => u.prog === commCalc.prog_banca);
            if (uscitaBanca) {
                if (uscitaBanca.mastrino_progs) {
                    uscitaBanca.mastrino_progs += ', ' + miglioreMatch.prog;
                } else {
                    uscitaBanca.mastrino_progs = miglioreMatch.prog;
                }

                if (uscitaBanca.commissione_calcolata) {
                    uscitaBanca.commissione_mastrino = uscitaBanca.commissione_calcolata;
                    uscitaBanca.commissione_calcolata = 0;
                }

                if (uscitaBanca.num_movimenti_mastrino) {
                    uscitaBanca.num_movimenti_mastrino++;
                } else {
                    uscitaBanca.num_movimenti_mastrino = 1;
                }
            }

            commissioniDaRegistrare.splice(i, 1);

            contatore++;
            console.log(`   ✓ ${commCalc.prog_banca}: Commissione €${formatImportoItaliano(commCalc.importo)} → ${miglioreMatch.prog} (diff: ${minDiffGiorni}gg)`);
        }
    }

    console.log(`   Totale commissioni riconciliate: ${contatore}`);
    console.log('');

    onProgress(100);

    const entrateConc = entrate.filter(e => e.conciliato).length;
    const usciteConc = uscite.filter(u => u.conciliato).length;

    console.log('═══════════════════════════════════════════');
    console.log('✅ CONCILIAZIONE COMPLETATA');
    console.log('═══════════════════════════════════════════');
    console.log(`📊 ENTRATE: ${entrateConc}/${entrate.length} conciliate (${(entrateConc/entrate.length*100).toFixed(1)}%)`);
    console.log(`📊 USCITE: ${usciteConc}/${uscite.length} conciliate (${(usciteConc/uscite.length*100).toFixed(1)}%)`);
    console.log('═══════════════════════════════════════════\n');

    return { entrate: estrattoFiltrato, uscite: [], dare: dare, avere: avere };
}

// === VISUALIZZAZIONE RISULTATI ===

let datiEstrattoConto = [];
let datiMastrino = [];
let filtriAttivi = {
    tipo: 'all',
    stato: 'all'
};

function mostraRisultati(ris) {
    console.log('📊 Visualizzazione risultati v4...');

    const entrate = ris.entrate.filter(m => m.entrate);
    const uscite = ris.entrate.filter(m => m.uscite);

    const entrateConc = entrate.filter(e => e.conciliato).length;
    const usciteConc = uscite.filter(u => u.conciliato).length;
    const commCalc = commissioniDaRegistrare.reduce((s, c) => s + c.importo, 0);

    const summaryHTML = `
        <div class="summary-card">
            <h3>Entrate Conciliate</h3>
            <div class="value">${entrateConc}</div>
            <div class="percentage">${entrate.length > 0 ? (entrateConc/entrate.length*100).toFixed(1) : 0}%</div>
        </div>
        <div class="summary-card">
            <h3>Entrate NON Conciliate</h3>
            <div class="value">${entrate.length - entrateConc}</div>
        </div>
        <div class="summary-card">
            <h3>Uscite Conciliate</h3>
            <div class="value">${usciteConc}</div>
            <div class="percentage">${uscite.length > 0 ? (usciteConc/uscite.length*100).toFixed(1) : 0}%</div>
        </div>
        <div class="summary-card">
            <h3>Uscite NON Conciliate</h3>
            <div class="value">${uscite.length - usciteConc}</div>
        </div>
        <div class="summary-card">
            <h3>Commissioni da registrare</h3>
            <div class="value">${commissioniDaRegistrare.length}</div>
            <div class="percentage">€ ${formatImportoItaliano(commCalc)}</div>
        </div>
    `;
    document.getElementById('summaryGrid').innerHTML = summaryHTML;

    datiEstrattoConto = [...entrate, ...uscite].sort((a, b) => {
        const progA = parseInt(a.prog.replace(/[^0-9]/g, '')) || 0;
        const progB = parseInt(b.prog.replace(/[^0-9]/g, '')) || 0;
        return progA - progB;
    });

    datiMastrino = [...ris.dare, ...ris.avere].sort((a, b) => {
        const progA = parseInt(a.prog.replace(/[^0-9]/g, '')) || 0;
        const progB = parseInt(b.prog.replace(/[^0-9]/g, '')) || 0;
        return progA - progB;
    });

    popolaTabellaEstrattoConto();
    popolaTabellaMastrino();
    popolaTabellaCommissioni();

    document.getElementById('results-section').style.display = 'block';
    mostraDownloadFab();

    verificaDiscordanzeMastrino(entrate, uscite, ris.dare, ris.avere);
}

function verificaDiscordanzeMastrino(entrate, uscite, dare, avere) {
    const totaleUsciteNonConc = uscite.filter(u => !u.conciliato).reduce((s, u) => s + (u.uscite || 0), 0);
    const totaleEntrateNonConc = entrate.filter(e => !e.conciliato).reduce((s, e) => s + (e.entrate || 0), 0);

    const totaleAvereNonConc = avere.filter(a => !a.conciliato).reduce((s, a) => s + (a.avere || 0), 0);
    const totaleDareNonConc = dare.filter(d => !d.conciliato).reduce((s, d) => s + (d.dare || 0), 0);

    const diffUsciteAvere = Math.round(Math.abs(totaleUsciteNonConc - totaleAvereNonConc) * 100) / 100;
    const diffEntrateDare = Math.round(Math.abs(totaleEntrateNonConc - totaleDareNonConc) * 100) / 100;

    const alertElement = document.getElementById('mastrino-alert');

    if (diffUsciteAvere > 0.01 || diffEntrateDare > 0.01) {
        alertElement.style.display = 'inline-flex';

        let tooltipText = 'DISCORDANZE RILEVATE:\n';
        if (diffUsciteAvere > 0.01) {
            tooltipText += `\nUscite non conc: €${formatImportoItaliano(totaleUsciteNonConc)}`;
            tooltipText += `\nAVERE non conc: €${formatImportoItaliano(totaleAvereNonConc)}`;
            tooltipText += `\nDifferenza: €${formatImportoItaliano(diffUsciteAvere)}`;
        }
        if (diffEntrateDare > 0.01) {
            if (diffUsciteAvere > 0.01) tooltipText += '\n---';
            tooltipText += `\nEntrate non conc: €${formatImportoItaliano(totaleEntrateNonConc)}`;
            tooltipText += `\nDARE non conc: €${formatImportoItaliano(totaleDareNonConc)}`;
            tooltipText += `\nDifferenza: €${formatImportoItaliano(diffEntrateDare)}`;
        }

        alertElement.title = tooltipText;
    } else {
        alertElement.style.display = 'none';
    }
}

function popolaTabellaEstrattoConto() {
    const table = document.getElementById('table-estratto-conto');

    if (datiEstrattoConto.length === 0) {
        table.innerHTML = `<div class="empty-state"><p>Nessun dato disponibile</p></div>`;
        return;
    }

    let html = `
        <thead>
            <tr>
                <th>Stato</th>
                <th>Prog Banca</th>
                <th>Data</th>
                <th>Descrizione</th>
                <th class="col-uscite" style="text-align: right;">Uscite</th>
                <th class="col-entrate" style="text-align: right;">Entrate</th>
                <th>Prog Mastrino</th>
                <th style="text-align: right;">Commissioni da registrare</th>
                <th style="text-align: center;">N.Mov Mastrino</th>
            </tr>
        </thead>
        <tbody>
    `;

    datiEstrattoConto.forEach(row => {
        const isEntrata = row.entrate > 0;
        const isUscita = row.uscite > 0;
        const isConciliato = row.conciliato || false;

        const badgeStato = isConciliato
            ? '<span class="badge-stato badge-conciliato">Conciliato</span>'
            : '<span class="badge-stato badge-non-conciliato">Non Conciliato</span>';

        const usciteStr = isUscita ? `<span class="importo-uscita">${formatImportoItaliano(row.uscite)}</span>` : '';
        const entrateStr = isEntrata ? `<span class="importo-entrata">${formatImportoItaliano(row.entrate)}</span>` : '';

        const commissioniStr = row.commissione_calcolata ? `<span class="importo-uscita">${formatImportoItaliano(row.commissione_calcolata)}</span>` : '';

        const progMastrinoStr = row.mastrino_progs || '';
        const nMovMastrinoStr = row.num_movimenti_mastrino || '';

        const tipoData = isEntrata ? 'entrata' : 'uscita';
        const statoData = isConciliato ? 'conciliato' : 'non-conciliato';

        html += `
            <tr data-tipo="${tipoData}" data-stato="${statoData}" data-uscite="${row.uscite || 0}" data-entrate="${row.entrate || 0}">
                <td>${badgeStato}</td>
                <td>${row.prog}</td>
                <td>${formatData(row.data)}</td>
                <td>${row.descrizione || ''}</td>
                <td class="col-uscite" style="text-align: right;">${usciteStr}</td>
                <td class="col-entrate" style="text-align: right;">${entrateStr}</td>
                <td>${progMastrinoStr}</td>
                <td style="text-align: right;">${commissioniStr}</td>
                <td style="text-align: center;">${nMovMastrinoStr}</td>
            </tr>
        `;
    });

    const totaleUscite = datiEstrattoConto.reduce((sum, row) => sum + (row.uscite || 0), 0);
    const totaleEntrate = datiEstrattoConto.reduce((sum, row) => sum + (row.entrate || 0), 0);

    html += `
        </tbody>
        <tfoot>
            <tr class="totals-row">
                <td colspan="4" style="text-align: right; font-weight: 700;">TOTALE:</td>
                <td class="col-uscite" style="text-align: right; font-weight: 700;"><span class="importo-uscita">${formatImportoItaliano(totaleUscite)}</span></td>
                <td class="col-entrate" style="text-align: right; font-weight: 700;"><span class="importo-entrata">${formatImportoItaliano(totaleEntrate)}</span></td>
                <td colspan="3"></td>
            </tr>
        </tfoot>
    `;
    table.innerHTML = html;
}

function popolaTabellaMastrino() {
    const table = document.getElementById('table-mastrino');

    if (datiMastrino.length === 0) {
        table.innerHTML = `<div class="empty-state"><p>Nessun dato disponibile</p></div>`;
        return;
    }

    let html = `
        <thead>
            <tr>
                <th>Stato</th>
                <th>Prog Mastrino</th>
                <th>Data Reg</th>
                <th>Num Doc</th>
                <th>Descrizione</th>
                <th class="col-dare" style="text-align: right;">Dare</th>
                <th class="col-avere" style="text-align: right;">Avere</th>
                <th>Prog Banca</th>
                <th style="text-align: center;">N.Mov Banca</th>
            </tr>
        </thead>
        <tbody>
    `;

    datiMastrino.forEach(row => {
        const isDare = row.dare > 0;
        const isAvere = row.avere > 0;
        const isConciliato = row.conciliato || false;

        const badgeStato = isConciliato
            ? '<span class="badge-stato badge-conciliato">Conciliato</span>'
            : '<span class="badge-stato badge-non-conciliato">Non Conciliato</span>';

        const dareStr = isDare ? `<span class="importo-dare">${formatImportoItaliano(row.dare)}</span>` : '';
        const avereStr = isAvere ? `<span class="importo-avere">${formatImportoItaliano(row.avere)}</span>` : '';

        const progBancaStr = row.banca_prog || '';
        const nMovBancaStr = row.num_movimenti_banca || '';

        const tipoData = isDare ? 'dare' : 'avere';
        const statoData = isConciliato ? 'conciliato' : 'non-conciliato';

        html += `
            <tr data-tipo="${tipoData}" data-stato="${statoData}" data-dare="${row.dare || 0}" data-avere="${row.avere || 0}">
                <td>${badgeStato}</td>
                <td>${row.prog}</td>
                <td>${formatData(row.data)}</td>
                <td>${row.num_doc || ''}</td>
                <td>${row.descrizione || ''}</td>
                <td class="col-dare" style="text-align: right;">${dareStr}</td>
                <td class="col-avere" style="text-align: right;">${avereStr}</td>
                <td>${progBancaStr}</td>
                <td style="text-align: center;">${nMovBancaStr}</td>
            </tr>
        `;
    });

    const totaleDare = datiMastrino.reduce((sum, row) => sum + (row.dare || 0), 0);
    const totaleAvere = datiMastrino.reduce((sum, row) => sum + (row.avere || 0), 0);

    html += `
        </tbody>
        <tfoot>
            <tr class="totals-row">
                <td colspan="5" style="text-align: right; font-weight: 700;">TOTALE:</td>
                <td class="col-dare" style="text-align: right; font-weight: 700;"><span class="importo-dare">${formatImportoItaliano(totaleDare)}</span></td>
                <td class="col-avere" style="text-align: right; font-weight: 700;"><span class="importo-avere">${formatImportoItaliano(totaleAvere)}</span></td>
                <td colspan="2"></td>
            </tr>
        </tfoot>
    `;
    table.innerHTML = html;
}

function popolaTabellaCommissioni() {
    const table = document.getElementById('table-commissioni');

    if (commissioniDaRegistrare.length === 0) {
        table.innerHTML = `<div class="empty-state"><p>Nessuna commissione da registrare</p></div>`;
        return;
    }

    let html = `
        <thead>
            <tr>
                <th style="width: 10%;">Prog Banca</th>
                <th style="width: 10%;">Data</th>
                <th style="width: 40%;">Descrizione</th>
                <th style="width: 15%; text-align: right;">Commissione</th>
                <th style="width: 25%; text-align: right; padding-right: 100px;">Importo Banca</th>
            </tr>
        </thead>
        <tbody>
    `;

    commissioniDaRegistrare.forEach(c => {
        html += `
            <tr>
                <td>${c.prog_banca}</td>
                <td>${formatData(c.data_banca)}</td>
                <td>${c.descrizione || ''}</td>
                <td style="text-align: right;"><span class="importo-uscita">${formatImportoItaliano(c.importo)}</span></td>
                <td style="text-align: right; padding-right: 100px;"><span class="importo-uscita">${formatImportoItaliano(c.importo_banca)}</span></td>
            </tr>
        `;
    });

    const totaleCommissioni = commissioniDaRegistrare.reduce((sum, c) => sum + c.importo, 0);

    html += `
        </tbody>
        <tfoot>
            <tr class="totals-row">
                <td colspan="3" style="text-align: right; font-weight: 700;">TOTALE:</td>
                <td style="text-align: right; font-weight: 700;"><span class="importo-uscita">${formatImportoItaliano(totaleCommissioni)}</span></td>
                <td></td>
            </tr>
        </tfoot>
    `;
    table.innerHTML = html;
}

function applyFilters() {
    const tipoFilter = document.querySelector('input[name="tipo-filter"]:checked')?.value || 'all';
    const statoFilter = document.querySelector('input[name="stato-filter"]:checked')?.value || 'all';

    filtriAttivi = { tipo: tipoFilter, stato: statoFilter };

    updateFiltersButtonState();

    const tableEstratto = document.getElementById('table-estratto-conto');
    const tableMastrino = document.getElementById('table-mastrino');

    if (tipoFilter === 'entrate-dare') {
        tableEstratto.classList.add('hide-column-uscite');
        tableEstratto.classList.remove('hide-column-entrate');
    } else if (tipoFilter === 'uscite-avere') {
        tableEstratto.classList.add('hide-column-entrate');
        tableEstratto.classList.remove('hide-column-uscite');
    } else {
        tableEstratto.classList.remove('hide-column-uscite');
        tableEstratto.classList.remove('hide-column-entrate');
    }

    if (tipoFilter === 'entrate-dare') {
        tableMastrino.classList.add('hide-column-avere');
        tableMastrino.classList.remove('hide-column-dare');
    } else if (tipoFilter === 'uscite-avere') {
        tableMastrino.classList.add('hide-column-dare');
        tableMastrino.classList.remove('hide-column-avere');
    } else {
        tableMastrino.classList.remove('hide-column-dare');
        tableMastrino.classList.remove('hide-column-avere');
    }

    const rowsEstratto = document.querySelectorAll('#table-estratto-conto tbody tr');
    rowsEstratto.forEach(row => {
        const rowTipo = row.dataset.tipo;
        const rowStato = row.dataset.stato;

        let show = true;

        if (tipoFilter === 'entrate-dare' && rowTipo !== 'entrata') show = false;
        if (tipoFilter === 'uscite-avere' && rowTipo !== 'uscita') show = false;

        if (statoFilter === 'conciliati' && rowStato !== 'conciliato') show = false;
        if (statoFilter === 'non-conciliati' && rowStato !== 'non-conciliato') show = false;

        if (show) {
            row.classList.remove('filter-hidden');
        } else {
            row.classList.add('filter-hidden');
        }
    });

    const rowsMastrino = document.querySelectorAll('#table-mastrino tbody tr');
    rowsMastrino.forEach(row => {
        const rowTipo = row.dataset.tipo;
        const rowStato = row.dataset.stato;

        let show = true;

        if (tipoFilter === 'entrate-dare' && rowTipo !== 'dare') show = false;
        if (tipoFilter === 'uscite-avere' && rowTipo !== 'avere') show = false;

        if (statoFilter === 'conciliati' && rowStato !== 'conciliato') show = false;
        if (statoFilter === 'non-conciliati' && rowStato !== 'non-conciliato') show = false;

        if (show) {
            row.classList.remove('filter-hidden');
        } else {
            row.classList.add('filter-hidden');
        }
    });

    ricalcolaTotaliFiltrati();
}

function ricalcolaTotaliFiltrati() {
    const rowsEstrattoVisibili = document.querySelectorAll('#table-estratto-conto tbody tr:not(.filter-hidden)');
    let totaleUsciteFiltrate = 0;
    let totaleEntrateFiltrate = 0;

    rowsEstrattoVisibili.forEach(row => {
        const uscite = parseFloat(row.dataset.uscite) || 0;
        const entrate = parseFloat(row.dataset.entrate) || 0;
        totaleUsciteFiltrate += uscite;
        totaleEntrateFiltrate += entrate;
    });

    const footerEstrattoUscite = document.querySelector('#table-estratto-conto tfoot .col-uscite');
    const footerEstrattoEntrate = document.querySelector('#table-estratto-conto tfoot .col-entrate');
    if (footerEstrattoUscite) {
        footerEstrattoUscite.innerHTML = `<span class="importo-uscita">${formatImportoItaliano(totaleUsciteFiltrate)}</span>`;
    }
    if (footerEstrattoEntrate) {
        footerEstrattoEntrate.innerHTML = `<span class="importo-entrata">${formatImportoItaliano(totaleEntrateFiltrate)}</span>`;
    }

    const rowsMastrinoVisibili = document.querySelectorAll('#table-mastrino tbody tr:not(.filter-hidden)');
    let totaleDareFiltrato = 0;
    let totaleAvereFiltrato = 0;

    rowsMastrinoVisibili.forEach(row => {
        const dare = parseFloat(row.dataset.dare) || 0;
        const avere = parseFloat(row.dataset.avere) || 0;
        totaleDareFiltrato += dare;
        totaleAvereFiltrato += avere;
    });

    const footerMastrinoDare = document.querySelector('#table-mastrino tfoot .col-dare');
    const footerMastrinoAvere = document.querySelector('#table-mastrino tfoot .col-avere');
    if (footerMastrinoDare) {
        footerMastrinoDare.innerHTML = `<span class="importo-dare">${formatImportoItaliano(totaleDareFiltrato)}</span>`;
    }
    if (footerMastrinoAvere) {
        footerMastrinoAvere.innerHTML = `<span class="importo-avere">${formatImportoItaliano(totaleAvereFiltrato)}</span>`;
    }
}

function updateFiltersButtonState() {
    const tipoFilter = document.querySelector('input[name="tipo-filter"]:checked')?.value || 'all';
    const statoFilter = document.querySelector('input[name="stato-filter"]:checked')?.value || 'all';

    const hasActiveFilters = (tipoFilter !== 'all' || statoFilter !== 'all');

    const btnReset = document.getElementById('btn-reset-filters');
    if (btnReset) {
        if (hasActiveFilters) {
            btnReset.classList.add('active');
            btnReset.title = 'Resetta filtri attivi';
        } else {
            btnReset.classList.remove('active');
            btnReset.title = 'Nessun filtro attivo';
        }
    }

    const btnToggle = document.getElementById('filters-toggle');
    if (btnToggle) {
        if (hasActiveFilters) {
            btnToggle.classList.add('has-filters');
        } else {
            btnToggle.classList.remove('has-filters');
        }
    }
}

function toggleFilters() {
    const menu = document.getElementById('filters-menu');
    const toggle = document.getElementById('filters-toggle');
    menu.classList.toggle('show');
    toggle.classList.toggle('active');
}

function resetFilters() {
    document.querySelectorAll('input[name="tipo-filter"]').forEach(input => {
        input.checked = (input.value === 'all');
    });
    document.querySelectorAll('input[name="stato-filter"]').forEach(input => {
        input.checked = (input.value === 'all');
    });

    applyFilters();

    const menu = document.getElementById('filters-menu');
    const toggle = document.getElementById('filters-toggle');
    if (menu) menu.classList.remove('show');
    if (toggle) toggle.classList.remove('active');
}

document.addEventListener('click', function(e) {
    const filtersControls = document.getElementById('filters-controls');
    const menu = document.getElementById('filters-menu');
    const toggle = document.getElementById('filters-toggle');

    if (filtersControls && !filtersControls.contains(e.target) && menu && menu.classList.contains('show')) {
        menu.classList.remove('show');
        if (toggle) toggle.classList.remove('active');
    }
});

// === EXPORT EXCEL ===

function generaExcel() {
    console.log('📥 Generazione file Excel...');

    const wb = XLSX.utils.book_new();

    const entrate = risultati.entrate.filter(m => m.entrate);
    const uscite = risultati.entrate.filter(m => m.uscite);

    const entrateConc = entrate.filter(e => e.conciliato).length;
    const usciteConc = uscite.filter(u => u.conciliato).length;
    const commMast = uscite.reduce((s, u) => s + (u.commissione_mastrino || 0), 0);
    const commCalc = commissioniDaRegistrare.reduce((s, c) => s + c.importo, 0);

    const riepilogo = [
        ['RIEPILOGO CONCILIAZIONE BANCARIA v4'],
        ['Anno di riferimento', annoRiferimento],
        [],
        ['Descrizione', 'Quantità', 'Importo'],
        ['MOVIMENTI BANCARI', '', ''],
        ['Entrate totali', entrate.length, entrate.reduce((s,e) => s + (e.entrate||0), 0)],
        ['  - Conciliate', entrateConc, entrate.filter(e=>e.conciliato).reduce((s,e) => s + (e.entrate||0), 0)],
        ['  - NON conciliate', entrate.length - entrateConc, entrate.filter(e=>!e.conciliato).reduce((s,e) => s + (e.entrate||0), 0)],
        [''],
        ['Uscite totali', uscite.length, uscite.reduce((s,u) => s + (u.uscite||0), 0)],
        ['  - Conciliate', usciteConc, uscite.filter(u=>u.conciliato).reduce((s,u) => s + (u.uscite||0), 0)],
        ['  - NON conciliate', uscite.length - usciteConc, uscite.filter(u=>!u.conciliato).reduce((s,u) => s + (u.uscite||0), 0)],
        [''],
        ['MOVIMENTI MASTRINO', '', ''],
        ['DARE totali', risultati.dare.length, risultati.dare.reduce((s,d) => s + (d.dare||0), 0)],
        ['  - Conciliati', risultati.dare.filter(d=>d.conciliato).length, risultati.dare.filter(d=>d.conciliato).reduce((s,d) => s + (d.dare||0), 0)],
        ['  - NON conciliati', risultati.dare.filter(d=>!d.conciliato).length, risultati.dare.filter(d=>!d.conciliato).reduce((s,d) => s + (d.dare||0), 0)],
        [''],
        ['AVERE totali', risultati.avere.length, risultati.avere.reduce((s,a) => s + (a.avere||0), 0)],
        ['  - Conciliati', risultati.avere.filter(a=>a.conciliato).length, risultati.avere.filter(a=>a.conciliato).reduce((s,a) => s + (a.avere||0), 0)],
        ['  - NON conciliati', risultati.avere.filter(a=>!a.conciliato).length, risultati.avere.filter(a=>!a.conciliato).reduce((s,a) => s + (a.avere||0), 0)],
        [''],
        ['COMMISSIONI', '', ''],
        ['Commissioni nel mastrino', uscite.filter(u => u.commissione_mastrino > 0).length, commMast],
        ['Commissioni da registrare', commissioniDaRegistrare.length, commCalc],
        ['Totale commissioni', '', commMast + commCalc]
    ];

    const wsRiep = XLSX.utils.aoa_to_sheet(riepilogo);
    XLSX.utils.book_append_sheet(wb, wsRiep, 'Riepilogo');

    const estrattoData = [...entrate, ...uscite].sort((a, b) => {
        const progA = parseInt(a.prog.replace(/[^0-9]/g, '')) || 0;
        const progB = parseInt(b.prog.replace(/[^0-9]/g, '')) || 0;
        return progA - progB;
    }).map(row => ({
        'Stato': row.conciliato ? 'Conciliato' : 'Non Conciliato',
        'Prog Banca': row.prog,
        'Data': formatData(row.data),
        'Descrizione': row.descrizione || '',
        'Uscite': row.uscite || '',
        'Entrate': row.entrate || '',
        'Prog Mastrino': row.mastrino_progs || '',
        'Commissioni da registrare': row.commissione_calcolata || '',
        'N.Mov Mastrino': row.num_movimenti_mastrino || ''
    }));

    if (estrattoData.length > 0) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(estrattoData), 'Estratto Conto');
    }

    const mastrinoData = [...risultati.dare, ...risultati.avere].sort((a, b) => {
        const progA = parseInt(a.prog.replace(/[^0-9]/g, '')) || 0;
        const progB = parseInt(b.prog.replace(/[^0-9]/g, '')) || 0;
        return progA - progB;
    }).map(row => ({
        'Stato': row.conciliato ? 'Conciliato' : 'Non Conciliato',
        'Prog Mastrino': row.prog,
        'Data Reg': formatData(row.data),
        'Num Doc': row.num_doc || '',
        'Descrizione': row.descrizione || '',
        'DARE': row.dare || '',
        'AVERE': row.avere || '',
        'Prog Banca': row.banca_prog || '',
        'N.Mov Banca': row.num_movimenti_banca || ''
    }));

    if (mastrinoData.length > 0) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mastrinoData), 'Mastrino');
    }

    if (commissioniDaRegistrare.length > 0) {
        const data = commissioniDaRegistrare.map(c => ({
            'Prog Banca': c.prog_banca,
            'Data': formatData(c.data_banca),
            'Descrizione': c.descrizione,
            'Commissione': c.importo,
            'Importo Banca': c.importo_banca
        }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), 'Commissioni da Registrare');
    }

    XLSX.writeFile(wb, `Conciliazione_Bancaria_${annoRiferimento || 'anno'}_v4.xlsx`);
    console.log('✅ File Excel generato');
}

// === EVENT LISTENERS ===

setupDropzone('dropzone-estratto-sidebar', 'file-estratto-sidebar', 'info-estratto-sidebar', 'estratto');
setupDropzone('dropzone-mastrino-sidebar', 'file-mastrino-sidebar', 'info-mastrino-sidebar', 'mastrino');

document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

        tab.classList.add('active');
        document.getElementById('tab-' + tab.dataset.tab).classList.add('active');

        const filtersControls = document.getElementById('filters-controls');
        const tabName = tab.dataset.tab;
        if (filtersControls) {
            if (tabName === 'estratto-conto' || tabName === 'mastrino') {
                filtersControls.classList.remove('hidden');
            } else {
                filtersControls.classList.add('hidden');
                const filtersMenu = document.getElementById('filters-menu');
                const filtersToggle = document.getElementById('filters-toggle');
                if (filtersMenu) filtersMenu.classList.remove('show');
                if (filtersToggle) filtersToggle.classList.remove('active');
            }
        }
    });
});

// === GESTIONE DROPDOWN E PARAMETRI ===

let parametriIniziali = null;
let hasConciliatoUnaVolta = false;
let periodiSalvati = [];

function toggleDropdown(type) {
    const content = document.getElementById(`dropdown-${type}-content`);
    content.classList.toggle('show');

    document.querySelectorAll('.dropdown-content').forEach(d => {
        if (d.id !== `dropdown-${type}-content`) {
            d.classList.remove('show');
        }
    });
}

document.addEventListener('click', function(e) {
    if (!e.target.closest('.custom-dropdown')) {
        document.querySelectorAll('.dropdown-content').forEach(d => d.classList.remove('show'));
    }
});

function updateRangeValue(type, value) {
    document.getElementById(`val-${type}`).textContent = value;
}

async function avviaConciliazioneAutomatica() {
    resetFilters();

    const periodi = hasConciliatoUnaVolta ? periodiSalvati : getPeriodi();

    if (periodi.length === 0) return;

    const config = getParametriConfig();
    COMMISSIONI_AMMESSE = config.commissioni;

    const inizioOverlay = Date.now();
    mostraLoading();

    setTimeout(() => {
        try {
            risultati = concilia(
                [...estrattoConto],
                [...mastrino],
                config,
                periodi,
                (progress) => {
                    // Progress callback
                },
                hasConciliatoUnaVolta
            );

            mostraRisultati(risultati);

            if (!hasConciliatoUnaVolta) {
                showAlert('✅ Prima conciliazione completata!', 'success');
                salvaParametriIniziali();
                mostraConfigFab();
            } else {
                showAlert('✅ Conciliazione aggiornata!', 'success');
            }

            const tempoTrascorso = Date.now() - inizioOverlay;
            const delayMinimo = 800;
            const ritardoNecessario = Math.max(0, delayMinimo - tempoTrascorso);

            setTimeout(() => {
                nascondiLoading();
            }, ritardoNecessario);

        } catch (error) {
            console.error('❌ Errore conciliazione:', error);
            showAlert('❌ Errore: ' + error.message, 'error');
            nascondiLoading();
        }
    }, 100);
}

function updatePeriodoText() {
    const checkboxes = document.querySelectorAll('#dropdown-periodo-content input[type="checkbox"]:checked');
    const selected = Array.from(checkboxes).map(cb => cb.value);
    const text = selected.length > 0 ? selected.join(', ') : 'Seleziona...';
    document.getElementById('periodo-text').textContent = text;

    if (selected.length > 0 && !hasConciliatoUnaVolta) {
        setTimeout(() => {
            document.querySelectorAll('.dropdown-content').forEach(d => d.classList.remove('show'));
            avviaConciliazioneAutomatica();
        }, 300);
    } else {
        onConfigChange();
    }
}

function updateCommissioniText() {
    const checkboxes = document.querySelectorAll('#dropdown-commissioni-content input[type="checkbox"]:checked');
    const selected = Array.from(checkboxes).map(cb => parseFloat(cb.value)).sort((a,b) => a-b);
    const text = selected.length > 0 ? selected.join(', ') : 'Nessuna';
    document.getElementById('commissioni-text').textContent = text;
    onConfigChange();
}

function showAddCommission() {
    document.getElementById('add-commission-input').classList.toggle('show');
    event.stopPropagation();
}

function addCommission() {
    const input = document.getElementById('new-commission-value');
    const value = parseFloat(input.value);

    if (!value || value <= 0) {
        alert('Inserisci un valore valido maggiore di 0');
        return;
    }

    const exists = document.querySelector(`#dropdown-commissioni-content input[value="${value}"]`);
    if (exists) {
        alert('Commissione già presente');
        return;
    }

    const container = document.getElementById('dropdown-commissioni-content');
    const newItem = document.createElement('div');
    newItem.className = 'dropdown-item';
    newItem.innerHTML = `
        <input type="checkbox" id="comm-${value.toString().replace('.', '')}" value="${value}" checked onchange="updateCommissioniText()">
        <label for="comm-${value.toString().replace('.', '')}">${value.toFixed(2)}</label>
    `;
    container.appendChild(newItem);

    input.value = '';
    document.getElementById('add-commission-input').classList.remove('show');
    updateCommissioniText();
}

function getCommissioniAmmesse() {
    return Array.from(document.querySelectorAll('#commission-list input:checked')).map(c => {
        const text = c.nextElementSibling.textContent.replace('€ ', '');
        return parseFloat(text);
    });
}

function getPeriodi() {
    return Array.from(document.querySelectorAll('.tag.selected')).map(t => t.dataset.value);
}

function onConfigChange() {
    // Gestito nella sidebar
}

function getParametriConfig() {
    return {
        maxMovimenti: sliderValues.movimenti,
        tolleranzaImporto: sliderValues.importo,
        tolleranzaGiorni: sliderValues.giorni,
        periodi: getPeriodi(),
        commissioni: getCommissioniAmmesse()
    };
}

function salvaParametriIniziali() {
    parametriIniziali = getParametriConfig();
    periodiSalvati = getPeriodi();
    hasConciliatoUnaVolta = true;
    console.log('📌 Periodi salvati:', periodiSalvati);
}

console.log('✅ Conciliator Banca v4 pronto');
