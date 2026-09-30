'use strict';

// ─── State ───────────────────────────────────────────────────────────────────
let allCards = [];
let deck = [];
let deckPos = 0;
let isFlipped = false;
let knownSet = new Set();
let unknownSet = new Set();
let showOnlyUnknown = false;

let config = { columns: [], frontColumns: [], backColumns: [] };

// ─── DOM ─────────────────────────────────────────────────────────────────────
const importView      = document.getElementById('import-view');
const cardView        = document.getElementById('card-view');
const settingsPanel   = document.getElementById('settings-panel');
const settingsOverlay = document.getElementById('settings-overlay');

const fileInput       = document.getElementById('file-input');
const dropZone        = document.getElementById('drop-zone');
const browseBtn       = document.getElementById('browse-btn');
const importError     = document.getElementById('import-error');

const card            = document.getElementById('card');
const frontContent    = document.getElementById('front-content');
const backContent     = document.getElementById('back-content');
const progressText    = document.getElementById('progress-text');
const progressBar     = document.getElementById('progress-bar');

const prevBtn         = document.getElementById('prev-btn');
const nextBtn         = document.getElementById('next-btn');
const shuffleBtn      = document.getElementById('shuffle-btn');
const markKnownBtn    = document.getElementById('mark-known-btn');
const markUnknownBtn  = document.getElementById('mark-unknown-btn');
const knownCount      = document.getElementById('known-count');
const unknownCount    = document.getElementById('unknown-count');
const resetBtn        = document.getElementById('reset-btn');
const reimportBtn     = document.getElementById('reimport-btn');
const settingsBtn     = document.getElementById('settings-btn');

const closeSettingsBtn   = document.getElementById('close-settings-btn');
const applySettingsBtn   = document.getElementById('apply-settings-btn');
const frontColsConfig    = document.getElementById('front-cols-config');
const backColsConfig     = document.getElementById('back-cols-config');
const showOnlyUnknownChk = document.getElementById('show-only-unknown');

// ─── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Register service worker
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  const stored = loadFromStorage();
  if (stored.cards && stored.cards.length > 0) {
    allCards = stored.cards;
    config = stored.config || config;
    knownSet = new Set(stored.known || []);
    unknownSet = new Set(stored.unknown || []);
    enterCardView();
  } else {
    show(importView);
  }

  setupImportListeners();
  setupCardListeners();
  setupSettingsListeners();
});

// ─── Storage (localStorage) ───────────────────────────────────────────────────
function persistAll() {
  try {
    localStorage.setItem('fc_cards', JSON.stringify(allCards));
    localStorage.setItem('fc_config', JSON.stringify(config));
    localStorage.setItem('fc_known', JSON.stringify([...knownSet]));
    localStorage.setItem('fc_unknown', JSON.stringify([...unknownSet]));
  } catch (e) {
    console.warn('Storage full:', e);
  }
}

function loadFromStorage() {
  try {
    return {
      cards:   JSON.parse(localStorage.getItem('fc_cards') || 'null'),
      config:  JSON.parse(localStorage.getItem('fc_config') || 'null'),
      known:   JSON.parse(localStorage.getItem('fc_known') || '[]'),
      unknown: JSON.parse(localStorage.getItem('fc_unknown') || '[]'),
    };
  } catch { return {}; }
}

// ─── Import ───────────────────────────────────────────────────────────────────
function setupImportListeners() {
  browseBtn.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('click', e => { if (e.target !== browseBtn) fileInput.click(); });
  fileInput.addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });

  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
}

function handleFile(file) {
  if (!file.name.match(/\.xlsx?$/i)) { showError('Chỉ hỗ trợ file .xlsx hoặc .xls'); return; }
  const reader = new FileReader();
  reader.onload = e => {
    try { parseExcel(e.target.result); }
    catch (err) { showError('Không đọc được file: ' + err.message); }
  };
  reader.readAsArrayBuffer(file);
}

function parseExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // Auto-detect header row
  let headerIdx = -1;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    if (rows[i].filter(c => c !== '').length >= 2) { headerIdx = i; break; }
  }
  if (headerIdx === -1) { showError('Không tìm thấy header trong file Excel'); return; }

  const colDefs = [];
  rows[headerIdx].forEach((h, idx) => {
    const name = String(h || '').trim();
    if (name) colDefs.push({ key: name, colIdx: idx });
  });
  if (!colDefs.length) { showError('Không tìm thấy cột dữ liệu'); return; }

  const parsed = rows.slice(headerIdx + 1)
    .filter(row => row.some(c => c !== ''))
    .map(row => {
      const obj = {};
      colDefs.forEach(({ key, colIdx }) => { obj[key] = String(row[colIdx] ?? '').trim(); });
      return obj;
    })
    .filter(obj => Object.values(obj).some(v => v));

  if (!parsed.length) { showError('Không có dữ liệu trong file'); return; }

  allCards = parsed;
  const colNames = colDefs.map(d => d.key);
  const vietCols = colNames.filter(c => /vi(et|[eệ]t)?/i.test(c));
  const frontDefault = vietCols.length ? vietCols : [colNames[0]];

  config = { columns: colNames, frontColumns: frontDefault, backColumns: colNames.filter(c => !frontDefault.includes(c)) };
  knownSet = new Set();
  unknownSet = new Set();

  persistAll();
  enterCardView();
}

function showError(msg) {
  importError.textContent = msg;
  importError.classList.remove('hidden');
}

// ─── Card View ─────────────────────────────────────────────────────────────────
function enterCardView() {
  show(cardView);
  hide(importView);
  buildDeck();
  renderCard();
}

function buildDeck() {
  let indices = allCards.map((_, i) => i).filter(i => !showOnlyUnknown || !knownSet.has(i));
  for (let i = indices.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [indices[i], indices[j]] = [indices[j], indices[i]];
  }
  deck = indices;
  deckPos = 0;
}

function renderCard() {
  if (!deck.length) {
    frontContent.innerHTML = '<div class="field-value" style="font-size:18px;opacity:.7">Không có thẻ nào!</div>';
    backContent.innerHTML = '';
    progressText.textContent = '0 / 0';
    progressBar.style.width = '0%';
    return;
  }

  isFlipped = false;
  card.classList.remove('flipped');

  const data = allCards[deck[deckPos]];

  frontContent.innerHTML = config.frontColumns
    .filter(col => data[col])
    .map(col => `<div class="field-block"><div class="field-value">${esc(data[col])}</div></div>`)
    .join('') || '<div class="field-value" style="opacity:.4">(trống)</div>';

  backContent.innerHTML = config.backColumns
    .filter(col => data[col])
    .map(col => `<div class="field-block">
      <div class="field-label">${esc(col)}</div>
      <div class="${fieldClass(col)}">${esc(data[col])}</div>
    </div>`)
    .join('') || '<div class="field-value-default" style="opacity:.4">(trống)</div>';

  const pos = deckPos + 1;
  progressText.textContent = `${pos} / ${deck.length}`;
  progressBar.style.width = `${(pos / deck.length) * 100}%`;
  updateStats();
}

function fieldClass(col) {
  const c = col.toLowerCase();
  if (/phien.?am|phonet|ipa|pronun/.test(c)) return 'field-value-phonetic';
  if (/vi.?du|example|ex\.?/.test(c))        return 'field-value-example';
  if (/anh|eng|english/.test(c))             return 'field-value-eng';
  return 'field-value-default';
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─── Card Listeners ────────────────────────────────────────────────────────────
function setupCardListeners() {
  card.addEventListener('click', () => {
    isFlipped = !isFlipped;
    card.classList.toggle('flipped', isFlipped);
  });

  nextBtn.addEventListener('click', () => {
    if (!deck.length) return;
    deckPos = (deckPos + 1) % deck.length;
    renderCard();
  });

  prevBtn.addEventListener('click', () => {
    if (!deck.length) return;
    deckPos = (deckPos - 1 + deck.length) % deck.length;
    renderCard();
  });

  shuffleBtn.addEventListener('click', () => { buildDeck(); renderCard(); });

  markKnownBtn.addEventListener('click', () => {
    if (!deck.length) return;
    const idx = deck[deckPos];
    knownSet.add(idx); unknownSet.delete(idx);
    persistAll(); goNext();
  });

  markUnknownBtn.addEventListener('click', () => {
    if (!deck.length) return;
    const idx = deck[deckPos];
    unknownSet.add(idx); knownSet.delete(idx);
    persistAll(); goNext();
  });

  resetBtn.addEventListener('click', () => {
    knownSet = new Set(); unknownSet = new Set();
    persistAll(); buildDeck(); renderCard();
  });

  reimportBtn.addEventListener('click', () => {
    show(importView); hide(cardView);
    importError.classList.add('hidden');
  });

  settingsBtn.addEventListener('click', openSettings);
}

function goNext() {
  if (showOnlyUnknown && knownSet.has(deck[deckPos])) {
    deck.splice(deckPos, 1);
    if (!deck.length) { renderCard(); return; }
    deckPos = deckPos % deck.length;
  } else {
    deckPos = (deckPos + 1) % deck.length;
  }
  renderCard();
}

function updateStats() {
  knownCount.textContent = `✓ ${knownSet.size}`;
  unknownCount.textContent = `✗ ${unknownSet.size}`;
}

// ─── Settings ──────────────────────────────────────────────────────────────────
function setupSettingsListeners() {
  closeSettingsBtn.addEventListener('click', closeSettings);
  settingsOverlay.addEventListener('click', closeSettings);
  applySettingsBtn.addEventListener('click', applySettings);
}

function openSettings() {
  frontColsConfig.innerHTML = config.columns.map(col => `
    <div class="col-item">
      <input type="checkbox" id="fr-${sid(col)}" data-col="${esc(col)}" class="front-col-chk" ${config.frontColumns.includes(col) ? 'checked' : ''}>
      <label for="fr-${sid(col)}">${esc(col)}</label>
    </div>`).join('');

  backColsConfig.innerHTML = config.columns.map(col => `
    <div class="col-item">
      <input type="checkbox" id="bk-${sid(col)}" data-col="${esc(col)}" class="back-col-chk" ${config.backColumns.includes(col) ? 'checked' : ''}>
      <label for="bk-${sid(col)}">${esc(col)}</label>
    </div>`).join('');

  showOnlyUnknownChk.checked = showOnlyUnknown;
  settingsPanel.classList.remove('hidden');
  settingsOverlay.classList.remove('hidden');
}

function closeSettings() {
  settingsPanel.classList.add('hidden');
  settingsOverlay.classList.add('hidden');
}

function applySettings() {
  const front = [...document.querySelectorAll('.front-col-chk:checked')].map(el => el.dataset.col);
  const back  = [...document.querySelectorAll('.back-col-chk:checked')].map(el => el.dataset.col);
  if (!front.length && !back.length) return;
  config.frontColumns = front;
  config.backColumns  = back;
  showOnlyUnknown = showOnlyUnknownChk.checked;
  persistAll(); buildDeck(); renderCard(); closeSettings();
}

function sid(s) { return s.replace(/[^a-z0-9]/gi, '_'); }

// ─── Helpers ───────────────────────────────────────────────────────────────────
function show(el) { el.classList.remove('hidden'); }
function hide(el) { el.classList.add('hidden'); }
