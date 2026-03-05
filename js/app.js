/* ============================================================================
   app.js - 主应用逻辑 (Main Application)
   依赖: config.js, core.js, i18n.js, license.js (按 index.html 引入顺序加载)
   ============================================================================ */

/* ----------------------------------------------------------------------------
   业务辅助 (依赖 appSettings，保留在 app.js)
   ---------------------------------------------------------------------------- */
function getGradeLabels() {
    const gl = appSettings.gradeLabels;
    if (Array.isArray(gl) && gl.length === 6) return gl;
    return [...GRADES];
}
function getPriceForGrade(grade) {
    if (!grade) return 0;
    const labels = getGradeLabels();
    const slotIndex = labels.indexOf(grade);
    if (slotIndex >= 0 && appSettings.priceByGrade) {
        return parseFloat(appSettings.priceByGrade[GRADES[slotIndex]]) || 0;
    }
    return parseFloat(appSettings.priceByGrade?.[grade]) || 0;
}

function addToHistoryCompany(c) {
    const name = getCompanyName(c).trim();
    if (!name) return;
    const idx = (histories.company || []).findIndex(x => getCompanyName(x) === name);
    if (idx >= 0) histories.company[idx] = normalizeCompany(c); else histories.company.push(normalizeCompany(c));
    localStorage.setItem(HIST_KEY, JSON.stringify(histories));
}
function addToHistorySeller(s) {
    const name = getCompanyName(s).trim();
    if (!name) return;
    const idx = (histories.seller || []).findIndex(x => getCompanyName(x) === name);
    if (idx >= 0) histories.seller[idx] = normalizeSeller(s); else histories.seller.push(normalizeSeller(s));
    localStorage.setItem(HIST_KEY, JSON.stringify(histories));
}
function updateSellerInHistory(originalSeller, newSeller) {
    const origName = getCompanyName(originalSeller).trim();
    if (!origName) return;
    const arr = histories.seller || [];
    const idx = arr.findIndex(x => getCompanyName(x) === origName);
    if (idx >= 0) arr[idx] = normalizeSeller(newSeller);
    else arr.push(normalizeSeller(newSeller));
    histories.seller = arr;
    localStorage.setItem(HIST_KEY, JSON.stringify(histories));
}

/* ============================================================================
   [2] THEME MANAGEMENT
   ============================================================================ */

function initTheme() {
    const savedTheme = localStorage.getItem('lmp_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);
    updateThemeToggleLabel();
}

function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
    const newTheme = currentTheme === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('lmp_theme', newTheme);
    updateThemeIcon(newTheme);
    updateThemeToggleLabel();
}

function updateThemeIcon(theme) {
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function updateThemeToggleLabel() {
    const theme = document.documentElement.getAttribute('data-theme') || 'dark';
    const label = document.getElementById('themeToggleLabel');
    const iconMoon = document.getElementById('themeIconMoon');
    const iconSun = document.getElementById('themeIconSun');
    if (label) label.textContent = theme === 'dark' ? (I18N[currentLang].theme_dark || '夜间') : (I18N[currentLang].theme_light || '白天');
    if (iconMoon) iconMoon.style.display = theme === 'dark' ? 'inline-block' : 'none';
    if (iconSun) iconSun.style.display = theme === 'light' ? 'inline-block' : 'none';
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

/* ============================================================================
   [3] STATE MANAGEMENT
   ============================================================================ */

/* ----------------------------------------------------------------------------
   [2.1] 核心数据状态 (Core Data State)
   ---------------------------------------------------------------------------- */
let logs = [];
let isMarkMode = false;
let isGroupMode = false;
let groupSelectIds = [];
let globalInfo = { container: '', note: '', description: '', measurer: '', location: '', company: defaultCompany(), seller: defaultSeller() };
let histories = { container: [], seller: [], measurer: [], location: [], company: [] };
/* ----------------------------------------------------------------------------
   [2.2] 应用设置状态 (Application Settings State)
   ---------------------------------------------------------------------------- */
let appSettings = {
    beginnerMode: false,
    deductLen: 0,
    deductDia: 0,
    showGrade: true,
    calcDia: false,
    roundMode: 'up',
    formulaEnabled: false,
    formula: 'huber',
    gradeLabels: null,
    statThresholdL4: 4,
    statThresholdL25: 2.5,
    statThresholdD30: 30,
    proKeyboard: false,
    useVirtualKeyboard: false,
    quickModeAutoDecimal: true,
    quickModeAutoJump: true,
    quickModeGradeAutoSave: true,
    quickModeAutoCode: true,
    priceEnabled: false,
    priceCurrency: 'EUR',
    priceMode: 'fixed',
    priceFixed: 0,
    priceByGrade: { 'F': 0, 'A+': 0, 'A': 0, 'B': 0, 'C': 0, 'D': 0 },
    taxPercent: 0,
    showPricePdf: false,
    showPriceCsv: false,
    showMarksInExport: true,
    showGroupInExport: true,
    showCompanyInPdf: true,
    keySound: true,
    volumeWarningEnabled: false,
    volumeWarningThreshold: 21
};

/* ----------------------------------------------------------------------------
   [2.3] UI 状态与模式 (UI State & Modes)
   ---------------------------------------------------------------------------- */
let activeFilter = null;
let currentLang = 'zh';
let isQuickMode = false;
let nextRoundUp = true;

/* ----------------------------------------------------------------------------
   [2.4] 会话与快照管理 (Session & Snapshot Management)
   ---------------------------------------------------------------------------- */
let currentSessionId = null;
let snapshots = [];
let historyViewerState = {
    snapshotId: null,
    logs: [],
    global: {},
    sourceDate: '',
    reversedOrder: false
};

/* ----------------------------------------------------------------------------
   [2.5] 专业键盘状态 (Professional Keyboard State)
   ---------------------------------------------------------------------------- */
let proState = {
    activeField: 'length',
    keypadMode: 'num',
    values: { code: '', length: '', dia: '', dia1: '', dia2: '', note: '' },
    gradeReady: false
};
let proKeyDebounceTimer = null;
let proKeyPending = null;

/* ============================================================================
   [3] INITIALIZATION (初始化)
   ============================================================================ */

function normalizeGlobalInfo(g) {
    if (!g) return;
    g.company = normalizeCompany(g.company);
    g.seller = normalizeSeller(g.seller);
}
function syncInfoModalCompanyDisplay() {
    const gCompany = document.getElementById('g_company');
    const gSeller = document.getElementById('g_seller');
    if (gCompany) gCompany.value = getCompanyName(globalInfo.company);
    if (gSeller) gSeller.value = getCompanyName(globalInfo.seller);
}
let companyModalSnapshot = null;
let sellerModalSnapshot = null;
function getCompanyFormValues(prefix) {
    const o = {};
    COMPANY_FIELDS.forEach(f => {
        const el = document.getElementById(prefix + f);
        if (el) o[f] = (el.value || '').trim();
    });
    return o;
}
function setCompanyFormValues(prefix, obj) {
    COMPANY_FIELDS.forEach(f => {
        const el = document.getElementById(prefix + f);
        if (el) el.value = (obj && obj[f]) || '';
    });
}
function isCompanyDataEmpty(obj) { return !obj || COMPANY_FIELDS.every(f => !(obj[f] && String(obj[f]).trim())); }
function isCompanyFormDirty(prefix, snapshot) {
    const cur = getCompanyFormValues(prefix);
    return COMPANY_FIELDS.some(f => (cur[f] || '') !== (snapshot[f] || ''));
}
function openCompanyDetailModal() {
    const c = globalInfo.company || defaultCompany();
    setCompanyFormValues('cd_', c);
    companyModalSnapshot = getCompanyFormValues('cd_');
    document.getElementById('companyDetailModal').style.display = 'flex';
}
function closeCompanyDetailModal(force) {
    if (!force && companyModalSnapshot && isCompanyDataEmpty(companyModalSnapshot) === false && isCompanyFormDirty('cd_', companyModalSnapshot)) {
        if (confirm(I18N[currentLang].confirm_company_save || '有未保存的修改，是否保存？')) {
            saveCompanyDetailModal();
        }
    }
    companyModalSnapshot = null;
    document.getElementById('companyDetailModal').style.display = 'none';
}
function saveCompanyDetailModal() {
    const c = defaultCompany();
    COMPANY_FIELDS.forEach(f => {
        const el = document.getElementById('cd_' + f);
        if (el) c[f] = (el.value || '').trim();
    });
    globalInfo.company = c;
    syncInfoModalCompanyDisplay();
    addToHistoryCompany(c);
    save();
}
let sellerModalContext = 'main';
let sellerModalIsNew = false;
let sellerModalOriginal = null;
function openSellerDetailModal() {
    sellerModalContext = 'main';
    sellerModalIsNew = false;
    const s = globalInfo.seller || defaultSeller();
    sellerModalOriginal = s && !isCompanyDataEmpty(s) ? Object.assign({}, s) : null;
    setCompanyFormValues('sd_', s);
    document.getElementById('sd_type_buyer').checked = (s.type === 'seller') ? false : true;
    document.getElementById('sd_type_seller').checked = (s.type === 'seller');
    sellerModalSnapshot = getCompanyFormValues('sd_');
    document.getElementById('sellerDetailModal').style.display = 'flex';
}
function openSellerDetailModalAsNew() {
    sellerModalContext = 'main';
    sellerModalIsNew = true;
    sellerModalOriginal = null;
    const empty = defaultSeller();
    setCompanyFormValues('sd_', empty);
    document.getElementById('sd_type_buyer').checked = true;
    document.getElementById('sd_type_seller').checked = false;
    sellerModalSnapshot = getCompanyFormValues('sd_');
    document.getElementById('sellerDetailModal').style.display = 'flex';
    if (typeof lucide !== 'undefined') lucide.createIcons();
}
function openSellerDetailModalForHistoryViewer() {
    sellerModalContext = 'historyViewer';
    sellerModalIsNew = false;
    const s = historyViewerState.global?.seller || defaultSeller();
    sellerModalOriginal = s && !isCompanyDataEmpty(s) ? Object.assign({}, s) : null;
    setCompanyFormValues('sd_', s);
    document.getElementById('sd_type_buyer').checked = (s.type === 'seller') ? false : true;
    document.getElementById('sd_type_seller').checked = (s.type === 'seller');
    sellerModalSnapshot = getCompanyFormValues('sd_');
    const modal = document.getElementById('sellerDetailModal');
    modal.classList.add('above-history-viewer');
    modal.style.display = 'flex';
}
function clearSellerFormAndNew() {
    sellerModalIsNew = true;
    sellerModalOriginal = null;
    const empty = defaultSeller();
    setCompanyFormValues('sd_', empty);
    document.getElementById('sd_type_buyer').checked = true;
    document.getElementById('sd_type_seller').checked = false;
    sellerModalSnapshot = getCompanyFormValues('sd_');
}
function closeSellerDetailModal(force) {
    if (!force && sellerModalSnapshot && isCompanyDataEmpty(sellerModalSnapshot) === false && isCompanyFormDirty('sd_', sellerModalSnapshot)) {
        if (confirm(I18N[currentLang].confirm_company_save || '有未保存的修改，是否保存？')) {
            saveSellerDetailModal();
        }
    }
    sellerModalSnapshot = null;
    sellerModalContext = 'main';
    sellerModalIsNew = false;
    sellerModalOriginal = null;
    const modal = document.getElementById('sellerDetailModal');
    modal.classList.remove('above-history-viewer');
    modal.style.display = 'none';
}
function saveSellerDetailModal() {
    const s = defaultSeller();
    COMPANY_FIELDS.forEach(f => {
        const el = document.getElementById('sd_' + f);
        if (el) s[f] = (el.value || '').trim();
    });
    s.type = document.getElementById('sd_type_seller').checked ? 'seller' : 'buyer';
    if (sellerModalIsNew) {
        addToHistorySeller(s);
    } else if (sellerModalOriginal) {
        updateSellerInHistory(sellerModalOriginal, s);
    } else if (getCompanyName(s).trim()) {
        addToHistorySeller(s);
    }
    if (sellerModalContext === 'historyViewer') {
        if (!historyViewerState.global) historyViewerState.global = {};
        historyViewerState.global.seller = s;
        const hvSeller = document.getElementById('hv_g_seller');
        if (hvSeller) hvSeller.value = getCompanyName(s);
    } else {
        globalInfo.seller = s;
        syncInfoModalCompanyDisplay();
        save();
    }
    sellerModalContext = 'main';
    sellerModalIsNew = false;
    sellerModalOriginal = null;
}

function onDomReady(fn) {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', fn);
    } else {
        fn();
    }
}
onDomReady(initLicenseCheck);

/* 触觉反馈 - 2025/2026 风格 */
function haptic() {
    try { if (navigator.vibrate) navigator.vibrate(10); } catch (e) { }
}
onDomReady(function () {
    document.addEventListener('click', function (e) {
        if (e.target.closest('button, .btn-header, .kb-key, .save-menu-item, .stat-tag, .btn-grade, .toggle-btn, .btn-tool, .btn-block, .btn-title-action')) {
            haptic();
        }
    }, { passive: true });
});

window.onload = () => {
    initTheme(); // Initialize theme
    const savedData = localStorage.getItem(STORAGE_KEY);
    if (savedData) {
        const p = JSON.parse(savedData);
        logs = p.logs || [];
        globalInfo = p.global || globalInfo;
        normalizeGlobalInfo(globalInfo);
        const gContainer = document.getElementById('g_container');
        const gNote = document.getElementById('g_note');
        const gCompany = document.getElementById('g_company');
        const gSeller = document.getElementById('g_seller');
        const gLocation = document.getElementById('g_location');
        const gMeasurer = document.getElementById('g_measurer');
        if (gContainer) gContainer.value = globalInfo.container || '';
        if (gNote) gNote.value = globalInfo.note || '';
        const gDesc = document.getElementById('g_description');
        if (gDesc) gDesc.value = globalInfo.description || '';
        if (gCompany) gCompany.value = getCompanyName(globalInfo.company);
        if (gSeller) gSeller.value = getCompanyName(globalInfo.seller);
        if (gLocation) gLocation.value = globalInfo.location || '';
        if (gMeasurer) gMeasurer.value = globalInfo.measurer || '';
    }
    const savedHist = localStorage.getItem(HIST_KEY);
    if (savedHist) histories = JSON.parse(savedHist);
    ['container', 'seller', 'measurer', 'location', 'company', 'description'].forEach(k => { if (!histories[k]) histories[k] = []; });
    migrateHistoriesSeller(histories);
    migrateHistoriesCompany(histories);

    const savedLang = localStorage.getItem(LANG_KEY);
    if (savedLang) { currentLang = savedLang; document.getElementById('langSelect').value = currentLang; }

    const savedQuick = localStorage.getItem(QUICK_KEY);
    if (savedQuick === 'true') { isQuickMode = true; }

    const savedSet = localStorage.getItem(SETTINGS_KEY);
    if (savedSet) { appSettings = Object.assign(appSettings, JSON.parse(savedSet)); }
    if (!Array.isArray(appSettings.gradeLabels) || appSettings.gradeLabels.length !== 6) {
        appSettings.gradeLabels = [...GRADES];
    }
    appSettings.useVirtualKeyboard = !!appSettings.proKeyboard;

    if (!appSettings.priceByGrade) appSettings.priceByGrade = {};
    GRADES.forEach(g => { if (appSettings.priceByGrade[g] === undefined) appSettings.priceByGrade[g] = 0; });
    if (appSettings.priceEnabled === undefined) {
        const hasPrice = (parseFloat(appSettings.priceFixed) || 0) > 0 ||
            (parseFloat(appSettings.taxPercent) || 0) > 0 ||
            GRADES.some(g => (parseFloat(appSettings.priceByGrade[g]) || 0) > 0);
        appSettings.priceEnabled = hasPrice;
    }

    const savedMix = localStorage.getItem(MIX_STATE_KEY);
    if (savedMix) { nextRoundUp = (savedMix === 'true'); }

    document.getElementById('setDeductLen').value = appSettings.deductLen || 0;
    document.getElementById('setDeductDia').value = appSettings.deductDia || 0;
    const vwEnabled = document.getElementById('volumeWarningEnabled');
    if (vwEnabled) vwEnabled.checked = !!appSettings.volumeWarningEnabled;
    const vwThreshold = document.getElementById('volumeWarningThreshold');
    if (vwThreshold) vwThreshold.value = appSettings.volumeWarningThreshold ?? 21;
    const cbBeginner = document.getElementById('checkBeginnerMode');
    if (cbBeginner) cbBeginner.checked = !!appSettings.beginnerMode;
    document.getElementById('checkShowGrade').checked = appSettings.showGrade;
    document.getElementById('checkCalcDia').checked = appSettings.calcDia || false;
    updateProKeyboardBtnUI();
    document.getElementById('priceEnabled').checked = !!appSettings.priceEnabled;
    document.getElementById('roundSelect').value = appSettings.roundMode || 'up';
    const formulaEnabledCb = document.getElementById('formulaEnabled');
    if (formulaEnabledCb) formulaEnabledCb.checked = !!appSettings.formulaEnabled;
    const formulaSelect = document.getElementById('formulaSelect');
    if (formulaSelect) formulaSelect.value = appSettings.formula || 'huber';
    toggleFormulaEnabledUI();
    document.getElementById('priceCurrency').value = appSettings.priceCurrency || 'EUR';
    document.getElementById('priceMode').value = appSettings.priceMode || 'fixed';
    document.getElementById('priceFixed').value = appSettings.priceFixed || 0;
    document.getElementById('priceTax').value = appSettings.taxPercent || 0;
    document.getElementById('priceShowPdf').checked = !!appSettings.showPricePdf;
    document.getElementById('priceShowCsv').checked = !!appSettings.showPriceCsv;
    const cbMarks = document.getElementById('exportShowMarks');
    const cbGroup = document.getElementById('exportShowGroup');
    if (cbMarks) cbMarks.checked = appSettings.showMarksInExport !== false;
    if (cbGroup) cbGroup.checked = appSettings.showGroupInExport !== false;
    const cbKeySound = document.getElementById('checkKeySound');
    if (cbKeySound) cbKeySound.checked = appSettings.keySound !== false;
    if (appSettings.quickModeAutoDecimal === undefined) appSettings.quickModeAutoDecimal = true;
    if (appSettings.quickModeAutoJump === undefined) appSettings.quickModeAutoJump = true;
    if (appSettings.quickModeGradeAutoSave === undefined) appSettings.quickModeGradeAutoSave = true;
    if (appSettings.quickModeAutoCode === undefined) appSettings.quickModeAutoCode = true;
    const qmAutoDec = document.getElementById('quickModeAutoDecimal');
    const qmAutoJump = document.getElementById('quickModeAutoJump');
    const qmGradeSave = document.getElementById('quickModeGradeAutoSave');
    const qmAutoCode = document.getElementById('quickModeAutoCode');
    if (qmAutoDec) qmAutoDec.checked = appSettings.quickModeAutoDecimal;
    if (qmAutoJump) qmAutoJump.checked = appSettings.quickModeAutoJump;
    if (qmGradeSave) qmGradeSave.checked = appSettings.quickModeGradeAutoSave;
    if (qmAutoCode) qmAutoCode.checked = appSettings.quickModeAutoCode;
    document.getElementById('price_grade_F').value = appSettings.priceByGrade['F'] || 0;
    document.getElementById('price_grade_Aplus').value = appSettings.priceByGrade['A+'] || 0;
    document.getElementById('price_grade_A').value = appSettings.priceByGrade['A'] || 0;
    document.getElementById('price_grade_B').value = appSettings.priceByGrade['B'] || 0;
    document.getElementById('price_grade_C').value = appSettings.priceByGrade['C'] || 0;
    document.getElementById('price_grade_D').value = appSettings.priceByGrade['D'] || 0;

    // 加载快照数据和会话 ID
    const savedSnapshots = localStorage.getItem(SNAPSHOTS_KEY);
    if (savedSnapshots) snapshots = JSON.parse(savedSnapshots);
    const savedSession = localStorage.getItem(SESSION_KEY);
    if (savedSession) currentSessionId = savedSession;

    applyLanguage();
    updateQuickBtnUI();
    updateBeginnerModeUI();
    toggleCalcDia();
    updatePriceModeUI();
    updatePriceEnabledUI();
    updateCurrencySymbols();
    applyProKeyboardUI();
    updateKeySoundRowVisibility();
    if (logs.length === 0) addNewLog();
    updateGroupBtnUI();
    renderAll();

    // 离线支持：注册 Service Worker
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./sw.js').catch(() => { });
    }
};

function changeLanguage(lang) { currentLang = lang; localStorage.setItem(LANG_KEY, lang); applyLanguage(); renderAll(); updateQuickBtnUI(); updateProKeyboardBtnUI(); updateThemeToggleLabel(); }
function applyLanguage() {
    const texts = I18N[currentLang];
    document.querySelectorAll('[data-i18n]').forEach(el => { const key = el.getAttribute('data-i18n'); if (texts[key]) el.innerText = texts[key]; });
    document.querySelectorAll('[data-title-i18n]').forEach(el => { const key = el.getAttribute('data-title-i18n'); if (texts[key]) { el.title = texts[key]; el.setAttribute('aria-label', texts[key]); } });
    // 内部记录模式：PDF/Excel/关闭 按钮仅图标，需动态更新 title/aria-label
    const hvPdf = document.querySelector('.hv-pdf'); if (hvPdf && texts.btn_pdf) { hvPdf.title = texts.btn_pdf; hvPdf.setAttribute('aria-label', texts.btn_pdf); }
    const hvExcel = document.querySelector('.hv-excel'); if (hvExcel && texts.btn_excel) { hvExcel.title = texts.btn_excel; hvExcel.setAttribute('aria-label', texts.btn_excel); }
    const hvClose = document.querySelector('.hv-close'); if (hvClose && texts.hv_close) { hvClose.title = texts.hv_close; hvClose.setAttribute('aria-label', texts.hv_close); }
    renderProKeyboardTopBar();
    renderProGradePanel();
    updateGroupBtnUI();
    if (typeof lucide !== 'undefined') lucide.createIcons();
    const viewer = document.getElementById('historyViewer');
    if (viewer && viewer.style.display === 'flex' && historyViewerState && historyViewerState.snapshotId) {
        renderHistoryViewer();
        const meta = document.getElementById('historyViewerMeta');
        const snap = snapshots.find(s => s.id === historyViewerState.snapshotId);
        if (meta && snap && texts) {
            const containerDisplay = (snap.container && snap.container !== '未命名') ? snap.container : texts.hv_unnamed;
            meta.innerText = `ID: ${snap.id} | ${texts.hv_container_label}: ${containerDisplay} | ${texts.hv_time}: ${snap.date || '-'}`;
        }
    }
}
function updateGlobal() {
    globalInfo.container = (document.getElementById('g_container')?.value || '').trim();
    globalInfo.note = (document.getElementById('g_note')?.value || '').trim();
    globalInfo.description = (document.getElementById('g_description')?.value || '').trim();
    globalInfo.measurer = (document.getElementById('g_measurer')?.value || '').trim();
    globalInfo.location = (document.getElementById('g_location')?.value || '').trim();
    save();
}
function save() { localStorage.setItem(STORAGE_KEY, JSON.stringify({ logs, global: globalInfo })); updateStats(); }

function downloadFullBackup() {
    updateGlobal();
    normalizeGlobalInfo(globalInfo);
    const normSnapshots = (snapshots || []).map(s => {
        const g = Object.assign({}, s.global || {});
        g.company = normalizeCompany(g.company);
        g.seller = normalizeSeller(g.seller);
        return Object.assign({}, s, { global: g });
    });
    const payload = {
        schema: 'logmetric_backup_v1',
        exportedAt: Date.now(),
        appName: 'LogMetric Pro',
        logs: JSON.parse(JSON.stringify(logs || [])),
        global: JSON.parse(JSON.stringify(globalInfo || {})),
        snapshots: JSON.parse(JSON.stringify(normSnapshots)),
        histories: JSON.parse(JSON.stringify(histories || {})),
        appSettings: JSON.parse(JSON.stringify(appSettings || {})),
        currentSessionId: currentSessionId || null
    };
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}-${String(d.getMinutes()).padStart(2, '0')}`;
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `LogMetricPro_Backup_${stamp}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 600);
    alert(currentLang === 'zh' ? '备份已下载到您的设备' : (currentLang === 'en' ? 'Backup downloaded to your device' : 'Varnostna kopija prenesena'));
}

function downloadExportData() {
    updateGlobal();
    const payload = {
        schema: 'oak_project_file_v1',
        exportedAt: Date.now(),
        currentSessionId: currentSessionId || null,
        logs: JSON.parse(JSON.stringify(logs || [])),
        global: JSON.parse(JSON.stringify(globalInfo || {}))
    };
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}-${String(d.getMinutes()).padStart(2, '0')}`;
    const container = (globalInfo.container || 'Project').replace(/[^\w\u4e00-\u9fa5-]+/g, '_');
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `LogMetricPro_Export_${container}_${stamp}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 600);
    alert(currentLang === 'zh' ? '导出已下载' : (currentLang === 'en' ? 'Export downloaded' : 'Izvoz prenesen'));
}

async function exportPackageAll() {
    updateGlobal();
    closeSaveMenu();
    const d = new Date();
    const timeStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}-${String(d.getMinutes()).padStart(2, '0')}`;
    const container = (globalInfo.container || 'Project').replace(/[^\w\u4e00-\u9fa5-]+/g, '_');
    try {
        const zip = new JSZip();
        const jsonPayload = {
            schema: 'oak_project_file_v1',
            exportedAt: Date.now(),
            currentSessionId: currentSessionId || null,
            logs: JSON.parse(JSON.stringify(logs || [])),
            global: JSON.parse(JSON.stringify(globalInfo || {}))
        };
        zip.file(`${container}_${timeStr}.json`, JSON.stringify(jsonPayload, null, 2));
        const pdfBlob = await generatePDF({ returnBlob: true });
        if (pdfBlob) zip.file(`${container}_${timeStr}.pdf`, pdfBlob);
        const excelBlob = await exportData({ returnBlob: true });
        if (excelBlob) zip.file(`${container}_${timeStr}.xlsx`, excelBlob);
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(zipBlob);
        link.download = `LogMetricPro_Package_${container}_${timeStr}.zip`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(link.href), 600);
        alert(currentLang === 'zh' ? '已打包导出 JSON、PDF、Excel' : (currentLang === 'en' ? 'Package exported: JSON, PDF, Excel' : 'Paket izvožen: JSON, PDF, Excel'));
    } catch (e) {
        alert(currentLang === 'zh' ? '打包导出失败，请重试' : (currentLang === 'en' ? 'Package export failed. Please try again.' : 'Paket izvoza ni uspel. Poskusite znova.'));
    }
}

function toggleShowCompanyInPdf() {
    const cb = document.getElementById('showCompanyInPdf');
    if (cb) { appSettings.showCompanyInPdf = cb.checked; localStorage.setItem(SETTINGS_KEY, JSON.stringify(appSettings)); }
}

// ========== 保存菜单和快照系统 ==========
function toggleSaveMenu() {
    const menu = document.getElementById('saveMenu');
    menu.classList.toggle('show');
    // 点击页面其他地方关闭菜单
    if (menu.classList.contains('show')) {
        setTimeout(() => {
            document.addEventListener('click', closeSaveMenuOnClickOutside);
        }, 0);
    }
}

function closeSaveMenu() {
    const menu = document.getElementById('saveMenu');
    menu.classList.remove('show');
    document.removeEventListener('click', closeSaveMenuOnClickOutside);
}

function closeSaveMenuOnClickOutside(e) {
    const menu = document.getElementById('saveMenu');
    const btn = document.querySelector('.btn-save-menu');
    if (!menu.contains(e.target) && !btn.contains(e.target)) {
        closeSaveMenu();
    }
}

function saveToSnapshot() {
    closeSaveMenu();

    // 如果没有数据，提示用户
    if (logs.length <= 1 && !logs[0]?.length && !logs[0]?.diameter) {
        alert(currentLang === 'zh' ? '没有数据可以保存' : (currentLang === 'en' ? 'No data to save' : 'Ni podatkov za shranjevanje'));
        return;
    }

    // 生成或使用现有的会话 ID
    if (!currentSessionId) {
        const today = new Date();
        const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
        // 找出今天已有的编号
        const todaySnapshots = snapshots.filter(s => s.id.startsWith(dateStr));
        const maxNum = todaySnapshots.length > 0
            ? Math.max(...todaySnapshots.map(s => parseInt(s.id.split('_')[1]) || 0))
            : 0;
        currentSessionId = `${dateStr}_${maxNum + 1}`;
    }

    // 创建快照数据
    const snapshotData = {
        id: currentSessionId,
        timestamp: Date.now(),
        date: new Date().toLocaleString(currentLang === 'zh' ? 'zh-CN' : currentLang === 'en' ? 'en-US' : 'sl-SI'),
        logs: JSON.parse(JSON.stringify(logs)),
        global: JSON.parse(JSON.stringify(globalInfo)),
        container: globalInfo.container || '未命名'
    };

    // 更新或添加快照
    const existingIndex = snapshots.findIndex(s => s.id === currentSessionId);
    if (existingIndex >= 0) {
        snapshots[existingIndex] = snapshotData;
    } else {
        snapshots.unshift(snapshotData);
    }

    // 保存到 localStorage
    localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));
    localStorage.setItem(SESSION_KEY, currentSessionId);

    // 提示用户
    const msg = currentLang === 'zh'
        ? `已暂存至内部记录\n记录编号: ${currentSessionId}\n集装箱: ${snapshotData.container}`
        : currentLang === 'en'
            ? `Saved to records\nRecord ID: ${currentSessionId}\nContainer: ${snapshotData.container}`
            : `Shranjeno v zapise\nID zapisa: ${currentSessionId}\nKontejner: ${snapshotData.container}`;
    alert(msg);
}

function openSnapshotHistory() {
    closeSaveMenu();

    if (snapshots.length === 0) {
        alert(currentLang === 'zh' ? '暂无历史记录' : (currentLang === 'en' ? 'No history records' : 'Ni zgodovinskih zapisov'));
        return;
    }

    // 创建历史记录列表
    const historyContent = snapshots.map(snap => {
        const isCurrent = snap.id === currentSessionId;
        const bgColor = isCurrent ? 'rgba(212,163,115,0.2)' : '#1a1a1a';
        const border = isCurrent ? '2px solid var(--accent-color)' : '1px solid #333';
        const badge = isCurrent ? '<span style="color:var(--accent-color);font-weight:600;margin-left:10px;"><i data-lucide="circle-dot" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> 当前</span>' : '';

        return `
                <div style="background:${bgColor};padding:15px;margin:10px 0;border-radius:8px;border:${border};cursor:pointer;" 
                     onclick="openHistoryViewer('${snap.id}')" 
                     onmouseover="this.style.background='rgba(212,163,115,0.1)'" 
                     onmouseout="this.style.background='${bgColor}'">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div>
                            <div style="font-weight:600;color:var(--accent-color);font-size:16px;">${snap.container || '未命名'}${badge}</div>
                            <div style="color:#999;font-size:12px;margin-top:5px;"><i data-lucide="calendar" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> ${snap.date}</div>
                            <div style="color:#999;font-size:12px;"><i data-lucide="hash" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> ${snap.id}</div>
                            <div style="color:#aaa;font-size:12px;margin-top:3px;"><i data-lucide="table" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> 共 ${snap.logs.length - 1} 根原木</div>
                        </div>
                        <span style="display:inline-flex;align-items:center;gap:6px;">
                            <button onclick="event.stopPropagation(); exportSingleRecord('${snap.id}')" 
                                    style="background:var(--accent-color);color:#111;border:none;padding:8px 12px;border-radius:6px;cursor:pointer;font-size:12px;display:inline-flex;align-items:center;gap:4px;font-weight:600;">
                                <i data-lucide="download" style="width:14px;height:14px;"></i> ${currentLang === 'zh' ? '导出' : (currentLang === 'en' ? 'Export' : 'Izvozi')}
                            </button>
                            <button onclick="event.stopPropagation(); deleteSnapshot('${snap.id}')" 
                                    style="background:#d32f2f;color:#fff;border:none;padding:8px 12px;border-radius:6px;cursor:pointer;font-size:12px;display:inline-flex;align-items:center;gap:4px;">
                                <i data-lucide="trash-2" style="width:14px;height:14px;"></i> ${currentLang === 'zh' ? '删除' : (currentLang === 'en' ? 'Delete' : 'Izbriši')}
                            </button>
                        </span>
                    </div>
                </div>
            `;
    }).join('');

    // 显示历史记录弹窗
    const modal = document.createElement('div');
    modal.id = 'snapshotHistoryModal';
    modal.className = 'modal-overlay';
    modal.style.display = 'flex';
    modal.innerHTML = `
            <div class="modal-card" onclick="event.stopPropagation()" style="max-width:600px;max-height:80vh;overflow-y:auto;">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
                    <div class="modal-title"><i data-lucide="folder-open" style="width:22px;height:22px;margin-right:6px;vertical-align:middle;"></i>${currentLang === 'zh' ? '内部记录' : (currentLang === 'en' ? 'Internal Records' : 'Interna zgodovina')}</div>
                    <span style="display:inline-flex;align-items:center;gap:4px;">
                        <button style="background:var(--accent-color);color:#111;border:none;padding:8px 14px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;white-space:nowrap;display:inline-flex;align-items:center;gap:6px;" onclick="openExportProjectModal();">
                            <i data-lucide="download" style="width:16px;height:16px;"></i> ${currentLang === 'zh' ? '批量导出' : (currentLang === 'en' ? 'Batch Export' : 'Paketni izvoz')}
                        </button>
                        <span class="btn-help" onclick="showHelp('menu_export_json');event.stopPropagation()" role="button">?</span>
                    </span>
                </div>
                <div style="margin-top:20px;">
                    ${historyContent}
                </div>
                <button class="btn-block" style="background:#333;color:#fff;margin-top:20px;display:flex;align-items:center;justify-content:center;gap:6px;" onclick="this.closest('.modal-overlay').remove()">
                    <i data-lucide="x" style="width:18px;height:18px;"></i> ${currentLang === 'zh' ? '关闭' : (currentLang === 'en' ? 'Close' : 'Zapri')}
                </button>
            </div>
        `;
    modal.onclick = () => modal.remove();
    document.body.appendChild(modal);
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function openHistoryViewer(id) {
    const snapshot = snapshots.find(s => s.id === id);
    if (!snapshot) {
        alert(currentLang === 'zh' ? '记录不存在' : (currentLang === 'en' ? 'Record not found' : 'Zapis ni najden'));
        return;
    }

    // 物理隔离：历史查看器仅使用快照副本，不触碰全局 logs/globalInfo
    const rawGlobal = JSON.parse(JSON.stringify(snapshot.global || {}));
    historyViewerState = {
        snapshotId: snapshot.id,
        logs: JSON.parse(JSON.stringify(snapshot.logs || [])),
        global: rawGlobal,
        sourceDate: snapshot.date || '',
        reversedOrder: false
    };
    historyViewerState.global.company = normalizeCompany(historyViewerState.global.company);
    historyViewerState.global.seller = normalizeSeller(historyViewerState.global.seller);

    const meta = document.getElementById('historyViewerMeta');
    if (meta) {
        const t = I18N[currentLang];
        const containerDisplay = (snapshot.container && snapshot.container !== '未命名') ? snapshot.container : t.hv_unnamed;
        meta.innerText = `ID: ${snapshot.id} | ${t.hv_container_label}: ${containerDisplay} | ${t.hv_time}: ${snapshot.date || '-'}`;
    }
    const historyModal = document.getElementById('snapshotHistoryModal');
    if (historyModal) historyModal.remove();
    const globalEl = document.getElementById('historyViewerGlobal');
    const infoBtn = document.getElementById('hvInfoBtn');
    if (globalEl) globalEl.style.display = 'none';
    if (infoBtn) infoBtn.classList.remove('active');
    renderHistoryViewer();
    const viewer = document.getElementById('historyViewer');
    if (viewer) viewer.style.display = 'flex';
    document.body.classList.add('history-viewer-open');
    document.documentElement.classList.add('history-viewer-open');
}

function closeHistoryViewer() {
    const viewer = document.getElementById('historyViewer');
    if (viewer) viewer.style.display = 'none';
    document.body.classList.remove('history-viewer-open');
    document.documentElement.classList.remove('history-viewer-open');
}
function toggleHistoryViewerInfo() {
    const el = document.getElementById('historyViewerGlobal');
    const btn = document.getElementById('hvInfoBtn');
    if (!el || !btn) return;
    const isHidden = el.style.display === 'none' || !el.style.display;
    el.style.display = isHidden ? 'block' : 'none';
    btn.classList.toggle('active', isHidden);
}

function renderHistoryViewer() {
    const g = historyViewerState.global || {};
    const list = historyViewerState.logs || [];
    const t = I18N[currentLang];

    const globalEl = document.getElementById('historyViewerGlobal');
    if (globalEl) {
        globalEl.innerHTML = `
                <div class="info-row info-row-inline">
                    <div class="info-field info-field-narrow">
                        <label>${t.container}</label>
                        <div class="info-controls"><input type="text" id="hv_g_container" value="${(g.container || '').replace(/"/g, '&quot;')}" oninput="updateHistoryViewerGlobal('container', this.value)"></div>
                    </div>
                    <div class="info-field info-field-wide">
                        <label>${t.note_global}</label>
                        <div class="info-controls"><input type="text" id="hv_g_note" value="${(g.note || '').replace(/"/g, '&quot;')}" oninput="updateHistoryViewerGlobal('note', this.value)"></div>
                    </div>
                </div>
                <div class="info-row"><label>${t.description}</label><div class="info-controls"><input type="text" id="hv_g_description" value="${(g.description || '').replace(/"/g, '&quot;')}" oninput="updateHistoryViewerGlobal('description', this.value)"></div></div>
                <div class="info-row"><label>${t.seller}</label><div class="info-controls info-controls-readonly"><input type="text" id="hv_g_seller" readonly class="input-readonly" value="${escapeAttr(getCompanyName(g.seller))}" placeholder="—" onclick="openSellerDetailModalForHistoryViewer()"><button class="btn-tool" onclick="saveToHistoryForHistoryViewer('seller')" title="${t.save_history || '保存到历史'}"><i data-lucide="save"></i></button><button class="btn-tool" onclick="showHistoryForHistoryViewer('seller')" title="${t.view_history || '查看历史记录'}"><i data-lucide="book-open"></i></button></div></div>
                <div class="info-row"><label>${t.location}</label><div class="info-controls"><input type="text" id="hv_g_location" value="${(g.location || '').replace(/"/g, '&quot;')}" oninput="updateHistoryViewerGlobal('location', this.value)"></div></div>
                <div class="info-row"><label>${t.measurer}</label><div class="info-controls"><input type="text" id="hv_g_measurer" value="${(g.measurer || '').replace(/"/g, '&quot;')}" oninput="updateHistoryViewerGlobal('measurer', this.value)"></div></div>
            `;
    }

    const listEl = document.getElementById('historyViewerList');
    if (listEl) {
        listEl.innerHTML = '';
        if (list.length > 1) {
            const header = document.createElement('div');
            header.className = 'list-header';
            const idxTitle = (t.hv_idx_click || (currentLang === 'zh' ? '点击颠倒顺序' : 'Click to reverse order'));
            header.innerHTML = `<div class="hv-idx-header" onclick="toggleHistoryViewerOrder()" title="${escapeAttr(idxTitle)}">${t.idx}</div><div>${t.code}</div><div>${t.grade}</div><div>${t.len}</div><div>${t.dia}</div><div>${t.vol}</div><div>${t.note}</div><div></div><div></div>`;
            listEl.appendChild(header);
        }
        const totalRows = list.length - 1;
        const reversed = !!historyViewerState.reversedOrder;
        for (let i = 1; i <= totalRows; i++) {
            const log = list[i];
            const displayIdx = reversed ? i : (list.length - i);
            listEl.appendChild(createHistoryViewerRow(log, displayIdx));
        }
    }
    updateHistoryViewerStats();
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function createHistoryViewerRow(log, idx) {
    const t = I18N[currentLang];
    const list = historyViewerState.logs || [];
    const gradeLabels = getGradeLabels();
    const customGrades = [...new Set(list.map(l => l.grade).filter(g => g && !gradeLabels.includes(g)))];
    const allGrades = [...gradeLabels, ...customGrades];
    let gradeOptions = `<option value="">-</option>`;
    gradeOptions += allGrades.map(g => `<option value="${g}" ${log.grade === g ? 'selected' : ''}>${g}</option>`).join('');

    const len_m = parseFloat(cleanInput((log.length || '').toString()));
    const d_cm = parseFloat(cleanInput((log.diameter || '').toString()));
    const warnClass = (len_m > 15) ? 'danger-text' : '';
    const diaDangerClass = (d_cm >= 200 && d_cm < 1000) ? 'dia-danger-text' : '';
    const displayLen = log.length ? (parseFloat(log.length) || log.length) : '';
    const displayDia = log.diameter ? (parseFloat(log.diameter) || log.diameter) : '';
    const vol = (log.volume != null && !isNaN(log.volume)) ? formatVolumeForDisplay(log.volume) : '0';

    const div = document.createElement('div');
    div.className = 'log-row';
    div.id = 'hv-row-' + log.id;
    div.setAttribute('data-grade', log.grade || '?');
    div.setAttribute('data-len', len_m || 0);
    div.setAttribute('data-dia', d_cm || 0);
    div.innerHTML = `
            <div class="row-index">${idx}</div>
            <input type="text" data-field="code" value="${escapeAttr(log.code || '')}" oninput="updateHistoryViewerItem(${log.id},'code',this.value)">
            <div class="hv-grade-cell" oncontextmenu="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event); return false" ondblclick="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event)" title="${currentLang === 'zh' ? '双击可修改等级按钮文字' : (currentLang === 'en' ? 'Double-click to change grade button label' : 'Dvojni klik za spremembo')}"><select onchange="updateHistoryViewerItem(${log.id},'grade',this.value)">${gradeOptions}</select></div>
            <input type="text" data-field="length" class="${warnClass}" inputmode="decimal" value="${escapeAttr(displayLen)}" oninput="updateHistoryViewerItem(${log.id},'length',this.value)" onblur="autoFixHistoryViewerInput(this)">
            <input type="text" data-field="diameter" class="${diaDangerClass}" inputmode="decimal" value="${escapeAttr(displayDia)}" oninput="updateHistoryViewerItem(${log.id},'diameter',this.value);toggleDiaDangerClass(this)" onblur="autoFixHistoryViewerInput(this)">
            <div class="col-vol" id="hv-v-row-${log.id}">${vol}</div>
            <input type="text" data-field="note" style="font-size:12px;color:#aaa;text-align:left;" value="${escapeAttr(log.note || '')}" oninput="updateHistoryViewerItem(${log.id},'note',this.value)">
            <div class="group-actions-cell"></div>
            <button type="button" class="btn-del-mini" onclick="delHistoryViewerRow(${log.id})">×</button>
        `;
    return div;
}
function escapeAttr(v) {
    if (v == null) return '';
    return String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function autoFixHistoryViewerInput(input) {
    let val = input.value;
    if (val && val.includes(',')) {
        input.value = val.replace(/,/g, '.');
        const id = input.closest('.log-row')?.id?.replace('hv-row-', '');
        const field = input.getAttribute('data-field');
        if (id && field) updateHistoryViewerItem(parseInt(id, 10), field === 'length' ? 'length' : 'diameter', input.value);
    }
}
function updateHistoryViewerGlobal(key, value) {
    if (!historyViewerState.global) historyViewerState.global = {};
    if (key === 'company') {
        historyViewerState.global.company = Object.assign(defaultCompany(), historyViewerState.global.company || {}, { name: value });
        return;
    }
    if (key === 'seller') {
        historyViewerState.global.seller = Object.assign(defaultSeller(), historyViewerState.global.seller || {}, { name: value });
        return;
    }
    historyViewerState.global[key] = value;
}
function updateHistoryViewerItem(id, field, value) {
    const item = (historyViewerState.logs || []).find(l => l.id === id);
    if (!item) return;
    item[field] = value;
    if (field === 'length' || field === 'diameter') {
        item.volume = calculateVolume(item.length, item.diameter);
        const vEl = document.getElementById('hv-v-row-' + id);
        if (vEl) vEl.innerText = (item.volume != null && !isNaN(item.volume)) ? formatVolumeForDisplay(item.volume) : '0';
        const row = document.getElementById('hv-row-' + id);
        if (row) {
            const lenInput = row.querySelector('input[data-field="length"]');
            if (lenInput) {
                const len_m = parseFloat(cleanInput((item.length || '').toString()));
                if (len_m > 15) lenInput.classList.add('danger-text');
                else lenInput.classList.remove('danger-text');
            }
        }
        updateHistoryViewerStats();
    }
}
function updateHistoryViewerStats() {
    const list = historyViewerState.logs || [];
    const validLogs = list.filter(l => parseFloat(l.volume) > 0);
    const totalV = validLogs.reduce((s, l) => s + (parseFloat(l.volume) || 0), 0);
    const statsEl = document.getElementById('historyViewerStats');
    if (statsEl) {
        const t = I18N[currentLang];
        statsEl.innerHTML = `<span>${t.total_count}: <strong id="hvTotalCount">${validLogs.length}</strong></span><span>${t.total_vol}: <strong id="hvTotalVol">${formatVolumeForDisplay(totalV)}</strong> m³</span>`;
    }
}
function toggleHistoryViewerOrder() {
    const list = historyViewerState.logs || [];
    if (list.length <= 2) return;
    const template = list[0];
    const dataRows = list.slice(1);
    dataRows.reverse();
    historyViewerState.logs = [template, ...dataRows];
    historyViewerState.reversedOrder = !historyViewerState.reversedOrder;
    renderHistoryViewer();
}
function delHistoryViewerRow(id) {
    if (!confirm(currentLang === 'zh' ? '删除此行？' : (currentLang === 'en' ? 'Delete this row?' : 'Izbriši to vrstico?'))) return;
    historyViewerState.logs = (historyViewerState.logs || []).filter(l => l.id !== id);
    renderHistoryViewer();
}
function addHistoryViewerRow() {
    const list = historyViewerState.logs || [];
    if (list.length === 0) {
        list.push({ id: 0, code: '', grade: '', length: '', diameter: '', volume: 0, note: '' });
    }
    list.push({ id: Date.now(), code: '', grade: '', length: '', diameter: '', volume: 0, note: '' });
    renderHistoryViewer();
}

function saveHistoryViewerChanges() {
    if (!historyViewerState.snapshotId) return;

    historyViewerState.global.container = (document.getElementById('hv_g_container')?.value || '').trim();
    historyViewerState.global.note = (document.getElementById('hv_g_note')?.value || '').trim();
    historyViewerState.global.description = (document.getElementById('hv_g_description')?.value || '').trim();
    historyViewerState.global.location = (document.getElementById('hv_g_location')?.value || '').trim();
    historyViewerState.global.measurer = (document.getElementById('hv_g_measurer')?.value || '').trim();
    historyViewerState.global.seller = historyViewerState.global.seller || defaultSeller();

    const idx = snapshots.findIndex(s => s.id === historyViewerState.snapshotId);
    if (idx < 0) {
        alert(currentLang === 'zh' ? '对应历史记录不存在，可能已被删除' : (currentLang === 'en' ? 'Snapshot no longer exists' : 'Posnetek ne obstaja več'));
        return;
    }

    snapshots[idx].logs = JSON.parse(JSON.stringify(historyViewerState.logs));
    snapshots[idx].global = JSON.parse(JSON.stringify(historyViewerState.global));
    snapshots[idx].container = (historyViewerState.global && historyViewerState.global.container) ? historyViewerState.global.container : (snapshots[idx].container || '未命名');
    snapshots[idx].date = new Date().toLocaleString(currentLang === 'zh' ? 'zh-CN' : currentLang === 'en' ? 'en-US' : 'sl-SI');
    snapshots[idx].timestamp = Date.now();
    localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));

    const meta = document.getElementById('historyViewerMeta');
    if (meta) {
        const t = I18N[currentLang];
        const snap = snapshots[idx];
        const containerDisplay = (snap.container && snap.container !== '未命名') ? snap.container : t.hv_unnamed;
        meta.innerText = `ID: ${snap.id} | ${t.hv_container_label}: ${containerDisplay} | ${t.hv_time}: ${snap.date || '-'}`;
    }
    alert(currentLang === 'zh' ? '历史快照已保存' : (currentLang === 'en' ? 'History snapshot saved' : 'Zgodovinski posnetek je shranjen'));
}

function generateHistoryViewerPDF() {
    generatePDF({
        logs: historyViewerState.logs || [],
        global: historyViewerState.global || {}
    });
}

function exportHistoryViewerExcel() {
    exportData({
        logs: historyViewerState.logs || [],
        global: historyViewerState.global || {}
    });
}

function loadSnapshot(id) {
    // 兼容旧入口：点击历史记录统一进入隔离查看器
    openHistoryViewer(id);
}

function exportSingleRecord(id) {
    const snap = snapshots.find(s => s.id === id);
    if (!snap) {
        alert(currentLang === 'zh' ? '记录不存在' : (currentLang === 'en' ? 'Record not found' : 'Zapis ni najden'));
        return;
    }
    const payload = {
        schema: 'oak_project_file_v1',
        exportedAt: Date.now(),
        currentSessionId: snap.id,
        logs: JSON.parse(JSON.stringify(snap.logs || [])),
        global: JSON.parse(JSON.stringify(snap.global || {}))
    };
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}-${String(d.getMinutes()).padStart(2, '0')}`;
    const container = (snap.container || 'Project').replace(/[^\w\u4e00-\u9fa5-]+/g, '_');
    const fileName = `OakProject_${container}_${stamp}.json`;
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 600);
}

function deleteSnapshot(id) {
    const msg = currentLang === 'zh' ? '确定要删除此记录吗？' : (currentLang === 'en' ? 'Delete this record?' : 'Izbriši ta zapis?');
    if (!confirm(msg)) return;

    snapshots = snapshots.filter(s => s.id !== id);
    localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));

    // 如果删除的是当前会话，清除会话 ID
    if (currentSessionId === id) {
        currentSessionId = null;
        localStorage.removeItem(SESSION_KEY);
    }

    // 刷新历史记录列表
    const historyModal = document.getElementById('snapshotHistoryModal');
    if (historyModal) historyModal.remove();
    openSnapshotHistory();
}

function openExportProjectModal() {
    closeSaveMenu();
    const historyModal = document.getElementById('snapshotHistoryModal');
    if (historyModal) historyModal.remove();
    if (snapshots.length === 0) {
        alert(currentLang === 'zh' ? '暂无内部记录可导出' : (currentLang === 'en' ? 'No internal records to export' : 'Ni notranjih zapisov za izvoz'));
        return;
    }
    const recordContent = snapshots.map(snap => {
        const isCurrent = snap.id === currentSessionId;
        const badge = isCurrent ? '<span style="color:var(--accent-color);font-weight:600;margin-left:8px;"><i data-lucide="circle-dot" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> 当前</span>' : '';
        return `
                <div class="export-record-item" style="display:flex;align-items:center;gap:12px;padding:12px 15px;margin:8px 0;background:#1a1a1a;border-radius:8px;border:1px solid #333;">
                    <label style="display:flex;align-items:center;cursor:pointer;flex-shrink:0;">
                        <input type="checkbox" class="export-record-cb export-select-cb" data-id="${snap.id}">
                    </label>
                    <div style="flex:1;cursor:pointer;" onclick="this.closest('.export-record-item').querySelector('.export-record-cb').click()">
                        <div style="font-weight:600;color:var(--accent-color);font-size:15px;">${(snap.container || '未命名').replace(/</g, '&lt;')}${badge}</div>
                        <div style="color:#999;font-size:12px;margin-top:4px;"><i data-lucide="calendar" style="width:12px;height:12px;display:inline-block;vertical-align:middle;"></i> ${(snap.date || '-').replace(/</g, '&lt;')} | <i data-lucide="hash" style="width:12px;height:12px;display:inline-block;vertical-align:middle;"></i> ${(snap.id || '').replace(/</g, '&lt;')}</div>
                        <div style="color:#aaa;font-size:12px;margin-top:2px;"><i data-lucide="table" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> 共 ${Math.max(0, (snap.logs?.length || 1) - 1)} 根原木</div>
                    </div>
                </div>
            `;
    }).join('');
    const modal = document.createElement('div');
    modal.id = 'exportProjectModal';
    modal.className = 'modal-overlay';
    modal.style.display = 'flex';
    modal.innerHTML = `
            <div class="modal-card" onclick="event.stopPropagation()" style="max-width:560px;max-height:85vh;overflow-y:auto;">
                <div class="modal-title"><i data-lucide="download" style="width:22px;height:22px;margin-right:6px;vertical-align:middle;"></i>${currentLang === 'zh' ? '批量导出' : (currentLang === 'en' ? 'Batch Export' : 'Paketni izvoz')}</div>
                <div class="export-select-row" style="margin:12px 0;display:flex;justify-content:space-between;align-items:center;gap:12px;">
                    <label class="export-select-all-label" style="cursor:pointer;color:#aaa;font-size:13px;display:inline-flex;align-items:center;gap:8px;">
                        <input type="checkbox" id="exportSelectAll" class="export-select-cb" onchange="toggleExportSelectAll(this)">
                        <span>${currentLang === 'zh' ? '全选' : (currentLang === 'en' ? 'Select All' : 'Izberi vse')}</span>
                    </label>
                    <button type="button" class="export-delete-selected-btn" onclick="doDeleteSelectedRecords()" style="background:rgba(211,47,47,0.2);color:#e57373;border:1px solid rgba(211,47,47,0.4);padding:6px 12px;border-radius:6px;cursor:pointer;font-size:12px;display:inline-flex;align-items:center;gap:4px;">
                        <i data-lucide="trash-2" style="width:14px;height:14px;"></i> ${currentLang === 'zh' ? '删除选中' : (currentLang === 'en' ? 'Delete Selected' : 'Izbriši izbrano')}
                    </button>
                </div>
                <div style="max-height:320px;overflow-y:auto;margin-bottom:16px;">
                    ${recordContent}
                </div>
                <div style="display:flex;gap:10px;">
                    <button class="btn-block" style="flex:1;background:var(--accent-color);color:#111;font-weight:600;display:flex;align-items:center;justify-content:center;gap:6px;" onclick="doExportSelectedRecords()">
                        <i data-lucide="check" style="width:18px;height:18px;"></i> ${currentLang === 'zh' ? '导出选中' : (currentLang === 'en' ? 'Export Selected' : 'Izvozi izbrano')}
                    </button>
                    <button class="btn-block" style="flex:1;background:#333;color:#fff;display:flex;align-items:center;justify-content:center;gap:6px;" onclick="document.getElementById('exportProjectModal')?.remove()">
                        <i data-lucide="x" style="width:18px;height:18px;"></i> ${currentLang === 'zh' ? '取消' : (currentLang === 'en' ? 'Cancel' : 'Prekliči')}
                    </button>
                </div>
            </div>
        `;
    modal.onclick = (e) => { if (e.target === modal) modal.remove(); };
    document.body.appendChild(modal);
    if (typeof lucide !== 'undefined') lucide.createIcons();
}
function toggleExportSelectAll(checkbox) {
    document.querySelectorAll('.export-record-cb').forEach(cb => cb.checked = checkbox.checked);
}
function doDeleteSelectedRecords() {
    const checked = document.querySelectorAll('.export-record-cb:checked');
    if (checked.length === 0) {
        alert(currentLang === 'zh' ? '请先选择要删除的记录' : (currentLang === 'en' ? 'Please select records to delete' : 'Izberite zapise za brisanje'));
        return;
    }
    const msg = currentLang === 'zh' ? `确定要删除选中的 ${checked.length} 条记录吗？此操作不可撤销。` : (currentLang === 'en' ? `Delete ${checked.length} selected record(s)? This cannot be undone.` : `Izbrišem ${checked.length} izbranih zapisov? Te operacije ni mogoče razveljaviti.`);
    if (!confirm(msg)) return;
    const ids = Array.from(checked).map(cb => cb.getAttribute('data-id'));
    ids.forEach(id => {
        if (currentSessionId === id) {
            currentSessionId = null;
            localStorage.removeItem(SESSION_KEY);
        }
        const idx = snapshots.findIndex(s => s.id === id);
        if (idx >= 0) snapshots.splice(idx, 1);
    });
    localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));
    document.getElementById('exportProjectModal')?.remove();
    openSnapshotHistory();
}
async function doExportSelectedRecords() {
    const checked = document.querySelectorAll('.export-record-cb:checked');
    if (checked.length === 0) {
        alert(currentLang === 'zh' ? '请至少选择一条记录' : (currentLang === 'en' ? 'Please select at least one record' : 'Izberite vsaj en zapis'));
        return;
    }
    const ids = Array.from(checked).map(cb => cb.getAttribute('data-id'));
    const records = snapshots.filter(s => ids.includes(s.id)).map(s => ({
        id: s.id,
        currentSessionId: s.id,
        logs: JSON.parse(JSON.stringify(s.logs || [])),
        global: JSON.parse(JSON.stringify(s.global || {})),
        container: (s.container || '未命名').replace(/[^\w\u4e00-\u9fa5-]+/g, '_')
    }));
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}-${String(d.getMinutes()).padStart(2, '0')}`;
    try {
        const zip = new JSZip();
        records.forEach((rec, i) => {
            const payload = {
                schema: 'oak_project_file_v1',
                exportedAt: Date.now(),
                currentSessionId: rec.currentSessionId || rec.id,
                logs: rec.logs,
                global: rec.global
            };
            const safeName = (rec.container || 'record').replace(/[^\w\u4e00-\u9fa5-]+/g, '_');
            const jsonName = records.length > 1 ? `${safeName}_${i + 1}.json` : `${safeName}.json`;
            zip.file(jsonName, JSON.stringify(payload, null, 2));
        });
        const blob = await zip.generateAsync({ type: 'blob' });
        const fileName = `OakProject_${records.length}records_${stamp}.zip`;
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(link.href), 600);
    } catch (e) {
        alert(currentLang === 'zh' ? '导出失败，请重试' : (currentLang === 'en' ? 'Export failed. Please try again.' : 'Izvoz ni uspel. Poskusite znova.'));
        return;
    }
    document.getElementById('exportProjectModal')?.remove();
    openSnapshotHistory();
}
function exportProjectFile() {
    const payload = {
        schema: 'oak_project_file_v1',
        exportedAt: Date.now(),
        currentSessionId: currentSessionId || null,
        logs: JSON.parse(JSON.stringify(logs || [])),
        global: JSON.parse(JSON.stringify(globalInfo || {}))
    };
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}-${String(d.getMinutes()).padStart(2, '0')}`;
    const container = (globalInfo.container || 'Project').replace(/[^\w\u4e00-\u9fa5-]+/g, '_');
    const fileName = `OakProject_${container}_${stamp}.json`;
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 600);
}

function triggerImportProjectFile() {
    const input = document.getElementById('projectFileInput');
    if (!input) return;
    input.value = '';
    input.click();
}
function triggerImportFromInfo() {
    triggerImportProjectFile();
}

async function handleImportProjectFile(event) {
    const file = event?.target?.files?.[0];
    if (!file) return;
    const isZip = file.name.toLowerCase().endsWith('.zip');
    if (isZip) {
        try {
            const zip = await JSZip.loadAsync(file);
            const jsonFiles = Object.keys(zip.files).filter(n => n.toLowerCase().endsWith('.json'));
            if (jsonFiles.length === 0) {
                alert(currentLang === 'zh' ? 'ZIP 中未找到 JSON 文件' : (currentLang === 'en' ? 'No JSON files found in ZIP' : 'V ZIP ni JSON datotek'));
                return;
            }
            let imported = 0;
            for (const name of jsonFiles) {
                const f = zip.files[name];
                if (f.dir) continue;
                const text = await f.async('string');
                let data;
                try { data = JSON.parse(text); } catch (_) { continue; }
                if (!Array.isArray(data?.logs) || !data?.global) continue;
                const containerName = (data.global && data.global.container) ? data.global.container : '未命名';
                const snapshot = {
                    id: 'zip_' + Date.now() + '_' + imported,
                    timestamp: Date.now(),
                    date: new Date().toLocaleString(currentLang === 'zh' ? 'zh-CN' : currentLang === 'en' ? 'en-US' : 'sl-SI'),
                    logs: JSON.parse(JSON.stringify(data.logs || [])),
                    global: JSON.parse(JSON.stringify(data.global || {})),
                    container: containerName
                };
                snapshots.unshift(snapshot);
                imported++;
            }
            localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));
            alert(currentLang === 'zh' ? `已从 ZIP 导入 ${imported} 条记录到内部记录` : (currentLang === 'en' ? `Imported ${imported} record(s) from ZIP to internal records` : `Uvoženo ${imported} zapisov iz ZIP v notranje zapise`));
            openSnapshotHistory();
        } catch (e) {
            alert(currentLang === 'zh' ? '读取 ZIP 失败' : (currentLang === 'en' ? 'Failed to read ZIP file' : 'Branje ZIP ni uspelo'));
        }
        return;
    }
    let text = '';
    try {
        text = await file.text();
    } catch (e) {
        alert(currentLang === 'zh' ? '读取文件失败' : (currentLang === 'en' ? 'Failed to read file' : 'Branje datoteke ni uspelo'));
        return;
    }

    let data = null;
    try {
        data = JSON.parse(text);
    } catch (e) {
        alert(currentLang === 'zh' ? '文件不是有效的 JSON' : (currentLang === 'en' ? 'Invalid JSON file' : 'Neveljavna JSON datoteka'));
        return;
    }

    let loadData = null;
    if (data.schema === 'logmetric_backup_v1') {
        const msg = currentLang === 'zh' ? '检测到完整备份文件，将恢复所有数据（原木、项目信息、历史记录、设置等）。是否继续？' : (currentLang === 'en' ? 'Full backup detected. This will restore all data (logs, project info, history, settings). Continue?' : 'Varnostna kopija. Obnovim vse podatke. Nadaljuj?');
        if (!confirm(msg)) return;
        logs = data.logs || [];
        globalInfo = data.global || globalInfo;
        normalizeGlobalInfo(globalInfo);
        snapshots = (data.snapshots || []).map(s => {
            const g = s.global || {};
            g.company = normalizeCompany(g.company);
            g.seller = normalizeSeller(g.seller);
            return Object.assign({}, s, { global: g });
        });
        histories = data.histories || histories;
        migrateHistoriesSeller(histories);
        migrateHistoriesCompany(histories);
        if (data.appSettings && typeof data.appSettings === 'object') {
            appSettings = Object.assign({}, appSettings, data.appSettings);
            localStorage.setItem(SETTINGS_KEY, JSON.stringify(appSettings));
        }
        currentSessionId = data.currentSessionId || null;
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ logs, global: globalInfo }));
        localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));
        localStorage.setItem(HIST_KEY, JSON.stringify(histories));
        if (currentSessionId) localStorage.setItem(SESSION_KEY, currentSessionId);
        else localStorage.removeItem(SESSION_KEY);
        save();
        location.reload();
        return;
    }
    if (data.schema === 'oak_project_file_v1_multi' && Array.isArray(data.records) && data.records.length > 0) {
        const rec = data.records[0];
        loadData = { logs: rec.logs, global: rec.global || {}, currentSessionId: rec.currentSessionId || rec.id };
    } else if (Array.isArray(data.logs) && data.global && typeof data.global === 'object') {
        loadData = { logs: data.logs, global: data.global, currentSessionId: data.currentSessionId || null };
    }
    if (!loadData) {
        alert(currentLang === 'zh' ? '项目文件缺少 logs/global 数据' : (currentLang === 'en' ? 'Project file missing logs/global data' : 'Datoteka nima logs/global podatkov'));
        return;
    }
    const btnReplace = currentLang === 'zh' ? '替换主界面当前数据' : (currentLang === 'en' ? 'Replace main screen data' : 'Zamenjaj glavni zaslon');
    const btnViewEdit = currentLang === 'zh' ? '进入快速查看修改模式' : (currentLang === 'en' ? 'Open in view/edit mode' : 'Odpri v načinu ogleda/urejanja');
    const promptTxt = currentLang === 'zh' ? '请选择导入方式：' : (currentLang === 'en' ? 'Choose import mode:' : 'Izberi način uvoza:');
    const modal = document.createElement('div');
    modal.id = 'importChoiceModal';
    modal.className = 'modal-overlay';
    modal.style.display = 'flex';
    modal.innerHTML = `
            <div class="modal-card" onclick="event.stopPropagation()" style="max-width:420px;">
                <div class="modal-title">${promptTxt}</div>
                <div style="padding:16px;color:#ccc;font-size:14px;">
                    ${currentLang === 'zh' ? '替换主界面：将文件数据加载到主界面，替换当前内容。' : (currentLang === 'en' ? 'Replace: Load file data into main screen, replacing current content.' : 'Zamenjaj: Naloži v glavni zaslon.')}
                    <br><br>
                    ${currentLang === 'zh' ? '快速查看修改：在内部记录编辑模式中打开，可先查看、修改后再决定是否保存。' : (currentLang === 'en' ? 'View/Edit: Open in internal record editor to preview and modify before saving.' : 'Ogled/Uredi: Odpri v urejevalniku za predogled.')}
                </div>
                <div style="display:flex;flex-direction:column;gap:10px;">
                    <button class="btn-block" style="background:var(--accent-color);color:#111;font-weight:600;" data-action="replace">${btnReplace}</button>
                    <button class="btn-block" style="background:#333;color:#fff;border:1px solid #555;" data-action="viewedit">${btnViewEdit}</button>
                    <button class="btn-block" style="background:#222;color:#888;" onclick="document.getElementById('importChoiceModal').remove()">${currentLang === 'zh' ? '取消' : (currentLang === 'en' ? 'Cancel' : 'Prekliči')}</button>
                </div>
            </div>
        `;
    modal.onclick = (e) => { if (e.target === modal) modal.remove(); };
    document.body.appendChild(modal);
    modal.querySelector('[data-action="replace"]').onclick = () => {
        modal.remove();
        doImportReplaceMain(loadData);
    };
    modal.querySelector('[data-action="viewedit"]').onclick = () => {
        modal.remove();
        doImportViewEditMode(loadData);
    };
}
function doImportReplaceMain(loadData) {
    const savedCompany = getCompanyName(globalInfo.company) ? Object.assign({}, defaultCompany(), globalInfo.company) : null;
    logs = JSON.parse(JSON.stringify(loadData.logs));
    globalInfo = JSON.parse(JSON.stringify(loadData.global));
    normalizeGlobalInfo(globalInfo);
    if (savedCompany) globalInfo.company = savedCompany;
    currentSessionId = loadData.currentSessionId || null;
    const gContainer = document.getElementById('g_container');
    const gNote = document.getElementById('g_note');
    const gCompany = document.getElementById('g_company');
    const gSeller = document.getElementById('g_seller');
    const gLocation = document.getElementById('g_location');
    const gMeasurer = document.getElementById('g_measurer');
    if (gContainer) gContainer.value = globalInfo.container || '';
    if (gNote) gNote.value = globalInfo.note || '';
    const gDesc = document.getElementById('g_description');
    if (gDesc) gDesc.value = globalInfo.description || '';
    if (gCompany) gCompany.value = getCompanyName(globalInfo.company);
    if (gSeller) gSeller.value = getCompanyName(globalInfo.seller);
    if (gLocation) gLocation.value = globalInfo.location || '';
    if (gMeasurer) gMeasurer.value = globalInfo.measurer || '';
    if (!Array.isArray(logs) || logs.length === 0) { logs = []; addNewLog(); }
    save();
    renderAll();
    if (currentSessionId) localStorage.setItem(SESSION_KEY, currentSessionId);
    else localStorage.removeItem(SESSION_KEY);
}
function doImportViewEditMode(loadData) {
    const tempId = 'imported_' + Date.now();
    const containerName = (loadData.global && loadData.global.container) ? loadData.global.container : '未命名';
    const snapshot = {
        id: tempId,
        timestamp: Date.now(),
        date: new Date().toLocaleString(currentLang === 'zh' ? 'zh-CN' : currentLang === 'en' ? 'en-US' : 'sl-SI'),
        logs: JSON.parse(JSON.stringify(loadData.logs || [])),
        global: JSON.parse(JSON.stringify(loadData.global || {})),
        container: containerName
    };
    snapshots.unshift(snapshot);
    localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));
    openHistoryViewer(tempId);
}
// ========== 快照系统结束 ==========

function saveSettings() {
    const cb = document.getElementById('checkBeginnerMode');
    if (cb) appSettings.beginnerMode = cb.checked;
    appSettings.deductLen = parseFloat(document.getElementById('setDeductLen').value) || 0;
    appSettings.deductDia = parseFloat(document.getElementById('setDeductDia').value) || 0;
    const vwEnabled = document.getElementById('volumeWarningEnabled');
    if (vwEnabled) appSettings.volumeWarningEnabled = vwEnabled.checked;
    const vwThreshold = document.getElementById('volumeWarningThreshold');
    if (vwThreshold) appSettings.volumeWarningThreshold = parseFloat(vwThreshold.value) || 21;
    appSettings.roundMode = document.getElementById('roundSelect').value;
    const formulaEnabledCb = document.getElementById('formulaEnabled');
    if (formulaEnabledCb) appSettings.formulaEnabled = formulaEnabledCb.checked;
    const formulaSelect = document.getElementById('formulaSelect');
    if (formulaSelect) appSettings.formula = formulaSelect.value;
    appSettings.useVirtualKeyboard = appSettings.proKeyboard;
    const priceEnabled = document.getElementById('priceEnabled').checked;
    appSettings.priceEnabled = priceEnabled;
    if (!priceEnabled) {
        appSettings.showPricePdf = false;
        appSettings.showPriceCsv = false;
    }
    appSettings.priceCurrency = document.getElementById('priceCurrency').value;
    appSettings.priceMode = document.getElementById('priceMode').value;
    appSettings.priceFixed = parseFloat(document.getElementById('priceFixed').value) || 0;
    appSettings.taxPercent = parseFloat(document.getElementById('priceTax').value) || 0;
    appSettings.showPricePdf = priceEnabled ? document.getElementById('priceShowPdf').checked : false;
    appSettings.showPriceCsv = priceEnabled ? document.getElementById('priceShowCsv').checked : false;
    const cbMarks = document.getElementById('exportShowMarks');
    const cbGroup = document.getElementById('exportShowGroup');
    if (cbMarks) appSettings.showMarksInExport = cbMarks.checked;
    if (cbGroup) appSettings.showGroupInExport = cbGroup.checked;
    const cbKeySound = document.getElementById('checkKeySound');
    if (cbKeySound) appSettings.keySound = cbKeySound.checked;
    appSettings.priceByGrade = {
        'F': parseFloat(document.getElementById('price_grade_F').value) || 0,
        'A+': parseFloat(document.getElementById('price_grade_Aplus').value) || 0,
        'A': parseFloat(document.getElementById('price_grade_A').value) || 0,
        'B': parseFloat(document.getElementById('price_grade_B').value) || 0,
        'C': parseFloat(document.getElementById('price_grade_C').value) || 0,
        'D': parseFloat(document.getElementById('price_grade_D').value) || 0
    };
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(appSettings));
    if (typeof updateStats === 'function') updateStats();
}
function toggleGradeDisplay() {
    appSettings.showGrade = document.getElementById('checkShowGrade').checked;
    saveSettings();
    renderAll();
    updateProSideState();
    if (!appSettings.showGrade && proState.keypadMode === 'grade') setProKeypadMode('num');
}
function toggleProKeyboard() {
    appSettings.proKeyboard = !appSettings.proKeyboard;
    updateProKeyboardBtnUI();
    updateKeySoundRowVisibility();
    saveSettings();
    applyProKeyboardUI();
    renderAll();
}
function updateKeySoundRowVisibility() {
    const row = document.getElementById('keySoundRow');
    const box = document.getElementById('proKbOptionsBox');
    if (row) row.style.display = appSettings.proKeyboard ? 'flex' : 'none';
    if (box) box.style.display = appSettings.proKeyboard ? 'block' : 'none';
}
function toggleKeySound() {
    appSettings.keySound = document.getElementById('checkKeySound').checked;
    saveSettings();
}
function toggleCalcDia() {
    appSettings.calcDia = document.getElementById('checkCalcDia').checked;
    document.getElementById('roundModeRow').style.display = appSettings.calcDia ? 'flex' : 'none';
    saveSettings();
    renderAll();
    if (appSettings.proKeyboard) {
        proState.values.dia1 = '';
        proState.values.dia2 = '';
    }
    renderProKeyboardTopBar();
    updateProSideState();
}

function updatePriceModeUI() {
    const mode = document.getElementById('priceMode').value;
    const fixedRow = document.getElementById('priceFixedRow');
    const gradeBox = document.getElementById('priceGradeBox');
    if (fixedRow) fixedRow.style.display = mode === 'fixed' ? 'flex' : 'none';
    if (gradeBox) gradeBox.style.display = mode === 'grade' ? 'block' : 'none';
}
function updatePriceEnabledUI() {
    const enabled = document.getElementById('priceEnabled').checked;
    const box = document.getElementById('priceSettingsBox');
    if (box) box.style.display = enabled ? 'block' : 'none';
}
function togglePriceEnabled() {
    const enabled = document.getElementById('priceEnabled').checked;
    if (!enabled) {
        appSettings.showPricePdf = false;
        appSettings.showPriceCsv = false;
        const pdfCb = document.getElementById('priceShowPdf');
        const csvCb = document.getElementById('priceShowCsv');
        if (pdfCb) pdfCb.checked = false;
        if (csvCb) csvCb.checked = false;
    }
    saveSettings();
    updatePriceEnabledUI();
}

function openSettingsModal() {
    const modal = document.getElementById('settingsModal');
    modal.style.display = 'flex';
    const card = modal.querySelector('.modal-card');
    if (card) card.scrollTop = 0;
    updateBeginnerModeUI();
    updateThemeToggleLabel();
    updateKeySoundRowVisibility();
    if (typeof lucide !== 'undefined') lucide.createIcons();
}
function showHiddenFeatures() {
    const h = HIDDEN_FEATURES;
    const lang = currentLang;
    const title = (h.title && (h.title[lang] || h.title.zh || h.title.en)) || '隐藏功能与操作说明';
    let body = '';
    (h.items || []).forEach((item, i) => {
        const t = (item.title && (item.title[lang] || item.title.zh || item.title.en)) || '';
        const c = (item[lang] || item.zh || item.en || '').replace(/</g, '&lt;').replace(/\n/g, '<br>');
        body += `<div class="hidden-feature-item"><div class="hidden-feature-title">${t}</div><div class="hidden-feature-content">${c}</div></div>`;
    });
    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.style.display = 'flex';
    modal.innerHTML = `<div class="modal-card hidden-features-modal" onclick="event.stopPropagation()">
            <div class="modal-title">${title}</div>
            <div class="hidden-features-body">${body}</div>
            <button class="btn-block" style="background:var(--accent-color);color:#111;" onclick="this.closest('.modal-overlay').remove()">${lang === 'zh' ? '关闭' : (lang === 'en' ? 'Close' : 'Zapri')}</button>
        </div>`;
    modal.onclick = () => modal.remove();
    document.body.appendChild(modal);
}
function showHelp(key) {
    const h = HELP_TEXTS[key];
    const titleBase = currentLang === 'zh' ? '功能介绍：' : (currentLang === 'en' ? 'Help: ' : 'Pomoč: ');
    const featureTitle = (h && h.title && (h.title[currentLang] || h.title.zh || h.title.en)) || key;
    let txt = (h && h[currentLang]) || (h && h.zh) || (h && h.en) || (currentLang === 'zh' ? '暂无说明' : 'No description');
    const parts = txt.split(/\*\*(.*?)\*\*/);
    let html = '';
    for (let i = 0; i < parts.length; i++) {
        const esc = parts[i].replace(/</g, '&lt;').replace(/\n/g, '<br>');
        html += (i % 2 === 1) ? '<strong>' + esc + '</strong>' : esc;
    }
    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.style.display = 'flex';
    modal.innerHTML = `<div class="modal-card" onclick="event.stopPropagation()" style="max-width:420px;">
            <div class="modal-title">${titleBase}${featureTitle}</div>
            <div style="padding:16px;line-height:1.6;color:#ccc;word-wrap:break-word;overflow-wrap:break-word;white-space:normal;">${html}</div>
            <button class="btn-block" style="background:var(--accent-color);color:#111;" onclick="this.closest('.modal-overlay').remove()">${currentLang === 'zh' ? '关闭' : (currentLang === 'en' ? 'Close' : 'Zapri')}</button>
        </div>`;
    modal.onclick = () => modal.remove();
    document.body.appendChild(modal);
}
function toggleBeginnerMode() {
    appSettings.beginnerMode = document.getElementById('checkBeginnerMode').checked;
    saveSettings();
    updateBeginnerModeUI();
}
function updateBeginnerModeUI() {
    document.body.classList.toggle('beginner-mode', !!appSettings.beginnerMode);
}
function cleanInput(val) { if (!val) return ''; return val.replace(/,/g, '.').replace(/[^0-9.]/g, ''); }
function autoFixInput(input) {
    let val = input.value;
    if (val.includes(',')) {
        input.value = val.replace(/,/g, '.');
        const id = input.getAttribute('data-id');
        const field = input.getAttribute('data-field');
        if (id && field) updateItem(parseInt(id), field, input.value);
    }
    const field = input.getAttribute('data-field');
    if (field === 'length' && isQuickMode && appSettings.quickModeAutoDecimal && /^\d+$/.test(val)) {
        const num = parseInt(val, 10);
        let newVal = null;
        if (num >= 20 && num <= 99) newVal = (num / 10).toString();
        else if (num >= 101 && num <= 159) newVal = (num / 10).toString();
        if (newVal) {
            input.value = newVal;
            const id = input.getAttribute('data-id');
            if (id) updateItem(parseInt(id), 'length', newVal);
        }
    }
}

function getCurrencySymbol() {
    if (appSettings.priceCurrency === 'EUR') return '€';
    if (appSettings.priceCurrency === 'USD') return '$';
    if (appSettings.priceCurrency === 'CNY') return '¥';
    return appSettings.priceCurrency || '';
}
function updateCurrencySymbols() {
    const symbol = getCurrencySymbol();
    const fixedSymbol = document.getElementById('currencySymbolFixed');
    if (fixedSymbol) fixedSymbol.innerText = symbol;
    document.querySelectorAll('.currencySymbolGrade').forEach(el => el.innerText = symbol);
}
function getUnitPriceForLog(log) {
    if (appSettings.priceMode === 'grade') {
        const p = getPriceForGrade(log.grade);
        return parseFloat(p) || 0;
    }
    return parseFloat(appSettings.priceFixed) || 0;
}
function calcLogAmountBeforeTax(log) {
    const unit = getUnitPriceForLog(log);
    const v = parseFloat(log.volume) || 0;
    return unit * v;
}
function formatMoney(n) { return (isNaN(n) ? '' : n.toFixed(2)); }

/**
 * 体积显示格式化：根据公式模式自适应精度，消除末位虚假 0。
 * Czech ČSN 48：标准表两位，但保留三位真实精度（如 1.281）。
 * Standard Huber：保留三位。一律不强制补 0。
 */
function formatVolumeForDisplay(volume) {
    if (volume == null || isNaN(volume)) return '0';
    const v = Number(volume);
    if (v === 0) return '0';
    const isCzech = appSettings.formulaEnabled && appSettings.formula === 'csn4800079';
    const decimals = isCzech ? 3 : 3;
    return parseFloat(v.toFixed(decimals)).toString();
}

function toggleQuickMode() { isQuickMode = !isQuickMode; localStorage.setItem(QUICK_KEY, isQuickMode); updateQuickBtnUI(); }
function updateQuickBtnUI() {
    const btn = document.getElementById('quickModeBtn');
    const t = I18N[currentLang];
    const label = btn?.querySelector('#quickModeLabel');
    if (btn) {
        btn.classList.toggle('active', isQuickMode);
        if (label) label.innerText = isQuickMode ? ' ' + t.quick_on : ' ' + t.quick_off;
    }
}
function toggleQuickModeExpand() {
    const opts = document.getElementById('quickModeOptions');
    const icon = document.getElementById('quickModeExpandIcon');
    if (opts && icon) {
        const show = opts.style.display !== 'block';
        opts.style.display = show ? 'block' : 'none';
        icon.style.transform = show ? 'rotate(180deg)' : '';
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }
}
function saveQuickModeOptions() {
    const a = document.getElementById('quickModeAutoDecimal');
    const b = document.getElementById('quickModeAutoJump');
    const c = document.getElementById('quickModeGradeAutoSave');
    const d = document.getElementById('quickModeAutoCode');
    if (a) appSettings.quickModeAutoDecimal = a.checked;
    if (b) appSettings.quickModeAutoJump = b.checked;
    if (c) appSettings.quickModeGradeAutoSave = c.checked;
    if (d) appSettings.quickModeAutoCode = d.checked;
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(appSettings));
}
function updateProKeyboardBtnUI() {
    const btn = document.getElementById('proKeyboardBtn');
    if (!btn) return;
    btn.classList.toggle('active', !!appSettings.proKeyboard);
    btn.innerText = appSettings.proKeyboard ? 'ON' : 'OFF';
}

function getActiveLog() { return logs.length > 0 ? logs[0] : null; }
function applyProKeyboardUI() {
    document.body.classList.toggle('pro-kb-enabled', !!appSettings.proKeyboard);
    appSettings.useVirtualKeyboard = !!appSettings.proKeyboard;
    if (appSettings.proKeyboard) {
        document.body.classList.remove('pro-kb-collapsed');
        renderProKeyboardTopBar();
        syncProStateFromLog();
        updateProSideState();
        setProKeypadMode('num');
        setProActiveField(proState.activeField || 'length');
    } else {
        document.body.classList.remove('pro-kb-collapsed');
    }
}
function renderProKeyboardTopBar() {
    const top = document.getElementById('proKbTop');
    if (!top) return;
    const t = I18N[currentLang];
    const fields = [
        { key: 'code', label: t.kb_code, clickable: true, readonly: true },
        { key: 'length', label: t.kb_len, clickable: true, readonly: true },
        { key: 'dia', label: t.kb_dia, clickable: true, readonly: true },
        { key: 'note', label: t.kb_note, clickable: true, readonly: true, noNativeKb: true }
    ];
    top.className = 'pro-kb-top four';
    top.innerHTML = fields.map(f => {
        const clickAttr = f.clickable ? `onclick="setProActiveField('${f.key}')"` : '';
        const readonlyAttr = (f.readonly || f.noNativeKb) ? 'readonly' : '';
        const noKbAttr = f.noNativeKb ? ' inputmode="none" tabindex="-1" autocomplete="off"' : '';
        const cls = f.clickable ? 'kb-field' : 'kb-field readonly';
        return `
            <div class="${cls}" data-field="${f.key}" ${clickAttr}>
                <div class="kb-field-label">${f.label}</div>
                <input id="kbv-${f.key}" class="kb-field-value" ${readonlyAttr}${noKbAttr} />
            </div>
        `;
    }).join('');
    if (!fields.some(f => f.key === proState.activeField) && !['dia1', 'dia2'].includes(proState.activeField)) {
        proState.activeField = 'length';
    }
    updateProTopBarValues();
    updateProActiveUI();
}
function syncProStateFromLog() {
    const log = getActiveLog();
    if (!log) return;
    proState.values.code = log.code || '';
    proState.values.length = log.length || '';
    proState.values.dia = log.diameter || '';
    proState.values.note = log.note || '';
    proState.gradeReady = false;
    updateProTopBarValues();
    renderProGradePanel();
    updateProVolumeDisplay();
}
function updateProTopBarValues() {
    const log = getActiveLog();
    const d1 = proState.values.dia1 || '';
    const d2 = proState.values.dia2 || '';
    const diaDisplay = appSettings.calcDia ? [d1, d2].filter(Boolean).join(' / ') : (proState.values.dia || '');
    const map = {
        code: proState.values.code,
        length: proState.values.length,
        dia: diaDisplay,
        dia1: proState.values.dia1,
        dia2: proState.values.dia2,
        note: proState.values.note
    };
    Object.keys(map).forEach(k => {
        const el = document.getElementById('kbv-' + k);
        if (el) el.value = map[k] || '';
    });
}
function updateProSideState() {
    const grdBtn = document.getElementById('kbSideGrd');
    if (grdBtn) grdBtn.classList.toggle('disabled', !appSettings.showGrade);
}
function normalizeProField(field) {
    const valid = ['code', 'length', 'dia', 'dia1', 'dia2', 'note'];
    if (!valid.includes(field)) field = 'length';
    if (appSettings.calcDia) {
        if (field === 'dia') field = 'dia1';
    } else {
        if (field === 'dia1' || field === 'dia2') field = 'dia';
    }
    return field;
}
function setProActiveField(field) {
    expandProKeyboard();
    if (appSettings.useVirtualKeyboard) flushProKeyPending();
    proState.activeField = normalizeProField(field);
    setProOkFocused(false);
    updateProActiveUI();
    focusProFieldInput();
}
function updateProActiveUI() {
    document.querySelectorAll('.kb-field').forEach(f => f.classList.remove('active'));
    const uiField = (proState.activeField === 'dia1' || proState.activeField === 'dia2') ? 'dia' : proState.activeField;
    const activeField = document.querySelector(`.kb-field[data-field="${uiField}"]`);
    if (activeField) activeField.classList.add('active');
}
function setProKeypadMode(mode) {
    proState.keypadMode = mode;
    const core = document.getElementById('kbCore');
    if (core) core.classList.toggle('grade-mode', mode === 'grade');
    const grdBtn = document.getElementById('kbSideGrd');
    if (grdBtn) grdBtn.classList.toggle('active', mode === 'grade');
    if (mode === 'grade') renderProGradePanel();
    if (mode === 'grade') setProOkFocused(false);
}
function handleProSide(type) {
    if (appSettings.useVirtualKeyboard) flushProKeyPending();
    if (type === 'hide') { collapseProKeyboard(); return; }
    if (type === 'minus') { handleProMinus(); return; }
    if (type === 'grd') {
        if (!appSettings.showGrade) return;
        setProOkFocused(false);
        if (proState.keypadMode === 'grade') setProKeypadMode('num');
        else {
            proState.gradeReady = false;
            setProKeypadMode('grade');
        }
        return;
    }
}
function focusProFieldInput() {
    if (proState.activeField === 'note') {
        if (document.activeElement && document.activeElement.tagName === 'INPUT') document.activeElement.blur();
        return;
    }
    const uiField = (proState.activeField === 'dia1' || proState.activeField === 'dia2') ? 'dia' : proState.activeField;
    const el = document.getElementById('kbv-' + uiField);
    if (el) { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
}
function setProOkFocused(isFocused) {
    const okBtn = document.querySelector('.kb-ok');
    if (!okBtn) return;
    if (isFocused) {
        okBtn.classList.add('focused');
        try { okBtn.focus({ preventScroll: true }); } catch (e) { okBtn.focus(); }
    } else {
        okBtn.classList.remove('focused');
    }
}
function collapseProKeyboard() { document.body.classList.add('pro-kb-collapsed'); }
function expandProKeyboard() { document.body.classList.remove('pro-kb-collapsed'); }
function handleProMinus() {
    const field = proState.activeField;
    if (!['code', 'note'].includes(field)) return;
    const current = proState.values[field] || '';
    updateProFieldValue(field, current + '-');
}
function updateProFieldValue(field, value) {
    const log = getActiveLog();
    if (!log) return;
    proState.values[field] = value;
    if (field === 'code') updateItem(log.id, 'code', value);
    if (field === 'length') updateItem(log.id, 'length', value);
    if (field === 'dia') updateItem(log.id, 'diameter', value);
    if (field === 'note') updateItem(log.id, 'note', value);
    if (field === 'note') {
        const noteInput = document.getElementById('kbv-note');
        if (noteInput && (document.activeElement !== noteInput || noteInput.value !== value)) noteInput.value = value || '';
    } else {
        updateProTopBarValues();
    }
    updateRowFieldDom(log.id, field, value);
    updateProVolumeDisplay();
}
function flushProKeyPending() {
    if (!proKeyPending) return;
    const { field, buffer } = proKeyPending;
    proKeyPending = null;
    if (!buffer) return;
    if (appSettings.keySound && typeof playKeySound === 'function') { for (let i = 0; i < buffer.length; i++) playKeySound(buffer[i]); }
    const current = proState.values[field] || '';
    const next = current + buffer;
    updateProFieldValue(field, next);
    if (['dia1', 'dia2'].includes(field) && appSettings.calcDia) updateProDualVolume();
    handleProQuickFlow(field);
    focusProFieldInput();
}
function handleProKey(ch) {
    const field = proState.activeField;
    if (!field) return;
    if (field === 'code' && ch === '.') return;
    if (['length', 'dia', 'dia1', 'dia2'].includes(field) && ch === '.' && (proState.values[field] || '').includes('.')) return;
    if (appSettings.useVirtualKeyboard) {
        if (!proKeyPending || proKeyPending.field !== field) {
            flushProKeyPending();
            proKeyPending = { field, buffer: '' };
        }
        proKeyPending.buffer += ch;
        if (proKeyDebounceTimer) clearTimeout(proKeyDebounceTimer);
        proKeyDebounceTimer = setTimeout(() => { flushProKeyPending(); }, 40);
        return;
    }
    if (appSettings.keySound && typeof playKeySound === 'function') playKeySound(ch);
    const current = proState.values[field] || '';
    const next = current + ch;
    updateProFieldValue(field, next);
    if (['dia1', 'dia2'].includes(field) && appSettings.calcDia) updateProDualVolume();
    handleProQuickFlow(field);
    focusProFieldInput();
}
function handleProDelete() {
    if (appSettings.keySound && typeof playKeySound === 'function') playKeySound('del');
    if (appSettings.useVirtualKeyboard) flushProKeyPending();
    const field = proState.activeField;
    if (!field) return;
    const current = proState.values[field] || '';
    const next = current.slice(0, -1);
    updateProFieldValue(field, next);
    if (['dia1', 'dia2'].includes(field) && appSettings.calcDia) updateProDualVolume();
    focusProFieldInput();
}
function handleProNext() {
    if (appSettings.useVirtualKeyboard) flushProKeyPending();
    const field = proState.activeField;
    if (field === 'code') return setProActiveField('length');
    if (field === 'length') return setProActiveField(appSettings.calcDia ? 'dia1' : 'dia');
    if (field === 'dia1') return setProActiveField('dia2');
    if (field === 'dia' || field === 'dia2') {
        if (appSettings.showGrade) {
            if (proState.gradeReady) {
                proState.gradeReady = false;
                return addNewLog();
            }
            proState.gradeReady = false;
            setProKeypadMode('grade');
            return;
        }
        return addNewLog();
    }
    if (field === 'note') return setProActiveField('length');
}
function handleProOk() { if (appSettings.keySound && typeof playKeySound === 'function') playKeySound('nxt'); handleProNext(); }
function handleProQuickFlow(field) {
    if (!isQuickMode) return;
    if (field === 'length') {
        const val = proState.values.length || '';
        if (!/^\d+$/.test(val)) return;
        if (val.length === 2 && val[0] !== '1') {
            updateProFieldValue('length', `${val[0]}.${val[1]}`);
            handleProNext();
            return;
        }
        if (val.length === 3 && val[0] === '1') {
            updateProFieldValue('length', `${val.slice(0, 2)}.${val[2]}`);
            handleProNext();
        }
        return;
    }
    if (!['dia', 'dia1', 'dia2'].includes(field)) return;
    const val = proState.values[field] || '';
    if (!/^\d+$/.test(val)) return;
    const shouldJump = (val[0] !== '1' && val.length === 2) || (val[0] === '1' && val.length === 3);
    if (!shouldJump) return;
    if (field === 'dia1') return setProActiveField('dia2');
    if (appSettings.showGrade) return setProKeypadMode('grade');
    handleProNext();
}
function updateProDualVolume() {
    const log = getActiveLog();
    if (!log) return;

    const length = parseFloat(cleanInput(log.length.toString())) || 0;
    const d1 = getDiameterValue('dia1');
    const d2 = getDiameterValue('dia2');

    // 计算有效直径
    let effectiveDiameter = 0;
    if (d1 > 0 && d2 > 0) {
        effectiveDiameter = (d1 + d2) / 2;
    } else if (d1 > 0) {
        effectiveDiameter = d1;
    } else if (d2 > 0) {
        effectiveDiameter = d2;
    }

    // 使用通用函数计算体积
    log.volume = calculateVolume(length, effectiveDiameter);

    // 更新显示
    const vDisplay = document.getElementById('v-' + log.id);
    if (vDisplay) vDisplay.innerText = formatVolumeForDisplay(log.volume);

    save();
    updateProVolumeDisplay();
}
function updateProVolumeDisplay() {
    const el = document.getElementById('kbVolumeDisplay');
    if (!el) return;
    const { totalV } = getValidLogsGroupStats();
    el.textContent = formatVolumeForDisplay(totalV);
    el.title = (I18N[currentLang].total_vol || '材积') + ': ' + formatVolumeForDisplay(totalV) + ' m³';
}
function updateRowFieldDom(id, field, value) {
    const row = document.getElementById('row-' + id);
    if (!row) return;
    const rowField = field === 'dia' ? 'diameter' : field;
    const input = row.querySelector(`input[data-field="${rowField}"]`);
    if (input) input.value = value || '';
}
function highlightNewestRow() {
    const rows = document.querySelectorAll('.log-row');
    rows.forEach(r => r.classList.remove('latest'));
    if (rows.length > 0) rows[rows.length - 1].classList.add('latest');
}
function scrollLogListToBottom() {
    const list = document.getElementById('logList');
    if (!list) return;
    setTimeout(() => {
        list.scrollTop = list.scrollHeight;
        requestAnimationFrame(() => {
            list.scrollTop = list.scrollHeight;
        });
    }, 0);
}
function renderProGradePanel() {
    const box = document.getElementById('proKbGrade');
    if (!box) return;
    const log = getActiveLog();
    const activeGrade = log ? log.grade : '';
    const gradeLabels = getGradeLabels();
    const esc = (s) => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    box.innerHTML = gradeLabels.map((g, idx) => `
            <button class="kb-key kb-grade-key ${activeGrade === g ? 'active' : ''}" onclick="selectProGrade('${esc(g)}')" oncontextmenu="handleGradeLabelEdit(${idx}, event); return false" ondblclick="handleGradeLabelEdit(${idx}, event)" title="${currentLang === 'zh' ? '双击可修改按钮文字' : (currentLang === 'en' ? 'Double-click to change button label' : 'Dvojni klik za spremembo')}">${g}</button>
        `).join('');
}
function selectProGrade(grade) {
    const log = getActiveLog();
    if (!log) return;
    if (appSettings.keySound && typeof playKeySound === 'function') playKeySound('grade');
    if (appSettings.useVirtualKeyboard) flushProKeyPending();
    setGrade(log.id, grade);
    // 在快速模式下且开启选择等级即保存，选择等级后自动提交并开始下一根
    if (isQuickMode && appSettings.quickModeGradeAutoSave) {
        proState.gradeReady = false;
        addNewLog();
        return;
    }
    proState.gradeReady = true;
    renderProGradePanel();
    setProOkFocused(true);
    // 快速模式且关闭选择等级即保存时：停留在等级页面，不切换回数字键盘，用户可点 NXT 保存
    if (!isQuickMode) setProKeypadMode('num');
}

/* ----------------------------------------------------------------------------
   通用快速输入工具函数 (Quick Input Utilities)
   ---------------------------------------------------------------------------- */

// 检查输入值是否满足自动跳转条件（2位非1开头或3位数字）
function shouldAutoJump(val) {
    if (!val || !/^\d+$/.test(val)) return false;
    return (val.length === 2 && !val.startsWith('1')) || val.length === 3;
}

// 通用的保存并跳转到下一根的逻辑
function saveAndJumpToNext() {
    setTimeout(() => {
        addNewLog();
        setTimeout(() => {
            const lenInput = document.querySelector(`.log-card .field input[data-field="length"]`);
            if (lenInput) lenInput.focus();
        }, 0);
    }, 150);
}

// 通用的跳转到指定输入框
function focusNextInput(targetId) {
    const target = document.getElementById(targetId);
    if (target) target.focus();
}

/* ----------------------------------------------------------------------------
   快速输入处理函数 (Quick Input Handlers)
   ---------------------------------------------------------------------------- */

function handleQuickLength(input) {
    if (!isQuickMode) return;
    let val = input.value;
    if (!val) return;

    // 情况1：纯数字（如 35、125）自动补小数点并跳转
    if (appSettings.quickModeAutoDecimal && /^\d+$/.test(val)) {
        if (val.startsWith('1')) return;
        let num = parseInt(val);
        let newVal = null;
        if (num >= 20 && num <= 99) newVal = (num / 10).toString();
        else if (num >= 101 && num <= 159) newVal = (num / 10).toString();

        if (newVal) {
            input.value = newVal;
            if (logs.length > 0) logs[0].length = newVal;
            updateItem(logs[0].id, 'length', newVal);
            if (appSettings.quickModeAutoJump) jumpLengthToDia();
        }
        return;
    }

    // 情况2：已带小数点（如 3.5、12.5）直接跳转到直径
    if (appSettings.quickModeAutoJump && /^\d+\.\d+$/.test(val)) {
        jumpLengthToDia();
    }
}

function jumpLengthToDia() {
    const nextTarget = appSettings.calcDia ? 'in-dia-1' : null;
    if (nextTarget) {
        focusNextInput(nextTarget);
    } else {
        const diaInput = document.querySelector(`.log-card .field input[data-field="diameter"]`);
        if (diaInput) diaInput.focus();
    }
}

function handleQuickDiameterSingle(input) {
    if (!isQuickMode || !appSettings.quickModeAutoJump || appSettings.showGrade) return;
    const val = input.value;
    if (!val) return;

    if (shouldAutoJump(val)) {
        updateItem(logs[0].id, 'diameter', val);
        saveAndJumpToNext();
    }
}

function handleQuickDia1(input) {
    if (!isQuickMode || !appSettings.quickModeAutoJump) return;
    const val = input.value;
    if (!val) return;

    if (shouldAutoJump(val)) {
        focusNextInput('in-dia-2');
    }
}

function handleQuickDia2(input) {
    if (!isQuickMode || !appSettings.quickModeAutoJump || appSettings.showGrade) return;
    const val = input.value;
    if (!val) return;

    if (shouldAutoJump(val)) {
        saveAndJumpToNext();
    }
}

function calculateDualDia(d1, d2) {
    const avg = (d1 + d2) / 2;
    if (avg % 1 === 0.5) {
        if (appSettings.roundMode === 'up') return Math.ceil(avg);
        if (appSettings.roundMode === 'down') return Math.floor(avg);
        if (appSettings.roundMode === 'mix') {
            const res = nextRoundUp ? Math.ceil(avg) : Math.floor(avg);
            nextRoundUp = !nextRoundUp;
            localStorage.setItem(MIX_STATE_KEY, nextRoundUp);
            return res;
        }
    }
    return Math.round(avg);
}

/* ----------------------------------------------------------------------------
   日志管理工具函数 (Log Management Utilities)
   ---------------------------------------------------------------------------- */

function getNextCode() {
    if (logs.length === 0) return '';
    const match = logs[0].code.match(/^(.*?)(\d+)$/);
    return match ? match[1] + (parseInt(match[2]) + 1).toString().padStart(match[2].length, '0') : logs[0].code;
}

// 获取直径值（支持专业键盘和普通输入）
function getDiameterValue(field) {
    const usePro = appSettings.proKeyboard;
    if (usePro) {
        const value = proState.values[field] || '';
        return parseFloat(cleanInput(value.toString())) || 0;
    } else {
        const inputId = field === 'dia1' ? 'in-dia-1' : 'in-dia-2';
        return parseFloat(document.getElementById(inputId)?.value) || 0;
    }
}

// 应用扣减（长度或直径）
function applyDeduction(value, deduction, isLength = false) {
    if (value <= 0 || deduction <= 0) return value;

    const deducted = isLength
        ? value - (deduction / 100)  // 长度扣减单位是厘米，需要转换
        : value - deduction;          // 直径扣减直接减

    return Math.max(0, deducted);
}

/* ----------------------------------------------------------------------------
   添加新日志函数 (Add New Log)
   ---------------------------------------------------------------------------- */

function addNewLog(forceSave = false) {
    if (forceSave && appSettings.keySound && typeof playKeySound === 'function') playKeySound('ok');
    if (logs.length > 0) {
        const current = logs[0];

        // 处理双径模式
        if (appSettings.calcDia) {
            const d1 = getDiameterValue('dia1');
            const d2 = getDiameterValue('dia2');

            // 双径模式：优先使用双径平均值，其次使用单个有效值
            if (d1 > 0 && d2 > 0) {
                current.diameter = calculateDualDia(d1, d2).toString();
            } else if (d1 > 0) {
                current.diameter = d1.toString();
            } else if (d2 > 0) {
                current.diameter = d2.toString();
            }
        }

        // 解析当前值
        let rawLen = parseFloat(cleanInput(current.length.toString()));
        let rawDia = parseFloat(cleanInput(current.diameter.toString()));

        // 虚拟键盘模式下验证数据有效性
        if (appSettings.useVirtualKeyboard && !forceSave) {
            const lenOk = !isNaN(rawLen) && rawLen > 0;
            const diaOk = !isNaN(rawDia) && rawDia > 0;
            if (!lenOk || !diaOk) return;
        }

        // 应用扣减
        let changed = false;
        if (appSettings.deductLen > 0) {
            const deducted = applyDeduction(rawLen, appSettings.deductLen, true);
            if (deducted !== rawLen) {
                rawLen = deducted;
                current.length = parseFloat(rawLen.toFixed(2)).toString();
                changed = true;
            }
        }
        if (appSettings.deductDia > 0) {
            const deducted = applyDeduction(rawDia, appSettings.deductDia, false);
            if (deducted !== rawDia) {
                rawDia = deducted;
                current.diameter = rawDia.toString();
                changed = true;
            }
        }

        // 重新计算体积
        if (changed || appSettings.calcDia) {
            current.volume = calculateVolume(rawLen, rawDia);
        }
    }

    document.querySelectorAll('input').forEach(i => autoFixInput(i));
    const savedLog = logs.length > 0 ? logs[0] : null;
    const nextCode = (isQuickMode && !appSettings.quickModeAutoCode) ? '' : getNextCode();
    const newLog = { id: Date.now(), code: nextCode, grade: '', length: '', diameter: '', volume: 0, note: '', markGrade: false, markLen: false, markDia: false };
    logs.unshift(newLog);
    save();

    const useVirtual = !!appSettings.useVirtualKeyboard;
    if (useVirtual && savedLog) {
        const container = document.getElementById('logList');
        if (!container.querySelector('.list-header')) {
            const t = I18N[currentLang];
            const header = document.createElement('div');
            header.className = 'list-header';
            header.innerHTML = `<div>${t.idx}</div><div>${t.code}</div><div>${t.grade}</div><div>${t.len}</div><div>${t.dia}</div><div>${t.vol}</div><div>${t.note}</div><div></div><div></div>`;
            container.insertBefore(header, container.firstChild);
        }
        const totalSaved = Math.max(logs.length - 1, 0);
        const realIndex = totalSaved;
        const newRow = createRow(savedLog, realIndex);
        newRow.classList.add('latest');
        container.appendChild(newRow);
        const prevLatest = container.querySelector('.log-row.latest:not(:last-child)');
        if (prevLatest) prevLatest.classList.remove('latest');
        updateStats();
        activeFilter = null;
        applyFilter();
        highlightNewestRow();
        scrollLogListToBottom();
    } else {
        renderAll();
        activeFilter = null;
        applyFilter();
    }

    setTimeout(() => {
        if (appSettings.proKeyboard) {
            proState.values.length = '';
            proState.values.dia = '';
            proState.values.dia1 = '';
            proState.values.dia2 = '';
            proState.gradeReady = false;
            syncProStateFromLog();
            setProKeypadMode('num');
            setProActiveField('length');
        } else {
            const lenInput = document.querySelector(`.log-card .field input[data-field="length"]`);
            if (lenInput) { lenInput.focus(); lenInput.click(); }
        }
    }, 10);
}

function renderAll() {
    const container = document.getElementById('logList');
    container.innerHTML = '';
    const t = I18N[currentLang];
    if (logs.length > 1) {
        const header = document.createElement('div');
        header.className = 'list-header';
        header.innerHTML = `<div>${t.idx}</div><div>${t.code}</div><div>${t.grade}</div><div>${t.len}</div><div>${t.dia}</div><div>${t.vol}</div><div>${t.note}</div><div></div><div></div>`;
        container.appendChild(header);
    }
    const useVirtual = !!appSettings.useVirtualKeyboard;
    if (useVirtual) {
        const totalSaved = Math.max(logs.length - 1, 0);
        for (let i = logs.length - 1; i >= 1; i--) {
            const log = logs[i];
            const realIndex = totalSaved - i + 1;
            container.appendChild(createRow(log, realIndex));
        }
    } else {
        logs.forEach((log, index) => {
            const realIndex = logs.length - index;
            if (index === 0) container.insertBefore(createCard(log, realIndex), container.firstChild);
            else container.appendChild(createRow(log, realIndex));
        });
    }
    updateStats();
    if (appSettings.proKeyboard) syncProStateFromLog();
    applyFilter();
    if (useVirtual) {
        highlightNewestRow();
        scrollLogListToBottom();
    }
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function createCard(log, idx) {
    const t = I18N[currentLang];
    let btnsHtml = '';
    const gradeLabels = getGradeLabels();
    gradeLabels.forEach((g, idx) => {
        const isActive = log.grade === g ? 'active' : '';
        const esc = (s) => String(s).replace(/'/g, "\\'").replace(/"/g, '&quot;');
        btnsHtml += `<button class="btn-grade ${isActive}" onmousedown="event.preventDefault()" onclick="setGrade(${log.id}, '${esc(g)}')" oncontextmenu="handleGradeLabelEdit(${idx}, event); return false" ondblclick="handleGradeLabelEdit(${idx}, event)" title="${currentLang === 'zh' ? '双击可修改按钮文字' : (currentLang === 'en' ? 'Double-click to change button label' : 'Dvojni klik za spremembo')}">${g}</button>`;
    });

    const gradeDisplayClass = appSettings.showGrade ? '' : 'hidden';

    let diaInputHtml = '';
    if (appSettings.calcDia) {
        diaInputHtml = `
                <div class="double-input-container">
                    <input type="text" inputmode="decimal" id="in-dia-1" oninput="handleQuickDia1(this)">
                    <input type="text" inputmode="decimal" id="in-dia-2" oninput="handleQuickDia2(this)">
                </div>
            `;
    } else {
        diaInputHtml = `
                <input type="text" inputmode="decimal" value="${log.diameter}" data-id="${log.id}" data-field="diameter" oninput="updateItem(${log.id},'diameter',this.value); handleQuickDiameterSingle(this)" onblur="autoFixInput(this)">
            `;
    }

    const div = document.createElement('div');
    div.className = 'log-card';
    div.innerHTML = `
            <div class="card-title"><span>${t.inputting}</span><span class="card-title-no">NO. ${idx}</span></div>
            <div class="card-row-1">
                <div class="field"><label>${t.code}</label><input type="text" inputmode="numeric" value="${log.code}" oninput="updateItem(${log.id},'code',this.value)"></div>
                <div class="field"><label>${t.note}</label><input type="text" class="input-note" value="${log.note || ''}" oninput="updateItem(${log.id},'note',this.value)"></div>
            </div>
            <div class="card-row-2">
                <div class="field"><label>${t.len}</label>
                    <input type="text" inputmode="decimal" value="${log.length}" data-id="${log.id}" data-field="length" oninput="updateItem(${log.id},'length',this.value); handleQuickLength(this)" onblur="autoFixInput(this)">
                </div>
                <div class="field"><label>${t.dia}</label>
                    ${diaInputHtml}
                </div>
                <div class="live-vol-display" id="v-${log.id}">${formatVolumeForDisplay(log.volume)}</div>
            </div>
            <div class="grade-section ${gradeDisplayClass}">
                <div class="grade-label">${t.select_grade}</div>
                <div class="grade-container">${btnsHtml}</div>
            </div>
            <button class="btn-add-inline" onclick="addNewLog()"><i data-lucide="plus" style="width:20px;height:20px;margin-right:4px;"></i>${t.add_next}</button>
        `;
    return div;
}

function createRow(log, idx) {
    const gradeLabels = getGradeLabels();
    const customGrades = [...new Set(logs.map(l => l.grade).filter(g => g && !gradeLabels.includes(g)))];
    const allGrades = [...gradeLabels, ...customGrades];
    let gradeOptions = `<option value="">-</option>`;
    gradeOptions += allGrades.map(g => `<option value="${g}" ${log.grade === g ? 'selected' : ''}>${g}</option>`).join('');

    const div = document.createElement('div');
    div.className = 'log-row' + (log.groupId ? ' grouped' : '');
    div.id = 'row-' + log.id;
    div.setAttribute('data-log-id', log.id);
    div.setAttribute('data-grade', log.grade || '?');
    const len_m = parseFloat(cleanInput(log.length.toString()));
    const d_cm = parseFloat(cleanInput(log.diameter.toString()));
    div.setAttribute('data-len', len_m || 0);
    div.setAttribute('data-dia', d_cm || 0);
    const warnClass = (len_m > 15) ? 'danger-text' : '';
    const diaDangerClass = (d_cm >= 200 && d_cm < 1000) ? 'dia-danger-text' : '';

    const displayLen = log.length ? parseFloat(log.length) : '';
    const displayDia = log.diameter ? parseFloat(log.diameter) : '';

    const markGrade = !!log.markGrade; const markLen = !!log.markLen; const markDia = !!log.markDia;
    div.innerHTML = `
            <div class="row-index" ondblclick="toggleMarkMode()" title="${currentLang === 'zh' ? '双击进入/退出标记模式' : (currentLang === 'en' ? 'Double-click to enter/exit mark mode' : 'Dvojni klik za označevanje')}">${idx}</div>
            <input type="text" data-field="code" value="${log.code}" oninput="updateItem(${log.id},'code',this.value)">
            <div class="mark-cell ${markGrade ? 'marked' : ''}" data-id="${log.id}" data-field="grade" oncontextmenu="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event); return false" ondblclick="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event)" title="${currentLang === 'zh' ? '双击可修改等级按钮文字' : (currentLang === 'en' ? 'Double-click to change grade button label' : 'Dvojni klik za spremembo')}">
                <span class="mark-symbol">↑</span>
                <select onchange="updateItem(${log.id},'grade',this.value)">${gradeOptions}</select>
                <div class="mark-overlay" onclick="handleMarkCellClick(event, ${log.id}, 'grade')"></div>
            </div>
            <div class="mark-cell ${markLen ? 'marked' : ''}" data-id="${log.id}" data-field="length">
                <span class="mark-symbol">↑</span>
                <input type="text" data-field="length" id="inp-len-${log.id}" class="${warnClass}" inputmode="decimal" value="${displayLen}" oninput="updateItem(${log.id},'length',this.value)" onblur="autoFixInput(this)">
                <div class="mark-overlay" onclick="handleMarkCellClick(event, ${log.id}, 'length')"></div>
            </div>
            <div class="mark-cell ${markDia ? 'marked' : ''}" data-id="${log.id}" data-field="diameter">
                <span class="mark-symbol">↑</span>
                <input type="text" data-field="diameter" class="${diaDangerClass}" inputmode="decimal" value="${displayDia}" oninput="updateItem(${log.id},'diameter',this.value);toggleDiaDangerClass(this)" onblur="autoFixInput(this)">
                <div class="mark-overlay" onclick="handleMarkCellClick(event, ${log.id}, 'diameter')"></div>
            </div>
            <div class="col-vol" id="v-row-${log.id}">${formatVolumeForDisplay(log.volume)}</div>
            <input type="text" data-field="note" style="font-size:12px;color:#aaa;" value="${log.note || ''}" oninput="updateItem(${log.id},'note',this.value)">
            <div class="group-actions-cell">${log.groupId ? `<button type="button" class="btn-ungroup" onclick="ungroupLog(${log.id});event.stopPropagation()" title="${currentLang === 'zh' ? '解除分组' : (currentLang === 'en' ? 'Ungroup' : 'Razdruži')}">⎋</button>` : ''}</div>
            <button class="btn-del-mini" onclick="delItem(${log.id})">×</button>
        `;
    div.insertAdjacentHTML('afterbegin', `<div class="group-select-overlay" onclick="handleGroupRowClick(event, ${log.id})"></div>`);
    return div;
}

function toggleDiaDangerClass(input) {
    const v = parseFloat(cleanInput((input.value || '').toString()));
    const isDanger = !isNaN(v) && v >= 200 && v < 1000;
    input.classList.toggle('dia-danger-text', isDanger);
}
function toggleMarkMode() { isMarkMode = !isMarkMode; document.body.classList.toggle('mark-mode', isMarkMode); }
function updateGroupBtnUI() {
    const btn = document.getElementById('btnGroupMode');
    if (!btn) return;
    btn.classList.toggle('active', isGroupMode);
    btn.innerText = isGroupMode ? (currentLang === 'zh' ? '取消' : (currentLang === 'en' ? 'Cancel' : 'Prekliči')) : (currentLang === 'zh' ? '分组' : (currentLang === 'en' ? 'Group' : 'Združi'));
    btn.title = currentLang === 'zh' ? '点击进入分组模式，选择两行合并为一组' : (currentLang === 'en' ? 'Click to enter group mode; select two rows to merge' : 'Klikni za način združevanja');
}
function toggleGroupMode() {
    isGroupMode = !isGroupMode;
    groupSelectIds = [];
    document.body.classList.toggle('group-mode', isGroupMode);
    updateGroupBtnUI();
    document.querySelectorAll('.log-row.group-selected').forEach(r => r.classList.remove('group-selected'));
    renderAll();
}
function handleGroupRowClick(e, id) {
    if (!isGroupMode) return;
    e.preventDefault(); e.stopPropagation();
    const row = document.getElementById('row-' + id);
    if (!row || row.classList.contains('grouped')) return;
    if (groupSelectIds.includes(id)) {
        groupSelectIds = groupSelectIds.filter(x => x !== id);
        row.classList.remove('group-selected');
        return;
    }
    groupSelectIds.push(id);
    row.classList.add('group-selected');
    if (groupSelectIds.length >= 2) {
        const id1 = groupSelectIds[0], id2 = groupSelectIds[1];
        const idx1 = logs.findIndex(l => l.id === id1);
        const idx2 = logs.findIndex(l => l.id === id2);
        const adjacent = idx1 >= 0 && idx2 >= 0 && Math.abs(idx1 - idx2) === 1;
        const prevOrder = !adjacent && idx1 >= 0 && idx2 >= 0 ? JSON.parse(JSON.stringify(logs)) : null;
        if (!adjacent && idx1 >= 0 && idx2 >= 0) {
            const log2 = logs[idx2];
            logs.splice(idx2, 1);
            const newIdx1 = logs.findIndex(l => l.id === id1);
            logs.splice(newIdx1 + 1, 0, log2);
            renderAll();
            document.getElementById('row-' + id1)?.classList.add('group-selected');
            document.getElementById('row-' + id2)?.classList.add('group-selected');
        }
        const msg = currentLang === 'zh' ? '是否将这两根合并为一组？' : (currentLang === 'en' ? 'Merge these two rows into one group?' : 'Združiti ti dve vrstici v eno skupino?');
        if (confirm(msg)) {
            const gid = 'group_' + Date.now();
            logs.forEach(l => { if (l.id === id1 || l.id === id2) l.groupId = gid; });
            save();
            toggleGroupMode();
            renderAll();
        } else {
            if (prevOrder) { logs.length = 0; logs.push(...prevOrder); save(); }
            groupSelectIds = [];
            document.querySelectorAll('.log-row.group-selected').forEach(r => r.classList.remove('group-selected'));
            renderAll();
        }
    }
}
function ungroupLog(id) {
    const log = logs.find(l => l.id === id); if (!log || !log.groupId) return;
    const gid = log.groupId;
    logs.forEach(l => { if (l.groupId === gid) delete l.groupId; });
    save();
    renderAll();
}
function handleMarkCellClick(e, id, field) {
    if (!isMarkMode) return;
    e.preventDefault(); e.stopPropagation();
    const log = logs.find(l => l.id === id); if (!log) return;
    const key = field === 'grade' ? 'markGrade' : (field === 'length' ? 'markLen' : 'markDia');
    log[key] = !log[key];
    save();
    const cell = e.currentTarget.closest('.mark-cell'); if (cell) cell.classList.toggle('marked', !!log[key]);
}
function handleGradeLabelEdit(slotIndex, e) {
    if (e) { e.preventDefault(); e.stopPropagation(); }
    const labels = getGradeLabels();
    const current = labels[slotIndex] || GRADES[slotIndex] || '';
    const msg = currentLang === 'zh' ? '修改此等级按钮的文字：' : (currentLang === 'en' ? 'Change this grade button label:' : 'Spremeni oznako gumba:');
    const val = prompt(msg, current);
    if (val !== null) {
        const trimmed = String(val).trim();
        if (trimmed) {
            if (!appSettings.gradeLabels || appSettings.gradeLabels.length !== 6) appSettings.gradeLabels = [...GRADES];
            appSettings.gradeLabels[slotIndex] = trimmed;
            localStorage.setItem(SETTINGS_KEY, JSON.stringify(appSettings));
            save();
            renderAll();
            if (appSettings.proKeyboard) renderProGradePanel();
        }
    }
}
function handleGradeLabelEditByGrade(grade, e) {
    if (!grade) return;
    const labels = getGradeLabels();
    let slotIndex = labels.indexOf(grade);
    if (slotIndex < 0) slotIndex = GRADES.indexOf(grade);
    if (slotIndex >= 0) handleGradeLabelEdit(slotIndex, e);
}
function setGrade(id, grade) {
    updateItem(id, 'grade', grade);
    const useVirtual = !!appSettings.useVirtualKeyboard;
    if (isQuickMode && appSettings.quickModeGradeAutoSave && !useVirtual) {
        addNewLog();
        setTimeout(() => {
            if (!appSettings.proKeyboard) {
                const lenInput = document.querySelector(`.log-card .field input[data-field="length"]`);
                if (lenInput) lenInput.focus();
            }
        }, 0);
    }
    else if (!useVirtual) { renderAll(); }
    if (appSettings.proKeyboard) renderProGradePanel();
}

/* ----------------------------------------------------------------------------
   数据更新工具函数 (Data Update Utilities)
   ---------------------------------------------------------------------------- */

// 显示或隐藏公式选项
function toggleFormulaEnabledUI() {
    const formulaEnabledCb = document.getElementById('formulaEnabled');
    const box = document.getElementById('formulaSelectBox');
    if (formulaEnabledCb && box) {
        box.style.display = formulaEnabledCb.checked ? 'block' : 'none';
    }
}

// 切换公式开关
window.toggleFormulaEnabled = function () {
    toggleFormulaEnabledUI();
    saveSettings();
    recalculateAllVolumes();
};

// 计算体积（通用函数）- 保持高精度，舍入仅在 UI 渲染时进行
function calculateVolume(length, diameter) {
    if (appSettings.formulaEnabled && appSettings.formula === 'csn4800079') {
        return calculateCzechVolume(length, diameter);
    }

    const l = parseFloat(String(length).replace(/,/g, '.'));
    const d = parseFloat(String(diameter).replace(/,/g, '.'));

    if (!isNaN(l) && !isNaN(d) && l > 0 && d > 0) {
        const radius_m = d / 200;
        return Math.PI * Math.pow(radius_m, 2) * l;
    }
    return 0;
}

/**
 * 捷克 ČSN 48 0007/9 (Huber Variant) 体积算法
 * 与图片实测数据严格对齐，规则如下：
 *
 * A. 输入清理:
 *    - 逗号替换为小数点: String(x).replace(/,/g, '.')
 *    - parseFloat 解析，无效则早退 0
 *
 * B. 长度扣除 (Nadměrek / Length Deduction) 四段式:
 *    - length >= 11:  lNet = l - 0.3
 *    - 10 <= length < 11: lNet = l - 0（无扣，与表格 10.1x32->0.67 对齐）
 *    - 8 <= length < 10: lNet = l - 0.1
 *    - length < 8:   lNet = l - 0.2
 *    - 底线保护: lNet = Math.max(0, lNet)，防止超短木材产生负数体积（ Critical Fix ）
 *
 * C. 树皮阶梯扣除 (Bark Deduction):
 *    - 直径向下取整: dFloored = Math.floor(d)
 *    - d <= 44: deduction = 3cm
 *    - 45 <= d <= 50: deduction = 4cm
 *    - d > 50: deduction = 5cm
 *    - dNet = Math.max(0, dFloored - deduction)
 *
 * D. 面积与精度:
 *    - area = 3位小数 floor 截断，与表格 8.3x39->0.83 对齐
 *    - volume = area * lNet
 *    - 返回高精度数值，舍入仅在 formatVolumeForDisplay 中进行；可安全传递给状态管理器
 *
 * 样本验证推导 (11.6, 41) 原始值:
 *   lNet = 11.6 - 0.3 = 11.3  (length >= 11)
 *   d = floor(41) = 41, d <= 44 => deduction = 3, dNet = 38
 *   area = floor(π*38²/40000 * 1000)/1000 = 0.113
 *   volume = 0.113 * 11.3 = 1.2769 -> 显示 1.277（formatVolumeForDisplay 三位）
 *
 * 样本2 (7.3, 43): lNet=7.1, area floor 3位=0.125, vol=0.8875 -> 0.89
 * 样本3 (10.1, 32): lNet=10.1(10-11无扣), area floor=0.066, vol=0.6666 -> 0.67
 * 样本4 (8.3, 39): lNet=8.2, area floor=0.101, vol=0.8282 -> 0.83
 * 超短样本 (0.1, 30): lNet = 0.1 - 0.2 = -0.1 -> Math.max(0,-0.1)=0，vol=0，防止负数漏洞
 */
function calculateCzechVolume(length, diameterOverBark) {
    // A. 输入清理：逗号替换为小数点
    const l = parseFloat(String(length).replace(/,/g, '.'));
    const d = parseFloat(String(diameterOverBark).replace(/,/g, '.'));

    if (isNaN(l) || isNaN(d) || l <= 0 || d <= 0) return 0;

    // B. 四段式长度扣除
    let lNet;
    if (l >= 11) {
        lNet = l - 0.3;
    } else if (l >= 10) {
        lNet = l;  // 10-11m 无扣
    } else if (l >= 8) {
        lNet = l - 0.1;
    } else {
        lNet = l - 0.2;
    }
    // Critical Fix: 底线保护，防止超短木材产生负数体积
    lNet = Math.max(0, lNet);

    // C. 树皮阶梯扣除：直径向下取整
    const dFloored = Math.floor(d);
    let deduction;
    if (dFloored <= 44) {
        deduction = 3;
    } else if (dFloored <= 50) {
        deduction = 4;
    } else {
        deduction = 5;
    }
    const dNet = Math.max(0, dFloored - deduction);

    // D. 横截面积：3 位小数 floor 截断，与表格 8.3x39->0.83 对齐
    const area = Math.floor((Math.PI * Math.pow(dNet, 2) / 40000) * 1000) / 1000;

    // 体积计算，返回高精度数值，舍入由 formatVolumeForDisplay 处理，可正确传递至状态管理器
    return area * lNet;
}

// 切换公式时重新计算全部
window.recalculateAllVolumes = function () {
    if (!logs || logs.length === 0) return;
    let modified = false;
    logs.forEach(item => {
        if (parseFloat(item.length) > 0 && parseFloat(item.diameter) > 0) {
            item.volume = calculateVolume(item.length, item.diameter);
            modified = true;
        }
    });
    if (modified) {
        save();
        renderAll();
    }
};

// 更新体积显示（DOM）
function updateVolumeDisplay(id, volume) {
    const volumeStr = formatVolumeForDisplay(volume);
    const vDisplay = document.getElementById('v-' + id);
    const vRowDisplay = document.getElementById('v-row-' + id);

    if (vDisplay) vDisplay.innerText = volumeStr;
    if (vRowDisplay) vRowDisplay.innerText = volumeStr;
    updateProVolumeDisplay();
}

// 更新长度危险提示
function updateLengthWarning(id, length) {
    const rowLenInput = document.getElementById('inp-len-' + id);
    if (rowLenInput) {
        const l = parseFloat(String(length).replace(/,/g, '.'));
        if (l > 15) {
            rowLenInput.classList.add('danger-text');
        } else {
            rowLenInput.classList.remove('danger-text');
        }
    }
}

// 处理双径模式的体积计算
function handleDualDiameterVolume(item, id) {
    if (!appSettings.calcDia || id !== logs[0].id) return;

    const usePro = appSettings.proKeyboard;
    const d1 = usePro
        ? (parseFloat(cleanInput((proState.values.dia1 || '').toString())) || 0)
        : (parseFloat(document.getElementById('in-dia-1')?.value) || 0);
    const d2 = usePro
        ? (parseFloat(cleanInput((proState.values.dia2 || '').toString())) || 0)
        : (parseFloat(document.getElementById('in-dia-2')?.value) || 0);

    if (d1 > 0 && d2 > 0) {
        const avgDiameter = (d1 + d2) / 2;
        const length = parseFloat(item.length) || 0;
        if (length > 0) {
            item.volume = calculateVolume(length, avgDiameter);
            const vDisplay = document.getElementById('v-' + id);
            if (vDisplay) vDisplay.innerText = formatVolumeForDisplay(item.volume);
        }
    }
}

/* ----------------------------------------------------------------------------
   主数据更新函数 (Main Update Function)
   ---------------------------------------------------------------------------- */

function updateItem(id, field, val) {
    const item = logs.find(l => l.id === id);
    if (!item) return;

    item[field] = val;

    // 处理长度或直径变化
    const isLengthOrDiameter = field === 'length' || field === 'diameter';
    if (isLengthOrDiameter) {
        item.volume = calculateVolume(item.length, item.diameter);
        updateVolumeDisplay(id, item.volume);

        if (field === 'length') {
            updateLengthWarning(id, item.length);
        }
    }

    // 处理等级变化
    if (field === 'grade') {
        if (appSettings.proKeyboard) renderProGradePanel();
    }

    // 双径模式特殊处理
    handleDualDiameterVolume(item, id);

    // 保存并应用筛选
    save();
    if (activeFilter) applyFilter();
}

function delItem(id) { if (confirm('Delete?')) { logs = logs.filter(l => l.id !== id); renderAll(); save(); } }

function resetLogOnly() {
    if (!confirm(I18N[currentLang].confirm_reset || I18N[currentLang].btn_new + '?')) return;
    // 确认后先保存当前数据到内部记录（如有数据）
    const hasData = logs.length > 1 || (logs.length === 1 && (logs[0]?.length || logs[0]?.diameter));
    if (hasData) {
        if (!currentSessionId) {
            const today = new Date();
            const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
            const todaySnapshots = snapshots.filter(s => s.id.startsWith(dateStr));
            const maxNum = todaySnapshots.length > 0 ? Math.max(...todaySnapshots.map(s => parseInt(s.id.split('_')[1]) || 0)) : 0;
            currentSessionId = `${dateStr}_${maxNum + 1}`;
        }
        const snapshotData = {
            id: currentSessionId,
            timestamp: Date.now(),
            date: new Date().toLocaleString(currentLang === 'zh' ? 'zh-CN' : currentLang === 'en' ? 'en-US' : 'sl-SI'),
            logs: JSON.parse(JSON.stringify(logs)),
            global: JSON.parse(JSON.stringify(globalInfo)),
            container: globalInfo.container || '未命名'
        };
        const existingIndex = snapshots.findIndex(s => s.id === currentSessionId);
        if (existingIndex >= 0) snapshots[existingIndex] = snapshotData;
        else snapshots.unshift(snapshotData);
        localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snapshots));
    }
    logs = [];
    globalInfo.container = '';
    globalInfo.note = '';
    globalInfo.description = '';
    document.getElementById('g_container').value = '';
    const gNote = document.getElementById('g_note');
    if (gNote) gNote.value = '';
    const gDesc = document.getElementById('g_description');
    if (gDesc) gDesc.value = '';
    currentSessionId = null;
    localStorage.removeItem(SESSION_KEY);
    activeFilter = null;
    closeHistoryViewer();
    save();
    addNewLog();
    renderAll();
}

function getValidLogsGroupStats(logList) {
    const list = logList || logs;
    const validLogs = list.filter(l => parseFloat(l.volume) > 0);
    const seenGroups = new Set();
    let totalV = 0;
    validLogs.forEach(l => {
        if (l.groupId) seenGroups.add(l.groupId);
        totalV += parseFloat(l.volume) || 0;
    });
    const rootCount = validLogs.filter(l => !l.groupId).length + seenGroups.size;
    return { validLogs, rowCount: validLogs.length, rootCount, totalV };
}
function updateStats() {
    const { validLogs, rowCount, rootCount, totalV } = getValidLogsGroupStats();
    const countEl = document.getElementById('totalCount');
    if (countEl) countEl.innerText = rowCount !== rootCount ? `${rowCount} (${rootCount})` : rowCount;
    document.getElementById('totalVol').innerText = formatVolumeForDisplay(totalV);

    const exceeded = !!appSettings.volumeWarningEnabled && (parseFloat(appSettings.volumeWarningThreshold) || 0) > 0 && totalV > (parseFloat(appSettings.volumeWarningThreshold) || 0);
    document.body.classList.toggle('volume-warning', exceeded);
    if (appSettings.proKeyboard) updateProVolumeDisplay();

    const gStats = {};
    const tL4 = parseFloat(appSettings.statThresholdL4) || 4;
    const tL25 = parseFloat(appSettings.statThresholdL25) || 2.5;
    const tD30 = parseFloat(appSettings.statThresholdD30) || 30;
    let l4 = 0, l25 = 0, d30 = 0;
    validLogs.forEach(l => {
        const g = l.grade || '?';
        gStats[g] = (gStats[g] || 0) + 1;
        const len_m = parseFloat(cleanInput(l.length.toString()));
        const d_cm = parseFloat(cleanInput(l.diameter.toString()));
        if (len_m < tL4) l4++;
        if (len_m < tL25) l25++;
        if (d_cm < tD30) d30++;
    });

    let html = '';
    const mkTag = (key, label, count) => {
        if (count === 0) return '';
        const isActive = activeFilter === key ? 'active-filter' : '';
        return `<div class="stat-tag ${isActive}" onclick="toggleFilter('${key}')" ondblclick="customizeStatThreshold('${key}', event)">${label}: <b>${count}</b></div>`;
    };
    for (let g in gStats) if (g !== '?' && g !== '') html += mkTag('g_' + g, g, gStats[g]);
    html += mkTag('l4', '< ' + tL4 + 'm', l4);
    html += mkTag('l25', '< ' + tL25 + 'm', l25);
    html += mkTag('d30', '< ' + tD30 + 'cm', d30);
    document.getElementById('statsDetail').innerHTML = html;
}

function toggleFilter(key) {
    if (activeFilter === key) activeFilter = null;
    else activeFilter = key;
    updateStats();
    applyFilter();
}
function customizeStatThreshold(key, e) {
    e.preventDefault(); e.stopPropagation();
    const labels = { l4: { zh: '小于X米', en: 'Less than X m', val: appSettings.statThresholdL4 }, l25: { zh: '小于X米', en: 'Less than X m', val: appSettings.statThresholdL25 }, d30: { zh: '小于X厘米', en: 'Less than X cm', val: appSettings.statThresholdD30 } };
    const lb = labels[key]; if (!lb) return;
    const unit = key === 'd30' ? (currentLang === 'zh' ? '厘米' : (currentLang === 'en' ? 'cm' : 'cm')) : (currentLang === 'zh' ? '米' : (currentLang === 'en' ? 'm' : 'm'));
    const promptMsg = (currentLang === 'zh' ? '输入数字（如 3 表示小于3' + unit + '）：' : (currentLang === 'en' ? 'Enter number (e.g. 3 = less than 3' + unit + '):' : 'Vnesite število (npr. 3 = manj kot 3' + unit + '):'));
    const v = prompt(promptMsg, String(lb.val));
    if (v === null || v === '') return;
    const num = parseFloat(v.replace(/,/g, '.'));
    if (isNaN(num) || num <= 0) return;
    if (key === 'l4') appSettings.statThresholdL4 = num;
    else if (key === 'l25') appSettings.statThresholdL25 = num;
    else if (key === 'd30') appSettings.statThresholdD30 = num;
    saveSettings();
    updateStats();
    applyFilter();
}

function applyFilter() {
    const rows = document.querySelectorAll('.log-row');
    if (!activeFilter) {
        rows.forEach(r => { r.classList.remove('dimmed', 'highlighted'); });
        return;
    }
    const tL4 = parseFloat(appSettings.statThresholdL4) || 4;
    const tL25 = parseFloat(appSettings.statThresholdL25) || 2.5;
    const tD30 = parseFloat(appSettings.statThresholdD30) || 30;
    rows.forEach(r => {
        let match = false;
        const g = r.getAttribute('data-grade');
        const l = parseFloat(r.getAttribute('data-len'));
        const d = parseFloat(r.getAttribute('data-dia'));
        if (activeFilter.startsWith('g_') && activeFilter === 'g_' + g) match = true;
        else if (activeFilter === 'l4' && l < tL4) match = true;
        else if (activeFilter === 'l25' && l < tL25) match = true;
        else if (activeFilter === 'd30' && d < tD30) match = true;
        if (match) { r.classList.add('highlighted'); r.classList.remove('dimmed'); }
        else { r.classList.add('dimmed'); r.classList.remove('highlighted'); }
    });
}

function openInfoModal() {
    syncInfoModalCompanyDisplay();
    const cb = document.getElementById('showCompanyInPdf');
    if (cb) cb.checked = appSettings.showCompanyInPdf !== false;
    document.getElementById('infoModal').style.display = 'flex';
}
function closeInfoModal() { document.getElementById('infoModal').style.display = 'none'; updateGlobal(); }

function togglePrintMode() {
    const body = document.body;
    if (body.classList.contains('print-mode')) {
        body.classList.remove('print-mode');
    } else {
        const d = new Date();
        const dateStr = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
        document.getElementById('p_company').innerText = getCompanyName(globalInfo.company);
        document.getElementById('print_row_company').style.display = getCompanyName(globalInfo.company) ? 'table-row' : 'none';
        document.getElementById('p_date').innerText = dateStr;
        document.getElementById('p_container').innerText = globalInfo.container;
        document.getElementById('p_seller').innerText = getCompanyName(globalInfo.seller);
        document.getElementById('p_location').innerText = globalInfo.location;
        document.getElementById('print_row_seller').style.display = (!getCompanyName(globalInfo.seller) && !globalInfo.location) ? 'none' : 'table-row';
        document.getElementById('p_measurer').innerText = globalInfo.measurer;
        document.getElementById('print_row_measurer').style.display = globalInfo.measurer ? 'table-row' : 'none';
        document.getElementById('p_note').innerText = globalInfo.note;
        document.getElementById('print_row_note').style.display = (!globalInfo.note) ? 'none' : 'table-row';
        body.classList.add('print-mode');
        setTimeout(() => { try { window.print(); } catch (e) { } }, 300);
    }
}

async function generatePDF(options = {}) {
    const { jsPDF } = window.jspdf;
    const sourceLogs = Array.isArray(options.logs) ? options.logs : logs;
    const sourceGlobal = (options.global && typeof options.global === 'object') ? options.global : globalInfo;

    // 创建一个临时的隐藏容器来渲染PDF内容
    const pdfContainer = document.createElement('div');
    pdfContainer.style.cssText = 'position:absolute;left:-9999px;width:800px;background:white;padding:40px;';
    document.body.appendChild(pdfContainer);

    const t = I18N[currentLang];
    const d = new Date();
    const dateStr = `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}-${d.getDate().toString().padStart(2, '0')}`;
    const timeStr = `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}_${d.getHours()}-${d.getMinutes()}`;

    function escapeHtml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
    function buildCompanyBlockHtml(co, label) {
        const hasAny = COMPANY_FIELDS.some(f => co[f] && String(co[f]).trim());
        if (!hasAny) return '';
        const lines = [];
        if (co.name) lines.push(`<div style="font-weight:bold; font-size:14px; margin-bottom:4px;">${escapeHtml(co.name)}</div>`);
        if (co.address || co.city || co.zip) {
            const addrLine = [co.address, co.city, co.zip].filter(Boolean).join(' ');
            lines.push(`<div style="margin-bottom:4px;">${escapeHtml(addrLine)}</div>`);
        }
        if (co.phone) lines.push(`<div style="margin-bottom:4px;">${t.company_phone}: ${escapeHtml(co.phone)}</div>`);
        if (co.email || co.website) {
            const parts = [co.email, co.website].filter(Boolean).map(s => escapeHtml(s));
            const emailWeb = parts.join('<span style="display:inline-block; width:10px;"></span>');
            lines.push(`<div style="margin-bottom:4px;">${emailWeb}</div>`);
        }
        if (co.taxId) lines.push(`<div style="margin-bottom:4px;">${t.company_taxId}: ${escapeHtml(co.taxId)}</div>`);
        if (co.bank) lines.push(`<div style="margin-bottom:4px;">${t.company_bank}: ${escapeHtml(co.bank)}</div>`);
        return `<div style="font-size:12px; line-height:1.4;"><div style="font-weight:600; margin-bottom:6px; font-size:11px; color:#555;">${escapeHtml(label)}</div>${lines.join('')}</div>`;
    }
    const company = normalizeCompany(sourceGlobal.company);
    const seller = normalizeSeller(sourceGlobal.seller);
    const showCompany = appSettings.showCompanyInPdf !== false;
    const companyBlockHtml = showCompany ? buildCompanyBlockHtml(company, t.my_company) : '';
    const sellerBlockHtml = buildCompanyBlockHtml(seller, (seller.type === 'seller' ? t.party_seller : t.party_buyer));
    const hasCompanyOrSeller = companyBlockHtml || sellerBlockHtml;
    const twoColBlockHtml = hasCompanyOrSeller ? `<table style="width:100%; margin-bottom:16px; border-collapse:collapse;"><tr><td style="width:35%; vertical-align:top; padding-right:16px;">${companyBlockHtml || '&nbsp;'}</td><td style="width:65%; vertical-align:top; padding-left:24px; text-align:right;">${sellerBlockHtml ? `<div style="display:inline-block; text-align:left;">${sellerBlockHtml}</div>` : '&nbsp;'}</td></tr></table>` : '';

    // 构建HTML内容
    let html = `
            <div style="font-family: Arial, sans-serif; color: #000;">
                <h2 style="text-align:center; margin-bottom:20px; font-size:20px; border-bottom:2px solid #000; padding-bottom:10px;">
                    PACKING LOG LIST / SHIPMENT
                </h2>
                ${twoColBlockHtml}
                <table style="width:100%; border-collapse:collapse; margin-bottom:20px; font-size:12px;">
        `;

    // 项目信息（专业双列布局）- 公司/卖方已移至上方独立块
    const infoItems = [];
    const pushInfo = (label, value) => {
        if (value === undefined || value === null) return;
        const v = String(value).trim();
        if (v === '') return;
        infoItems.push([label, v]);
    };
    infoItems.push([t.p_date, dateStr]);
    pushInfo(t.container, sourceGlobal.container);
    pushInfo(t.description, sourceGlobal.description);
    pushInfo(t.location, sourceGlobal.location);
    pushInfo(t.measurer, sourceGlobal.measurer);
    pushInfo(t.note_global, sourceGlobal.note);

    for (let i = 0; i < infoItems.length; i += 2) {
        const left = infoItems[i];
        const right = infoItems[i + 1];
        if (right) {
            html += `
                    <tr>
                        <td style="border:0.6px solid #999; padding:8px; background:#f7f7f7; width:18%; font-weight:600;">${left[0]}</td>
                        <td style="border:0.6px solid #999; padding:8px; width:32%;">${left[1]}</td>
                        <td style="border:0.6px solid #999; padding:8px; background:#f7f7f7; width:18%; font-weight:600;">${right[0]}</td>
                        <td style="border:0.6px solid #999; padding:8px; width:32%;">${right[1]}</td>
                    </tr>
                `;
        } else {
            html += `
                    <tr>
                        <td style="border:0.6px solid #999; padding:8px; background:#f7f7f7; width:18%; font-weight:600;">${left[0]}</td>
                        <td style="border:0.6px solid #999; padding:8px;" colspan="3">${left[1]}</td>
                    </tr>
                `;
        }
    }

    html += `</table>`;

    // 原木数据表格（显示金额列）
    const { validLogs, rowCount, rootCount, totalV: totalVGroup } = getValidLogsGroupStats(sourceLogs);
    const showPricePdf = !!appSettings.priceEnabled && !!appSettings.showPricePdf;
    const currencySymbol = getCurrencySymbol();
    const amountHeadHtml = showPricePdf
        ? `<th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.price_amount}</th>`
        : '';

    const dataTableHeaderHtml = `
            <table style="width:100%; border-collapse:collapse; font-size:11px; margin-bottom:20px; color:#222;">
                <thead>
                    <tr style="background:#e9e9e9;">
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.idx}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.code}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.grade}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.len}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.dia}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.vol} (m³)</th>
                        ${amountHeadHtml}
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.note}</th>
                    </tr>
                </thead>
                <tbody>`;

    const dataRowHtmls = [];
    [...validLogs].reverse().forEach((log, index) => {
        const len = parseFloat(cleanInput(log.length.toString())) || '';
        const dia = parseFloat(cleanInput(log.diameter.toString())) || '';
        const vol = log.volume ? formatVolumeForDisplay(log.volume) : '0';
        const bgColor = index % 2 === 0 ? '#f7f7f7' : '#ffffff';
        const showMarks = appSettings.showMarksInExport !== false;
        const showGroup = appSettings.showGroupInExport !== false;
        const groupBorder = (showGroup && log.groupId) ? 'border-left:4px solid #1976D2; font-weight:600;' : '';
        const amount = calcLogAmountBeforeTax(log);
        const amountCellHtml = showPricePdf
            ? `<td style="border:0.6px solid #999; padding:6px; text-align:center;">${amount ? formatMoney(amount) : ''}</td>`
            : '';
        const gradeDisp = (log.grade || '-') + (showMarks && log.markGrade ? '↑' : '');
        const lenDisp = len + (showMarks && log.markLen ? '↑' : '');
        const diaDisp = dia + (showMarks && log.markDia ? '↑' : '');
        dataRowHtmls.push(`
                <tr style="background:${bgColor};${groupBorder}">
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${index + 1}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${log.code || ''}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${gradeDisp}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${lenDisp}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${diaDisp}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${vol}</td>
                    ${amountCellHtml}
                    <td style="border:0.6px solid #999; padding:6px; text-align:left; font-size:10px;">${log.note || ''}</td>
                </tr>`);
    });

    // 根据头部内容动态调整首页行数：无公司/买方/测量人等时减少留白
    const infoRowCount = Math.ceil(infoItems.length / 2);
    const ROWS_FIRST_CHUNK = hasCompanyOrSeller ? 25 : (infoRowCount >= 3 ? 28 : 35);
    const ROWS_PER_CHUNK = 35;
    const FIRST_PAGE_BOTTOM_SPACER = '<div style="height:56px; min-height:56px;"></div>';

    const headerHtml = html;
    const mainHtmlChunks = [];
    let start = 0;
    let chunkIndex = 0;
    if (dataRowHtmls.length === 0) {
        mainHtmlChunks.push(headerHtml + dataTableHeaderHtml + '</tbody></table>' + FIRST_PAGE_BOTTOM_SPACER + '</div>');
    } else {
        while (start < dataRowHtmls.length) {
            const count = chunkIndex === 0 ? ROWS_FIRST_CHUNK : ROWS_PER_CHUNK;
            const end = Math.min(start + count, dataRowHtmls.length);
            const rowsHtml = dataRowHtmls.slice(start, end).join('');
            if (chunkIndex === 0) {
                mainHtmlChunks.push(headerHtml + dataTableHeaderHtml + rowsHtml + '</tbody></table>' + FIRST_PAGE_BOTTOM_SPACER + '</div>');
            } else {
                mainHtmlChunks.push('<div style="font-family: Arial, sans-serif; color: #000;">' + dataTableHeaderHtml + rowsHtml + '</tbody></table></div>');
            }
            start = end;
            chunkIndex++;
        }
    }

    // 统计汇总（分组全部计入）
    const totalV = totalVGroup;
    const gStats = {};
    const pdfT4 = parseFloat(appSettings.statThresholdL4) || 4;
    const pdfT25 = parseFloat(appSettings.statThresholdL25) || 2.5;
    const pdfT30 = parseFloat(appSettings.statThresholdD30) || 30;
    let l4 = 0, l25 = 0, d30 = 0;
    validLogs.forEach(l => {
        const g = l.grade || '?';
        gStats[g] = (gStats[g] || 0) + 1;
        const len_m = parseFloat(cleanInput(l.length.toString()));
        const dia_cm = parseFloat(cleanInput(l.diameter.toString()));
        if (len_m < pdfT4) l4++;
        if (len_m < pdfT25) l25++;
        if (dia_cm < pdfT30) d30++;
    });

    // 价格汇总 - 独立的6列表格（分组全部计入）
    if (showPricePdf) {
        const taxPercent = parseFloat(appSettings.taxPercent) || 0;
        let totalBeforeTax = 0;
        let avgUnitPrice = 0;
        validLogs.forEach(l => {
            const unit = getUnitPriceForLog(l);
            const v = parseFloat(l.volume) || 0;
            totalBeforeTax += unit * v;
        });
        if (totalV > 0) avgUnitPrice = totalBeforeTax / totalV;

        const taxAmount = taxPercent > 0 ? totalBeforeTax * (taxPercent / 100) : 0;
        const totalAfterTax = totalBeforeTax + taxAmount;

        const skupajLabel = currentLang === 'sl' ? 'Skupaj' : (currentLang === 'zh' ? '总计' : 'Total');
        const taxLabel = currentLang === 'sl' ? 'DDV' : (currentLang === 'zh' ? '增值税' : 'TAX');
        const totalLabel = currentLang === 'sl' ? 'Bruto znesek' : (currentLang === 'zh' ? '含税总额' : 'Total with Tax');

        const gradeLabel = currentLang === 'sl' ? 'Razred' : (currentLang === 'zh' ? '等级' : 'Grade');
        const kolicinaLabel = currentLang === 'sl' ? 'Količina' : (currentLang === 'zh' ? '数量' : 'Quantity');
        const znesekLabel = currentLang === 'sl' ? 'Znesek' : (currentLang === 'zh' ? '金额' : 'Amount');

        // 计算每个等级的详细信息（分组全部计入）
        const gradeDetails = {};
        validLogs.forEach(l => {
            const g = l.grade || '?';
            if (g === '?' || g === '') return;
            if (!gradeDetails[g]) {
                gradeDetails[g] = { count: 0, volume: 0, amount: 0 };
            }
            gradeDetails[g].count++;
            gradeDetails[g].volume += parseFloat(l.volume) || 0;
            const unitPrice = getUnitPriceForLog(l);
            const v = parseFloat(l.volume) || 0;
            gradeDetails[g].amount += unitPrice * v;
        });

        // 判断是否按等级定价
        const isByGrade = appSettings.priceMode === 'grade';
        const totalLabel2 = currentLang === 'sl' ? 'Skupaj' : (currentLang === 'zh' ? '总计' : 'Total');
        const unitPriceLabel = currentLang === 'sl' ? `Cena (${currencySymbol}/m³)` : (currentLang === 'zh' ? `单价 (${currencySymbol}/m³)` : `Unit Price (${currencySymbol}/m³)`);

        let priceTableHtml = `
                <div style="font-family: Arial, sans-serif; color: #000;">
                <table style="width:100%; border-collapse:collapse; font-size:11px; margin-top:20px; color:#222;">
                    <thead>
                        <tr style="background:#e9e9e9;">
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${gradeLabel}</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${kolicinaLabel}</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">m³</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${unitPriceLabel}</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${znesekLabel}</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

        // 如果按等级定价，先显示等级明细（按 gradeLabels 顺序）
        if (isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    priceTableHtml += `
                            <tr style="background:#f9f9f9;">
                                <td style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:bold;">${g}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${detail.count}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatVolumeForDisplay(detail.volume)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(unitPrice)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(detail.amount)} ${currencySymbol}</td>
                            </tr>
                        `;
                }
            });
        }

        // 总计行
        priceTableHtml += `
                        <tr style="background:#f0f0f0; font-weight:bold; border-top:2px solid #666;">
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${totalLabel2}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${rootCount}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatVolumeForDisplay(totalV)}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(avgUnitPrice)}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(totalBeforeTax)} ${currencySymbol}</td>
                        </tr>
            `;
        if (taxPercent > 0) {
            priceTableHtml += `
                        <tr style="background:#f9f9f9;">
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:right;">${taxLabel} ${taxPercent.toFixed(1)} %</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(taxAmount)} ${currencySymbol}</td>
                        </tr>
                        <tr style="background:#f0f0f0; font-weight:bold;">
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:right;">${totalLabel}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(totalAfterTax)} ${currencySymbol}</td>
                        </tr>
                `;
        }

        // 如果不是按等级定价，在最后显示等级明细
        if (!isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    priceTableHtml += `
                            <tr style="background:#f9f9f9;">
                                <td style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:bold;">${g}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${detail.count}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatVolumeForDisplay(detail.volume)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(unitPrice)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(detail.amount)} ${currencySymbol}</td>
                            </tr>
                        `;
                }
            });
        }

        priceTableHtml += `
                    </tbody>
                </table>
                </div>
            `;
        mainHtmlChunks.push(priceTableHtml);
    }

    // 统计汇总（单独渲染，避免被分页截断时直接整块移到下一页）
    const summaryTitle = currentLang === 'zh' ? '统计汇总' : (currentLang === 'en' ? 'SUMMARY' : 'POVZETEK');
    let summaryHtml = `
            <div style="font-family: Arial, sans-serif; color: #000; width:100%; box-sizing:border-box;">
            <div style="background:#f0f0f0; padding:15px; border:2px solid #000; margin-top:8px; width:100%; box-sizing:border-box; overflow-wrap:break-word; word-wrap:break-word;">
                <h3 style="margin:0 0 10px 0; font-size:14px;">${summaryTitle}</h3>
                <p style="margin:5px 0; font-size:12px;">
                    <strong>${t.total_count}:</strong> ${rowCount !== rootCount ? `${rowCount} (${rootCount})` : validLogs.length} &nbsp;&nbsp;
                    <strong>${t.total_vol}:</strong> <span style="font-weight:bold;">${formatVolumeForDisplay(totalV)} m³</span>
                </p>
        `;
    if (Object.keys(gStats).length > 0) {
        summaryHtml += `<p style="margin:5px 0; font-size:11px;"><strong>Grade:</strong> `;
        for (let g in gStats) {
            if (g !== '?' && g !== '') summaryHtml += `${g}: ${gStats[g]} &nbsp;&nbsp; `;
        }
        summaryHtml += `</p>`;
    }
    if (l4 > 0 || l25 > 0 || d30 > 0) {
        summaryHtml += `<p style="margin:5px 0; font-size:10px; color:#666;">`;
        if (l4 > 0) summaryHtml += `&lt; ${pdfT4}m: ${l4} &nbsp;&nbsp; `;
        if (l25 > 0) summaryHtml += `&lt; ${pdfT25}m: ${l25} &nbsp;&nbsp; `;
        if (d30 > 0) summaryHtml += `&lt; ${pdfT30}cm: ${d30}`;
        summaryHtml += `</p>`;
    }
    summaryHtml += `
                <p style="margin:10px 0 0 0; font-size:9px; color:#999; border-top:1px solid #ccc; padding-top:10px;">
                    LogMetric Pro | ${dateStr}
                </p>
            </div></div>
        `;

    try {
        if (!window.html2canvas) {
            await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
        }

        const doc = new jsPDF('p', 'mm', 'a4');
        const imgWidth = 210;
        const pageHeight = 297;
        const bottomMargin = 12;
        const contentHeight = pageHeight - bottomMargin;
        let currentY = 0;
        let position = 0;

        function addPageBottomMargin() {
            doc.setFillColor(255, 255, 255);
            doc.rect(0, contentHeight, imgWidth, bottomMargin, 'F');
        }

        for (const chunkHtml of mainHtmlChunks) {
            pdfContainer.innerHTML = chunkHtml;
            const canvas = await html2canvas(pdfContainer, {
                scale: 2,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff'
            });
            const chunkImgHeight = (canvas.height * imgWidth) / canvas.width;
            const imgData = canvas.toDataURL('image/png');
            if (currentY + chunkImgHeight > contentHeight && currentY > 0) {
                addPageBottomMargin();
                doc.addPage();
                currentY = 0;
            }
            doc.addImage(imgData, 'PNG', 0, currentY, imgWidth, chunkImgHeight);
            currentY += chunkImgHeight;
        }
        addPageBottomMargin();

        // 统计汇总单独渲染：若当前页剩余空间不足则整块移到下一页
        const summaryContainer = document.createElement('div');
        summaryContainer.style.cssText = 'position:absolute;left:-9999px;width:760px;background:white;padding:8px 40px 40px 40px;box-sizing:border-box;overflow:hidden;';
        document.body.appendChild(summaryContainer);
        summaryContainer.innerHTML = summaryHtml;
        const canvasSummary = await html2canvas(summaryContainer, {
            scale: 2,
            useCORS: true,
            logging: false,
            backgroundColor: '#ffffff'
        });
        document.body.removeChild(summaryContainer);

        const imgHeightSummary = (canvasSummary.height * imgWidth) / canvasSummary.width;
        const imgDataSummary = canvasSummary.toDataURL('image/png');
        const spaceLeftOnLastPage = contentHeight - currentY;

        if (spaceLeftOnLastPage < imgHeightSummary && currentY > 0) {
            doc.addPage();
            doc.addImage(imgDataSummary, 'PNG', 0, 0, imgWidth, imgHeightSummary);
        } else {
            doc.addImage(imgDataSummary, 'PNG', 0, currentY, imgWidth, imgHeightSummary);
        }
        addPageBottomMargin();

        const fileName = `OakLog_${sourceGlobal.container || 'Export'}_${timeStr}.pdf`;
        if (options.returnBlob) {
            const blob = doc.output('blob');
            return blob;
        }
        doc.save(fileName);

    } catch (error) {
        alert('PDF generation failed. Please try again.');
    } finally {
        document.body.removeChild(pdfContainer);
    }
}

// 辅助函数：动态加载脚本
function loadScript(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
}

async function exportData(options = {}) {
    const d = new Date();
    const timeStr = `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}_${d.getHours()}-${d.getMinutes()}`;
    const ds = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
    const t = I18N[currentLang];
    const sourceLogs = Array.isArray(options.logs) ? options.logs : logs;
    const sourceGlobal = (options.global && typeof options.global === 'object') ? options.global : globalInfo;
    const showPriceCsv = !!appSettings.priceEnabled && !!appSettings.showPriceCsv;
    const currencySymbol = getCurrencySymbol();
    const taxPercent = parseFloat(appSettings.taxPercent) || 0;

    const wb = XLSX.utils.book_new();
    const wsData = [];

    // 项目信息 - 横向专业布局
    const showCompany = appSettings.showCompanyInPdf !== false;
    const infoRow1 = [t.p_date, ds];
    if (showCompany && getCompanyName(sourceGlobal.company)) { infoRow1.push(t.my_company, getCompanyName(sourceGlobal.company)); }
    wsData.push(infoRow1);

    const infoRow2 = [];
    if (sourceGlobal.container) { infoRow2.push(t.container, sourceGlobal.container); }
    if (getCompanyName(sourceGlobal.seller)) { infoRow2.push(t.seller, getCompanyName(sourceGlobal.seller)); }
    if (infoRow2.length > 0) wsData.push(infoRow2);

    const infoRow3 = [];
    if (sourceGlobal.description) { infoRow3.push(t.description, sourceGlobal.description); }
    if (sourceGlobal.location) { infoRow3.push(t.location, sourceGlobal.location); }
    if (sourceGlobal.measurer) { infoRow3.push(t.measurer, sourceGlobal.measurer); }
    if (infoRow3.length > 0) wsData.push(infoRow3);

    if (sourceGlobal.note) wsData.push([t.note_global, sourceGlobal.note]);

    wsData.push([]);

    // 表头（包含金额列）
    const headerRow = [t.idx, t.code, t.grade, t.len, t.dia, t.vol];
    if (showPriceCsv) {
        headerRow.push(t.price_amount);
    }
    headerRow.push(t.note);
    wsData.push(headerRow);
    const dataStartRowIndex = wsData.length;

    // 数据行
    const { validLogs, rowCount: excelRowCount, rootCount: excelRootCount, totalV: excelTotalV } = getValidLogsGroupStats(sourceLogs);
    let totalBeforeTax = 0;
    const reversedLogs = [...validLogs].reverse();
    reversedLogs.forEach((l, idx) => {
        const len = parseFloat(cleanInput(l.length.toString())) || '';
        const dia = parseFloat(cleanInput(l.diameter.toString())) || '';
        const vol = l.volume ? parseFloat(formatVolumeForDisplay(l.volume)) : 0;
        const showMarks = appSettings.showMarksInExport !== false;
        const gradeDisp = (l.grade || '-') + (showMarks && l.markGrade ? '↑' : '');
        const lenDisp = len + (showMarks && l.markLen ? '↑' : '');
        const diaDisp = dia + (showMarks && l.markDia ? '↑' : '');

        const dataRow = [idx + 1, l.code || '', gradeDisp, lenDisp, diaDisp, vol];

        // 添加金额列（分组全部计入）
        if (showPriceCsv) {
            const amount = calcLogAmountBeforeTax(l);
            totalBeforeTax += amount;
            dataRow.push(amount ? parseFloat(amount.toFixed(2)) : '');
        }

        dataRow.push(l.note || '');
        wsData.push(dataRow);
    });

    // 统计汇总 - 独立的6列表格（分组全部计入）
    wsData.push([]);
    const totalV = excelTotalV;

    if (showPriceCsv && totalBeforeTax > 0) {
        let avgUnitPrice = totalV > 0 ? totalBeforeTax / totalV : 0;
        const taxAmount = taxPercent > 0 ? totalBeforeTax * (taxPercent / 100) : 0;
        const totalAfterTax = totalBeforeTax + taxAmount;

        const gradeLabel = currentLang === 'sl' ? 'Razred' : (currentLang === 'zh' ? '等级' : 'Grade');
        const kolicinaLabel = currentLang === 'sl' ? 'Količina' : (currentLang === 'zh' ? '数量' : 'Quantity');
        const znesekLabel = currentLang === 'sl' ? 'Znesek' : (currentLang === 'zh' ? '金额' : 'Amount');
        const taxLabel = currentLang === 'sl' ? 'DDV' : (currentLang === 'zh' ? '增值税' : 'TAX');
        const totalLabel = currentLang === 'sl' ? 'Bruto znesek' : (currentLang === 'zh' ? '含税总额' : 'Total with Tax');

        // 计算每个等级的详细信息（分组全部计入）
        const gradeDetails = {};
        validLogs.forEach(l => {
            const g = l.grade || '?';
            if (g === '?' || g === '') return;
            if (!gradeDetails[g]) {
                gradeDetails[g] = { count: 0, volume: 0, amount: 0 };
            }
            gradeDetails[g].count++;
            gradeDetails[g].volume += parseFloat(l.volume) || 0;
            const unitPrice = getUnitPriceForLog(l);
            const v = parseFloat(l.volume) || 0;
            gradeDetails[g].amount += unitPrice * v;
        });

        // 判断是否按等级定价
        const isByGrade = appSettings.priceMode === 'grade';
        const totalLabel2 = currentLang === 'sl' ? 'Skupaj' : (currentLang === 'zh' ? '总计' : 'Total');

        // 5列表格：等级 | 数量 | m³ | 单价(€/m³) | 金额
        // 表头行
        wsData.push([]);
        const unitPriceLabel = currentLang === 'sl' ? `Cena (${currencySymbol}/m³)` : (currentLang === 'zh' ? `单价 (${currencySymbol}/m³)` : `Unit Price (${currencySymbol}/m³)`);
        const headerRow = [gradeLabel, kolicinaLabel, 'm³', unitPriceLabel, znesekLabel];
        wsData.push(headerRow);

        // 如果按等级定价，先显示等级明细（按 gradeLabels 顺序）
        if (isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    wsData.push([
                        g,
                        detail.count,
                        parseFloat(formatVolumeForDisplay(detail.volume)),
                        parseFloat(formatMoney(unitPrice)),
                        `${formatMoney(detail.amount)} ${currencySymbol}`
                    ]);
                }
            });
        }

        // 总计行
        const row1 = [
            totalLabel2,
            excelRootCount,
            parseFloat(formatVolumeForDisplay(totalV)),
            parseFloat(formatMoney(avgUnitPrice)),
            `${formatMoney(totalBeforeTax)} ${currencySymbol}`
        ];
        wsData.push(row1);
        if (excelRowCount !== excelRootCount) {
            wsData.push([t.total_rows, excelRowCount, '', '', '']);
        }

        // 如果有税收
        if (taxPercent > 0) {
            // 税额行
            const row2 = ['', '', '', `${taxLabel} ${taxPercent.toFixed(1)} %`, `${formatMoney(taxAmount)} ${currencySymbol}`];
            wsData.push(row2);

            // 含税总额行
            const row3 = ['', '', '', totalLabel, `${formatMoney(totalAfterTax)} ${currencySymbol}`];
            wsData.push(row3);
        }

        // 如果不是按等级定价，在最后显示等级明细
        if (!isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    wsData.push([
                        g,
                        detail.count,
                        parseFloat(formatVolumeForDisplay(detail.volume)),
                        parseFloat(formatMoney(unitPrice)),
                        `${formatMoney(detail.amount)} ${currencySymbol}`
                    ]);
                }
            });
        }
    } else {
        // 不显示价格时的简单汇总
        wsData.push([t.total_count, excelRootCount, '', t.total_vol, formatVolumeForDisplay(totalV) + ' m³']);
        if (excelRowCount !== excelRootCount) {
            wsData.push([t.total_rows, excelRowCount, '', '', '']);
        }
    }

    // 其他统计
    const exT4 = parseFloat(appSettings.statThresholdL4) || 4;
    const exT25 = parseFloat(appSettings.statThresholdL25) || 2.5;
    const exT30 = parseFloat(appSettings.statThresholdD30) || 30;
    let l4 = 0, l25 = 0, d30 = 0;
    validLogs.forEach(l => {
        const len_m = parseFloat(cleanInput(l.length.toString()));
        const dia_cm = parseFloat(cleanInput(l.diameter.toString()));
        if (len_m < exT4) l4++; if (len_m < exT25) l25++; if (dia_cm < exT30) d30++;
    });
    wsData.push([]);
    if (l4 > 0) wsData.push(['< ' + exT4 + 'm', l4]);
    if (l25 > 0) wsData.push(['< ' + exT25 + 'm', l25]);
    if (d30 > 0) wsData.push(['< ' + exT30 + 'cm', d30]);

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // 分组行样式（xlsx-js-style）：浅蓝背景 + 加粗；码号列左侧蓝框、备注列右侧蓝框
    const showGroup = appSettings.showGroupInExport !== false;
    const numDataCols = 6 + (showPriceCsv ? 1 : 0) + 1;
    const codeCol = 1;
    const noteCol = numDataCols - 1;
    const blueBorder = { style: 'medium', color: { rgb: 'FF1976D2' } };
    const baseStyle = {
        fill: { patternType: 'solid', fgColor: { rgb: 'FFE3F2FD' } },
        font: { bold: true }
    };
    if (showGroup) {
        reversedLogs.forEach((l, idx) => {
            if (!l.groupId) return;
            const row = dataStartRowIndex + idx;
            for (let c = 0; c < numDataCols; c++) {
                const ref = XLSX.utils.encode_cell({ r: row, c });
                if (!ws[ref]) continue;
                const style = { ...baseStyle };
                if (c === codeCol) style.border = { left: blueBorder };
                else if (c === noteCol) style.border = { right: blueBorder };
                ws[ref].s = style;
            }
        });
    }

    // 设置列宽
    const colWidths = [
        { wch: 6 }, { wch: 16 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 12 }
    ];
    if (showPriceCsv) {
        colWidths.push({ wch: 16 }, { wch: 16 });
    }
    colWidths.push({ wch: 25 });
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, 'Oak Logs');

    const fileName = `OakLog_${sourceGlobal.container || 'Export'}_${timeStr}.xlsx`;
    if (options.returnBlob) {
        const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
    XLSX.writeFile(wb, fileName);
}

function saveToHistory(t) {
    if (t === 'company') {
        const c = normalizeCompany(globalInfo.company || defaultCompany());
        const name = getCompanyName(c).trim();
        if (!name) return;
        const idx = (histories.company || []).findIndex(x => getCompanyName(x) === name);
        if (idx >= 0) histories.company[idx] = c; else histories.company.push(c);
        localStorage.setItem(HIST_KEY, JSON.stringify(histories));
        alert(currentLang === 'zh' ? '已保存' : (currentLang === 'en' ? 'Saved' : 'Shranjeno'));
        return;
    }
    if (t === 'seller') {
        const s = normalizeSeller(globalInfo.seller || defaultSeller());
        const name = getCompanyName(s).trim();
        if (!name) return;
        const idx = (histories.seller || []).findIndex(x => getCompanyName(x) === name);
        if (idx >= 0) histories.seller[idx] = s; else histories.seller.push(s);
        localStorage.setItem(HIST_KEY, JSON.stringify(histories));
        alert(currentLang === 'zh' ? '已保存' : (currentLang === 'en' ? 'Saved' : 'Shranjeno'));
        return;
    }
    const v = (document.getElementById('g_' + t)?.value || '').trim();
    if (v && !histories[t].includes(v)) { histories[t].push(v); localStorage.setItem(HIST_KEY, JSON.stringify(histories)); alert(currentLang === 'zh' ? '已保存' : (currentLang === 'en' ? 'Saved' : 'Shranjeno')); }
}
function saveToHistoryForHistoryViewer(t) {
    if (t === 'seller') {
        const s = normalizeSeller(historyViewerState.global?.seller || defaultSeller());
        const name = getCompanyName(s).trim();
        if (!name) return;
        const idx = (histories.seller || []).findIndex(x => getCompanyName(x) === name);
        if (idx >= 0) histories.seller[idx] = s; else histories.seller.push(s);
        localStorage.setItem(HIST_KEY, JSON.stringify(histories));
        alert(currentLang === 'zh' ? '已保存' : (currentLang === 'en' ? 'Saved' : 'Shranjeno'));
    }
}
let historyPopContext = 'main';
let historyPopType = '';
function showHistory(t) {
    historyPopContext = 'main';
    historyPopType = t;
    const pop = document.getElementById('historyPop');
    pop.classList.remove('above-history-viewer');
    const l = document.getElementById('historyList'); l.innerHTML = '';
    const items = histories[t] || [];
    if (items.length === 0) l.innerHTML = '<div style="padding:10px;color:#666;">No Record</div>';
    items.forEach((i, x) => {
        const d = document.createElement('div');
        d.style.cssText = 'padding:12px;border-bottom:1px solid #333;display:flex;justify-content:space-between;align-items:center;';
        const label = (t === 'seller' || t === 'company') ? getCompanyName(i) : (typeof i === 'object' ? (i.name || '') : String(i));
        d.innerHTML = `<div onclick="selectHistoryByIndex(${x})" style="flex:1;color:#fff;cursor:pointer;">${escapeHtml(label)}</div><div onclick="event.stopPropagation();delHistory('${t}',${x})" style="color:red;padding:5px 10px;font-size:18px;cursor:pointer;">×</div>`;
        l.appendChild(d);
    });
    pop.style.display = 'flex';
}
function showHistoryForHistoryViewer(t) {
    historyPopContext = 'historyViewer';
    historyPopType = t;
    const pop = document.getElementById('historyPop');
    pop.classList.add('above-history-viewer');
    const l = document.getElementById('historyList'); l.innerHTML = '';
    const items = histories[t] || [];
    if (items.length === 0) l.innerHTML = '<div style="padding:10px;color:#666;">No Record</div>';
    items.forEach((i, x) => {
        const d = document.createElement('div');
        d.style.cssText = 'padding:12px;border-bottom:1px solid #333;display:flex;justify-content:space-between;align-items:center;';
        const label = (t === 'seller' || t === 'company') ? getCompanyName(i) : (typeof i === 'object' ? (i.name || '') : String(i));
        d.innerHTML = `<div onclick="selectHistoryByIndex(${x})" style="flex:1;color:#fff;cursor:pointer;">${escapeHtml(label)}</div><div onclick="event.stopPropagation();delHistory('${t}',${x})" style="color:red;padding:5px 10px;font-size:18px;cursor:pointer;">×</div>`;
        l.appendChild(d);
    });
    pop.style.display = 'flex';
}
function escapeHtml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function selectHistoryByIndex(idx) {
    const t = historyPopType;
    const items = histories[t] || [];
    const v = items[idx];
    if (v == null) return;
    selectHistory(t, v);
}
function selectHistory(t, v) {
    if (historyPopContext === 'historyViewer' && t === 'seller') {
        if (!historyViewerState.global) historyViewerState.global = {};
        const s = typeof v === 'object' ? normalizeSeller(v) : Object.assign(defaultSeller(), { name: v });
        historyViewerState.global.seller = s;
        const hvSeller = document.getElementById('hv_g_seller');
        if (hvSeller) hvSeller.value = getCompanyName(s);
        document.getElementById('historyPop').style.display = 'none';
        document.getElementById('historyPop').classList.remove('above-history-viewer');
        historyPopContext = 'main';
        return;
    }
    if (t === 'company') {
        const c = typeof v === 'object' ? normalizeCompany(v) : Object.assign(defaultCompany(), { name: v });
        globalInfo.company = c;
        document.getElementById('g_company').value = getCompanyName(c);
    } else if (t === 'seller') {
        const s = typeof v === 'object' ? normalizeSeller(v) : Object.assign(defaultSeller(), { name: v });
        globalInfo.seller = s;
        syncInfoModalCompanyDisplay();
    } else {
        globalInfo[t] = v;
        const el = document.getElementById('g_' + t);
        if (el) el.value = v;
    }
    document.getElementById('historyPop').style.display = 'none';
    save();
}
function delHistory(t, i) {
    const msg = currentLang === 'zh' ? '确定删除此条历史记录？' : (currentLang === 'en' ? 'Delete this history item?' : 'Izbriši ta zapis?');
    if (!confirm(msg)) return;
    histories[t].splice(i, 1); localStorage.setItem(HIST_KEY, JSON.stringify(histories)); if (historyPopContext === 'historyViewer') showHistoryForHistoryViewer(t); else showHistory(t);
}
