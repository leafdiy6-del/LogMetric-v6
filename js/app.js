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
    lsSet(HIST_KEY, JSON.stringify(histories));
}
function addToHistorySeller(s) {
    const name = getCompanyName(s).trim();
    if (!name) return;
    const idx = (histories.seller || []).findIndex(x => getCompanyName(x) === name);
    if (idx >= 0) histories.seller[idx] = normalizeSeller(s); else histories.seller.push(normalizeSeller(s));
    lsSet(HIST_KEY, JSON.stringify(histories));
}
function updateSellerInHistory(originalSeller, newSeller) {
    const origName = getCompanyName(originalSeller).trim();
    if (!origName) return;
    const arr = histories.seller || [];
    const idx = arr.findIndex(x => getCompanyName(x) === origName);
    if (idx >= 0) arr[idx] = normalizeSeller(newSeller);
    else arr.push(normalizeSeller(newSeller));
    histories.seller = arr;
    lsSet(HIST_KEY, JSON.stringify(histories));
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
    lsSet('lmp_theme', newTheme);
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
let hvGroupMode = false;
let hvGroupSelectIds = [];
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
    volumeWarningThreshold: 21,
    volumeDecimals: 3
};

/* ----------------------------------------------------------------------------
   [2.3] UI 状态与模式 (UI State & Modes)
   ---------------------------------------------------------------------------- */
let activeFilter = null;
let currentLang = 'zh';
let isQuickMode = false;
let nextRoundUp = true;
let isContainerNumberMode = false; // 集装箱号模式(false) / 号模式(true)

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

/* ----------------------------------------------------------------------------
   [2.6] 行操作菜单状态 (Row Action Sheet State)
   ---------------------------------------------------------------------------- */
let currentActionLogId = null;
let currentActionContext = 'main'; // 'main' | 'hv'（区分主列表和内部记录编辑模式）

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

function toggleContainerMode() {
    isContainerNumberMode = !isContainerNumberMode;
    const btn = document.getElementById('containerToggleBtn');
    const label = document.getElementById('containerToggleLabel');
    const containerLabel = document.querySelector('label[data-i18n="container"]');
    
    if (btn && label) {
        if (isContainerNumberMode) {
            btn.classList.add('active');
            label.textContent = '号';
            // 修改标签为"号"
            if (containerLabel) {
                const helpBtn = containerLabel.querySelector('.btn-help');
                containerLabel.innerHTML = `号 ${helpBtn ? helpBtn.outerHTML : ''}`;
            }
        } else {
            btn.classList.remove('active');
            label.textContent = '号';
            // 恢复标签为"集装箱号"
            if (containerLabel) {
                const helpBtn = containerLabel.querySelector('.btn-help');
                containerLabel.innerHTML = `集装箱号 ${helpBtn ? helpBtn.outerHTML : ''}`;
            }
        }
    }
    haptic();
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
    validateSlovakTableMonotonicity();
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
    appSettings.useVirtualKeyboard = false; // will be set by updateOrientation()

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
    if (formulaSelect) {
        let f = appSettings.formula || 'huber';
        if (f === 'csn4800079') f = 'csn480009';  // 迁移：0007/9 → 0009
        formulaSelect.value = f;
    }
    const volumeDecimalsSelect = document.getElementById('volumeDecimalsSelect');
    if (volumeDecimalsSelect) volumeDecimalsSelect.value = String(appSettings.volumeDecimals === 2 ? 2 : 3);
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
    window.addEventListener('orientationchange', updateOrientation);
    window.addEventListener('resize', updateOrientation);
    updateOrientation();
    setTimeout(updateOrientation, 300);
    updateKeySoundRowVisibility();
    if (logs.length === 0) addNewLog();
    updateGroupBtnUI();
    renderAll();

    // 离线支持：注册 Service Worker
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./sw.js').catch(() => { });
    }
};

function changeLanguage(lang) { currentLang = lang; lsSet(LANG_KEY, lang); applyLanguage(); renderAll(); updateQuickBtnUI(); updateProKeyboardBtnUI(); updateThemeToggleLabel(); }
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
// localStorage 安全写入：存储配额满时弹提示，不崩溃
function lsSet(key, value) {
    try {
        localStorage.setItem(key, value);
    } catch (e) {
        const msg = currentLang === 'zh'
            ? '存储空间不足，请先导出备份再清理数据'
            : (currentLang === 'en' ? 'Storage full. Export a backup first.' : 'Pomnilnik je poln. Najprej izvozite varnostno kopijo.');
        alert(msg);
    }
}

function save() { lsSet(STORAGE_KEY, JSON.stringify({ logs, global: globalInfo })); updateStats(); }

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
    if (cb) { appSettings.showCompanyInPdf = cb.checked; lsSet(SETTINGS_KEY, JSON.stringify(appSettings)); }
}

// ========== 保存菜单和快照系统 ==========
// 保存菜单 + 快照 → js/modules/snapshot-manager.js
// 历史记录查看器 → js/modules/history-viewer.js

// 项目导出/导入 → js/modules/snapshot-manager.js

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
    const volumeDecimalsSelect = document.getElementById('volumeDecimalsSelect');
    if (volumeDecimalsSelect) appSettings.volumeDecimals = parseInt(volumeDecimalsSelect.value, 10) || 3;
    appSettings.useVirtualKeyboard = isProKeyboardEffective();
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
    lsSet(SETTINGS_KEY, JSON.stringify(appSettings));
    if (typeof updateStats === 'function') updateStats();
}
function toggleGradeDisplay() {
    appSettings.showGrade = document.getElementById('checkShowGrade').checked;
    saveSettings();
    renderAll();
    updateProSideState();
    if (!appSettings.showGrade && proState.keypadMode === 'grade') setProKeypadMode('num');
}
// Pro Keyboard 辅助函数 → js/modules/pro-keyboard.js
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
    const decimals = (appSettings.volumeDecimals === 2 || appSettings.volumeDecimals === 3) ? appSettings.volumeDecimals : 3;
    return parseFloat(v.toFixed(decimals)).toString();
}

function toggleQuickMode() { isQuickMode = !isQuickMode; lsSet(QUICK_KEY, isQuickMode); updateQuickBtnUI(); }
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
    lsSet(SETTINGS_KEY, JSON.stringify(appSettings));
}
// Pro Keyboard 主模块 → js/modules/pro-keyboard.js

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

    // 情况1：纯数字自动补小数点并跳转
    if (appSettings.quickModeAutoDecimal && /^\d+$/.test(val)) {
        let newVal = null;
        if (!val.startsWith('1')) {
            // 非1开头：2位数（20-99）→ 自动加小数点，如 35→3.5
            let num = parseInt(val);
            if (val.length === 2 && num >= 20 && num <= 99) newVal = (num / 10).toString();
        } else {
            // 1开头：3位数 → 第2位后插入小数点，如 123→12.3
            if (val.length === 3) newVal = val[0] + val[1] + '.' + val[2];
            // 1-2位时继续等待，不触发
        }

        if (newVal) {
            input.value = newVal;
            const logId = parseInt(input.getAttribute('data-id'));
            if (!isNaN(logId)) updateItem(logId, 'length', newVal);
            if (appSettings.quickModeAutoJump && input.closest('.log-card')) jumpLengthToDia();
        }
        return;
    }

    // 情况2：已带小数点（如 3.5、12.5）直接跳转到直径
    if (appSettings.quickModeAutoJump && input.closest('.log-card') && /^\d+\.\d+$/.test(val)) {
        jumpLengthToDia();
    }
}

// 历史记录列表编辑模式的长度快速输入（负责全部更新逻辑，仅此一次调用 updateHistoryViewerItem）
function handleHvQuickLength(input, logId) {
    let val = input.value;
    if (!val) return;

    let finalVal = val;

    if (appSettings.quickModeAutoDecimal && /^\d+$/.test(val)) {
        let newVal = null;
        if (!val.startsWith('1')) {
            let num = parseInt(val);
            if (val.length === 2 && num >= 20 && num <= 99) newVal = (num / 10).toString();
        } else {
            if (val.length === 3) newVal = val[0] + val[1] + '.' + val[2];
        }
        if (newVal) {
            input.value = newVal;
            finalVal = newVal;
        }
    }

    updateHistoryViewerItem(logId, 'length', finalVal);
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
            lsSet(MIX_STATE_KEY, nextRoundUp);
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
        const firstRow = container.querySelector('.log-row');
        if (firstRow) container.insertBefore(newRow, firstRow);
        else container.appendChild(newRow);
        updateStats();
        activeFilter = null;
        applyFilter();
        highlightNewestRow();
        container.scrollTop = 0;
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
        for (let i = 1; i <= logs.length - 1; i++) {
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
        const list = document.getElementById('logList');
        if (list) list.scrollTop = 0;
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
                <input type="text" inputmode="decimal" value="${log.diameter}" data-id="${log.id}" data-field="diameter" oninput="updateItem(${log.id},'diameter',this.value); handleQuickDiameterSingle(this); toggleDiaDangerClass(this)" onblur="autoFixInput(this)">
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
    const warnClass = (len_m > 15 && !log.length.toString().trim().startsWith('1')) ? 'danger-text' : '';
    const diaStr2 = log.diameter.toString().trim();
    const diaDangerClass = ((d_cm >= 200 && d_cm < 1000) || (/^\d+$/.test(diaStr2) && (diaStr2.length === 1 || (diaStr2.length === 3 && !diaStr2.startsWith('1'))))) ? 'dia-danger-text' : '';

    const displayLen = log.length ? parseFloat(log.length) : '';
    const displayDia = log.diameter ? parseFloat(log.diameter) : '';

    const markGrade = !!log.markGrade; const markLen = !!log.markLen; const markDia = !!log.markDia;
    div.innerHTML = `
            <div class="row-index" ondblclick="openRowActionSheet(${log.id})" title="${currentLang === 'zh' ? '双击操作菜单' : (currentLang === 'en' ? 'Double-tap for row actions' : 'Dvojni klik za dejanja')}" style="cursor:pointer;">${idx}</div>
            <input type="text" data-field="code" value="${log.code}" oninput="updateItem(${log.id},'code',this.value)">
            <div class="mark-cell ${markGrade ? 'marked' : ''}" data-id="${log.id}" data-field="grade" oncontextmenu="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event); return false" ondblclick="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event)" title="${currentLang === 'zh' ? '双击可修改等级按钮文字' : (currentLang === 'en' ? 'Double-click to change grade button label' : 'Dvojni klik za spremembo')}">
                <span class="mark-symbol">↑</span>
                <select onchange="updateItem(${log.id},'grade',this.value)">${gradeOptions}</select>
                <div class="mark-overlay" onclick="handleMarkCellClick(event, ${log.id}, 'grade')"></div>
            </div>
            <div class="mark-cell ${markLen ? 'marked' : ''}" data-id="${log.id}" data-field="length">
                <span class="mark-symbol">↑</span>
                <input type="text" data-id="${log.id}" data-field="length" id="inp-len-${log.id}" class="${warnClass}" inputmode="decimal" value="${displayLen}" oninput="updateItem(${log.id},'length',this.value); handleQuickLength(this)" onblur="autoFixInput(this)">
                <div class="mark-overlay" onclick="handleMarkCellClick(event, ${log.id}, 'length')"></div>
            </div>
            <div class="mark-cell ${markDia ? 'marked' : ''}" data-id="${log.id}" data-field="diameter">
                <span class="mark-symbol">↑</span>
                <input type="text" data-id="${log.id}" data-field="diameter" class="${diaDangerClass}" inputmode="decimal" value="${displayDia}" oninput="updateItem(${log.id},'diameter',this.value);toggleDiaDangerClass(this)" onblur="autoFixInput(this)">
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
    const val = (input.value || '').toString().trim();
    const v = parseFloat(cleanInput(val));
    const isPureDanger = !isNaN(v) && v >= 200 && v < 1000;
    const isMissingDecimal = /^\d+$/.test(val) && (val.length === 1 || (val.length === 3 && !val.startsWith('1')));
    input.classList.toggle('dia-danger-text', isPureDanger || isMissingDecimal);
}
// 分组 + 标记 → js/modules/grouping-marking.js
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
// 体积计算函数 → js/modules/volume-calc.js

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
    
    // 保存当前集装箱号/号值用于后续自动递增
    const currentContainerValue = document.getElementById('g_container')?.value || '';
    const nextContainerValue = isContainerNumberMode && currentContainerValue ? (parseInt(currentContainerValue) + 1).toString() : '';
    
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
        lsSet(SNAPSHOTS_KEY, JSON.stringify(snapshots));
    }
    logs = [];
    globalInfo.container = '';
    globalInfo.note = '';
    globalInfo.description = '';
    
    // 如果是号模式，自动设置下一个号
    const gContainer = document.getElementById('g_container');
    if (gContainer) {
        if (isContainerNumberMode && nextContainerValue) {
            gContainer.value = nextContainerValue;
            globalInfo.container = nextContainerValue;
        } else {
            gContainer.value = '';
        }
    }
    
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

// 统计 + 过滤 → js/modules/stats.js

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

// PDF 生成函数 → js/modules/pdf-generator.js

// Excel 导出函数 → js/modules/excel-exporter.js

// 历史输入自动补全 → js/modules/history-autocomplete.js
// 行操作菜单 + 重编号 → js/modules/row-actions.js
