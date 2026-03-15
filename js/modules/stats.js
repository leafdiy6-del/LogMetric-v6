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
