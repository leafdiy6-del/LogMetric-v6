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

