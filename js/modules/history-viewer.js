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
    const lenStr = (log.length || '').toString().trim();
    const diaStr = (log.diameter || '').toString().trim();
    const lenMissingDecimal = /^\d+$/.test(lenStr) && lenStr.length >= 2 && lenStr.startsWith('2');
    const diaMissingDecimal = /^\d+$/.test(diaStr) && (diaStr.length === 1 || (diaStr.length === 3 && !diaStr.startsWith('1')));
    const warnClass = ((len_m > 15 && !lenStr.startsWith('1')) || lenMissingDecimal) ? 'danger-text' : '';
    const diaDangerClass = ((d_cm >= 200 && d_cm < 1000) || diaMissingDecimal) ? 'dia-danger-text' : '';
    const displayLen = log.length ? (parseFloat(log.length) || log.length) : '';
    const displayDia = log.diameter ? (parseFloat(log.diameter) || log.diameter) : '';
    const vol = (log.volume != null && !isNaN(log.volume)) ? formatVolumeForDisplay(log.volume) : '0';

    const div = document.createElement('div');
    div.className = 'log-row' + (log.groupId ? ' grouped' : '');
    div.id = 'hv-row-' + log.id;
    div.setAttribute('data-grade', log.grade || '?');
    div.setAttribute('data-len', len_m || 0);
    div.setAttribute('data-dia', d_cm || 0);
    div.innerHTML = `
            <div class="row-index" ondblclick="openRowActionSheet(${log.id},'hv')" title="${currentLang === 'zh' ? '双击操作菜单' : (currentLang === 'en' ? 'Double-tap for row actions' : 'Dvojni klik za dejanja')}" style="cursor:pointer;">${idx}</div>
            <input type="text" data-field="code" value="${escapeAttr(log.code || '')}" oninput="updateHistoryViewerItem(${log.id},'code',this.value)">
            <div class="hv-grade-cell" oncontextmenu="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event); return false" ondblclick="handleGradeLabelEditByGrade('${(log.grade || '').replace(/'/g, "\\'")}', event)" title="${currentLang === 'zh' ? '双击可修改等级按钮文字' : (currentLang === 'en' ? 'Double-click to change grade button label' : 'Dvojni klik za spremembo')}"><select onchange="updateHistoryViewerItem(${log.id},'grade',this.value)">${gradeOptions}</select></div>
            <input type="text" data-field="length" class="${warnClass}" inputmode="decimal" value="${escapeAttr(displayLen)}" oninput="handleHvQuickLength(this,${log.id})" onblur="autoFixHistoryViewerInput(this)">
            <input type="text" data-field="diameter" class="${diaDangerClass}" inputmode="decimal" value="${escapeAttr(displayDia)}" oninput="updateHistoryViewerItem(${log.id},'diameter',this.value);toggleDiaDangerClass(this)" onblur="autoFixHistoryViewerInput(this)">
            <div class="col-vol" id="hv-v-row-${log.id}">${vol}</div>
            <input type="text" data-field="note" style="font-size:12px;color:#aaa;text-align:left;" value="${escapeAttr(log.note || '')}" oninput="updateHistoryViewerItem(${log.id},'note',this.value)">
            <div class="group-actions-cell">${log.groupId ? `<button type="button" class="btn-ungroup" onclick="hvUngroupLog(${log.id});event.stopPropagation()" title="${currentLang === 'zh' ? '解除分组' : (currentLang === 'en' ? 'Ungroup' : 'Razdruži')}">⎋</button>` : ''}</div>
            <button type="button" class="btn-del-mini" onclick="delHistoryViewerRow(${log.id})">×</button>
            <div class="group-select-overlay" onclick="handleHvGroupRowClick(event,${log.id})"></div>
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
                const lenStr2 = (item.length || '').toString().trim();
                const len_m2 = parseFloat(cleanInput(lenStr2));
                const lenWarn = (len_m2 > 15 && !lenStr2.startsWith('1')) || (/^\d+$/.test(lenStr2) && lenStr2.length >= 2 && lenStr2.startsWith('2'));
                lenInput.classList.toggle('danger-text', lenWarn);
            }
        }
        updateHistoryViewerStats();
    }
}
function updateHistoryViewerStats() {
    const list = historyViewerState.logs || [];
    const validLogs = list.filter(l => parseFloat(l.volume) > 0);
    const rowCount = validLogs.length;
    const totalV = validLogs.reduce((s, l) => s + (parseFloat(l.volume) || 0), 0);
    const statsEl = document.getElementById('historyViewerStats');
    if (statsEl) {
        const t = I18N[currentLang];
        const groupLabel = t.row_action_insert_above ? (currentLang === 'zh' ? '分组' : (currentLang === 'en' ? 'Group' : 'Skupina')) : '分组';
        statsEl.innerHTML = `
            <span class="stats-group-cell"><button type="button" class="btn-group-mode${hvGroupMode ? ' active' : ''}" id="btnHvGroupMode" onclick="toggleHvGroupMode()">${groupLabel}</button></span>
            <span class="stats-total-cell">${t.total_count}:<strong id="hvTotalCount">${rowCount}</strong></span>
            <span class="stats-vol-cell">${t.total_vol}:<strong id="hvTotalVol">${formatVolumeForDisplay(totalV)}</strong> m³</span>`;
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
