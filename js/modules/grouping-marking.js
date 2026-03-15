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
    const msg = currentLang === 'zh' ? '取消此分组？' : (currentLang === 'en' ? 'Remove this group?' : 'Odstraniti skupino?');
    if (!confirm(msg)) return;
    const gid = log.groupId;
    logs.forEach(l => { if (l.groupId === gid) delete l.groupId; });
    save();
    renderAll();
}

/* ---- 内部记录编辑模式 分组功能 (HV Group Mode) ---- */
function toggleHvGroupMode() {
    hvGroupMode = !hvGroupMode;
    hvGroupSelectIds = [];
    document.body.classList.toggle('hv-group-mode', hvGroupMode);
    const btn = document.getElementById('btnHvGroupMode');
    if (btn) btn.classList.toggle('active', hvGroupMode);
    document.querySelectorAll('#historyViewerList .log-row.group-selected').forEach(r => r.classList.remove('group-selected'));
}

function handleHvGroupRowClick(e, id) {
    if (!hvGroupMode) return;
    e.preventDefault(); e.stopPropagation();
    const row = document.getElementById('hv-row-' + id);
    if (!row || row.classList.contains('grouped')) return;
    if (hvGroupSelectIds.includes(id)) {
        hvGroupSelectIds = hvGroupSelectIds.filter(x => x !== id);
        row.classList.remove('group-selected');
        return;
    }
    hvGroupSelectIds.push(id);
    row.classList.add('group-selected');
    if (hvGroupSelectIds.length >= 2) {
        const id1 = hvGroupSelectIds[0], id2 = hvGroupSelectIds[1];
        const hvLogs = historyViewerState.logs || [];
        const idx1 = hvLogs.findIndex(l => l.id === id1);
        const idx2 = hvLogs.findIndex(l => l.id === id2);
        const adjacent = idx1 >= 0 && idx2 >= 0 && Math.abs(idx1 - idx2) === 1;
        const prevOrder = !adjacent && idx1 >= 0 && idx2 >= 0 ? JSON.parse(JSON.stringify(hvLogs)) : null;
        if (!adjacent && idx1 >= 0 && idx2 >= 0) {
            const log2 = hvLogs[idx2];
            hvLogs.splice(idx2, 1);
            const newIdx1 = hvLogs.findIndex(l => l.id === id1);
            hvLogs.splice(newIdx1 + 1, 0, log2);
            renderHistoryViewer();
            document.getElementById('hv-row-' + id1)?.classList.add('group-selected');
            document.getElementById('hv-row-' + id2)?.classList.add('group-selected');
        }
        const msg = currentLang === 'zh' ? '是否将这两根合并为一组？' : (currentLang === 'en' ? 'Merge these two rows into one group?' : 'Združiti ti dve vrstici v eno skupino?');
        if (confirm(msg)) {
            const gid = 'group_' + Date.now();
            hvLogs.forEach(l => { if (l.id === id1 || l.id === id2) l.groupId = gid; });
            toggleHvGroupMode();
            renderHistoryViewer();
        } else {
            if (prevOrder) { historyViewerState.logs = prevOrder; }
            hvGroupSelectIds = [];
            document.querySelectorAll('#historyViewerList .log-row.group-selected').forEach(r => r.classList.remove('group-selected'));
            renderHistoryViewer();
        }
    }
}

function hvUngroupLog(id) {
    const hvLogs = historyViewerState.logs || [];
    const log = hvLogs.find(l => l.id === id); if (!log || !log.groupId) return;
    const msg = currentLang === 'zh' ? '取消此分组？' : (currentLang === 'en' ? 'Remove this group?' : 'Odstraniti skupino?');
    if (!confirm(msg)) return;
    const gid = log.groupId;
    hvLogs.forEach(l => { if (l.groupId === gid) delete l.groupId; });
    renderHistoryViewer();
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
            if (appSettings.proKeyboard) renderProSideGradeButtons();
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
    }
    else if (!useVirtual) { renderAll(); }
    if (appSettings.proKeyboard) renderProSideGradeButtons();
}

/* ----------------------------------------------------------------------------
   数据更新工具函数 (Data Update Utilities)
   ---------------------------------------------------------------------------- */

