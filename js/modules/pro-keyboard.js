// Pro Keyboard 辅助函数
function isLandscape() {
    return window.matchMedia('(orientation: landscape)').matches;
}
function isProKeyboardEffective() {
    return !!appSettings.proKeyboard && isLandscape();
}
function updateOrientation() {
    document.body.classList.toggle('landscape', isLandscape());
    appSettings.useVirtualKeyboard = isProKeyboardEffective();
    applyProKeyboardUI();
    renderAll();
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

// Pro Keyboard 主模块
function updateProKeyboardBtnUI() {
    const btn = document.getElementById('proKeyboardBtn');
    if (!btn) return;
    btn.classList.toggle('active', !!appSettings.proKeyboard);
    btn.innerText = appSettings.proKeyboard ? 'ON' : 'OFF';
}

function getActiveLog() { return logs.length > 0 ? logs[0] : null; }
function applyProKeyboardUI() {
    const effective = isProKeyboardEffective();
    document.body.classList.toggle('pro-kb-enabled', effective);
    appSettings.useVirtualKeyboard = effective;
    if (effective) {
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
    renderProSideGradeButtons();
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
    renderProSideGradeButtons();
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
}
function handleProSide(type) {
    if (appSettings.useVirtualKeyboard) flushProKeyPending();
    if (type === 'hide') { collapseProKeyboard(); return; }
    if (type === 'minus') { handleProMinus(); return; }
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
            setProOkFocused(true);
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
    if (appSettings.showGrade) { setProOkFocused(true); return; }
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
    if (rows.length > 0) rows[0].classList.add('latest');
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
function renderProSideGradeButtons() {
    const box = document.getElementById('kbSideGrades');
    if (!box) return;
    if (!appSettings.showGrade) {
        box.innerHTML = '';
        box.style.display = 'none';
        return;
    }
    box.style.display = '';
    const log = getActiveLog();
    const activeGrade = log ? log.grade : '';
    const gradeLabels = getGradeLabels();
    const labelsToShow = gradeLabels.filter(g => g !== 'F');
    const esc = (s) => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const titleAttr = currentLang === 'zh' ? '双击可修改按钮文字' : (currentLang === 'en' ? 'Double-click to change button label' : 'Dvojni klik za spremembo');
    box.innerHTML = labelsToShow.map((g) => {
        const realIdx = gradeLabels.indexOf(g);
        return `<button class="kb-side-btn ${activeGrade === g ? 'active' : ''}" onclick="selectProGrade('${esc(g)}')" oncontextmenu="handleGradeLabelEdit(${realIdx}, event); return false" ondblclick="handleGradeLabelEdit(${realIdx}, event)" title="${titleAttr}">${g}</button>`;
    }).join('');
}
function renderProGradePanel() {
    const box = document.getElementById('proKbGrade');
    if (!box) return;
    box.innerHTML = '';
}
function selectProGrade(grade) {
    const log = getActiveLog();
    if (!log) return;
    if (appSettings.keySound && typeof playKeySound === 'function') playKeySound('grade');
    if (appSettings.useVirtualKeyboard) flushProKeyPending();
    setGrade(log.id, grade);
    if (isQuickMode && appSettings.quickModeGradeAutoSave) {
        proState.gradeReady = false;
        addNewLog();
        return;
    }
    proState.gradeReady = true;
    renderProSideGradeButtons();
    setProOkFocused(true);
}
