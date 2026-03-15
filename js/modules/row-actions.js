/* ============================================================================
   行操作底部菜单 (Row Action Bottom Sheet)
   新增功能 (2026-03):
     - 双击序号 → 底部操作菜单
     - 在上方插入新行 / 在下方插入新行
     - 自动修改码号（扫描下方连续段，支持只改连续段或全部修改）
   ============================================================================ */

function openRowActionSheet(logId, context) {
    currentActionLogId = logId;
    currentActionContext = context || 'main';
    const t = I18N[currentLang];
    const targetLogs = currentActionContext === 'hv' ? (historyViewerState.logs || []) : logs;
    const idx = targetLogs.findIndex(l => l.id === logId);
    const displayNum = idx >= 0 ? (targetLogs.length - idx) : '';
    const titleEl = document.getElementById('rowActionTitle');
    const labelAbove = document.getElementById('rowActionLabelAbove');
    const labelBelow = document.getElementById('rowActionLabelBelow');
    const labelRenumber = document.getElementById('rowActionLabelRenumber');
    const labelCancel = document.getElementById('rowActionLabelCancel');
    if (titleEl) titleEl.textContent = displayNum ? `No. ${displayNum}` : '';
    if (labelAbove) labelAbove.textContent = t.row_action_insert_above || '在上方插入新行';
    if (labelBelow) labelBelow.textContent = t.row_action_insert_below || '在下方插入新行';
    if (labelRenumber) labelRenumber.textContent = t.row_action_renumber || '自动修改码号';
    if (labelCancel) labelCancel.textContent = t.row_action_cancel || '取消';
    const labelSplit = document.getElementById('rowActionLabelSplit');
    if (labelSplit) labelSplit.textContent = t.rowActionSplit || '从此行开始新页面';
    const splitBtn = document.getElementById('rowActionSplitBtn');
    if (splitBtn) splitBtn.style.display = (currentActionContext === 'main' && idx > 0 && idx < logs.length - 1) ? '' : 'none';
    const overlay = document.getElementById('rowActionSheet');
    if (overlay) overlay.classList.add('show');
    if (typeof lucide !== 'undefined') lucide.createIcons();
    if (appSettings.keySound && typeof playKeySound === 'function') playKeySound('nxt');
}

function closeRowActionSheet() {
    const overlay = document.getElementById('rowActionSheet');
    if (overlay) overlay.classList.remove('show');
    currentActionLogId = null;
}

function rowActionInsertAbove() {
    if (currentActionLogId == null) return;
    const isHV = currentActionContext === 'hv';
    const targetLogs = isHV ? (historyViewerState.logs || []) : logs;
    const i = targetLogs.findIndex(l => l.id === currentActionLogId);
    closeRowActionSheet();
    if (i < 1) return; // i=0 是编辑卡，不在此插入
    const newLog = { id: Date.now(), code: '', grade: '', length: '', diameter: '', volume: 0, note: '', markGrade: false, markLen: false, markDia: false };
    targetLogs.splice(i, 0, newLog);
    if (isHV) { renderHistoryViewer(); } else { save(); renderAll(); }
}

function rowActionInsertBelow() {
    if (currentActionLogId == null) return;
    const isHV = currentActionContext === 'hv';
    const targetLogs = isHV ? (historyViewerState.logs || []) : logs;
    const i = targetLogs.findIndex(l => l.id === currentActionLogId);
    closeRowActionSheet();
    if (i < 1) return;
    const newLog = { id: Date.now(), code: '', grade: '', length: '', diameter: '', volume: 0, note: '', markGrade: false, markLen: false, markDia: false };
    targetLogs.splice(i + 1, 0, newLog);
    if (isHV) { renderHistoryViewer(); } else { save(); renderAll(); }
}

// 触发自动修改码号：扫描后打开第二层菜单
function rowActionRenumber() {
    if (currentActionLogId == null) return;
    const scan = scanForRenumber(currentActionLogId, currentActionContext);
    closeRowActionSheet();
    if (!scan || scan.totalBelow === 0) return;
    openRenumberSheet(scan);
}

// 扫描当前行以下的码号连续情况（按下方行自身内部顺序判断，非依赖当前行码号）
function scanForRenumber(logId, context) {
    const targetLogs = (context === 'hv') ? (historyViewerState.logs || []) : logs;
    const i = targetLogs.findIndex(l => l.id === logId);
    if (i < 1) return null;
    const currentCode = (targetLogs[i].code || '').trim();
    const currentNum = parseInt(currentCode, 10);
    if (isNaN(currentNum)) return null;

    const padLen = (currentCode.startsWith('0') && currentCode.length > 1) ? currentCode.length : 0;

    const belowIndices = [];
    for (let j = i + 1; j < targetLogs.length; j++) belowIndices.push(j);
    if (belowIndices.length === 0) return null;

    // Count consecutive segment among rows BELOW by their internal sequence
    let consecutiveCount = 0;
    const firstCode = (targetLogs[belowIndices[0]].code || '').trim();
    const firstNum = parseInt(firstCode, 10);
    if (!isNaN(firstNum)) {
        consecutiveCount = 1;
        let expectedNext = firstNum - 1;
        for (let k = 1; k < belowIndices.length; k++) {
            const c = (targetLogs[belowIndices[k]].code || '').trim();
            const n = parseInt(c, 10);
            if (!isNaN(n) && n === expectedNext) { consecutiveCount++; expectedNext--; }
            else break;
        }
    }

    return { context: context || 'main', logIdx: i, currentNum, padLen, belowIndices, consecutiveCount, totalBelow: belowIndices.length };
}

function openRenumberSheet(scan) {
    const t = I18N[currentLang];
    const { consecutiveCount, totalBelow } = scan;
    const hasBreak = consecutiveCount < totalBelow;
    window._renumberScan = scan;

    const titleEl = document.getElementById('renumberSheetTitle');
    if (titleEl) titleEl.textContent = t.renumber_sheet_title || '码号自动修改';

    const noticeEl = document.getElementById('renumberNotice');
    if (noticeEl) {
        const template = hasBreak
            ? (t.renumber_notice_break || '第 {n} 行起码号中断')
            : (t.renumber_notice_consecutive || '下方 {n} 行码号全部连续，确认修改');
        noticeEl.textContent = template.replace('{n}', hasBreak ? (consecutiveCount + 1) : totalBelow);
        noticeEl.style.display = 'block';
    }

    const btnCon = document.getElementById('renumberBtnConsecutive');
    const labelCon = document.getElementById('renumberLabelConsecutive');
    if (btnCon && labelCon) {
        if (hasBreak && consecutiveCount > 0) {
            btnCon.style.display = 'flex';
            labelCon.textContent = `${t.renumber_btn_consecutive || '只修改连续段'}（${consecutiveCount} 行）`;
        } else {
            btnCon.style.display = 'none';
        }
    }

    const labelAll = document.getElementById('renumberLabelAll');
    if (labelAll) labelAll.textContent = `${t.renumber_btn_all || '全部修改'}（${totalBelow} 行）`;

    const labelCancel = document.getElementById('renumberLabelCancel');
    if (labelCancel) labelCancel.textContent = t.row_action_cancel || '取消';

    const overlay = document.getElementById('renumberSheet');
    if (overlay) overlay.classList.add('show');
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function closeRenumberSheet() {
    const overlay = document.getElementById('renumberSheet');
    if (overlay) overlay.classList.remove('show');
    window._renumberScan = null;
}

function doRenumber(mode) {
    const scan = window._renumberScan;
    closeRenumberSheet();
    if (!scan) return;
    const { context, currentNum, padLen, belowIndices, consecutiveCount } = scan;
    const isHV = context === 'hv';
    const targetLogs = isHV ? (historyViewerState.logs || []) : logs;
    const targets = mode === 'consecutive' ? belowIndices.slice(0, consecutiveCount) : belowIndices;
    targets.forEach((j, offset) => {
        const newNum = currentNum - offset - 1;
        targetLogs[j].code = padLen > 0 ? newNum.toString().padStart(padLen, '0') : newNum.toString();
    });
    if (isHV) { renderHistoryViewer(); } else { save(); renderAll(); }
}

function rowActionSplitNewPage() {
    if (!currentActionLogId) return;
    doSplitNewPage(currentActionLogId);
}

function doSplitNewPage(logId) {
    const idx = logs.findIndex(l => l.id === logId);
    if (idx <= 0 || idx >= logs.length - 1) return;

    const savedLogs  = logs.slice(idx + 1);    // 较旧的行（将保存为快照）
    const displayNum = logs.length - idx;      // 被点击行的当前显示序号
    const prevRows   = savedLogs.length;       // 将被保存的行数

    const msg = (I18N[currentLang].confirm_split || '')
        .replace('{prev}', displayNum - 1)
        .replace('{n}', prevRows)
        .replace('{from}', displayNum);

    // 先关闭菜单再弹确认，避免 confirm 关闭时触发遮罩 onclick
    closeRowActionSheet();
    if (!confirm(msg)) return;

    // 生成 sessionId（与 resetLogOnly 相同逻辑）
    let snapId = currentSessionId;
    if (!snapId) {
        const today = new Date();
        const dateStr = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
        const todaySnaps = snapshots.filter(s => s.id.startsWith(dateStr));
        const maxNum = todaySnaps.length > 0 ? Math.max(...todaySnaps.map(s => parseInt(s.id.split('_')[1])||0)) : 0;
        snapId = `${dateStr}_${maxNum + 1}`;
    }

    // 快照 logs 需带空输入卡（historyViewer 将 [0] 作为输入卡处理）
    const emptyCard = { id: 'card_' + Date.now(), code: '', grade: '', length: '', diameter: '', volume: 0, note: '', groupId: '', markGrade: false, markLen: false, markDia: false };
    const snapLogs = [emptyCard].concat(savedLogs);

    // 保存旧行为快照
    const snap = {
        id: snapId,
        timestamp: Date.now(),
        date: new Date().toLocaleString(currentLang === 'zh' ? 'zh-CN' : currentLang === 'en' ? 'en-US' : 'sl-SI'),
        logs: JSON.parse(JSON.stringify(snapLogs)),
        global: JSON.parse(JSON.stringify(globalInfo)),
        container: globalInfo.container || '未命名'
    };
    const existingIdx = snapshots.findIndex(s => s.id === snapId);
    if (existingIdx >= 0) snapshots[existingIdx] = snap;
    else snapshots.unshift(snap);
    lsSet(SNAPSHOTS_KEY, JSON.stringify(snapshots));

    // 更新主界面：移除已保存的行，保留从第 9 行开始的内容（相当于新建但保留后续行）
    logs.splice(idx + 1, logs.length - (idx + 1));

    // 将输入卡置空，相当于新建：清空已保存到内部记录的数据，只保留第 9 行及之后的列表内容
    const nextCode = (typeof isQuickMode !== 'undefined' && isQuickMode && typeof appSettings !== 'undefined' && !appSettings.quickModeAutoCode) ? '' : (typeof getNextCode === 'function' ? getNextCode() : '');
    logs[0] = { id: Date.now(), code: nextCode, grade: '', length: '', diameter: '', volume: 0, note: '', groupId: '', markGrade: false, markLen: false, markDia: false };

    // 与主页面「新柜」一致：号模式时项目信息里的号自动 +1（如输入框为 1 则变为 2，以此类推）
    const currentContainerValue = (document.getElementById('g_container')?.value || '').trim() || (globalInfo.container || '').trim();
    const num = parseInt(currentContainerValue, 10);
    const nextContainerValue = (typeof isContainerNumberMode !== 'undefined' && isContainerNumberMode && currentContainerValue && !isNaN(num)) ? (num + 1).toString() : currentContainerValue;
    if (nextContainerValue !== currentContainerValue) {
        globalInfo.container = nextContainerValue;
        const gContainer = document.getElementById('g_container');
        if (gContainer) gContainer.value = nextContainerValue;
    }

    currentSessionId = null;
    localStorage.removeItem(SESSION_KEY);

    save();
    renderAll();
}
