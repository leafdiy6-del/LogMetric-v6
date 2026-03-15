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
