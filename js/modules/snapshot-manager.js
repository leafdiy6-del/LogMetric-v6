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
        const bgColor = isCurrent ? 'rgba(212,163,115,0.2)' : 'var(--surface-secondary)';
        const hoverBg = isCurrent ? 'rgba(212,163,115,0.3)' : 'var(--surface-primary)';
        const border = isCurrent ? '2px solid var(--accent-color)' : '1px solid var(--border-color)';
        const badge = isCurrent ? '<span style="color:var(--accent-color);font-weight:600;margin-left:10px;"><i data-lucide="circle-dot" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> 当前</span>' : '';

        return `
                <div style="background:${bgColor};padding:15px;margin:10px 0;border-radius:8px;border:${border};cursor:pointer;" 
                     onclick="openHistoryViewer('${snap.id}')" 
                     onmouseover="this.style.background='${hoverBg}'" 
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
                <div class="export-record-item" style="display:flex;align-items:center;gap:12px;padding:12px 15px;margin:8px 0;background:var(--surface-secondary);border-radius:8px;border:1px solid var(--border-color);">
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
