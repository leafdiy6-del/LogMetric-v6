# 代码混淆说明

## 快速结论：哪些文件需要混淆（比较安全）

- **建议混淆**：`js/license.js`、`js/config.js`、`js/core.js`、`js/app.js`、`js/sounds.js`、`js/modules/*.js`、`js/data/slovakTable.js`
- **仅压缩、不混淆 key**：`js/i18n.js`（否则会破坏 `I18N[key]` 文案）
- **混淆时**：必须保留 HTML 与动态 HTML 里用到的全部全局函数名（见下文保留名单），否则点击/输入会报错

---

## 一、建议混淆的文件（安全优先）

| 文件 | 说明 |
|------|------|
| **js/license.js** | 授权/激活逻辑，优先混淆 |
| **js/config.js** | 含 Supabase URL/Key，混淆时需保留 `LMP_CONFIG` 及属性名（见下方保留名单） |
| **js/core.js** | 核心工具，可混淆 |
| **js/app.js** | 主逻辑，**必须保留** HTML 与动态 HTML 中调用的全局函数名（见下方保留名单） |
| **js/sounds.js** | 可混淆 |
| **js/modules/volume-calc.js** | 体积计算，可混淆 |
| **js/modules/stats.js** | 可混淆 |
| **js/modules/grouping-marking.js** | 需保留 `setGrade`（被 app 与 pro-keyboard 调用） |
| **js/modules/history-autocomplete.js** | 可混淆 |
| **js/modules/row-actions.js** | 可混淆 |
| **js/modules/history-viewer.js** | 可混淆 |
| **js/modules/snapshot-manager.js** | 可混淆 |
| **js/modules/pro-keyboard.js** | 可混淆 |
| **js/modules/pdf-generator.js** | 可混淆 |
| **js/modules/excel-exporter.js** | 可混淆 |
| **js/data/slovakTable.js** | 若为纯数据且无动态 key 访问，可混淆；否则保留属性名 |

## 二、不建议混淆或需特殊处理的文件

| 文件 | 说明 |
|------|------|
| **js/i18n.js** | 大量 `I18N[currentLang].xxx` 动态 key 访问，混淆 key 会破坏功能。**建议只压缩不混淆**，或混淆时保留 `I18N`、`HELP_TEXTS` 及所有文案 key（zh、en、sl 及各 key 名）。 |

## 三、必须保留的全局名称（否则页面/逻辑会报错）

### 1. 配置与国际化（config / i18n）

- `LMP_CONFIG`、`SUPABASE_URL`、`SUPABASE_ANON_KEY`（若混淆 config.js）
- `I18N`、`HELP_TEXTS`、`currentLang`（若混淆 i18n 或 app）

### 2. HTML 内联调用的函数（index.html + app 内动态 HTML）

以下名称在 `onclick` / `onchange` / `oninput` / `onblur` / `ondblclick` / `oncontextmenu` 等中被使用，**必须保留**：

```
handleLicenseActivate, resetLogOnly, showHelp, openInfoModal, openSettingsModal,
toggleSaveMenu, saveToSnapshot, openSnapshotHistory, triggerImportProjectFile, closeSaveMenu,
exportPackageAll, generatePDF, exportData, handleImportProjectFile,
saveHistoryViewerChanges, generateHistoryViewerPDF, exportHistoryViewerExcel,
toggleHistoryViewerInfo, closeHistoryViewer, addHistoryViewerRow,
handleProKey, addNewLog, handleProDelete, handleProSide, handleProOk,
toggleGroupMode, toggleContainerMode, downloadFullBackup, triggerImportFromInfo,
updateGlobal, saveToHistory, showHistory, openCompanyDetailModal, toggleShowCompanyInPdf,
openSellerDetailModal, openSellerDetailModalAsNew, closeInfoModal,
closeCompanyDetailModal, saveCompanyDetailModal, clearSellerFormAndNew,
closeSellerDetailModal, saveSellerDetailModal, toggleTheme, updateThemeToggleLabel,
changeLanguage, toggleGradeDisplay, toggleCalcDia, saveSettings, toggleQuickModeExpand,
toggleQuickMode, saveQuickModeOptions, toggleFormulaEnabled, recalculateAllVolumes,
setGrade, updateItem, handleQuickLength, handleQuickDia1, handleQuickDia2,
handleQuickDiameterSingle, toggleDiaDangerClass, autoFixInput,
openRowActionSheet, handleGradeLabelEdit, handleGradeLabelEditByGrade,
handleMarkCellClick, ungroupLog, delItem, handleGroupRowClick
```

### 3. 其他脚本依赖的全局变量/函数（按需保留）

- `appSettings`, `logs`, `globalInfo`, `currentSessionId`, `snapshots`, `histories`
- `currentLang`, `isQuickMode`, `isContainerNumberMode`
- 以及各模块之间通过 `window.xxx` 或全局调用的函数（如 `setGrade` 在 grouping-marking 与 pro-keyboard 中调用）

混淆时建议使用 **reservedNames**（或等价配置）把以上名称全部加入保留列表。

## 四、推荐工具与用法

- **javascript-obfuscator**（Node）
  - 安装：`npm install -g javascript-obfuscator` 或项目内 `npm install --save-dev javascript-obfuscator`
  - 单文件示例：
    ```bash
    javascript-obfuscator js/license.js --output js/license.obf.js --config '{"reservedNames":["LMP_CONFIG","handleLicenseActivate"]}'
    ```
  - 混淆时把「三、必须保留的全局名称」中与当前文件相关的部分写入 `reservedNames`，避免破坏 HTML 与模块调用。

- **只压缩不混淆**（适合 i18n.js）
  - 使用 Terser 等仅做压缩，不修改变量/属性名，避免破坏 `I18N[key]` 访问。

## 五、部署流程建议

1. 保留一份**未混淆源码**（仅本地或私有仓库），用于排查问题和后续更新。
2. 使用**单独目录或构建产物**存放混淆后的 JS（例如 `dist/` 或 `js-built/`），再在 HTML 中引用混淆后的脚本。
3. **config.js**：若含敏感信息，建议部署时用环境变量或构建脚本替换内容，混淆只作为额外保护。
4. 混淆后**务必在真机/多浏览器**测试：授权、设置、列表、导出 PDF/Excel、历史记录等，确保无报错。

按上述范围混淆并保留名单后，可以在「比较安全」的前提下尽量保护业务与授权逻辑。

---

## 六、仅混淆 license / config / core 时的保留名单（可直接复制到工具）

以下按文件列出**必须保留**的标识符，混淆时在工具的 Reserved Names 中填入对应列表即可。

### js/license.js（High）

- **必须保留**：只有 `handleLicenseActivate` 被 index.html 的 `onclick` / `onkeydown` 调用，其它名均可被混淆。
- license.js 内部会访问 `window.LMP_CONFIG.SUPABASE_URL`、`window.LMP_CONFIG.SUPABASE_ANON_KEY`（来自 config.js），以及 `window.supabase`、`window.FingerprintJS`（来自 CDN），这些是**外部**对象，不需要在 license 的保留名单里写；只要 config.js 保留 `LMP_CONFIG` 与属性名即可。

**Reserved Names（license.js）：**
```
handleLicenseActivate
```

### js/config.js（中等）

- **必须保留**：`LMP_CONFIG` 是挂到 `window` 上的对象名，license.js 通过 `window.LMP_CONFIG.SUPABASE_URL` / `SUPABASE_ANON_KEY` 读取，因此对象名与这两个**属性名**都不能被混淆。

**Reserved Names（config.js）：**
```
LMP_CONFIG
SUPABASE_URL
SUPABASE_ANON_KEY
```

### js/core.js（中等）

- **必须保留**：core.js 定义的常量和函数被 app.js、i18n.js、license 以外的多个模块直接使用（按名引用），以下全部需要保留。

**Reserved Names（core.js）：**
```
STORAGE_KEY
HIST_KEY
LANG_KEY
QUICK_KEY
SETTINGS_KEY
MIX_STATE_KEY
SNAPSHOTS_KEY
SESSION_KEY
GRADES
COMPANY_FIELDS
defaultCompany
defaultSeller
normalizeCompany
normalizeSeller
getCompanyName
migrateHistoriesSeller
migrateHistoriesCompany
```

---

**检查结论**：  
- license.js：只保留 `handleLicenseActivate` 即可，其它逻辑名可混淆。  
- config.js：保留 `LMP_CONFIG`、`SUPABASE_URL`、`SUPABASE_ANON_KEY` 三个。  
- core.js：保留上面 17 个常量和函数名，避免 app、history-autocomplete、pdf-generator、excel-exporter、grouping-marking、history-viewer、snapshot-manager 等报错。  

若使用 obfuscator.io：在对应文件的设置里找到 Reserved Names / reservedNames，将上面各文件的列表逐行或逗号粘贴进去即可。
