# LogMetric Pro — 架构说明（AI 参考文档）

> **AI 使用说明：** 修改前先读本文件，找到对应模块再读源码，不要直接读 app.js 全文。

---

## 1. 技术栈与架构

- **纯原生 HTML + CSS + JS**，无框架、无构建工具、无 npm
- 所有 JS 文件共享**全局作用域**（无 `import/export`）
- 入口：`index.html`，按 `<script>` 顺序加载，`app.js` 最后加载
- 数据持久化：`localStorage`（无后端）
- 外部库（CDN）：`jsPDF`、`html2canvas`（PDF）、`XLSX`（Excel）、`lucide`（图标）、`Supabase`（授权验证）

---

## 2. 核心全局状态（定义在 `app.js` 顶部）

修改任何功能前，先了解这几个全局变量：

| 变量 | 类型 | 说明 |
|------|------|------|
| `logs` | `Array<Log>` | 当前主界面所有原木记录（见下方数据模型） |
| `globalInfo` | `Object` | 当前项目信息（集装箱号、备注、公司、买卖方等） |
| `appSettings` | `Object` | 所有设置项（公式、键盘、价格、语言等） |
| `snapshots` | `Array` | 已保存的快照列表（存 localStorage） |
| `historyViewerState` | `Object` | 历史查看器的临时状态（logs、global、id 等） |
| `proState` | `Object` | 专业键盘当前状态（activeField、values、keypadMode 等） |
| `currentLang` | `string` | 当前语言：`'zh'` / `'en'` / `'sl'` |
| `isMarkMode` | `boolean` | 是否处于标记模式 |
| `isGroupMode` | `boolean` | 是否处于分组模式 |
| `isQuickMode` | `boolean` | 是否处于快速输入模式 |

### Log 数据模型（单条原木记录）

```js
{
  id: string,        // 唯一 ID（时间戳）
  code: string,      // 码号（如 "001"）
  grade: string,     // 等级（"A" / "B" / "C" / "D" / "A+" / "F"）
  length: string,    // 长度（米，字符串存储）
  diameter: string,  // 直径（厘米，字符串存储）
  volume: number,    // 材积（m³，计算值）
  note: string,      // 备注
  groupId: string,   // 分组 ID（有分组时非空）
  markGrade/markLen/markDia: boolean  // 字段标记
}
```

### appSettings 关键字段

```js
appSettings.formula        // 'huber' | 'csn480009' | 'lookupIgor' | ''（默认圆柱）
appSettings.formulaEnabled // boolean，是否启用自定义公式
appSettings.proKeyboard    // boolean，专业键盘开关
appSettings.priceEnabled   // boolean，价格功能开关
appSettings.priceMode      // 'fixed' | 'grade'，定价方式
appSettings.currentLang    // 语言（同 currentLang 全局变量）
appSettings.showCompanyInPdf // boolean，PDF 中显示公司信息
```

---

## 3. 快速定位表（先查这里）

### JS 修改

| 要修改的功能 | 文件 | 关键函数 |
|-------------|------|---------|
| 体积计算公式 | `js/modules/volume-calc.js` | `calculateVolume`, `calculateHuberVolume`, `calculateCzechVolume`, `calculateSlovakTableVolume` |
| PDF 样式/内容/布局 | `js/modules/pdf-generator.js` | `generatePDF(options)` |
| Excel 导出格式/列 | `js/modules/excel-exporter.js` | `exportData(options)` |
| 底部统计显示 | `js/modules/stats.js` | `updateStats`, `getValidLogsGroupStats` |
| 分组/标记逻辑 | `js/modules/grouping-marking.js` | `toggleGroupMode`, `handleGroupRowClick`, `setGrade` |
| 行操作菜单/重编号 | `js/modules/row-actions.js` | `openRowActionSheet`, `doRenumber` |
| 历史查看器（查看/编辑快照） | `js/modules/history-viewer.js` | `openHistoryViewer`, `renderHistoryViewer`, `saveHistoryViewerChanges` |
| 快照保存/加载/删除 | `js/modules/snapshot-manager.js` | `saveToSnapshot`, `loadSnapshot`, `deleteSnapshot` |
| 项目文件导出/导入 | `js/modules/snapshot-manager.js` | `exportProjectFile`, `handleImportProjectFile` |
| 批量导出快照 | `js/modules/snapshot-manager.js` | `openExportProjectModal`, `doExportSelectedRecords` |
| 专业键盘按键逻辑 | `js/modules/pro-keyboard.js` | `handleProKey`, `handleProNext`, `handleProOk` |
| 专业键盘字段切换 | `js/modules/pro-keyboard.js` | `setProActiveField`, `handleProQuickFlow` |
| 专业键盘渲染/UI | `js/modules/pro-keyboard.js` | `applyProKeyboardUI`, `renderProKeyboardTopBar`, `renderProGradePanel` |
| 输入历史下拉补全 | `js/modules/history-autocomplete.js` | `showHistory`, `selectHistory`, `saveToHistory` |
| 添加新记录 | `js/app.js` | `addNewLog` |
| 渲染整个列表 | `js/app.js` | `renderAll`, `createCard`, `createRow` |
| 修改/删除记录 | `js/app.js` | `updateItem`, `delItem` |
| 快速输入自动跳转 | `js/app.js` | `handleQuickLength`, `handleQuickDia1`, `handleQuickDia2`, `jumpLengthToDia` |
| 设置弹窗 | `js/app.js` | `openSettingsModal`, `saveSettings` |
| 公司/买卖方弹窗 | `js/app.js` | `openCompanyDetailModal`, `openSellerDetailModal`, `saveCompanyDetailModal` |
| 主题切换（深/浅色） | `js/app.js` | `initTheme`, `toggleTheme` |
| 多语言文字 | `js/i18n.js` | `I18N['zh'/'en'/'sl']` 对象 |
| 全局常量/默认值 | `js/config.js` | `STORAGE_KEY`, `defaultCompany()`, `COMPANY_FIELDS` 等 |

### CSS 修改

| 要修改的样式 | 文件 |
|-------------|------|
| 颜色变量、间距、字体大小 | `css/variables.css` |
| 页面整体布局、header、容器 | `css/base.css` |
| 按钮、输入框、下拉、表单 | `css/components.css` |
| 卡片行、列表行、标记/分组样式 | `css/cards-lists.css` |
| 所有弹窗、设置界面 | `css/modals.css` |
| 专业键盘、横屏分屏布局 | `css/pro-keyboard.css` |
| 激活遮罩、滚动条、警告、行操作菜单 | `css/extras.css` |
| 浅色（白天）主题所有覆盖 | `css/theme-addon.css` |

---

## 4. 文件结构

```
index.html                     # 唯一入口，含所有 modal HTML 和 <script>/<link> 引入

css/
├── variables.css              # CSS 自定义变量（颜色/间距/字体，改颜色先看这里）
├── base.css                   # 基础重置 + 整体布局容器
├── components.css             # 按钮 / 下拉菜单 / 输入框 / 表单
├── cards-lists.css            # 卡片 / 列表行 / 统计显示
├── modals.css                 # 弹窗 / 模态框 / 设置界面
├── pro-keyboard.css           # 专业键盘 + 横屏分屏布局 + 动画
├── extras.css                 # 激活层 / 滚动条 / 警告 / 行操作菜单 / 杂项
└── theme-addon.css            # 浅色主题（必须最后加载，覆盖深色默认值）

js/
├── config.js                  # 常量、localStorage key、默认数据结构
├── core.js                    # 基础工具（normalizeCompany, getCompanyName 等）
├── i18n.js                    # 多语言：I18N['zh'/'en'/'sl'] 对象
├── license.js                 # 授权验证（Supabase）
├── sounds.js                  # 按键音效（playKeySound）
├── data/slovakTable.js        # 斯洛伐克材积查表数据（SLOVAK_TABLE 全局对象）
├── modules/
│   ├── volume-calc.js         # 体积计算（4种公式 + 查表插值）
│   ├── stats.js               # 底部统计 + 过滤
│   ├── grouping-marking.js    # 分组选中 + 字段标记 + 等级标签
│   ├── history-autocomplete.js# 输入历史下拉（公司/地点/测量人等）
│   ├── row-actions.js         # 双击行序号菜单 + 码号重编号
│   ├── history-viewer.js      # 历史快照查看/编辑/导出（全屏覆盖层）
│   ├── snapshot-manager.js    # 快照 CRUD + 项目文件导入导出 + 批量操作
│   ├── pro-keyboard.js        # 专业虚拟键盘（状态、渲染、按键处理）
│   ├── pdf-generator.js       # PDF 生成（html2canvas 渲染）
│   └── excel-exporter.js      # Excel 生成（SheetJS/XLSX）
└── app.js                     # 主协调器（全局状态、渲染、设置、初始化）
```

---

## 5. 模块间依赖关系

```
config.js, core.js, i18n.js          ← 无依赖，最基础
    ↓
volume-calc.js                        ← 依赖 appSettings（app.js 的全局变量）
stats.js                              ← 依赖 logs, appSettings, renderAll
grouping-marking.js                   ← 依赖 logs, save(), renderAll(), setGrade→updateItem
history-autocomplete.js               ← 依赖 histories（全局变量）
row-actions.js                        ← 依赖 logs, renderAll, historyViewerState
history-viewer.js                     ← 依赖 snapshots, generatePDF, exportData
snapshot-manager.js                   ← 依赖 logs, globalInfo, openHistoryViewer
pro-keyboard.js                       ← 依赖 logs, appSettings, updateItem, addNewLog
pdf-generator.js                      ← 依赖 logs, globalInfo, appSettings, window.jspdf
excel-exporter.js                     ← 依赖 logs, globalInfo, appSettings, window.XLSX
    ↓
app.js                                ← 依赖所有模块，最后加载
```

> **关键规则：** 所有模块都能直接访问 `app.js` 定义的全局变量（`logs`, `appSettings` 等），
> 因为 `app.js` 先于模块执行时的函数调用（模块只定义函数，不立即执行）。

---

## 6. 数据流

```
用户输入
  → updateItem(id, field, val)       [app.js]
      → calculateVolume()            [volume-calc.js]
      → updateVolumeDisplay()        [app.js]
      → save()                       [app.js → localStorage]
      → updateStats()                [stats.js]

用户点击"添加并下一根"
  → addNewLog()                      [app.js]
      → getNextCode()                [app.js]
      → calculateVolume()            [volume-calc.js]
      → save()
      → renderAll()
      → updateStats()

用户点击"保存/导出" → saveToSnapshot()  [snapshot-manager.js]
用户点击历史记录   → openHistoryViewer() [history-viewer.js]
用户导出 PDF      → generatePDF()        [pdf-generator.js]
用户导出 Excel    → exportData()         [excel-exporter.js]
```

---

## 7. 重要约束（修改前必读）

1. **无模块系统**：所有函数都是全局函数，跨文件调用直接写函数名即可
2. **CSS 加载顺序**：`variables.css` 必须第一个，`theme-addon.css` 必须最后一个（覆盖深色默认）
3. **JS 加载顺序**：`app.js` 必须最后加载；模块之间的函数调用在运行时才发生，不受定义顺序影响
4. **logs 数组**：`logs[0]` 始终是"当前正在输入的那条"（最新的在最前面），渲染时用 `[...logs].reverse()` 倒序显示
5. **体积精度**：计算时保持高精度，仅在 `formatVolumeForDisplay()` 时做 UI 舍入，不要在计算层舍入
6. **历史查看器**：操作的是 `historyViewerState.logs`，不是全局 `logs`；保存后才写回 `snapshots`
7. **多语言**：所有面向用户的文字必须通过 `I18N[currentLang].xxx` 获取，不能硬编码中文
8. **localStorage key**：在 `config.js` 中定义，不要在模块里硬编码

---

## 8. 添加新功能的标准步骤

**添加新 JS 功能（>100 行）：**
1. 新建 `js/modules/xxx.js`
2. 在 `index.html` 中 `js/app.js` 的 `<script>` 之前添加 `<script src="js/modules/xxx.js">`
3. 如有多语言文字，在 `js/i18n.js` 的三个语言对象中同步添加
4. 如有新设置项，在 `js/config.js` 的默认值中添加，并在 `app.js` 的 `saveSettings()` 中处理

**添加新 CSS 样式：**
1. 颜色/间距：先在 `css/variables.css` 添加变量，然后引用变量
2. 新组件样式：添加到最相关的 CSS 文件末尾
3. 浅色主题覆盖：**只加在** `css/theme-addon.css`，不要分散到其他文件

**修改多语言文字：**
- 只改 `js/i18n.js`，三种语言（`zh`/`en`/`sl`）必须同步修改
