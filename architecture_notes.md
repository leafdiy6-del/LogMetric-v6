# LogMetric Pro — 架构与算法笔记

> 本文档记录项目核心业务逻辑与算法实现，供维护与回归测试参考。  
> 最后更新：2025-03

---

## 一、捷克 ČSN 48 0009 体积算法（橡木树皮公式）

### 1.1 概述

捷克 ČSN 48 0009 为 LogMetric Pro 可选体积公式，采用 STN/ČSN 橡木树皮公式计算净直径。启用后在设置中选择 `csn480009` 即可切换。（兼容旧设置值 `csn4800079`）

**实现位置**：`js/app.js` 中的 `calculateCzechVolume(length, diameterOverBark)` 函数。

### 1.2 算法流程

#### A. 输入清理

- 将输入中的逗号替换为小数点：`String(x).replace(/,/g, '.')`
- 使用 `parseFloat` 解析；若结果为 `NaN` 或 `l ≤ 0` 或 `d ≤ 0`，直接返回 `0`

#### B. STN/ČSN 48 0009 橡木树皮公式

- 公式：`2k = p0 + p1 × D^p2`
- 橡木系数：`p0 = 1.2474`, `p1 = 0.042323`, `p2 = 1.0623`
- `bark2k = 1.2474 + 0.042323 × d^1.0623`
- 净直径：`dNet = Math.max(0, d - bark2k)`

#### C. 横截面积

- 公式：`area = π × dNet² / 40000`（m²）
- 截断：`area = Math.floor(area × 1000) / 1000`

#### D. 体积计算

- `volume = area × l`（无长度扣除，使用原始长度）

### 1.3 调用链

```
calculateVolume(length, diameter)
  └─ 若 appSettings.formulaEnabled && (formula === 'csn480009' || formula === 'csn4800079')
       └─ calculateCzechVolume(length, diameter)
  └─ 否则：标准 Huber 公式 (π × (d/200)² × l)
```

`recalculateAllVolumes` 在公式切换时重新计算全部材积并保存。

---

## 二、UI 与样式规范（本次收尾不涉及修改）

- **主题**：Soft Industrial + Glassmorphism（参考 `UI-UPGRADE.md`）
- **CSS 变量**：`css/app.css` 中的 `:root` 定义
- **主题扩展**：`css/theme-addon.css` 负责浅色/深色切换
- **设置界面**：通过 `index.html` 与 `js/app.js` 中的模态逻辑渲染

**本次里程碑收尾仅涉及文档与代码清理，不修改任何 CSS、HTML 结构或设置界面渲染逻辑。**

---

## 三、项目结构

```
v6/
├── index.html           # 入口页
├── manifest.json        # PWA 配置
├── architecture_notes.md # 本架构笔记
├── UI-UPGRADE.md        # UI 升级说明
├── css/
│   ├── app.css          # 主样式
│   └── theme-addon.css  # 主题扩展
└── js/
    ├── app.js           # 主应用逻辑（含捷克公式）
    ├── core.js          # 常量与工具
    ├── config.js        # 外部配置
    ├── i18n.js          # 多语言
    ├── license.js       # 许可与激活
    └── sounds.js        # 音效
```
