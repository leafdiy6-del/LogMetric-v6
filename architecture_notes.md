# LogMetric Pro — 架构与算法笔记

> 本文档记录项目核心业务逻辑与算法实现，供维护与回归测试参考。  
> 最后更新：2025-03

---

## 一、捷克 ČSN 48 0007/9 实战变体（定稿版）

### 1.1 概述

捷克 ČSN 48 0007/9 (Huber Variant) 为 LogMetric Pro 可选体积公式，与标准 Huber 公式并存。启用后在设置中选择 `csn4800079` 即可切换。

**实现位置**：`js/app.js` 中的 `calculateCzechVolume(length, diameterOverBark)` 函数。

### 1.2 算法流程（定稿逻辑）

#### A. 输入清理

- 将输入中的逗号替换为小数点：`String(x).replace(/,/g, '.')`
- 使用 `parseFloat` 解析；若结果为 `NaN` 或 `l ≤ 0` 或 `d ≤ 0`，直接返回 `0`

#### B. 四段式长度扣除 (Nadměrek / Length Deduction)

| 长度区间           | 扣除规则        | lNet 公式  |
|--------------------|-----------------|------------|
| `length ≥ 11`      | 扣 0.3 m        | lNet = l - 0.3 |
| `10 ≤ length < 11` | 无扣除          | lNet = l |
| `8 ≤ length < 10`  | 扣 0.1 m        | lNet = l - 0.1 |
| `length < 8`       | 扣 0.2 m        | lNet = l - 0.2 |

**底线保护（Critical Fix）**：计算 lNet 后执行 `lNet = Math.max(0, lNet)`，防止超短木材（如 0.1 m）产生负数体积。

#### C. 三阶树皮扣除 (Bark Deduction)

- 直径向下取整：`dFloored = Math.floor(d)`
- 阶梯扣除规则：

| 直径区间         | 扣除量 |
|------------------|--------|
| `d ≤ 44`         | 3 cm   |
| `45 ≤ d ≤ 50`    | 4 cm   |
| `d > 50`         | 5 cm   |

- 净直径：`dNet = Math.max(0, dFloored - deduction)`

#### D. 横截面积（3 位小数 floor 截断）

- 公式：`area = π × dNet² / 40000`（面积单位 m²）
- 截断：`area = Math.floor(area × 1000) / 1000`
- 与 ČSN 标准表数值对齐

#### E. 体积计算与返回

- `volume = area × lNet`
- 返回高精度数值，舍入仅在 `formatVolumeForDisplay` 中进行
- 返回值直接写入 `item.volume` / `log.volume`，供状态管理与 UI 使用

---

### 1.3 金标准验证数据（回归测试依据）

以下 4 组为核心金标准，用于未来回归测试。实现变更后必须保证输出一致。

| 序号 | 输入 (长度 m × 直径 cm) | lNet | dNet | area | 体积（原始） | 显示（formatVolumeForDisplay 三位） |
|------|--------------------------|------|------|------|--------------|-------------------------------------|
| 1    | **11.6 × 41**            | 11.3 | 38   | 0.113 | 1.2769      | **1.277**                           |
| 2    | **7.3 × 43**             | 7.1  | 40   | 0.125 | 0.8875      | **0.888**                           |
| 3    | **10.1 × 32**            | 10.1 | 29   | 0.066 | 0.6666      | **0.667**                           |
| 4    | **8.3 × 39**             | 8.2  | 36   | 0.101 | 0.8282      | **0.828**                           |

> 注：显示值由 `formatVolumeForDisplay` 按三位小数 `toFixed(3)` 后 `parseFloat().toString()` 输出，会去掉末位无意义的 0。

#### 推导校验（便于手工复算）

1. **11.6 × 41**  
   - lNet = 11.6 - 0.3 = 11.3  
   - dFloored = 41, deduction = 3, dNet = 38  
   - area = floor(π×38²/40000 × 1000)/1000 = 0.113  
   - volume = 0.113 × 11.3 = 1.2769  

2. **7.3 × 43**  
   - lNet = 7.3 - 0.1 = 7.1  
   - dFloored = 43, deduction = 3, dNet = 40  
   - area = floor(π×40²/40000 × 1000)/1000 = 0.125  
   - volume = 0.125 × 7.1 = 0.8875  

3. **10.1 × 32**  
   - lNet = 10.1（10–11 m 无扣）  
   - dFloored = 32, deduction = 3, dNet = 29  
   - area = floor(π×29²/40000 × 1000)/1000 = 0.066  
   - volume = 0.066 × 10.1 = 0.6666  

4. **8.3 × 39**  
   - lNet = 8.3 - 0.1 = 8.2  
   - dFloored = 39, deduction = 3, dNet = 36  
   - area = floor(π×36²/40000 × 1000)/1000 = 0.101  
   - volume = 0.101 × 8.2 = 0.8282  

#### 边界与漏洞防护测试

- **超短样本 (0.1 × 30)**  
  - 原始 lNet = 0.1 - 0.2 = -0.1  
  - 底线保护后：lNet = max(0, -0.1) = 0  
  - 体积 = 0，避免负数漏洞  

---

### 1.4 调用链

```
calculateVolume(length, diameter)
  └─ 若 appSettings.formulaEnabled && appSettings.formula === 'csn4800079'
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
