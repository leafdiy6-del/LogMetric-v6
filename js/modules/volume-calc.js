function calculateVolume(length, diameter) {
    if (appSettings.formulaEnabled && appSettings.formula === 'huber') {
        return calculateHuberVolume(length, diameter);
    }
    if (appSettings.formulaEnabled && (appSettings.formula === 'csn480009' || appSettings.formula === 'csn4800079')) {
        return calculateCzechVolume(length, diameter);
    }
    if (appSettings.formulaEnabled && appSettings.formula === 'lookupIgor') {
        return calculateSlovakTableVolume(length, diameter);
    }

    const l = parseFloat(String(length).replace(/,/g, '.'));
    const d = parseFloat(String(diameter).replace(/,/g, '.'));

    if (!isNaN(l) && !isNaN(d) && l > 0 && d > 0) {
        const radius_m = d / 200;
        return Math.PI * Math.pow(radius_m, 2) * l;
    }
    return 0;
}

/**
 * 欧洲标准 Huber 公式 (EN 1309-2)
 * V = (π/4) × (d/100)² × L
 * d = 中央直径（去皮，cm），L = 原木长度（m），V = 体积（m³）
 */
function calculateHuberVolume(length, diameterMidpoint) {
    const l = parseFloat(String(length).replace(/,/g, '.'));
    const d = parseFloat(String(diameterMidpoint).replace(/,/g, '.'));
    if (isNaN(l) || isNaN(d) || l <= 0 || d <= 0) return 0;
    const volume = (Math.PI / 4) * Math.pow(d / 100, 2) * l;
    return Math.round(volume * 1000) / 1000;
}

/**
 * 捷克 ČSN 48 0009 体积算法（STN/ČSN 橡木 Dub 官方高精度系数）
 * - 树皮公式 2k = p0 + p1 × D^p2
 * - 面积直接计算，不用 floor 截断
 * - 体积四舍五入到 2 位小数
 */
function calculateCzechVolume(length, diameterOverBark) {
    const l = parseFloat(String(length).replace(/,/g, '.'));
    const d = parseFloat(String(diameterOverBark).replace(/,/g, '.'));
    if (isNaN(l) || isNaN(d) || l <= 0 || d <= 0) return 0;

    // 1. ČSN 48 0009 官方高精度系数 (针对橡木 Dub)
    const p0 = 1.2474;
    const p1 = 0.042323;
    const p2 = 1.0623;

    // 2. 计算树皮厚度 (2k)
    const barkDeduction = p0 + p1 * Math.pow(d, p2);

    // 3. 计算去皮直径
    const dNet = Math.max(0, d - barkDeduction);

    // 4. 计算截面积 (直接计算，严禁使用 Math.floor)
    const area = (Math.PI / 4) * Math.pow(dNet / 100, 2);

    // 5. 计算体积 (直接用原始长度 l，不要减去 0.1 或 0.3)
    const volume = area * l;

    // 6. 最终体积四舍五入到 2 位小数
    return Math.round(volume * 100) / 100;
}

/**
 * 校验 SLOVAK_TABLE 单调性：同直径时长度越长体积越大，同长度时直径越大体积越大。
 * 若发现违反则仅在控制台打印警告，不阻断运行。
 */
function validateSlovakTableMonotonicity() {
    const tbl = typeof window !== 'undefined' && window.SLOVAK_TABLE;
    if (!tbl || !tbl.table || !tbl.lengths || !tbl.diameters) return;
    const lengths = tbl.lengths;
    const diameters = tbl.diameters;
    let hasWarn = false;
    for (let i = 0; i < lengths.length; i++) {
        const L = lengths[i];
        const row = tbl.table[String(L)];
        if (!row) continue;
        for (let j = 0; j < diameters.length; j++) {
            const D = diameters[j];
            const v = row[String(D)];
            if (v == null) continue;
            for (let i2 = 0; i2 < i; i2++) {
                const L2 = lengths[i2];
                const row2 = tbl.table[String(L2)];
                if (!row2) continue;
                const v2 = row2[String(D)];
                if (v2 != null && v < v2) {
                    if (!hasWarn) hasWarn = true;
                    console.warn('[SLOVAK_TABLE] 单调性异常：长度 ' + L + 'm、直径 ' + D + 'cm 体积 ' + v + ' 小于 长度 ' + L2 + 'm、直径 ' + D + 'cm 体积 ' + v2);
                }
            }
            for (let j2 = 0; j2 < j; j2++) {
                const D2 = diameters[j2];
                const v2 = row[String(D2)];
                if (v2 != null && v < v2) {
                    if (!hasWarn) hasWarn = true;
                    console.warn('[SLOVAK_TABLE] 单调性异常：长度 ' + L + 'm、直径 ' + D + 'cm 体积 ' + v + ' 小于 长度 ' + L + 'm、直径 ' + D2 + 'cm 体积 ' + v2);
                }
            }
        }
    }
    if (!hasWarn && lengths.length > 0 && diameters.length > 0) {
        console.log('[SLOVAK_TABLE] 单调性校验通过');
    }
}

/**
 * Igor Števlík DUB 橡木材积查表法 (Lookup Table-igor)
 * 表数据来源: js/data/slovakTable.json
 * 长度或直径非整数时使用线性插值
 */
function calculateSlovakTableVolume(length, diameter) {
    const l = parseFloat(String(length).replace(/,/g, '.'));
    const d = parseFloat(String(diameter).replace(/,/g, '.'));
    if (isNaN(l) || isNaN(d) || l <= 0 || d <= 0) return 0;

    const tbl = typeof window !== 'undefined' && window.SLOVAK_TABLE;
    if (!tbl || !tbl.table || !tbl.lengths || !tbl.diameters) return 0;

    const lenArr = tbl.lengths;
    const diaArr = tbl.diameters;
    const lenMin = lenArr[0], lenMax = lenArr[lenArr.length - 1];
    const diaMin = diaArr[0], diaMax = diaArr[diaArr.length - 1];

    if (l < lenMin || l > lenMax || d < diaMin || d > diaMax) return 0;

    const L1 = Math.floor(l), L2 = Math.ceil(l);
    const D1 = Math.floor(d), D2 = Math.ceil(d);
    const lFrac = (L1 === L2) ? 0 : (l - L1) / (L2 - L1);
    const dFrac = (D1 === D2) ? 0 : (d - D1) / (D2 - D1);

    const getV = (len, dia) => {
        const row = tbl.table[String(len)];
        if (!row) return null;
        const v = row[String(dia)];
        return v != null ? v : null;
    };

    const vL1D1 = getV(L1, D1), vL1D2 = getV(L1, D2), vL2D1 = getV(L2, D1), vL2D2 = getV(L2, D2);
    if (vL1D1 == null || vL1D2 == null || vL2D1 == null || vL2D2 == null) return 0;

    const vL1 = dFrac === 0 ? vL1D1 : vL1D1 + (vL1D2 - vL1D1) * dFrac;
    const vL2 = dFrac === 0 ? vL2D1 : vL2D1 + (vL2D2 - vL2D1) * dFrac;
    const volume = lFrac === 0 ? vL1 : vL1 + (vL2 - vL1) * lFrac;
    return Math.round(volume * 100) / 100;
}

// 切换公式时重新计算全部
window.recalculateAllVolumes = function () {
    if (!logs || logs.length === 0) return;
    let modified = false;
    logs.forEach(item => {
        if (parseFloat(item.length) > 0 && parseFloat(item.diameter) > 0) {
            item.volume = calculateVolume(item.length, item.diameter);
            modified = true;
        }
    });
    if (modified) {
        save();
        renderAll();
    }
};
