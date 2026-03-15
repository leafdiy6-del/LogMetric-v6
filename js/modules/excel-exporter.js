async function exportData(options = {}) {
    const d = new Date();
    const timeStr = `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}_${d.getHours()}-${d.getMinutes()}`;
    const ds = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
    const t = I18N[currentLang];
    const sourceLogs = Array.isArray(options.logs) ? options.logs : logs;
    const sourceGlobal = (options.global && typeof options.global === 'object') ? options.global : globalInfo;
    const showPriceCsv = !!appSettings.priceEnabled && !!appSettings.showPriceCsv;
    const currencySymbol = getCurrencySymbol();
    const taxPercent = parseFloat(appSettings.taxPercent) || 0;

    const wb = XLSX.utils.book_new();
    const wsData = [];

    // 项目信息 - 横向专业布局
    const showCompany = appSettings.showCompanyInPdf !== false;
    const infoRow1 = [t.p_date, ds];
    if (showCompany && getCompanyName(sourceGlobal.company)) { infoRow1.push(t.my_company, getCompanyName(sourceGlobal.company)); }
    wsData.push(infoRow1);

    const infoRow2 = [];
    if (sourceGlobal.container) { infoRow2.push(t.container, sourceGlobal.container); }
    if (getCompanyName(sourceGlobal.seller)) { infoRow2.push(t.seller, getCompanyName(sourceGlobal.seller)); }
    if (infoRow2.length > 0) wsData.push(infoRow2);

    const infoRow3 = [];
    if (sourceGlobal.description) { infoRow3.push(t.description, sourceGlobal.description); }
    if (sourceGlobal.location) { infoRow3.push(t.location, sourceGlobal.location); }
    if (sourceGlobal.measurer) { infoRow3.push(t.measurer, sourceGlobal.measurer); }
    if (infoRow3.length > 0) wsData.push(infoRow3);

    if (sourceGlobal.note) wsData.push([t.note_global, sourceGlobal.note]);

    wsData.push([]);

    // 表头（包含金额列）
    const headerRow = [t.idx, t.code, t.grade, t.len, t.dia, t.vol];
    if (showPriceCsv) {
        headerRow.push(t.price_amount);
    }
    headerRow.push(t.note);
    wsData.push(headerRow);
    const dataStartRowIndex = wsData.length;

    // 数据行
    const { validLogs, rowCount: excelRowCount, rootCount: excelRootCount, totalV: excelTotalV } = getValidLogsGroupStats(sourceLogs);
    let totalBeforeTax = 0;
    const reversedLogs = [...validLogs].reverse();
    reversedLogs.forEach((l, idx) => {
        const len = parseFloat(cleanInput(l.length.toString())) || '';
        const dia = parseFloat(cleanInput(l.diameter.toString())) || '';
        const vol = l.volume ? parseFloat(formatVolumeForDisplay(l.volume)) : 0;
        const showMarks = appSettings.showMarksInExport !== false;
        const gradeDisp = (l.grade || '-') + (showMarks && l.markGrade ? '↑' : '');
        const lenDisp = len + (showMarks && l.markLen ? '↑' : '');
        const diaDisp = dia + (showMarks && l.markDia ? '↑' : '');

        const dataRow = [idx + 1, l.code || '', gradeDisp, lenDisp, diaDisp, vol];

        // 添加金额列（分组全部计入）
        if (showPriceCsv) {
            const amount = calcLogAmountBeforeTax(l);
            totalBeforeTax += amount;
            dataRow.push(amount ? parseFloat(amount.toFixed(2)) : '');
        }

        dataRow.push(l.note || '');
        wsData.push(dataRow);
    });

    // 统计汇总 - 独立的6列表格（分组全部计入）
    wsData.push([]);
    const totalV = excelTotalV;

    if (showPriceCsv && totalBeforeTax > 0) {
        let avgUnitPrice = totalV > 0 ? totalBeforeTax / totalV : 0;
        const taxAmount = taxPercent > 0 ? totalBeforeTax * (taxPercent / 100) : 0;
        const totalAfterTax = totalBeforeTax + taxAmount;

        const gradeLabel = currentLang === 'sl' ? 'Razred' : (currentLang === 'zh' ? '等级' : 'Grade');
        const kolicinaLabel = currentLang === 'sl' ? 'Količina' : (currentLang === 'zh' ? '数量' : 'Quantity');
        const znesekLabel = currentLang === 'sl' ? 'Znesek' : (currentLang === 'zh' ? '金额' : 'Amount');
        const taxLabel = currentLang === 'sl' ? 'DDV' : (currentLang === 'zh' ? '增值税' : 'TAX');
        const totalLabel = currentLang === 'sl' ? 'Bruto znesek' : (currentLang === 'zh' ? '含税总额' : 'Total with Tax');

        // 计算每个等级的详细信息（分组全部计入）
        const gradeDetails = {};
        validLogs.forEach(l => {
            const g = l.grade || '?';
            if (g === '?' || g === '') return;
            if (!gradeDetails[g]) {
                gradeDetails[g] = { count: 0, volume: 0, amount: 0 };
            }
            gradeDetails[g].count++;
            gradeDetails[g].volume += parseFloat(l.volume) || 0;
            const unitPrice = getUnitPriceForLog(l);
            const v = parseFloat(l.volume) || 0;
            gradeDetails[g].amount += unitPrice * v;
        });

        // 判断是否按等级定价
        const isByGrade = appSettings.priceMode === 'grade';
        const totalLabel2 = currentLang === 'sl' ? 'Skupaj' : (currentLang === 'zh' ? '总计' : 'Total');

        // 5列表格：等级 | 数量 | m³ | 单价(€/m³) | 金额
        // 表头行
        wsData.push([]);
        const unitPriceLabel = currentLang === 'sl' ? `Cena (${currencySymbol}/m³)` : (currentLang === 'zh' ? `单价 (${currencySymbol}/m³)` : `Unit Price (${currencySymbol}/m³)`);
        const headerRow = [gradeLabel, kolicinaLabel, 'm³', unitPriceLabel, znesekLabel];
        wsData.push(headerRow);

        // 如果按等级定价，先显示等级明细（按 gradeLabels 顺序）
        if (isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    wsData.push([
                        g,
                        detail.count,
                        parseFloat(formatVolumeForDisplay(detail.volume)),
                        parseFloat(formatMoney(unitPrice)),
                        `${formatMoney(detail.amount)} ${currencySymbol}`
                    ]);
                }
            });
        }

        // 总计行
        const row1 = [
            totalLabel2,
            excelRootCount,
            parseFloat(formatVolumeForDisplay(totalV)),
            parseFloat(formatMoney(avgUnitPrice)),
            `${formatMoney(totalBeforeTax)} ${currencySymbol}`
        ];
        wsData.push(row1);
        if (excelRowCount !== excelRootCount) {
            wsData.push([t.total_rows, excelRowCount, '', '', '']);
        }

        // 如果有税收
        if (taxPercent > 0) {
            // 税额行
            const row2 = ['', '', '', `${taxLabel} ${taxPercent.toFixed(1)} %`, `${formatMoney(taxAmount)} ${currencySymbol}`];
            wsData.push(row2);

            // 含税总额行
            const row3 = ['', '', '', totalLabel, `${formatMoney(totalAfterTax)} ${currencySymbol}`];
            wsData.push(row3);
        }

        // 如果不是按等级定价，在最后显示等级明细
        if (!isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    wsData.push([
                        g,
                        detail.count,
                        parseFloat(formatVolumeForDisplay(detail.volume)),
                        parseFloat(formatMoney(unitPrice)),
                        `${formatMoney(detail.amount)} ${currencySymbol}`
                    ]);
                }
            });
        }
    } else {
        // 不显示价格时的简单汇总
        wsData.push([t.total_count, excelRootCount, '', t.total_vol, formatVolumeForDisplay(totalV) + ' m³']);
        if (excelRowCount !== excelRootCount) {
            wsData.push([t.total_rows, excelRowCount, '', '', '']);
        }
    }

    // 其他统计
    const exT4 = parseFloat(appSettings.statThresholdL4) || 4;
    const exT25 = parseFloat(appSettings.statThresholdL25) || 2.5;
    const exT30 = parseFloat(appSettings.statThresholdD30) || 30;
    let l4 = 0, l25 = 0, d30 = 0;
    validLogs.forEach(l => {
        const len_m = parseFloat(cleanInput(l.length.toString()));
        const dia_cm = parseFloat(cleanInput(l.diameter.toString()));
        if (len_m < exT4) l4++; if (len_m < exT25) l25++; if (dia_cm < exT30) d30++;
    });
    wsData.push([]);
    if (l4 > 0) wsData.push(['< ' + exT4 + 'm', l4]);
    if (l25 > 0) wsData.push(['< ' + exT25 + 'm', l25]);
    if (d30 > 0) wsData.push(['< ' + exT30 + 'cm', d30]);

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // 分组行样式（xlsx-js-style）：浅蓝背景 + 加粗；码号列左侧蓝框、备注列右侧蓝框
    const showGroup = appSettings.showGroupInExport !== false;
    const numDataCols = 6 + (showPriceCsv ? 1 : 0) + 1;
    const codeCol = 1;
    const noteCol = numDataCols - 1;
    const blueBorder = { style: 'medium', color: { rgb: 'FF1976D2' } };
    const baseStyle = {
        fill: { patternType: 'solid', fgColor: { rgb: 'FFE3F2FD' } },
        font: { bold: true }
    };
    if (showGroup) {
        reversedLogs.forEach((l, idx) => {
            if (!l.groupId) return;
            const row = dataStartRowIndex + idx;
            for (let c = 0; c < numDataCols; c++) {
                const ref = XLSX.utils.encode_cell({ r: row, c });
                if (!ws[ref]) continue;
                const style = { ...baseStyle };
                if (c === codeCol) style.border = { left: blueBorder };
                else if (c === noteCol) style.border = { right: blueBorder };
                ws[ref].s = style;
            }
        });
    }

    // 设置列宽
    const colWidths = [
        { wch: 6 }, { wch: 16 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 12 }
    ];
    if (showPriceCsv) {
        colWidths.push({ wch: 16 }, { wch: 16 });
    }
    colWidths.push({ wch: 25 });
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, 'Oak Logs');

    const fileName = `OakLog_${sourceGlobal.container || 'Export'}_${timeStr}.xlsx`;
    if (options.returnBlob) {
        const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
    XLSX.writeFile(wb, fileName);
}
