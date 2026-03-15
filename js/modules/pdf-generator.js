async function generatePDF(options = {}) {
    const { jsPDF } = window.jspdf;
    const sourceLogs = Array.isArray(options.logs) ? options.logs : logs;
    const sourceGlobal = (options.global && typeof options.global === 'object') ? options.global : globalInfo;

    // 创建一个临时的隐藏容器来渲染PDF内容
    const pdfContainer = document.createElement('div');
    pdfContainer.style.cssText = 'position:absolute;left:-9999px;width:800px;background:white;padding:40px;';
    document.body.appendChild(pdfContainer);

    const t = I18N[currentLang];
    const d = new Date();
    const dateStr = `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}-${d.getDate().toString().padStart(2, '0')}`;
    const timeStr = `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}_${d.getHours()}-${d.getMinutes()}`;

    function escapeHtml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
    function buildCompanyBlockHtml(co, label) {
        const hasAny = COMPANY_FIELDS.some(f => co[f] && String(co[f]).trim());
        if (!hasAny) return '';
        const lines = [];
        if (co.name) lines.push(`<div style="font-weight:bold; font-size:14px; margin-bottom:4px;">${escapeHtml(co.name)}</div>`);
        if (co.address || co.city || co.zip) {
            const addrLine = [co.address, co.city, co.zip].filter(Boolean).join(' ');
            lines.push(`<div style="margin-bottom:4px;">${escapeHtml(addrLine)}</div>`);
        }
        if (co.phone) lines.push(`<div style="margin-bottom:4px;">${t.company_phone}: ${escapeHtml(co.phone)}</div>`);
        if (co.email || co.website) {
            const parts = [co.email, co.website].filter(Boolean).map(s => escapeHtml(s));
            const emailWeb = parts.join('<span style="display:inline-block; width:10px;"></span>');
            lines.push(`<div style="margin-bottom:4px;">${emailWeb}</div>`);
        }
        if (co.taxId) lines.push(`<div style="margin-bottom:4px;">${t.company_taxId}: ${escapeHtml(co.taxId)}</div>`);
        if (co.bank) lines.push(`<div style="margin-bottom:4px;">${t.company_bank}: ${escapeHtml(co.bank)}</div>`);
        return `<div style="font-size:12px; line-height:1.4;"><div style="font-weight:600; margin-bottom:6px; font-size:11px; color:#555;">${escapeHtml(label)}</div>${lines.join('')}</div>`;
    }
    const company = normalizeCompany(sourceGlobal.company);
    const seller = normalizeSeller(sourceGlobal.seller);
    const showCompany = appSettings.showCompanyInPdf !== false;
    const companyBlockHtml = showCompany ? buildCompanyBlockHtml(company, t.my_company) : '';
    const sellerBlockHtml = buildCompanyBlockHtml(seller, (seller.type === 'seller' ? t.party_seller : t.party_buyer));
    const hasCompanyOrSeller = companyBlockHtml || sellerBlockHtml;
    const twoColBlockHtml = hasCompanyOrSeller ? `<table style="width:100%; margin-bottom:16px; border-collapse:collapse;"><tr><td style="width:35%; vertical-align:top; padding-right:16px;">${companyBlockHtml || '&nbsp;'}</td><td style="width:65%; vertical-align:top; padding-left:24px; text-align:right;">${sellerBlockHtml ? `<div style="display:inline-block; text-align:left;">${sellerBlockHtml}</div>` : '&nbsp;'}</td></tr></table>` : '';

    // 构建HTML内容
    let html = `
            <div style="font-family: Arial, sans-serif; color: #000;">
                <h2 style="text-align:center; margin-bottom:20px; font-size:20px; border-bottom:2px solid #000; padding-bottom:10px;">
                    PACKING LOG LIST / SHIPMENT
                </h2>
                ${twoColBlockHtml}
                <table style="width:100%; border-collapse:collapse; margin-bottom:20px; font-size:12px;">
        `;

    // 项目信息（专业双列布局）- 公司/卖方已移至上方独立块
    const infoItems = [];
    const pushInfo = (label, value) => {
        if (value === undefined || value === null) return;
        const v = String(value).trim();
        if (v === '') return;
        infoItems.push([label, v]);
    };
    infoItems.push([t.p_date, dateStr]);
    pushInfo(t.container, sourceGlobal.container);
    pushInfo(t.description, sourceGlobal.description);
    pushInfo(t.location, sourceGlobal.location);
    pushInfo(t.measurer, sourceGlobal.measurer);
    pushInfo(t.note_global, sourceGlobal.note);

    for (let i = 0; i < infoItems.length; i += 2) {
        const left = infoItems[i];
        const right = infoItems[i + 1];
        if (right) {
            html += `
                    <tr>
                        <td style="border:0.6px solid #999; padding:8px; background:#f7f7f7; width:18%; font-weight:600;">${left[0]}</td>
                        <td style="border:0.6px solid #999; padding:8px; width:32%;">${left[1]}</td>
                        <td style="border:0.6px solid #999; padding:8px; background:#f7f7f7; width:18%; font-weight:600;">${right[0]}</td>
                        <td style="border:0.6px solid #999; padding:8px; width:32%;">${right[1]}</td>
                    </tr>
                `;
        } else {
            html += `
                    <tr>
                        <td style="border:0.6px solid #999; padding:8px; background:#f7f7f7; width:18%; font-weight:600;">${left[0]}</td>
                        <td style="border:0.6px solid #999; padding:8px;" colspan="3">${left[1]}</td>
                    </tr>
                `;
        }
    }

    html += `</table>`;

    // 原木数据表格（显示金额列）
    const { validLogs, rowCount, rootCount, totalV: totalVGroup } = getValidLogsGroupStats(sourceLogs);
    const showPricePdf = !!appSettings.priceEnabled && !!appSettings.showPricePdf;
    const currencySymbol = getCurrencySymbol();
    const amountHeadHtml = showPricePdf
        ? `<th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.price_amount}</th>`
        : '';

    const dataTableHeaderHtml = `
            <table style="width:100%; border-collapse:collapse; font-size:11px; margin-bottom:20px; color:#222;">
                <thead>
                    <tr style="background:#e9e9e9;">
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.idx}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.code}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.grade}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.len}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.dia}</th>
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.vol} (m³)</th>
                        ${amountHeadHtml}
                        <th style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:600;">${t.note}</th>
                    </tr>
                </thead>
                <tbody>`;

    const dataRowHtmls = [];
    [...validLogs].reverse().forEach((log, index) => {
        const len = parseFloat(cleanInput(log.length.toString())) || '';
        const dia = parseFloat(cleanInput(log.diameter.toString())) || '';
        const vol = log.volume ? formatVolumeForDisplay(log.volume) : '0';
        const bgColor = index % 2 === 0 ? '#f7f7f7' : '#ffffff';
        const showMarks = appSettings.showMarksInExport !== false;
        const showGroup = appSettings.showGroupInExport !== false;
        const groupBorder = (showGroup && log.groupId) ? 'border-left:4px solid #1976D2; font-weight:600;' : '';
        const amount = calcLogAmountBeforeTax(log);
        const amountCellHtml = showPricePdf
            ? `<td style="border:0.6px solid #999; padding:6px; text-align:center;">${amount ? formatMoney(amount) : ''}</td>`
            : '';
        const gradeDisp = (log.grade || '-') + (showMarks && log.markGrade ? '↑' : '');
        const lenDisp = len + (showMarks && log.markLen ? '↑' : '');
        const diaDisp = dia + (showMarks && log.markDia ? '↑' : '');
        dataRowHtmls.push(`
                <tr style="background:${bgColor};${groupBorder}">
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${index + 1}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${log.code || ''}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${gradeDisp}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${lenDisp}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${diaDisp}</td>
                    <td style="border:0.6px solid #999; padding:6px; text-align:center;">${vol}</td>
                    ${amountCellHtml}
                    <td style="border:0.6px solid #999; padding:6px; text-align:left; font-size:10px;">${log.note || ''}</td>
                </tr>`);
    });

    // 根据头部内容动态调整首页行数：无公司/买方/测量人等时减少留白
    const infoRowCount = Math.ceil(infoItems.length / 2);
    const ROWS_FIRST_CHUNK = hasCompanyOrSeller ? 25 : (infoRowCount >= 3 ? 28 : 35);
    const ROWS_PER_CHUNK = 35;
    const FIRST_PAGE_BOTTOM_SPACER = '<div style="height:56px; min-height:56px;"></div>';

    const headerHtml = html;
    const mainHtmlChunks = [];
    let start = 0;
    let chunkIndex = 0;
    if (dataRowHtmls.length === 0) {
        mainHtmlChunks.push(headerHtml + dataTableHeaderHtml + '</tbody></table>' + FIRST_PAGE_BOTTOM_SPACER + '</div>');
    } else {
        while (start < dataRowHtmls.length) {
            const count = chunkIndex === 0 ? ROWS_FIRST_CHUNK : ROWS_PER_CHUNK;
            const end = Math.min(start + count, dataRowHtmls.length);
            const rowsHtml = dataRowHtmls.slice(start, end).join('');
            if (chunkIndex === 0) {
                mainHtmlChunks.push(headerHtml + dataTableHeaderHtml + rowsHtml + '</tbody></table>' + FIRST_PAGE_BOTTOM_SPACER + '</div>');
            } else {
                mainHtmlChunks.push('<div style="font-family: Arial, sans-serif; color: #000;">' + dataTableHeaderHtml + rowsHtml + '</tbody></table></div>');
            }
            start = end;
            chunkIndex++;
        }
    }

    // 统计汇总（分组全部计入）
    const totalV = totalVGroup;
    const gStats = {};
    const pdfT4 = parseFloat(appSettings.statThresholdL4) || 4;
    const pdfT25 = parseFloat(appSettings.statThresholdL25) || 2.5;
    const pdfT30 = parseFloat(appSettings.statThresholdD30) || 30;
    let l4 = 0, l25 = 0, d30 = 0;
    validLogs.forEach(l => {
        const g = l.grade || '?';
        gStats[g] = (gStats[g] || 0) + 1;
        const len_m = parseFloat(cleanInput(l.length.toString()));
        const dia_cm = parseFloat(cleanInput(l.diameter.toString()));
        if (len_m < pdfT4) l4++;
        if (len_m < pdfT25) l25++;
        if (dia_cm < pdfT30) d30++;
    });

    // 价格汇总 - 独立的6列表格（分组全部计入）
    if (showPricePdf) {
        const taxPercent = parseFloat(appSettings.taxPercent) || 0;
        let totalBeforeTax = 0;
        let avgUnitPrice = 0;
        validLogs.forEach(l => {
            const unit = getUnitPriceForLog(l);
            const v = parseFloat(l.volume) || 0;
            totalBeforeTax += unit * v;
        });
        if (totalV > 0) avgUnitPrice = totalBeforeTax / totalV;

        const taxAmount = taxPercent > 0 ? totalBeforeTax * (taxPercent / 100) : 0;
        const totalAfterTax = totalBeforeTax + taxAmount;

        const skupajLabel = currentLang === 'sl' ? 'Skupaj' : (currentLang === 'zh' ? '总计' : 'Total');
        const taxLabel = currentLang === 'sl' ? 'DDV' : (currentLang === 'zh' ? '增值税' : 'TAX');
        const totalLabel = currentLang === 'sl' ? 'Bruto znesek' : (currentLang === 'zh' ? '含税总额' : 'Total with Tax');

        const gradeLabel = currentLang === 'sl' ? 'Razred' : (currentLang === 'zh' ? '等级' : 'Grade');
        const kolicinaLabel = currentLang === 'sl' ? 'Količina' : (currentLang === 'zh' ? '数量' : 'Quantity');
        const znesekLabel = currentLang === 'sl' ? 'Znesek' : (currentLang === 'zh' ? '金额' : 'Amount');

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
        const unitPriceLabel = currentLang === 'sl' ? `Cena (${currencySymbol}/m³)` : (currentLang === 'zh' ? `单价 (${currencySymbol}/m³)` : `Unit Price (${currencySymbol}/m³)`);

        let priceTableHtml = `
                <div style="font-family: Arial, sans-serif; color: #000;">
                <table style="width:100%; border-collapse:collapse; font-size:11px; margin-top:20px; color:#222;">
                    <thead>
                        <tr style="background:#e9e9e9;">
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${gradeLabel}</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${kolicinaLabel}</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">m³</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${unitPriceLabel}</th>
                            <th style="border:0.6px solid #999; padding:8px; width:20%; text-align:center; font-weight:600;">${znesekLabel}</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

        // 如果按等级定价，先显示等级明细（按 gradeLabels 顺序）
        if (isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    priceTableHtml += `
                            <tr style="background:#f9f9f9;">
                                <td style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:bold;">${g}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${detail.count}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatVolumeForDisplay(detail.volume)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(unitPrice)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(detail.amount)} ${currencySymbol}</td>
                            </tr>
                        `;
                }
            });
        }

        // 总计行
        priceTableHtml += `
                        <tr style="background:#f0f0f0; font-weight:bold; border-top:2px solid #666;">
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${totalLabel2}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${rootCount}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatVolumeForDisplay(totalV)}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(avgUnitPrice)}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(totalBeforeTax)} ${currencySymbol}</td>
                        </tr>
            `;
        if (taxPercent > 0) {
            priceTableHtml += `
                        <tr style="background:#f9f9f9;">
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:right;">${taxLabel} ${taxPercent.toFixed(1)} %</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(taxAmount)} ${currencySymbol}</td>
                        </tr>
                        <tr style="background:#f0f0f0; font-weight:bold;">
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px;"></td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:right;">${totalLabel}</td>
                            <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(totalAfterTax)} ${currencySymbol}</td>
                        </tr>
                `;
        }

        // 如果不是按等级定价，在最后显示等级明细
        if (!isByGrade) {
            const labels = getGradeLabels();
            labels.forEach(g => {
                if (gradeDetails[g]) {
                    const detail = gradeDetails[g];
                    const unitPrice = getPriceForGrade(g);
                    priceTableHtml += `
                            <tr style="background:#f9f9f9;">
                                <td style="border:0.6px solid #999; padding:8px; text-align:center; font-weight:bold;">${g}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${detail.count}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatVolumeForDisplay(detail.volume)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(unitPrice)}</td>
                                <td style="border:0.6px solid #999; padding:8px; text-align:center;">${formatMoney(detail.amount)} ${currencySymbol}</td>
                            </tr>
                        `;
                }
            });
        }

        priceTableHtml += `
                    </tbody>
                </table>
                </div>
            `;
        mainHtmlChunks.push(priceTableHtml);
    }

    // 统计汇总（单独渲染，避免被分页截断时直接整块移到下一页）
    const summaryTitle = currentLang === 'zh' ? '统计汇总' : (currentLang === 'en' ? 'SUMMARY' : 'POVZETEK');
    let summaryHtml = `
            <div style="font-family: Arial, sans-serif; color: #000; width:100%; box-sizing:border-box;">
            <div style="background:#f0f0f0; padding:15px; border:2px solid #000; margin-top:8px; width:100%; box-sizing:border-box; overflow-wrap:break-word; word-wrap:break-word;">
                <h3 style="margin:0 0 10px 0; font-size:14px;">${summaryTitle}</h3>
                <p style="margin:5px 0; font-size:12px;">
                    <strong>${t.total_count}:</strong> ${rowCount !== rootCount ? `${rowCount} (${rootCount})` : validLogs.length} &nbsp;&nbsp;
                    <strong>${t.total_vol}:</strong> <span style="font-weight:bold;">${formatVolumeForDisplay(totalV)} m³</span>
                </p>
        `;
    if (Object.keys(gStats).length > 0) {
        summaryHtml += `<p style="margin:5px 0; font-size:11px;"><strong>Grade:</strong> `;
        for (let g in gStats) {
            if (g !== '?' && g !== '') summaryHtml += `${g}: ${gStats[g]} &nbsp;&nbsp; `;
        }
        summaryHtml += `</p>`;
    }
    if (l4 > 0 || l25 > 0 || d30 > 0) {
        summaryHtml += `<p style="margin:5px 0; font-size:10px; color:#666;">`;
        if (l4 > 0) summaryHtml += `&lt; ${pdfT4}m: ${l4} &nbsp;&nbsp; `;
        if (l25 > 0) summaryHtml += `&lt; ${pdfT25}m: ${l25} &nbsp;&nbsp; `;
        if (d30 > 0) summaryHtml += `&lt; ${pdfT30}cm: ${d30}`;
        summaryHtml += `</p>`;
    }
    summaryHtml += `
                <p style="margin:10px 0 0 0; font-size:9px; color:#999; border-top:1px solid #ccc; padding-top:10px;">
                    LogMetric Pro | ${dateStr}
                </p>
            </div></div>
        `;

    try {
        if (!window.html2canvas) {
            await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
        }

        const doc = new jsPDF('p', 'mm', 'a4');
        const imgWidth = 210;
        const pageHeight = 297;
        const bottomMargin = 12;
        const contentHeight = pageHeight - bottomMargin;
        let currentY = 0;
        let position = 0;

        function addPageBottomMargin() {
            doc.setFillColor(255, 255, 255);
            doc.rect(0, contentHeight, imgWidth, bottomMargin, 'F');
        }

        for (const chunkHtml of mainHtmlChunks) {
            pdfContainer.innerHTML = chunkHtml;
            const canvas = await html2canvas(pdfContainer, {
                scale: 2,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff'
            });
            const chunkImgHeight = (canvas.height * imgWidth) / canvas.width;
            const imgData = canvas.toDataURL('image/png');
            if (currentY + chunkImgHeight > contentHeight && currentY > 0) {
                addPageBottomMargin();
                doc.addPage();
                currentY = 0;
            }
            doc.addImage(imgData, 'PNG', 0, currentY, imgWidth, chunkImgHeight);
            currentY += chunkImgHeight;
        }
        addPageBottomMargin();

        // 统计汇总单独渲染：若当前页剩余空间不足则整块移到下一页
        const summaryContainer = document.createElement('div');
        summaryContainer.style.cssText = 'position:absolute;left:-9999px;width:760px;background:white;padding:8px 40px 40px 40px;box-sizing:border-box;overflow:hidden;';
        document.body.appendChild(summaryContainer);
        summaryContainer.innerHTML = summaryHtml;
        const canvasSummary = await html2canvas(summaryContainer, {
            scale: 2,
            useCORS: true,
            logging: false,
            backgroundColor: '#ffffff'
        });
        document.body.removeChild(summaryContainer);

        const imgHeightSummary = (canvasSummary.height * imgWidth) / canvasSummary.width;
        const imgDataSummary = canvasSummary.toDataURL('image/png');
        const spaceLeftOnLastPage = contentHeight - currentY;

        if (spaceLeftOnLastPage < imgHeightSummary && currentY > 0) {
            doc.addPage();
            doc.addImage(imgDataSummary, 'PNG', 0, 0, imgWidth, imgHeightSummary);
        } else {
            doc.addImage(imgDataSummary, 'PNG', 0, currentY, imgWidth, imgHeightSummary);
        }
        addPageBottomMargin();

        const fileName = `OakLog_${sourceGlobal.container || 'Export'}_${timeStr}.pdf`;
        if (options.returnBlob) {
            const blob = doc.output('blob');
            return blob;
        }
        doc.save(fileName);

    } catch (error) {
        alert('PDF generation failed. Please try again.');
    } finally {
        document.body.removeChild(pdfContainer);
    }
}

// 辅助函数：动态加载脚本
function loadScript(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
}
