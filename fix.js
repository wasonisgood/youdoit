// 主要的 PDF 生成函數
async function generatePDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 10;
    let yOffset = margin;

    // 添加標題
    doc.setFontSize(20);
    doc.text("中央政府預算宣導經費報告", margin, yOffset);
    yOffset += 10;

    // 獲取所有年份
    const years = [...new Set(budgetData.map(item => item['季度'].split('-')[0]))].sort();

    // 為每個年份生成數據
    for (const year of years) {
        const yearData = budgetData.filter(item => item['季度'].startsWith(year));

        // 添加年份標題
        doc.setFontSize(16);
        doc.text(`${year}年度分析`, margin, yOffset);
        yOffset += 10;

        // 添加年度數據表格
        yOffset = await addTableToPDF(doc, createYearDataTable(yearData), margin, yOffset);

        // 添加年度圖表
        yOffset = await addChartToPDF(doc, createAgencyBudgetChart(yearData), '機關預算分佈', margin, yOffset);
        yOffset = await addChartToPDF(doc, createMediaTypeChart(yearData), '媒體類型預算分佈', margin, yOffset);
        yOffset = await addChartToPDF(doc, createPublishTargetChart(yearData), '刊登或託播對象分佈', margin, yOffset);

        // 檢查頁面空間，如果需要就添加新頁面
        if (yOffset > pageHeight - 20) {
            doc.addPage();
            yOffset = margin;
        }
    }

    // 添加 LINE 相關數據表格
    doc.addPage();
    yOffset = margin;
    doc.setFontSize(16);
    doc.text("LINE 相關宣導數據", margin, yOffset);
    yOffset += 10;
    yOffset = await addTableToPDF(doc, createLineDataTable(), margin, yOffset);

    // 添加 YouTube 相關數據表格
    if (yOffset > pageHeight - 60) {
        doc.addPage();
        yOffset = margin;
    }
    doc.setFontSize(16);
    doc.text("YouTube 相關宣導數據", margin, yOffset);
    yOffset += 10;
    yOffset = await addTableToPDF(doc, createYoutubeDataTable(), margin, yOffset);

    // 添加總覽圖表
    doc.addPage();
    yOffset = margin;
    doc.setFontSize(16);
    doc.text("總覽圖表", margin, yOffset);
    yOffset += 10;
    yOffset = await addChartToPDF(doc, createAgencyBudgetChart(budgetData), '總機關預算分佈', margin, yOffset);
    yOffset = await addChartToPDF(doc, createMediaTypeChart(budgetData), '總媒體類型預算分佈', margin, yOffset);
    yOffset = await addChartToPDF(doc, createPublishTargetChart(budgetData), '總刊登或託播對象分佈', margin, yOffset);

    // 添加年度增長折線圖
    doc.addPage();
    yOffset = margin;
    doc.setFontSize(16);
    doc.text("年度增長趨勢", margin, yOffset);
    yOffset += 10;
    yOffset = await addChartToPDF(doc, createYearlyGrowthChart(), '年度預算增長趨勢', margin, yOffset);

    // 保存 PDF
    doc.save("budget_report.pdf");
}

// 輔助函數：創建年度數據表格
function createYearDataTable(yearData) {
    const headers = ['機關名稱', '執行金額', '媒體類型', '刊登或託播對象'];
    const rows = yearData.map(item => [
        item['機關名稱'],
        Number(item['執行金額']).toLocaleString(),
        item['媒體類型'],
        item['刊登或託播對象']
    ]);
    return { headers, rows };
}

// 輔助函數：創建 LINE 數據表格
function createLineDataTable() {
    const lineData = budgetData.filter(item => 
        item['宣導項目、標題及內容'].toLowerCase().includes('line')
    );
    const headers = ['宣導標題', '執行金額', '委託廠商'];
    const rows = lineData.map(item => [
        item['宣導項目、標題及內容'],
        Number(item['執行金額']).toLocaleString(),
        item['執行單位']
    ]);
    return { headers, rows };
}

// 輔助函數：創建 YouTube 數據表格
function createYoutubeDataTable() {
    const youtubeData = budgetData.filter(item => 
        item['刊登或託播對象'].toLowerCase().includes('youtube') ||
        item['媒體類型'].toLowerCase().includes('youtube')
    );
    const headers = ['委託廠商', '宣導標題', '執行金額'];
    const rows = youtubeData.map(item => [
        item['執行單位'],
        item['宣導項目、標題及內容'],
        Number(item['執行金額']).toLocaleString()
    ]);
    return { headers, rows };
}

// 輔助函數：將表格添加到 PDF
async function addTableToPDF(doc, tableData, x, y) {
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);

    const { headers, rows } = tableData;
    const cellWidth = (doc.internal.pageSize.width - 2 * x) / headers.length;
    const cellHeight = 10;

    // 繪製表頭
    headers.forEach((header, i) => {
        doc.setFillColor(200, 200, 200);
        doc.rect(x + i * cellWidth, y, cellWidth, cellHeight, 'F');
        doc.text(header, x + i * cellWidth + 2, y + cellHeight - 2);
    });
    y += cellHeight;

    // 繪製數據行
    rows.forEach((row, rowIndex) => {
        row.forEach((cell, cellIndex) => {
            doc.setFillColor(255, 255, 255);
            doc.rect(x + cellIndex * cellWidth, y, cellWidth, cellHeight, 'F');
            doc.text(cell.toString(), x + cellIndex * cellWidth + 2, y + cellHeight - 2);
        });
        y += cellHeight;

        // 如果到達頁面底部，添加新頁面
        if (y > doc.internal.pageSize.height - 20) {
            doc.addPage();
            y = 20;
        }
    });

    return y + 10; // 返回新的 y 偏移量
}

// 輔助函數：將圖表添加到 PDF
async function addChartToPDF(doc, chartData, title, x, y) {
    const chartWidth = doc.internal.pageSize.width - 2 * x;
    const chartHeight = 60;

    // 添加標題
    doc.setFontSize(12);
    doc.text(title, x, y);
    y += 10;

    // 創建圓餅圖
    const pieChartCanvas = await createChartCanvas(chartData, 'pie');
    doc.addImage(pieChartCanvas, 'PNG', x, y, chartWidth / 2 - 5, chartHeight);

    // 創建柱狀圖
    const barChartCanvas = await createChartCanvas(chartData, 'bar');
    doc.addImage(barChartCanvas, 'PNG', x + chartWidth / 2 + 5, y, chartWidth / 2 - 5, chartHeight);

    y += chartHeight + 10;

    // 添加表格形式的數據
    const tableData = {
        headers: ['項目', '金額', '百分比'],
        rows: chartData.labels.map((label, index) => [
            label,
            chartData.datasets[0].data[index].toLocaleString(),
            `${((chartData.datasets[0].data[index] / chartData.datasets[0].data.reduce((a, b) => a + b, 0)) * 100).toFixed(2)}%`
        ])
    };

    y = await addTableToPDF(doc, tableData, x, y);

    return y + 10; // 返回新的 y 偏移量
}

// 輔助函數：創建圖表 Canvas
async function createChartCanvas(chartData, type) {
    const canvas = document.createElement('canvas');
    canvas.width = 400;
    canvas.height = 300;
    const ctx = canvas.getContext('2d');

    new Chart(ctx, {
        type: type,
        data: chartData,
        options: {
            responsive: false,
            plugins: {
                legend: {
                    display: false
                }
            }
        }
    });

    return canvas;
}

// 輔助函數：創建年度增長折線圖
function createYearlyGrowthChart() {
    const yearlyTotals = {};
    budgetData.forEach(item => {
        const year = item['季度'].split('-')[0];
        const amount = Number(item['執行金額']);
        yearlyTotals[year] = (yearlyTotals[year] || 0) + amount;
    });

    const years = Object.keys(yearlyTotals).sort();
    const totals = years.map(year => yearlyTotals[year]);

    return {
        labels: years,
        datasets: [{
            label: '年度總預算',
            data: totals,
            borderColor: 'rgb(75, 192, 192)',
            tension: 0.1
        }]
    };
}

// 修改後的機關預算圖表函數
function createAgencyBudgetChart(data, type = 'pie') {
    const agencyBudgets = {};
    data.forEach(item => {
        const agency = item['機關名稱'];
        const amount = Number(item['執行金額']);
        agencyBudgets[agency] = (agencyBudgets[agency] || 0) + amount;
    });

    const sortedData = Object.entries(agencyBudgets)
        .sort((a, b) => b[1] - a[1])
        .reduce((obj, [key, value]) => ({ ...obj, [key]: value }), {});

    const chartData = {
        labels: Object.keys(sortedData),
        datasets: [{
            label: '執行金額',
            data: Object.values(sortedData),
            backgroundColor: [
                '#ff6384', '#36a2eb', '#ffce56', '#4bc0c0', '#9966ff',
                '#ff9f40', '#ffcd56', '#4bc0c0', '#36a2eb', '#ff6384'
            ]
        }]
    };

    const options = {
        responsive: true,
        plugins: {
            title: {
                display: true,
                text: '各機關預算執行金額'
            },
            tooltip: {
                callbacks: {
                    label: function (context) {
                        const label = context.label || '';
                        const value = context.raw;
                        const total = context.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                        const percentage = total > 0 ? ((value / total) * 100).toFixed(2) : 0;
                        return `${label}: ${value.toLocaleString()} (${percentage}%)`;
                    }
                }
            }
        }
    };

    if (type === 'bar') {
        options.indexAxis = 'y';
    }

    return { type, data: chartData, options };
}

// 修改後的媒體類型圖表函數
function createMediaTypeChart(data, type = 'pie') {
    const mediaTypes = {};
    data.forEach(item => {
        const mediaType = item['媒體類型'];
        const amount = Number(item['執行金額']);
        if (!mediaTypes[mediaType]) {
            mediaTypes[mediaType] = { total: 0, count: 0 };
        }
        mediaTypes[mediaType].total += amount;
        mediaTypes[mediaType].count++;
    });

    const sortedData = Object.entries(mediaTypes)
        .sort((a, b) => b[1].total - a[1].total)
        .reduce((obj, [key, value]) => ({ ...obj, [key]: value }), {});

    const chartData = {
        labels: Object.keys(sortedData),
        datasets: [{
            label: '執行金額',
            data: Object.values(sortedData).map(v => v.total),
            backgroundColor: [
                '#ff6384', '#36a2eb', '#ffce56', '#4bc0c0', '#9966ff',
                '#ff9f40', '#ffcd56', '#4bc0c0', '#36a2eb', '#ff6384'
            ]
        }]
    };

    const options = {
        responsive: true,
        plugins: {
            title: {
                display: true,
                text: '媒體類型預算分布'
            },
            tooltip: {
                callbacks: {
                    label: function (context) {
                        const label = context.label || '';
                        const value = context.raw;
                        const count = sortedData[label].count;
                        const total = context.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                        const percentage = total > 0 ? ((value / total) * 100).toFixed(2) : 0;
                        return `${label}: ${value.toLocaleString()} (${percentage}%, 數量: ${count})`;
                    }
                }
            }
        }
    };

    if (type === 'bar') {
        options.indexAxis = 'y';
    }

    return { type, data: chartData, options };
}

// 修改後的刊登或託播對象圖表函數
function createPublishTargetChart(data, type = 'pie') {
    const publishTargets = {};

    data.forEach(item => {
        const targets = splitPublishTargets(item['刊登或託播對象']);
        const amount = Number(item['執行金額']);
        const amountPerTarget = amount / targets.length;

        targets.forEach(target => {
            target = normalizeChannelName(target);
            if (!publishTargets[target]) {
                publishTargets[target] = { total: 0, count: 0 };
            }
            publishTargets[target].total += amountPerTarget;
            publishTargets[target].count++;
        });
    });

    const sortedData = Object.entries(publishTargets)
        .sort((a, b) => b[1].total - a[1].total)
        .reduce((obj, [key, value]) => ({ ...obj, [key]: value }), {});

    const chartData = {
        labels: Object.keys(sortedData),
        datasets: [{
            label: '執行金額',
            data: Object.values(sortedData).map(v => v.total),
            backgroundColor: [
                '#ff6384', '#36a2eb', '#ffce56', '#4bc0c0', '#9966ff',
                '#ff9f40', '#ffcd56', '#4bc0c0', '#36a2eb', '#ff6384'
            ]
        }]
    };

    const options = {
        responsive: true,
        plugins: {
            title: {
                display: true,
                text: '刊登或託播對象分布'
            },
            tooltip: {
                callbacks: {
                    label: function (context) {
                        const label = context.label || '';
                        const value = context.raw;
                        const count = sortedData[label].count;
                        const total = context.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                        const percentage = ((value / total) * 100).toFixed(2);
                        return `${label}: ${value.toLocaleString()} (${percentage}%, 數量: ${count})`;
                    }
                }
            }
        }
    };

    if (type === 'bar') {
        options.indexAxis = 'y';
    }

    return { type, data: chartData, options };
}


// 新增年度增長折線圖函數
function createYearlyGrowthChart() {
    const yearlyTotals = {};
    budgetData.forEach(item => {
        const year = item['季度'].split('-')[0];
        const amount = Number(item['執行金額']);
        yearlyTotals[year] = (yearlyTotals[year] || 0) + amount;
    });

    const years = Object.keys(yearlyTotals).sort();
    const totals = years.map(year => yearlyTotals[year]);

    const chartData = {
        labels: years,
        datasets: [{
            label: '年度總預算',
            data: totals,
            borderColor: 'rgb(75, 192, 192)',
            tension: 0.1
        }]
    };

    const options = {
        responsive: true,
        plugins: {
            title: {
                display: true,
                text: '年度預算增長趨勢'
            },
            tooltip: {
                callbacks: {
                    label: function(context) {
                        const label = context.dataset.label || '';
                        const value = context.parsed.y;
                        return `${label}: ${value.toLocaleString()}`;
                    }
                }
            }
        },
        scales: {
            y: {
                beginAtZero: true,
                title: {
                    display: true,
                    text: '預算金額'
                }
            },
            x: {
                title: {
                    display: true,
                    text: '年度'
                }
            }
        }
    };

    return { type: 'line', data: chartData, options };
}