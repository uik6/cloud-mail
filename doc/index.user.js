// ==UserScript==
// @name         物流订单上网率统计助手v8.3
// @namespace    http://tampermonkey.net/
// @version      8.5
// @description  统计OMP物流上网率，支持Excel导出，含五大多维ECharts看板，悬浮按钮支持自由拖拽与开关切换
// @author       AI Assistant
// @match        *://*.xlwms.com/*
// @updateURL    https://github.com/uik6/cloud-mail/raw/refs/heads/main/doc/index.user.js
// @downloadURL  https://github.com/uik6/cloud-mail/raw/refs/heads/main/doc/index.user.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/echarts/5.4.3/echarts.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
// @grant        none
// @run-at       document-start
// ==/UserScript==


(function() {
    'use strict';

    const ENABLE_ANALYSIS_DEV_MODE = true;
    const ANALYSIS_DEV_CACHE_PREFIX = 'sr_dev_cache_v1';

    function paginateRows(rows, page, pageSize) {
        const safePageSize = Math.max(1, Number(pageSize) || 50);
        const total = rows.length;
        const totalPages = Math.max(1, Math.ceil(total / safePageSize));
        const currentPage = Math.min(Math.max(1, Number(page) || 1), totalPages);
        const startIndex = (currentPage - 1) * safePageSize;
        return {
            pageRows: rows.slice(startIndex, startIndex + safePageSize),
            total,
            totalPages,
            currentPage,
            pageSize: safePageSize
        };
    }

    function pickSkuField(row, keys, fallback = '') {
        for (const key of keys) {
            if (row && row[key] !== undefined) return row[key];
        }
        return fallback;
    }

    function normalizeSkuInventoryRow(row, dim) {
        const normalized = {
            customerName: pickSkuField(row, ['customerName', '客户名称'], ''),
            customerCode: pickSkuField(row, ['customerCode', '客户编码'], ''),
            whName: pickSkuField(row, ['whName', '发货仓库'], ''),
            whCode: pickSkuField(row, ['whCode', '仓库代码'], ''),
            sku: pickSkuField(row, ['sku', 'SKU'], ''),
            productName: pickSkuField(row, ['productName', '产品名称'], ''),
            skuCount: Number(pickSkuField(row, ['skuCount', 'SKU数'], 0)),
            warehouseCount: Number(pickSkuField(row, ['warehouseCount', '仓库数'], 0)),
            customerCount: Number(pickSkuField(row, ['customerCount', '客户数'], 0)),
            preStockQty: Number(pickSkuField(row, ['preStockQty', '期初库存'], 0)),
            closeStockQty: Number(pickSkuField(row, ['closeStockQty', '期末库存'], 0)),
            outboundBookQty: Number(pickSkuField(row, ['outboundBookQty', '出库预占'], 0)),
            stockTurnoverRate: Number(pickSkuField(row, ['stockTurnoverRate', '库存周转率'], 0)),
            stockTurnoverDays: Number(pickSkuField(row, ['stockTurnoverDays', '库存周转天数'], 0)),
            stockSaleRate: Number(pickSkuField(row, ['stockSaleRate', '库存售罄率'], 0))
        };
        return dim ? { ...row, ...normalized, __dim: dim } : { ...row, ...normalized };
    }

    function normalizeSkuInventoryReport(report) {
        if (!report) return report;
        const detailRows = (report.detailRows || []).map((row) => normalizeSkuInventoryRow(row, 'detail'));
        const customerRows = (report.customerRows || []).map((row) => normalizeSkuInventoryRow(row, 'customer'));
        const skuRows = (report.skuRows || []).map((row) => normalizeSkuInventoryRow(row, 'sku'));
        return {
            ...report,
            detailRows,
            customerRows,
            skuRows,
            topSkuRow: report.topSkuRow ? normalizeSkuInventoryRow(report.topSkuRow, 'sku') : null
        };
    }

    function getSkuInventoryDimensionMeta() {
        const dim = document.getElementById('sr-sku-inventory-dim')?.value || 'detail';
        if (!skuInventoryReportData) return { dim, rows: [] };
        if (dim === 'customer') return { dim, rows: skuInventoryReportData.customerRows };
        if (dim === 'sku') return { dim, rows: skuInventoryReportData.skuRows };
        return { dim, rows: skuInventoryReportData.detailRows };
    }

    function getSkuInventoryRowLabel(row, dim) {
        if (dim === 'customer') return row["客户名称"];
        if (dim === 'sku') return row["SKU"];
        return `${row["客户名称"]} / ${row["SKU"]}`;
    }

    function renderSkuInventoryCustomerCharts() {
        const container = document.getElementById('sr-sku-inventory-customer-charts');
        if (!container) return;

        container.innerHTML = '';
        charts.skuInventoryCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.skuInventoryCustomerCharts = [];

        if (!window.echarts || !skuInventoryReportData) return;

        const detailGroupMap = {};
        skuInventoryReportData.detailRows.forEach((row) => {
            const customerKey = `${row.customerCode || ''}__${row.customerName || ''}`;
            if (!detailGroupMap[customerKey]) detailGroupMap[customerKey] = [];
            detailGroupMap[customerKey].push(row);
        });

        const groups = skuInventoryReportData.customerRows
            .slice()
            .sort((a, b) => (a.customerName || '').localeCompare(b.customerName || '', 'zh-Hans-CN') || (a.customerCode || '').localeCompare(b.customerCode || ''))
            .map((customerRow) => {
                const customerKey = `${customerRow.customerCode || ''}__${customerRow.customerName || ''}`;
                return { customerRow, rows: detailGroupMap[customerKey] || [] };
            })
            .filter((group) => group.rows.length > 0);

        if (groups.length === 0) {
            container.innerHTML = '<div style="color:#999; padding:12px 4px;">暂无客户图表数据</div>';
            return;
        }

        const renderQueue = [];
        groups.forEach(({ customerRow, rows }, index) => {
            const card = document.createElement('div');
            card.className = 'sr-inventory-card';
            const pieId = `sr-sku-inventory-customer-pie-${index}`;
            const barId = `sr-sku-inventory-customer-bar-${index}`;
            card.innerHTML = `
                <div class="sr-inventory-card-title">${customerRow.customerName || '-'}</div>
                <div class="sr-inventory-card-meta">SKU数 ${customerRow.skuCount || 0} ｜ 期末库存 ${customerRow.closeStockQty || 0}</div>
                <div class="sr-inventory-card-charts">
                    <div id="${pieId}" class="sr-inventory-card-box-half"></div>
                    <div id="${barId}" class="sr-inventory-card-box-half"></div>
                </div>
            `;
            container.appendChild(card);
            renderQueue.push({ pieId, barId, rows });
        });

        requestAnimationFrame(() => {
            renderQueue.forEach(({ pieId, barId, rows }) => {
                const pieEl = document.getElementById(pieId);
                const barEl = document.getElementById(barId);
                if (!pieEl || !barEl) return;

                const sortByMetricDesc = (list, field) => list
                    .slice()
                    .sort((a, b) => {
                        const aValue = Number(a[field] || 0);
                        const bValue = Number(b[field] || 0);
                        const aHasValue = aValue > 0 ? 1 : 0;
                        const bHasValue = bValue > 0 ? 1 : 0;
                        if (bHasValue !== aHasValue) return bHasValue - aHasValue;
                        return bValue - aValue || b.closeStockQty - a.closeStockQty || a.sku.localeCompare(b.sku, 'zh-Hans-CN');
                    });

                const turnoverRateRows = sortByMetricDesc(rows, 'stockTurnoverRate')
                    .filter((row) => Number(row.stockTurnoverRate || 0) > 0)
                    .slice(0, SKU_INVENTORY_CHART_LIMIT);
                const turnoverDaysRows = sortByMetricDesc(rows, 'stockTurnoverDays')
                    .filter((row) => Number(row.stockTurnoverDays || 0) > 0)
                    .slice(0, SKU_INVENTORY_CHART_LIMIT);

                const pieChart = echarts.init(pieEl);
                pieChart.setOption({
                    title: { text: `Top ${SKU_INVENTORY_CHART_LIMIT} SKU库存周转率`, left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'item', formatter: (params) => `${params.name}<br/>周转率: ${params.value}%<br/>占比: ${params.percent}%` },
                    legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 10 },
                    series: [{
                        type: 'pie',
                        radius: ['35%', '65%'],
                        center: ['38%', '50%'],
                        data: rows.map((row) => ({ name: row["SKU"], value: Number(row["库存周转率"] || 0) })).filter((item) => item.value > 0),
                        label: { formatter: '{b}\n{d}%' }
                    }]
                }, true);
                charts.skuInventoryCustomerCharts.push(pieChart);

                const barChart = echarts.init(barEl);
                barChart.setOption({
                    title: { text: 'SKU库存周转天数', left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                    grid: { left: '8%', right: '6%', bottom: 72, top: 48, containLabel: true },
                    xAxis: {
                        type: 'category',
                        data: rows.map((row) => row["SKU"]),
                        axisLabel: { interval: 0, rotate: 28, margin: 14 }
                    },
                    yAxis: { type: 'value', name: '天数' },
                    series: [{
                        name: '库存周转天数',
                        type: 'bar',
                        data: rows.map((row) => Number(row["库存周转天数"] || 0)),
                        itemStyle: { color: '#91cc75' },
                        barMaxWidth: 42,
                        label: { show: true, position: 'top', formatter: ({ value }) => value == null ? '' : value }
                    }]
                }, true);
                charts.skuInventoryCustomerCharts.push(barChart);
            });
        });
    }

    function renderSkuInventoryAnalysisTable() {
        const thead = document.querySelector('#sr-sku-inventory-table thead');
        const tbody = document.querySelector('#sr-sku-inventory-table tbody');
        const { dim, rows } = getSkuInventoryDimensionMeta();
        const pagination = paginateRows(rows, skuInventoryTableState.page, skuInventoryTableState.pageSize);
        skuInventoryTableState.page = pagination.currentPage;
        const pageRows = pagination.pageRows;

        if (dim === 'customer') {
            thead.innerHTML = '<tr><th>客户名称</th><th>SKU数</th><th>仓库数</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>';
        } else if (dim === 'sku') {
            thead.innerHTML = '<tr><th>SKU</th><th>产品名称</th><th>客户数</th><th>仓库数</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>';
        } else {
            thead.innerHTML = '<tr><th>客户名称</th><th>发货仓库</th><th>SKU</th><th>产品名称</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>';
        }

        tbody.innerHTML = '';
        if (!skuInventoryReportData || rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="${dim === 'customer' ? 9 : 10}" style="color:#999; padding:30px 0;">暂无SKU库存分析结果</td></tr>`;
            document.getElementById('sr-sku-inventory-page-info').innerText = '共 0 条';
            document.getElementById('sr-sku-inventory-prev').disabled = true;
            document.getElementById('sr-sku-inventory-next').disabled = true;
            syncPageJumpControl('sr-sku-inventory-page-jump', 'sr-sku-inventory-page-go', 1, 1, true);
            return;
        }

        pageRows.forEach((row) => {
            const tr = document.createElement('tr');
            if (dim === 'customer') {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row["客户名称"]}</td><td>${row["SKU数"]}</td><td>${row["仓库数"]}</td><td>${row["期初库存"]}</td><td>${row["期末库存"]}</td><td>${row["出库预占"]}</td><td>${row["库存周转率"]}%</td><td>${row["库存周转天数"]}</td><td>${row["库存售罄率"]}</td>`;
            } else if (dim === 'sku') {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row["SKU"]}</td><td style="text-align:left;">${row["产品名称"]}</td><td>${row["客户数"]}</td><td>${row["仓库数"]}</td><td>${row["期初库存"]}</td><td>${row["期末库存"]}</td><td>${row["出库预占"]}</td><td>${row["库存周转率"]}%</td><td>${row["库存周转天数"]}</td><td>${row["库存售罄率"]}</td>`;
            } else {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row["客户名称"]}</td><td>${row["发货仓库"]}</td><td>${row["SKU"]}</td><td style="text-align:left;">${row["产品名称"]}</td><td>${row["期初库存"]}</td><td>${row["期末库存"]}</td><td>${row["出库预占"]}</td><td>${row["库存周转率"]}%</td><td>${row["库存周转天数"]}</td><td>${row["库存售罄率"]}</td>`;
            }
            tbody.appendChild(tr);
        });

        document.getElementById('sr-sku-inventory-page-info').innerText =
            `第 ${pagination.currentPage}/${pagination.totalPages} 页，共 ${pagination.total} 条`;
        document.getElementById('sr-sku-inventory-prev').disabled = pagination.currentPage <= 1;
        document.getElementById('sr-sku-inventory-next').disabled = pagination.currentPage >= pagination.totalPages;
    }

    function renderSkuInventoryAnalysisCharts() {
        if (!window.echarts || !skuInventoryReportData) return;

        const { dim, rows } = getSkuInventoryDimensionMeta();
        const rankedRows = rows
            .slice()
            .sort((a, b) => b["期末库存"] - a["期末库存"] || b["库存周转率"] - a["库存周转率"]);
        const stockRows = rankedRows.slice(0, SKU_INVENTORY_CHART_LIMIT);

        if (!charts.skuInventoryPie) charts.skuInventoryPie = echarts.init(document.getElementById('chart-sku-inventory-pie'));
        charts.skuInventoryPie.setOption({
            title: { text: dim === 'sku' ? `Top ${SKU_INVENTORY_CHART_LIMIT} SKU库存对比` : `Top ${SKU_INVENTORY_CHART_LIMIT} 库存对比`, left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            legend: { top: 28 },
            grid: { left: '3%', right: '3%', bottom: '10%', top: 72, containLabel: true },
            xAxis: { type: 'category', data: stockRows.map((row) => getSkuInventoryRowLabel(row, dim)), axisLabel: { interval: 0, rotate: 18 } },
            yAxis: { type: 'value', name: '库存' },
            series: [
                { name: '期初库存', type: 'bar', data: stockRows.map((row) => Number(row["期初库存"] || 0)), itemStyle: { color: '#69c0ff' }, barMaxWidth: 28 },
                { name: '期末库存', type: 'bar', data: stockRows.map((row) => Number(row["期末库存"] || 0)), itemStyle: { color: '#722ed1' }, barMaxWidth: 28 }
            ]
        }, true);

        const barRows = rows
            .slice()
            .sort((a, b) => Number(b.stockTurnoverDays || 0) - Number(a.stockTurnoverDays || 0) || Number(b.closeStockQty || 0) - Number(a.closeStockQty || 0))
            .slice(0, SKU_INVENTORY_CHART_LIMIT);
        if (!charts.skuInventoryBar) charts.skuInventoryBar = echarts.init(document.getElementById('chart-sku-inventory-bar'));
        charts.skuInventoryBar.setOption({
            title: { text: dim === 'sku' ? `Top ${SKU_INVENTORY_CHART_LIMIT} SKU周转天数` : `Top ${SKU_INVENTORY_CHART_LIMIT} 周转天数`, left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            grid: { left: '3%', right: '3%', bottom: '10%', containLabel: true },
            xAxis: {
                type: 'category',
                data: barRows.map((row) => getSkuInventoryRowLabel(row, dim)),
                axisLabel: { interval: 0, rotate: 18 }
            },
            yAxis: { type: 'value', name: '天数' },
            series: [{
                name: '库存周转天数',
                type: 'bar',
                data: barRows.map((row) => Number(row["库存周转天数"] || 0)),
                itemStyle: { color: '#b37feb' }
            }]
        }, true);
    }

    function renderSkuInventoryAnalysisView() {
        if (!skuInventoryReportData) return;

        document.getElementById('sib-customer-count').innerText = skuInventoryReportData.customerCount;
        document.getElementById('sib-sku-count').innerText = skuInventoryReportData.skuCount;
        document.getElementById('sib-detail-count').innerText = skuInventoryReportData.detailCount;
        document.getElementById('sib-close-stock').innerText = skuInventoryReportData.totalCloseStockQty;

        const topSkuText = skuInventoryReportData.topSkuRow
            ? `${skuInventoryReportData.topSkuRow["SKU"]}（${skuInventoryReportData.topSkuRow["期末库存"]}）`
            : '-';
        document.getElementById('sr-sku-inventory-summary').innerText =
            `${skuInventoryReportData.startStr} 至 ${skuInventoryReportData.endStr}，刷新日期 ${skuInventoryReportData.refreshTime || '-'}，总周转率 ${skuInventoryReportData.totalTurnoverRate}% ，总周转天数 ${skuInventoryReportData.totalTurnoverDays}，库存最高SKU：${topSkuText}`;

        renderSkuInventoryAnalysisCharts();
        renderSkuInventoryCustomerCharts();
        renderSkuInventoryAnalysisTable();
    }

    function resetSkuInventoryAnalysisView() {
        document.getElementById('sib-customer-count').innerText = '-';
        document.getElementById('sib-sku-count').innerText = '-';
        document.getElementById('sib-detail-count').innerText = '-';
        document.getElementById('sib-close-stock').innerText = '-';
        document.getElementById('sr-sku-inventory-summary').innerText = '请先点击【开始SKU库存分析】。';
        document.querySelector('#sr-sku-inventory-table thead').innerHTML = '';
        document.querySelector('#sr-sku-inventory-table tbody').innerHTML = '<tr><td colspan="10" style="color:#999; padding:30px 0;">暂无SKU库存分析结果</td></tr>';
        document.getElementById('sr-sku-inventory-page-info').innerText = '共 0 条';
        if (charts.skuInventoryPie) charts.skuInventoryPie.clear();
        if (charts.skuInventoryBar) charts.skuInventoryBar.clear();
        const customerCharts = document.getElementById('sr-sku-inventory-customer-charts');
        if (customerCharts) customerCharts.innerHTML = '';
        charts.skuInventoryCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.skuInventoryCustomerCharts = [];
    }

    async function computeSkuInventoryAnalysisReport(startStr, endStr, skipWeekends, progress) {
        const startTime = `${startStr} 00:00:00`;
        const endTime = `${endStr} 23:59:59`;
        const records = [];
        let current = 1;
        let pages = 1;
        let refreshTime = '';
        let totalData = null;

        while (current <= pages) {
            if (progress) progress(current, pages, records.length);
            const data = await fetchStockAnalysisPage(startTime, endTime, current, { aggregateMode: 1 });
            const pageData = data.page || {};
            const pageRecords = (pageData.records || []).filter((record) =>
                !skipWeekends || !isWeekendDate(record.statisticDate)
            );
            const pageSize = Number(pageData.size || 500);
            const pageTotal = Number(pageData.total || 0);
            records.push(...pageRecords);
            refreshTime = data.refreshTime || refreshTime;
            totalData = data.total || totalData;
            pages = pageData.pages
                || data.pages
                || (pageTotal > 0 ? Math.ceil(pageTotal / pageSize) : (pageRecords.length < pageSize ? current : current + 1));
            current++;
            await sleep(150);
        }

        return buildSkuInventoryReport(records, totalData, refreshTime, startStr, endStr);
    }

    async function startSkuInventoryAnalysisProcess() {
        const btn = document.getElementById('sr-sku-inventory-start-btn');
        const exportBtn = document.getElementById('sr-sku-inventory-export');
        const startStr = document.getElementById('sr-sku-inventory-start').value;
        const endStr = document.getElementById('sr-sku-inventory-end').value;
        const skipWeekends = document.getElementById('sr-sku-inventory-skip-weekends').checked;
        const statusEl = document.getElementById('sr-status');

        if (!startStr || !endStr) return alert('请先选择完整的SKU库存分析周期！');
        if (new Date(startStr) > new Date(endStr)) return alert('开始日期不能晚于结束日期！');

        btn.disabled = true;
        exportBtn.disabled = true;
        btn.innerText = 'SKU库存分析中...';
        skuInventoryTableState.page = 1;
        resetSkuInventoryAnalysisView();

        try {
            skuInventoryReportData = await computeSkuInventoryAnalysisReport(startStr, endStr, (current, pages, count) => {
                btn.innerText = `SKU库存分析中... (${current}/${pages}页)`;
                statusEl.innerText = `正在拉取SKU库存分析数据... 已累计 ${count} 条`;
            });

            renderSkuInventoryAnalysisView();
            statusEl.innerText = `✅ SKU库存分析完成，共汇总 ${skuInventoryReportData.skuCount} 个SKU`;
        } catch (error) {
            skuInventoryReportData = null;
            resetSkuInventoryAnalysisView();
            statusEl.innerText = `❌ SKU库存分析失败: ${error.message}`;
        } finally {
            btn.disabled = false;
            btn.innerText = '开始SKU库存分析';
        }
    }

    // ==========================================
    // 0. 日志工具
    // ==========================================
    const LOG_PREFIX = '📦 [物流统计助手]';
    const logger = {
        info: (...args) => console.log(`%c${LOG_PREFIX} [INFO]`, 'color: #1890ff; font-weight: bold;', ...args),
        success: (...args) => console.log(`%c${LOG_PREFIX} [SUCCESS]`, 'color: #52c41a; font-weight: bold;', ...args),
        warn: (...args) => console.warn(`%c${LOG_PREFIX} [WARN]`, 'color: #faad14; font-weight: bold;', ...args),
        error: (...args) => console.error(`%c${LOG_PREFIX} [ERROR]`, 'color: #ff4d4f; font-weight: bold;', ...args)
    };

    // ==========================================
    // 1. OMP Track-Key 签名算法及 Payload 格式化
    // ==========================================
    const makeTrackKey = function b(e){const t=0,n="",i=8;function a(e){const t=g(e);return r(f(t),t.length*i)}function o(e){return b(a(e))}function s(e){return v(a(e))}function r(e,t){e[t>>5]|=128<<t%32,e[14+(t+64>>>9<<4)]=t;let n=1732584193,i=3989678985,a=2562383614,o=271733878;for(let s=0;s<e.length;s+=16){const t=n,r=i,c=a,h=o;n=l(n,i,a,o,e[s+0],7,-680876936),o=l(o,n,i,a,e[s+1],12,-389564586),a=l(a,o,n,i,e[s+2],17,606105819),i=l(i,a,o,n,e[s+3],22,-1044525330),n=l(n,i,a,o,e[s+4],7,-176418897),o=l(o,n,i,a,e[s+5],12,1200080426),a=l(a,o,n,i,e[s+6],17,-1473231341),i=l(i,a,o,n,e[s+7],22,-45705983),n=l(n,i,a,o,e[s+8],7,1770035416),o=l(o,n,i,a,e[s+9],12,-1958414417),a=l(a,o,n,i,e[s+10],17,-42063),i=l(i,a,o,n,e[s+11],22,-1990404162),n=l(n,i,a,o,e[s+12],7,1804603682),o=l(o,n,i,a,e[s+13],12,-40341101),a=l(a,o,n,i,e[s+14],17,-1502002290),i=l(i,a,o,n,e[s+15],22,1236535329),n=u(n,i,a,o,e[s+1],5,-165796510),o=u(o,n,i,a,e[s+6],9,-1069501632),a=u(a,o,n,i,e[s+11],14,643717713),i=u(i,a,o,n,e[s+0],20,-373897302),n=u(n,i,a,o,e[s+5],5,-701558691),o=u(o,n,i,a,e[s+10],9,38016083),a=u(a,o,n,i,e[s+15],14,-660478335),i=u(i,a,o,n,e[s+4],20,-405537848),n=u(n,i,a,o,e[s+9],5,568446438),o=u(o,n,i,a,e[s+14],9,-1019803690),a=u(a,o,n,i,e[s+3],14,-187363961),i=u(i,a,o,n,e[s+8],20,1163531501),n=u(n,i,a,o,e[s+13],5,-1444681467),o=u(o,n,i,a,e[s+2],9,-51403784),a=u(a,o,n,i,e[s+7],14,1735328473),i=u(i,a,o,n,e[s+12],20,-1926607734),n=d(n,i,a,o,e[s+5],4,-378558),o=d(o,n,i,a,e[s+8],11,-2022574463),a=d(a,o,n,i,e[s+11],16,1839030562),i=d(i,a,o,n,e[s+14],23,-35309556),n=d(n,i,a,o,e[s+1],4,-1530992060),o=d(o,n,i,a,e[s+4],11,1272893353),a=d(a,o,n,i,e[s+7],16,-155497632),i=d(i,a,o,n,e[s+10],23,-1094730640),n=d(n,i,a,o,e[s+13],4,681279174),o=d(o,n,i,a,e[s+0],11,-358537222),a=d(a,o,n,i,e[s+3],16,-722521979),i=d(i,a,o,n,e[s+6],23,76029189),n=d(n,i,a,o,e[s+9],4,-640364487),o=d(o,n,i,a,e[s+12],11,-421815835),a=d(a,o,n,i,e[s+15],16,530742520),i=d(i,a,o,n,e[s+2],23,-995338651),n=p(n,i,a,o,e[s+0],6,-198630844),o=p(o,n,i,a,e[s+7],10,1126891415),a=p(a,o,n,i,e[s+14],15,-1416354905),i=p(i,a,o,n,e[s+5],21,-57434055),n=p(n,i,a,o,e[s+12],6,1700485571),o=p(o,n,i,a,e[s+3],10,-1894986606),a=p(a,o,n,i,e[s+10],15,-1051523),i=p(i,a,o,n,e[s+1],21,-2054922799),n=p(n,i,a,o,e[s+8],6,1873313359),o=p(o,n,i,a,e[s+15],10,-30611744),a=p(a,o,n,i,e[s+6],15,-1560198380),i=p(i,a,o,n,e[s+13],21,1309151649),n=p(n,i,a,o,e[s+4],6,-145523070),o=p(o,n,i,a,e[s+11],10,-1120210379),a=p(a,o,n,i,e[s+2],15,718787259),i=p(i,a,o,n,e[s+9],21,-343485551),n=m(n,t),i=m(i,r),a=m(a,c),o=m(o,h)}return[n,i,a,o]}function c(e,t,n,i,a,o){return m(h(m(m(t,e),m(i,o)),a),n)}function l(e,t,n,i,a,o,s){return c(t&n|~t&i,e,t,a,o,s)}function u(e,t,n,i,a,o,s){return c(t&i|n&~i,e,t,a,o,s)}function d(e,t,n,i,a,o,s){return c(t^n^i,e,t,a,o,s)}function p(e,t,n,i,a,o,s){return c(n^(t|~i),e,t,a,o,s)}function m(e,t){const n=(65535&e)+(65535&t),i=(e>>16)+(t>>16)+(n>>16);return i<<16|65535&n}function h(e,t){return e<<t|e>>>32-t}function g(e){let t="";for(const n of e){const e=n.codePointAt(0);t+=e<=127?String.fromCharCode(e):e<=2047?String.fromCharCode(192|e>>6,128|63&e):e<=65535?String.fromCharCode(224|e>>12,128|e>>6&63,128|63&e):String.fromCharCode(240|e>>18,128|e>>12&63,128|e>>6&63,128|63&e)}return t}function f(e){const t=[],n=(1<<i)-1;for(let a=0;a<e.length*i;a+=i)t[a>>5]|=(e.charCodeAt(a/i)&n)<<a%32;return t}function b(e){const n=t?[48,49,50,51,52,53,54,55,56,57,65,66,67,68,69,70].map((function(e){return String.fromCharCode(e)})).join(""):[48,49,50,51,52,53,54,55,56,57,97,98,99,100,101,102].map((function(e){return String.fromCharCode(e)})).join("");let i="";for(let t=0;t<4*e.length;t++)i+=n.charAt(e[t>>2]>>t%4*8+4&15)+n.charAt(e[t>>2]>>t%4*8&15);return i}function v(e){const t=[65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,48,49,50,51,52,53,54,55,56,57,43,47].map((function(e){return String.fromCharCode(e)})).join("");let i="";for(let a=0;a<4*e.length;a+=3){const o=(e[a>>2]>>a%4*8&255)<<16|(e[a+1>>2]>>(a+1)%4*8&255)<<8|e[a+2>>2]>>(a+2)%4*8&255;for(let s=0;s<4;s++)8*a+6*s>32*e.length?i+=n:i+=t.charAt(o>>6*(3-s)&63)}return i}const w=Date.now().toString(),y=s(w),C=o(w),x="string"===typeof e?e:JSON.stringify(null==e?{}:e),k=o(x);return y+k+C};

    function isPlainObject(value) {
        if (Object.prototype.toString.call(value) !== "[object Object]") return false;
        const proto = Object.getPrototypeOf(value);
        return proto === Object.prototype || proto === null;
    }

    // 格式化 Payload：保证 JSON Keys 按字母排序，确保与 Track-Key 校验算法匹配
    function normalizePayload(value) {
        if (value === undefined) return undefined;
        if (value === null) return null;

        if (Array.isArray(value)) {
            return value.map((item) => {
                const normalized = normalizePayload(item);
                return normalized === undefined ? null : normalized;
            });
        }

        if (!isPlainObject(value)) {
            return value;
        }

        const output = {};
        for (const key of Object.keys(value).sort()) {
            const normalized = normalizePayload(value[key]);
            if (normalized !== undefined) {
                output[key] = normalized;
            }
        }
        return output;
    }

    function stringifyPayload(payload) {
        if (typeof payload === "string") return payload;
        return JSON.stringify(normalizePayload(payload));
    }


    // ==========================================
    // 2. 拦截器：自动提取系统鉴权 Token
    // ==========================================
    let sysHeaders = {
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json, text/plain, */*"
    };

    const origOpen = XMLHttpRequest.prototype.open;
    const origSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;

    XMLHttpRequest.prototype.open = function() {
        this._url = arguments[1];
        origOpen.apply(this, arguments);
    };

    let headerIntercepted = false;
    XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
        if(!this._headers) this._headers = {};
        this._headers[header] = value;

        if (this._url && (this._url.includes('/gateway/omp/') || this._url.includes('/gateway/wms/'))) {
            // 过滤掉原请求中的动态 Track-Key，避免污染公共 Header 池
            if (header.toLowerCase() !== 'track-key') {
                sysHeaders[header] = value;
            }
            if (!headerIntercepted && (header.toLowerCase().includes('auth') || header.toLowerCase().includes('token') || header === 'Tenant-Id')) {
                headerIntercepted = true;
            }
        }
        origSetRequestHeader.apply(this, arguments);
    };


    // ==========================================
    // 3. 时间工具
    // ==========================================
    function formatDateStandard(dateObj) {
        const yyyy = dateObj.getFullYear();
        const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
        const dd = String(dateObj.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    function isWeekendDate(dateLike) {
        if (!dateLike) return false;
        const text = String(dateLike).trim().slice(0, 10);
        if (!text) return false;
        const date = new Date(text.replace(/-/g, '/'));
        if (Number.isNaN(date.getTime())) return false;
        const day = date.getDay();
        return day === 0 || day === 6;
    }

    function getDateRangeList(startStr, endStr, skipWeekends = false) {
        const dates = [];
        const current = new Date(String(startStr).replace(/-/g, '/'));
        const end = new Date(String(endStr).replace(/-/g, '/'));
        if (Number.isNaN(current.getTime()) || Number.isNaN(end.getTime())) return dates;
        while (current <= end) {
            const dateStr = formatDateStandard(current);
            if (!skipWeekends || !isWeekendDate(dateStr)) dates.push(dateStr);
            current.setDate(current.getDate() + 1);
        }
        return dates;
    }

    function getTargetDate(daysAgo, skipWeekends) {
        let d = new Date();
        if (skipWeekends) { while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() - 1); }
        let count = 0;
        while (count < daysAgo) {
            d.setDate(d.getDate() - 1);
            if (skipWeekends) { while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() - 1); }
            count++;
        }
        const dateStr = formatDateStandard(d);
        return { dateStr: dateStr, startTime: `${dateStr} 00:00:00`, endTime: `${dateStr} 23:59:59` };
    }

    // 获取上周的周一和周五
    function getLastWeekMonFri() {
        let d = new Date();
        let dayOfWeek = d.getDay() || 7; // 1-7
        d.setDate(d.getDate() - (dayOfWeek - 1) - 7); // 上周一
        const start = formatDateStandard(d);
        d.setDate(d.getDate() + 4); // 上周五
        const end = formatDateStandard(d);
        return { start, end };
    }

    function getHoursDiff(startTimeStr, endTimeStr) {
        if (!startTimeStr || !endTimeStr) return Infinity;
        const s = new Date(startTimeStr.replace(/-/g, '/')).getTime();
        const e = new Date(endTimeStr.replace(/-/g, '/')).getTime();
        if (Number.isNaN(s) || Number.isNaN(e)) return Infinity;
        return (e - s) / (1000 * 60 * 60);
    }

    function isReceiptWithinHours(createTimeStr, receiptTimeStr, thresholdHours) {
        const hoursDiff = getHoursDiff(createTimeStr, receiptTimeStr);
        return Number.isFinite(hoursDiff) && hoursDiff <= Number(thresholdHours || 0);
    }

    function isOnlineInSampleWindow(record, thresholdHours) {
        const receiptTime = String(record?.receiptTime || '').trim();
        if (receiptTime) {
            return isReceiptWithinHours(record?.createTime, receiptTime, thresholdHours);
        }
        const outboundTime = String(record?.outboundTime || '').trim();
        return Boolean(outboundTime);
    }

    function isCancelledOrder(record) {
        return String(record?.status ?? '').trim() === '99';
    }

    function isOnlineOver72Hours(record) {
        const receiptTime = String(record?.receiptTime || '').trim();
        if (!receiptTime) return false;
        const hoursDiff = getHoursDiff(record?.createTime, receiptTime);
        return Number.isFinite(hoursDiff) && hoursDiff > 72;
    }

    function shouldIncludeInOnlineRateStats(record) {
        return !isCancelledOrder(record) && !isOnlineOver72Hours(record);
    }

    function formatOnlineDuration(hoursDiff) {
        if (!Number.isFinite(hoursDiff)) return '-';
        const sign = hoursDiff < 0 ? '-' : '';
        const totalMinutes = Math.abs(Math.round(hoursDiff * 60));
        const hours = Math.floor(totalMinutes / 60);
        const minutes = totalMinutes % 60;
        if (hours === 0) return `${sign}${minutes}分`;
        if (minutes === 0) return `${sign}${hours}小时`;
        return `${sign}${hours}小时${minutes}分`;
    }

    function buildOnlineRateWindow(startDateStr, endDateStr) {
        return {
            startTime: `${startDateStr} 00:00:00`,
            endTime: `${endDateStr} 23:59:59`,
            label: startDateStr === endDateStr ? startDateStr : `${startDateStr} ~ ${endDateStr}`
        };
    }

    function getOnlineRateSampleRange(skipWeekends) {
        const startInput = document.getElementById('sr-online-start');
        const endInput = document.getElementById('sr-online-end');
        const startStr = startInput?.value || '';
        const endStr = endInput?.value || '';
        if (startStr || endStr) {
            if (!startStr || !endStr) throw new Error('请同时选择上网率样本开始和结束日期');
            return { startStr, endStr, custom: true };
        }

        const day24 = getTargetDate(1, skipWeekends);
        const day72 = getTargetDate(3, skipWeekends);
        return { startStr: day72.dateStr, endStr: day24.dateStr, custom: false };
    }

    // ==========================================
    // 4. API 请求封装 (已添加 Track-Key 支持)
    // ==========================================
    const API_ONLINE_RATE = "https://omp.xlwms.com/gateway/omp/order/delivery/page";
    const API_ORDER_ADDRESS_INFO = "https://omp.xlwms.com/gateway/omp/order/getAddressInfo";
    const API_STOCK_ANALYSIS = "https://omp.xlwms.com/gateway/omp/analysis/stock-flow/listAnalysisPage";
    const API_CUSTOMER_LIST = "https://omp.xlwms.com/gateway/omp/customer/list";
    const API_WAREHOUSE = "https://omp.xlwms.com/gateway/wms/warehouse/page?size=999";
    const API_OUTBOUND_RATE = "https://omp.xlwms.com/gateway/wms/report/deliveryEfficiencySummary";
    const UNKNOWN_REGION = '未知';
    const REGION_COLUMNS = ['美东北', '美东', '美东南', '美中', '美中南', '美中北', '美西', '美西南', '美西北'];
    const ZIP_REGION_RULES = [
        '邮编以 0 开头 → 归类为 美东北',
        '邮编以 1 开头 → 归类为 美东北',
        '邮编以 2 开头 → 归类为 美东',
        '邮编以 3 开头，且完整数字 < 35000 → 美东',
        '邮编以 3 开头，且完整数字 ≥ 35000 → 美东南',
        '邮编以 4 开头，且完整数字 < 45000 → 美中',
        '邮编以 4 开头，且完整数字 ≥ 45000 → 美中南',
        '邮编以 5 开头，且完整数字 < 55000 → 美中北',
        '邮编以 5 开头，且完整数字 ≥ 55000 → 美中',
        '邮编以 6 开头 → 归类为 美中',
        '邮编以 7 开头，且完整数字 < 75000 → 美中南',
        '邮编以 7 开头，且完整数字 ≥ 75000 → 美西南',
        '邮编以 8 开头，且完整数字 < 85000 → 美西',
        '邮编以 8 开头，且完整数字 ≥ 85000 → 美西南',
        '邮编以 9 开头，且完整数字 < 95000 → 美西北',
        '邮编以 9 开头，且完整数字 ≥ 95000 → 美西',
        '所有其他情况（比如首数字不是 0-9，或者格式不对）→ 判定为 未知，不纳入统计'
    ];

    async function fetchOnlinePage(startTime, endTime, page) {
        // 与 Web 端发送的默认 Payload 字段结构严格对齐
        const payload = {
            "appendixFlag": "",
            "categoryIdList": [],
            "countKind": "orderWeight",
            "current": page,
            "endTime": endTime,
            "expressFlag": "",
            "morePkgFlag": "",
            "noTypeNo": "",
            "orderSourceList": [],
            "platformCode": 1,
            "relatedReturnOrder": "",
            "reveiverInput": "",
            "size": 500, // 保留大容量抓取
            "startTime": startTime,
            "status": "",
            "storeName": "",
            "timeType": "createTime",
            "transitStatusList": [],
            "unitMark": 0,
            "whCode": "",
            "withVas": ""
        };

        const bodyStr = stringifyPayload(payload);
        const headers = {
            ...sysHeaders,
            "track-key": makeTrackKey(bodyStr)
        };

        const response = await fetch(API_ONLINE_RATE, {
            method: 'POST',
            headers: headers,
            body: bodyStr
        });

        if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
        const res = await response.json();
        if (res.code !== 200) throw new Error(res.msg || '获取上网数据失败');
        return res.data;
    }

    async function fetchWarehouses() {
        const res = await fetch(API_WAREHOUSE, { method: 'GET', headers: sysHeaders });
        const json = await res.json();
        return json.code === 200 ? (json.data.records || []) : [];
    }

    async function fetchCustomerListPage(page = 1, size = 200) {
        const params = new URLSearchParams({
            current: String(page),
            size: String(size),
            status: '0',
            type: '1'
        });
        const response = await fetch(`${API_CUSTOMER_LIST}?${params.toString()}`, { method: 'GET', headers: sysHeaders });
        if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
        const json = await response.json();
        if (json.code !== 200) throw new Error(json.msg || '获取客户列表失败');
        return json.data || {};
    }

    async function fetchAllCustomerNames(progress) {
        const records = [];
        let current = 1;
        let pages = 1;
        const size = 200;
        while (current <= pages) {
            if (progress) progress(current, pages, records.length);
            const data = await fetchCustomerListPage(current, size);
            const pageRecords = data.records || [];
            records.push(...pageRecords);
            const pageSize = Number(data.size || size);
            const pageTotal = Number(data.total || 0);
            pages = Number(data.pages || 0)
                || (pageTotal > 0 ? Math.ceil(pageTotal / pageSize) : (pageRecords.length < pageSize ? current : current + 1));
            current++;
            await sleep(100);
        }
        return records;
    }

    async function fetchStockAnalysisPage(startTime, endTime, page, options = {}) {
        const payload = {
            "aggregateMode": Number(options.aggregateMode ?? 0),
            "current": page,
            "customerCodeList": [],
            "endTime": endTime,
            "searchType": "0",
            "searchValue": options.searchValue || "",
            "size": 500,
            "startTime": startTime,
            "stockItemType": 0,
            "unitMark": 0,
            "whCodeList": [],
            "Barcode": options.barcode || ""
        };

        const bodyStr = stringifyPayload(payload);
        const headers = {
            ...sysHeaders,
            "track-key": makeTrackKey(bodyStr)
        };

        const response = await fetch(API_STOCK_ANALYSIS, {
            method: 'POST',
            headers,
            body: bodyStr
        });

        if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
        const res = await response.json();
        if (res.code !== 200) throw new Error(res.msg || '获取库存分析数据失败');
        return res.data || {};
    }

    async function fetchOrderAddressInfo(orderNo) {
        const params = new URLSearchParams({
            type: '1',
            orderNo: orderNo || ''
        });
        const res = await fetch(`${API_ORDER_ADDRESS_INFO}?${params.toString()}`, { method: 'GET', headers: sysHeaders });
        if (!res.ok) throw new Error(`HTTP Error: ${res.status}`);
        const json = await res.json();
        if (json.code !== 200) throw new Error(json.msg || '获取订单地址信息失败');
        return json.data || {};
    }

    async function fetchOutboundRate(whCode, tenantCode, startDate, endDate) {
        const payload = { whCode, startDate, endDate, tenantCode: tenantCode || "2836" };
        const bodyStr = stringifyPayload(payload);
        const headers = {
            ...sysHeaders,
            "whcode": whCode,
            "track-key": makeTrackKey(bodyStr)
        };
        const res = await fetch(API_OUTBOUND_RATE, { method: 'POST', headers: headers, body: bodyStr });
        const json = await res.json();
        return json.data || {};
    }

    function sleep(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    function createRegionCounter() {
        return REGION_COLUMNS.reduce((acc, region) => {
            acc[region] = 0;
            return acc;
        }, {});
    }

    function getRegionByPostCode(postCode) {
        if (!postCode) return UNKNOWN_REGION;
        const digits = String(postCode).replace(/\D/g, '').slice(0, 5);
        if (!/^\d{5}$/.test(digits)) return UNKNOWN_REGION;

        const first = digits[0];
        const numeric = Number(digits);

        if (first === '0' || first === '1') return '美东北';
        if (first === '2') return '美东';
        if (first === '3') return numeric < 35000 ? '美东' : '美东南';
        if (first === '4') return numeric < 45000 ? '美中' : '美中南';
        if (first === '5') return numeric < 55000 ? '美中北' : '美中';
        if (first === '6') return '美中';
        if (first === '7') return numeric < 75000 ? '美中南' : '美西南';
        if (first === '8') return numeric < 85000 ? '美西' : '美西南';
        if (first === '9') return numeric < 95000 ? '美西北' : '美西';
        return UNKNOWN_REGION;
    }

    async function forEachPeriodOrder(startStr, endStr, skipWeekends, onRecord, onProgress, pageDelay = 150) {
        const startTime = `${startStr} 00:00:00`;
        const endTime = `${endStr} 23:59:59`;
        let current = 1, pages = 1, matchedCount = 0;

        while (current <= pages) {
            if (onProgress) onProgress(current, pages, matchedCount);
            const data = await fetchOnlinePage(startTime, endTime, current);
            pages = data.pages || 1;

            for (const r of (data.records || [])) {
                if (!r.createTime) continue;

                if (skipWeekends) {
                    const day = new Date(r.createTime.replace(/-/g, '/')).getDay();
                    if (day === 0 || day === 6) continue;
                }

                matchedCount++;
                const result = onRecord && onRecord(r, matchedCount, current, pages);
                if (result && typeof result.then === 'function') {
                    await result;
                }
            }

            current++;
            await sleep(pageDelay);
        }

        return matchedCount;
    }

    async function runWithConcurrency(items, limit, handler) {
        let cursor = 0;
        const workerCount = Math.min(limit, items.length);
        const workers = Array.from({ length: workerCount }, async () => {
            while (true) {
                const index = cursor++;
                if (index >= items.length) break;
                await handler(items[index], index);
            }
        });
        await Promise.all(workers);
    }

    function upsertRegionStats(container, order, region, postCode) {
        if (!REGION_COLUMNS.includes(region)) return false;

        const customerKey = `${order.customerCode || ''}__${order.customerName || '未知客户'}`;
        const warehouseKey = `${order.whCode || ''}__${order.whName || '未知仓'}`;

        if (!container[customerKey]) {
            container[customerKey] = {
                customerCode: order.customerCode || '',
                customerName: order.customerName || '未知客户',
                total: 0,
                regions: createRegionCounter(),
                warehouses: {}
            };
        }

        const customerNode = container[customerKey];
        if (!customerNode.warehouses[warehouseKey]) {
            customerNode.warehouses[warehouseKey] = {
                whCode: order.whCode || '',
                whName: order.whName || '未知仓',
                total: 0,
                regions: createRegionCounter(),
                samplePostCodes: []
            };
        }

        const warehouseNode = customerNode.warehouses[warehouseKey];
        customerNode.total++;
        customerNode.regions[region] += 1;
        warehouseNode.total++;
        warehouseNode.regions[region] += 1;

        if (postCode && warehouseNode.samplePostCodes.length < 3 && !warehouseNode.samplePostCodes.includes(postCode)) {
            warehouseNode.samplePostCodes.push(postCode);
        }

        return true;
    }

    function buildRegionRows(regionStats, includeWarehouse) {
        const sortByCustomerName = (a, b) =>
            a.customerName.localeCompare(b.customerName, 'zh-Hans-CN') ||
            (a.customerCode || '').localeCompare(b.customerCode || '');
        const sortByWarehouseName = (a, b) =>
            a.whName.localeCompare(b.whName, 'zh-Hans-CN') ||
            (a.whCode || '').localeCompare(b.whCode || '');

        const rows = [];
        Object.values(regionStats)
            .sort(sortByCustomerName)
            .forEach((customerNode) => {
                if (includeWarehouse) {
                    Object.values(customerNode.warehouses)
                        .sort(sortByWarehouseName)
                        .forEach((warehouseNode) => {
                            const row = {
                                "客户名称": customerNode.customerName,
                                "客户编码": customerNode.customerCode,
                                "发货仓库": warehouseNode.whName,
                                "仓库代码": warehouseNode.whCode,
                                "总订单量": warehouseNode.total
                            };
                            REGION_COLUMNS.forEach((region) => { row[region] = warehouseNode.regions[region]; });
                            row["样例邮编"] = warehouseNode.samplePostCodes.join(' / ');
                            rows.push(row);
                        });
                } else {
                    const row = {
                        "客户名称": customerNode.customerName,
                        "客户编码": customerNode.customerCode,
                        "总订单量": customerNode.total
                    };
                    REGION_COLUMNS.forEach((region) => { row[region] = customerNode.regions[region]; });
                    rows.push(row);
                }
            });
        return rows;
    }

    function buildRegionReport(regionStats, detailFailed, sourceTotalOrders, startStr, endStr, excludedUnknown = 0) {
        const detailRows = buildRegionRows(regionStats, true);
        const customerRows = buildRegionRows(regionStats, false);
        const ruleRows = ZIP_REGION_RULES.map((rule, index) => ({ "序号": index + 1, "邮编地区规则": rule }));
        const regionTotals = createRegionCounter();

        customerRows.forEach((row) => {
            REGION_COLUMNS.forEach((region) => {
                regionTotals[region] += Number(row[region] || 0);
            });
        });

        const totalOrders = customerRows.reduce((sum, row) => sum + Number(row["总订单量"] || 0), 0);
        const customerCount = customerRows.length;
        const warehouseCount = detailRows.length;
        let topRegion = REGION_COLUMNS
            .map((region) => ({ region, value: regionTotals[region] }))
            .sort((a, b) => b.value - a.value)[0] || { region: '-', value: 0 };
        if (!topRegion.value) topRegion = { region: '-', value: 0 };

        return {
            startStr,
            endStr,
            totalOrders,
            sourceTotalOrders,
            detailFailed,
            excludedUnknown,
            customerCount,
            warehouseCount,
            regionTotals,
            topRegion,
            detailRows,
            customerRows,
            ruleRows
        };
    }

    function buildInventoryReport(records, totalData, refreshTime, startStr, endStr) {
        const warehouseMap = {};

        records.forEach((record) => {
            const whCode = record.whCode || '';
            const whName = record.whName || '未知仓';
            const key = `${whCode}__${whName}`;
            const closeStockQty = Number(record.closeStockQty || 0);
            const analysis = record.analysisResult || {};
            const weight = closeStockQty > 0 ? closeStockQty : 1;

            if (!warehouseMap[key]) {
                warehouseMap[key] = {
                    whCode,
                    whName,
                    customerSet: new Set(),
                    preStockQty: 0,
                    closeStockQty: 0,
                    outboundBookQty: 0,
                    weightTotal: 0,
                    stockTurnoverRateWeighted: 0,
                    stockTurnoverDaysWeighted: 0,
                    stockSaleRateWeighted: 0
                };
            }

            const node = warehouseMap[key];
            node.customerSet.add(record.customerName || '未知客户');
            node.preStockQty += Number(record.preStockQty || 0);
            node.closeStockQty += closeStockQty;
            node.outboundBookQty += Math.abs(Number(record.outboundBookQty || 0));
            node.weightTotal += weight;
            node.stockTurnoverRateWeighted += Number(analysis.stockTurnoverRate || 0) * weight;
            node.stockTurnoverDaysWeighted += Number(analysis.stockTurnoverDays || 0) * weight;
            node.stockSaleRateWeighted += Number(analysis.stockSaleRate || 0) * weight;
        });

        const warehouseRows = Object.values(warehouseMap)
            .map((node) => ({
                "仓库名称": node.whName,
                "仓库代码": node.whCode,
                "客户数": node.customerSet.size,
                "期初库存": Number(node.preStockQty.toFixed(2)),
                "期末库存": Number(node.closeStockQty.toFixed(2)),
                "出库预占": Number(node.outboundBookQty.toFixed(2)),
                "库存周转率": node.weightTotal === 0 ? 0 : Number((node.stockTurnoverRateWeighted / node.weightTotal).toFixed(2)),
                "库存周转天数": node.weightTotal === 0 ? 0 : Number((node.stockTurnoverDaysWeighted / node.weightTotal).toFixed(2)),
                "库存销率": node.weightTotal === 0 ? 0 : Number((node.stockSaleRateWeighted / node.weightTotal).toFixed(2))
            }))
            .sort((a, b) =>
                a["仓库名称"].localeCompare(b["仓库名称"], 'zh-Hans-CN') ||
                a["仓库代码"].localeCompare(b["仓库代码"], 'zh-Hans-CN')
            );

        const totalAnalysis = (totalData && totalData.analysisResult) || {};
        const topTurnoverWarehouse = warehouseRows
            .slice()
            .sort((a, b) => b["库存周转率"] - a["库存周转率"])[0] || null;

        return {
            startStr,
            endStr,
            refreshTime: refreshTime || '',
            totalRows: records.length,
            warehouseRows,
            totalWarehouseCount: warehouseRows.length,
            totalCloseStockQty: Number(totalData?.closeStockQty || 0),
            totalPreStockQty: Number(totalData?.preStockQty || 0),
            totalTurnoverRate: Number(totalAnalysis.stockTurnoverRate || 0),
            totalTurnoverDays: Number(totalAnalysis.stockTurnoverDays || 0),
            totalStockSaleRate: Number(totalAnalysis.stockSaleRate || 0),
            topTurnoverWarehouse
        };
    }

    function createInventoryMetricsNode(extra = {}) {
        return {
            preStockQty: 0,
            closeStockQty: 0,
            outboundBookQty: 0,
            weightTotal: 0,
            stockTurnoverRateWeighted: 0,
            stockTurnoverDaysWeighted: 0,
            stockSaleRateWeighted: 0,
            ...extra
        };
    }

    function appendInventoryMetrics(node, record) {
        const closeStockQty = Number(record.closeStockQty || 0);
        const analysis = record.analysisResult || {};
        const weight = closeStockQty > 0 ? closeStockQty : 1;

        node.preStockQty += Number(record.preStockQty || 0);
        node.closeStockQty += closeStockQty;
        node.outboundBookQty += Math.abs(Number(record.outboundBookQty || 0));
        node.weightTotal += weight;
        node.stockTurnoverRateWeighted += Number(analysis.stockTurnoverRate || 0) * weight;
        node.stockTurnoverDaysWeighted += Number(analysis.stockTurnoverDays || 0) * weight;
        node.stockSaleRateWeighted += Number(analysis.stockSaleRate || 0) * weight;
    }

    function finalizeInventoryMetrics(node) {
        return {
            preStockQty: Number(node.preStockQty.toFixed(2)),
            closeStockQty: Number(node.closeStockQty.toFixed(2)),
            outboundBookQty: Number(node.outboundBookQty.toFixed(2)),
            stockTurnoverRate: node.weightTotal === 0 ? 0 : Number((node.stockTurnoverRateWeighted / node.weightTotal).toFixed(2)),
            stockTurnoverDays: node.weightTotal === 0 ? 0 : Number((node.stockTurnoverDaysWeighted / node.weightTotal).toFixed(2)),
            stockSaleRate: node.weightTotal === 0 ? 0 : Number((node.stockSaleRateWeighted / node.weightTotal).toFixed(2))
        };
    }

    function buildInventoryReport(records, totalData, refreshTime, startStr, endStr) {
        const customerMap = {};

        records.forEach((record) => {
            const customerCode = record.customerCode || '';
            const customerName = record.customerName || '未知客户';
            const whCode = record.whCode || '';
            const whName = record.whName || '未知仓库';
            const customerKey = `${customerCode}__${customerName}`;
            const warehouseKey = `${whCode}__${whName}`;

            if (!customerMap[customerKey]) {
                customerMap[customerKey] = createInventoryMetricsNode({
                    customerCode,
                    customerName,
                    warehouseSet: new Set(),
                    warehouses: {}
                });
            }

            const customerNode = customerMap[customerKey];
            customerNode.warehouseSet.add(warehouseKey);
            appendInventoryMetrics(customerNode, record);

            if (!customerNode.warehouses[warehouseKey]) {
                customerNode.warehouses[warehouseKey] = createInventoryMetricsNode({
                    whCode,
                    whName
                });
            }

            appendInventoryMetrics(customerNode.warehouses[warehouseKey], record);
        });

        const sortByCustomerName = (a, b) =>
            a.customerName.localeCompare(b.customerName, 'zh-Hans-CN') ||
            (a.customerCode || '').localeCompare(b.customerCode || '');
        const sortByWarehouseName = (a, b) =>
            a.whName.localeCompare(b.whName, 'zh-Hans-CN') ||
            (a.whCode || '').localeCompare(b.whCode || '');

        const detailRows = [];
        const customerRows = Object.values(customerMap)
            .sort(sortByCustomerName)
            .map((customerNode) => {
                Object.values(customerNode.warehouses)
                    .sort(sortByWarehouseName)
                    .forEach((warehouseNode) => {
                        const metrics = finalizeInventoryMetrics(warehouseNode);
                        detailRows.push({
                            "客户名称": customerNode.customerName,
                            "客户编码": customerNode.customerCode,
                            "发货仓库": warehouseNode.whName,
                            "仓库代码": warehouseNode.whCode,
                            "期初库存": metrics.preStockQty,
                            "期末库存": metrics.closeStockQty,
                            "出库预占": metrics.outboundBookQty,
                            "库存周转率": metrics.stockTurnoverRate,
                            "库存周转天数": metrics.stockTurnoverDays,
                            "库存售罄率": metrics.stockSaleRate
                        });
                    });

                const metrics = finalizeInventoryMetrics(customerNode);
                return {
                    "客户名称": customerNode.customerName,
                    "客户编码": customerNode.customerCode,
                    "分仓数": customerNode.warehouseSet.size,
                    "期初库存": metrics.preStockQty,
                    "期末库存": metrics.closeStockQty,
                    "出库预占": metrics.outboundBookQty,
                    "库存周转率": metrics.stockTurnoverRate,
                    "库存周转天数": metrics.stockTurnoverDays,
                    "库存售罄率": metrics.stockSaleRate
                };
            });

        detailRows.sort((a, b) =>
            a["客户名称"].localeCompare(b["客户名称"], 'zh-Hans-CN') ||
            a["发货仓库"].localeCompare(b["发货仓库"], 'zh-Hans-CN') ||
            a["仓库代码"].localeCompare(b["仓库代码"], 'zh-Hans-CN')
        );

        const totalAnalysis = (totalData && totalData.analysisResult) || {};
        const totalCloseStockQty = Number(totalData?.closeStockQty || customerRows.reduce((sum, row) => sum + Number(row["期末库存"] || 0), 0).toFixed(2));
        const totalPreStockQty = Number(totalData?.preStockQty || customerRows.reduce((sum, row) => sum + Number(row["期初库存"] || 0), 0).toFixed(2));
        const totalWarehouseCount = new Set(detailRows.map((row) => `${row["仓库代码"] || ''}__${row["发货仓库"] || ''}`)).size;
        const topTurnoverWarehouse = detailRows
            .slice()
            .sort((a, b) => b["库存周转率"] - a["库存周转率"] || b["期末库存"] - a["期末库存"])[0] || null;
        const topStockCustomer = customerRows
            .slice()
            .sort((a, b) => b["期末库存"] - a["期末库存"] || b["库存周转率"] - a["库存周转率"])[0] || null;

        return {
            startStr,
            endStr,
            refreshTime: refreshTime || '',
            totalRows: records.length,
            detailRows,
            customerRows,
            warehouseRows: detailRows,
            customerCount: customerRows.length,
            detailCount: detailRows.length,
            totalWarehouseCount,
            totalCloseStockQty,
            totalPreStockQty,
            totalTurnoverRate: Number(totalAnalysis.stockTurnoverRate || 0),
            totalTurnoverDays: Number(totalAnalysis.stockTurnoverDays || 0),
            totalStockSaleRate: Number(totalAnalysis.stockSaleRate || 0),
            topTurnoverWarehouse,
            topStockCustomer
        };
    }

    function normalizeInventoryIdentity(value) {
        return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
    }

    function getInventoryIdentityKey(type, code, name, fallback = 'unknown') {
        const normalizedCode = normalizeInventoryIdentity(code);
        const normalizedName = normalizeInventoryIdentity(name);
        if (normalizedCode) return `${type}:code:${normalizedCode}`;
        if (normalizedName) return `${type}:name:${normalizedName.toLowerCase()}`;
        return `${type}:${fallback}`;
    }

    function pickPreferredInventoryText(currentValue, nextValue, fallback = '') {
        const current = normalizeInventoryIdentity(currentValue);
        const next = normalizeInventoryIdentity(nextValue);
        if (!current) return next || fallback;
        if (!next) return current || fallback;
        if (current === next) return current;
        return next.length > current.length ? next : current;
    }

    function resolveSkuInventoryRecord(record) {
        const skuProperty = record.stockItemProperty || {};
        return {
            customerCode: normalizeInventoryIdentity(record.customerCode),
            customerName: pickPreferredInventoryText('', record.customerName, '未知客户'),
            whCode: normalizeInventoryIdentity(record.whCode),
            whName: pickPreferredInventoryText('', record.whName, '未知仓库'),
            sku: pickPreferredInventoryText(
                '',
                record.stockItemFirstCode || skuProperty.sku || record.stockItemSecondCode || skuProperty.barcode || skuProperty.customizeBarcode,
                '-'
            ),
            productName: pickPreferredInventoryText('', skuProperty.productName, '-')
        };
    }

    function syncSkuInventoryNodeIdentity(node, resolved) {
        node.customerCode = node.customerCode || resolved.customerCode;
        node.customerName = pickPreferredInventoryText(node.customerName, resolved.customerName, '未知客户');
        node.whCode = node.whCode || resolved.whCode;
        node.whName = pickPreferredInventoryText(node.whName, resolved.whName, '未知仓库');
        node.sku = pickPreferredInventoryText(node.sku, resolved.sku, '-');
        node.productName = pickPreferredInventoryText(node.productName, resolved.productName, '-');
    }

    function buildSkuInventoryReport(records, totalData, refreshTime, startStr, endStr) {
        const detailMap = {};
        const customerMap = {};
        const skuMap = {};
        const seenRecordKeys = new Set();

        records.forEach((record) => {
            const resolved = resolveSkuInventoryRecord(record);
            const customerKey = getInventoryIdentityKey('customer', resolved.customerCode, resolved.customerName);
            const warehouseKey = getInventoryIdentityKey('warehouse', resolved.whCode, resolved.whName);
            const skuKey = getInventoryIdentityKey('sku', resolved.sku, resolved.productName, 'missing-sku');
            const detailKey = `${customerKey}__${warehouseKey}__${skuKey}`;
            const recordKey = normalizeInventoryIdentity(record.id)
                || `${detailKey}__${normalizeInventoryIdentity(record.statisticDate) || endStr}`;

            if (seenRecordKeys.has(recordKey)) return;
            seenRecordKeys.add(recordKey);

            if (!detailMap[detailKey]) {
                detailMap[detailKey] = createInventoryMetricsNode({
                    customerCode: resolved.customerCode,
                    customerName: resolved.customerName,
                    whCode: resolved.whCode,
                    whName: resolved.whName,
                    sku: resolved.sku,
                    productName: resolved.productName
                });
            }
            syncSkuInventoryNodeIdentity(detailMap[detailKey], resolved);
            appendInventoryMetrics(detailMap[detailKey], record);

            if (!customerMap[customerKey]) {
                customerMap[customerKey] = createInventoryMetricsNode({
                    customerCode: resolved.customerCode,
                    customerName: resolved.customerName,
                    skuSet: new Set(),
                    warehouseSet: new Set()
                });
            }
            syncSkuInventoryNodeIdentity(customerMap[customerKey], resolved);
            customerMap[customerKey].skuSet.add(skuKey);
            customerMap[customerKey].warehouseSet.add(warehouseKey);
            appendInventoryMetrics(customerMap[customerKey], record);

            if (!skuMap[skuKey]) {
                skuMap[skuKey] = createInventoryMetricsNode({
                    sku: resolved.sku,
                    productName: resolved.productName,
                    customerSet: new Set(),
                    warehouseSet: new Set()
                });
            }
            syncSkuInventoryNodeIdentity(skuMap[skuKey], resolved);
            skuMap[skuKey].customerSet.add(customerKey);
            skuMap[skuKey].warehouseSet.add(warehouseKey);
            appendInventoryMetrics(skuMap[skuKey], record);
        });

        const detailRows = Object.values(detailMap)
            .map((node) => {
                const metrics = finalizeInventoryMetrics(node);
                return {
                    "客户名称": node.customerName,
                    "客户编码": node.customerCode,
                    "发货仓库": node.whName,
                    "仓库代码": node.whCode,
                    "SKU": node.sku,
                    "产品名称": node.productName,
                    "期初库存": metrics.preStockQty,
                    "期末库存": metrics.closeStockQty,
                    "出库预占": metrics.outboundBookQty,
                    "库存周转率": metrics.stockTurnoverRate,
                    "库存周转天数": metrics.stockTurnoverDays,
                    "库存售罄率": metrics.stockSaleRate
                };
            })
            .sort((a, b) =>
                a["客户名称"].localeCompare(b["客户名称"], 'zh-Hans-CN') ||
                a["SKU"].localeCompare(b["SKU"], 'zh-Hans-CN') ||
                a["发货仓库"].localeCompare(b["发货仓库"], 'zh-Hans-CN')
            );

        const customerRows = Object.values(customerMap)
            .map((node) => {
                const metrics = finalizeInventoryMetrics(node);
                return {
                    "客户名称": node.customerName,
                    "客户编码": node.customerCode,
                    "SKU数": node.skuSet.size,
                    "仓库数": node.warehouseSet.size,
                    "期初库存": metrics.preStockQty,
                    "期末库存": metrics.closeStockQty,
                    "出库预占": metrics.outboundBookQty,
                    "库存周转率": metrics.stockTurnoverRate,
                    "库存周转天数": metrics.stockTurnoverDays,
                    "库存售罄率": metrics.stockSaleRate
                };
            })
            .sort((a, b) =>
                a["客户名称"].localeCompare(b["客户名称"], 'zh-Hans-CN') ||
                (a["客户编码"] || '').localeCompare(b["客户编码"] || '')
            );

        const skuRows = Object.values(skuMap)
            .map((node) => {
                const metrics = finalizeInventoryMetrics(node);
                return {
                    "SKU": node.sku,
                    "产品名称": node.productName,
                    "客户数": node.customerSet.size,
                    "仓库数": node.warehouseSet.size,
                    "期初库存": metrics.preStockQty,
                    "期末库存": metrics.closeStockQty,
                    "出库预占": metrics.outboundBookQty,
                    "库存周转率": metrics.stockTurnoverRate,
                    "库存周转天数": metrics.stockTurnoverDays,
                    "库存售罄率": metrics.stockSaleRate
                };
            })
            .sort((a, b) =>
                a["SKU"].localeCompare(b["SKU"], 'zh-Hans-CN') ||
                a["产品名称"].localeCompare(b["产品名称"], 'zh-Hans-CN')
            );

        const totalAnalysis = (totalData && totalData.analysisResult) || {};
        const topSkuRow = skuRows
            .slice()
            .sort((a, b) => b["期末库存"] - a["期末库存"] || b["库存周转率"] - a["库存周转率"])[0] || null;

        return {
            startStr,
            endStr,
            refreshTime: refreshTime || '',
            totalRows: seenRecordKeys.size,
            detailRows,
            customerRows,
            skuRows,
            customerCount: customerRows.length,
            skuCount: skuRows.length,
            detailCount: detailRows.length,
            totalCloseStockQty: Number(totalData?.closeStockQty || skuRows.reduce((sum, row) => sum + Number(row["期末库存"] || 0), 0).toFixed(2)),
            totalPreStockQty: Number(totalData?.preStockQty || skuRows.reduce((sum, row) => sum + Number(row["期初库存"] || 0), 0).toFixed(2)),
            totalTurnoverRate: Number(totalAnalysis.stockTurnoverRate || 0),
            totalTurnoverDays: Number(totalAnalysis.stockTurnoverDays || 0),
            totalStockSaleRate: Number(totalAnalysis.stockSaleRate || 0),
            topSkuRow
        };
    }

    function pickSkuFieldSafe(row, keys, fallback = '') {
        for (const key of keys) {
            if (row && row[key] !== undefined) return row[key];
        }
        return fallback;
    }

    function normalizeSkuInventoryRow(row, dim) {
        return {
            ...row,
            __dim: dim || row.__dim || 'detail',
            customerName: pickSkuFieldSafe(row, ['customerName', '客户名称'], ''),
            customerCode: pickSkuFieldSafe(row, ['customerCode', '客户编码'], ''),
            whName: pickSkuFieldSafe(row, ['whName', '发货仓库'], ''),
            whCode: pickSkuFieldSafe(row, ['whCode', '仓库代码'], ''),
            sku: pickSkuFieldSafe(row, ['sku', 'SKU'], ''),
            productName: pickSkuFieldSafe(row, ['productName', '产品名称'], ''),
            skuCount: Number(pickSkuFieldSafe(row, ['skuCount', 'SKU数'], 0)),
            warehouseCount: Number(pickSkuFieldSafe(row, ['warehouseCount', '仓库数'], 0)),
            customerCount: Number(pickSkuFieldSafe(row, ['customerCount', '客户数'], 0)),
            preStockQty: Number(pickSkuFieldSafe(row, ['preStockQty', '期初库存'], 0)),
            closeStockQty: Number(pickSkuFieldSafe(row, ['closeStockQty', '期末库存'], 0)),
            outboundBookQty: Number(pickSkuFieldSafe(row, ['outboundBookQty', '出库预占'], 0)),
            stockTurnoverRate: Number(pickSkuFieldSafe(row, ['stockTurnoverRate', '库存周转率'], 0)),
            stockTurnoverDays: Number(pickSkuFieldSafe(row, ['stockTurnoverDays', '库存周转天数'], 0)),
            stockSaleRate: Number(pickSkuFieldSafe(row, ['stockSaleRate', '库存售罄率'], 0))
        };
    }

    function normalizeSkuInventoryReport(report) {
        if (!report) return report;
        return {
            ...report,
            detailRows: (report.detailRows || []).map((row) => normalizeSkuInventoryRow(row, 'detail')),
            customerRows: (report.customerRows || []).map((row) => normalizeSkuInventoryRow(row, 'customer')),
            skuRows: (report.skuRows || []).map((row) => normalizeSkuInventoryRow(row, 'sku')),
            topSkuRow: report.topSkuRow ? normalizeSkuInventoryRow(report.topSkuRow, 'sku') : null
        };
    }

    function reportSkuInventoryRenderError(stage, error) {
        logger.error(`SKU库存分析${stage}失败`, error);
        const message = error && error.message ? error.message : String(error);
        const statusEl = document.getElementById('sr-status');
        const summaryEl = document.getElementById('sr-sku-inventory-summary');
        if (statusEl) statusEl.innerText = `❌ SKU库存分析${stage}失败: ${message}`;
        if (summaryEl) summaryEl.innerText = `渲染异常: ${message}`;
        const tbody = document.querySelector('#sr-sku-inventory-table tbody');
        if (tbody) tbody.innerHTML = `<tr><td colspan="10" style="color:#ff4d4f; padding:30px 0;">渲染异常: ${message}</td></tr>`;
    }

    function normalizeSalesText(value, fallback = '') {
        const text = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
        return text || fallback;
    }

    function formatChartNumber(value) {
        const num = Number(value || 0);
        return Number.isInteger(num) ? `${num}` : num.toFixed(2);
    }

    function sanitizeFileName(name, fallback = 'report') {
        const text = String(name == null ? '' : name).trim() || fallback;
        return text.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
    }

    function escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function escapeAttribute(value) {
        return escapeHtml(value).replace(/`/g, '&#96;');
    }

    function normalizeEditableRateValue(value) {
        const text = String(value == null ? '' : value).trim();
        if (!text) return 0;
        const numeric = Number(text.replace(/%/g, ''));
        if (!Number.isFinite(numeric)) throw new Error(`无法识别的出库效率数值：${value}`);
        return Number(numeric.toFixed(2));
    }

    function cloneOutboundRows(rows = []) {
        return (rows || []).map((row) => ({ ...row }));
    }

    function normalizeAdviceText(value, fallback = '') {
        return String(value == null ? fallback : value).trim();
    }

    function downloadBlob(blob, filename) {
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function getAnalysisDevCacheKey(moduleKey) {
        return `${ANALYSIS_DEV_CACHE_PREFIX}_${moduleKey}`;
    }

    function getAnalysisDevModuleLabel(moduleKey) {
        return {
            outbound: '出库发货率',
            region: '订单分布分析',
            inventory: '按客户库存分析',
            skuInventory: 'SKU库存分析',
            skuSales: 'SKU销量分析'
        }[moduleKey] || moduleKey;
    }

    function getAnalysisDevCachePayload(moduleKey) {
        if (moduleKey === 'outbound') return outboundReportData?.rows?.length ? outboundReportData : null;
        if (moduleKey === 'region') return regionReportData;
        if (moduleKey === 'inventory') return inventoryReportData;
        if (moduleKey === 'skuInventory') return skuInventoryReportData;
        if (moduleKey === 'skuSales') return skuSalesReportData;
        return null;
    }

    function setAnalysisDevCachePayload(moduleKey, payload) {
        if (moduleKey === 'outbound') {
            outboundReportData = payload || { rows: [], startDate: '', endDate: '', skipWeekends: false, summary: { warehouseCount: 0, avg24: 0, avg48: 0, avg72: 0 } };
            return;
        }
        if (moduleKey === 'region') {
            regionReportData = payload;
            return;
        }
        if (moduleKey === 'inventory') {
            inventoryReportData = payload;
            return;
        }
        if (moduleKey === 'skuInventory') {
            skuInventoryReportData = payload ? normalizeSkuInventoryReport(payload) : payload;
            return;
        }
        if (moduleKey === 'skuSales') {
            skuSalesReportData = payload;
        }
    }

    function renderAnalysisDevModule(moduleKey) {
        if (moduleKey === 'outbound') {
            if (outboundReportData?.rows?.length) renderOutboundRateView();
            else resetOutboundRateView();
        } else if (moduleKey === 'region') {
            if (regionReportData) renderRegionDistributionView();
            else resetRegionDistributionView();
        } else if (moduleKey === 'inventory') {
            if (inventoryReportData) renderInventoryAnalysisView();
            else resetInventoryAnalysisView();
        } else if (moduleKey === 'skuInventory') {
            if (skuInventoryReportData) renderSkuInventoryAnalysisView();
            else resetSkuInventoryAnalysisView();
        } else if (moduleKey === 'skuSales') {
            if (skuSalesReportData) renderSkuSalesAnalysisView();
            else resetSkuSalesAnalysisView();
        }
        updateSalesReportReadyState();
    }

    function saveAnalysisDevCache(moduleKey) {
        const payload = getAnalysisDevCachePayload(moduleKey);
        if (!payload) {
            alert(`当前没有可保存的${getAnalysisDevModuleLabel(moduleKey)}数据`);
            return;
        }
        const body = {
            moduleKey,
            savedAt: new Date().toISOString(),
            payload
        };
        localStorage.setItem(getAnalysisDevCacheKey(moduleKey), JSON.stringify(body));
        const statusEl = document.getElementById('sr-status');
        if (statusEl) statusEl.innerText = `已保存开发模式缓存：${getAnalysisDevModuleLabel(moduleKey)}`;
    }

    function loadAnalysisDevCache(moduleKey) {
        const raw = localStorage.getItem(getAnalysisDevCacheKey(moduleKey));
        if (!raw) {
            alert(`本地没有已保存的${getAnalysisDevModuleLabel(moduleKey)}缓存数据`);
            return;
        }
        try {
            const parsed = JSON.parse(raw);
            setAnalysisDevCachePayload(moduleKey, parsed?.payload || null);
            renderAnalysisDevModule(moduleKey);
            const statusEl = document.getElementById('sr-status');
            const timeText = parsed?.savedAt ? `，保存时间：${new Date(parsed.savedAt).toLocaleString('zh-CN')}` : '';
            if (statusEl) statusEl.innerText = `已加载开发模式缓存：${getAnalysisDevModuleLabel(moduleKey)}${timeText}`;
        } catch (error) {
            logger.error(`加载开发模式缓存失败: ${moduleKey}`, error);
            alert(`加载开发模式缓存失败: ${error.message}`);
        }
    }

    function clearAnalysisDevCache(moduleKey) {
        localStorage.removeItem(getAnalysisDevCacheKey(moduleKey));
        const statusEl = document.getElementById('sr-status');
        if (statusEl) statusEl.innerText = `已清空开发模式缓存：${getAnalysisDevModuleLabel(moduleKey)}`;
    }

    function bindAnalysisDevControls(moduleKey) {
        const saveBtn = document.getElementById(`sr-dev-save-${moduleKey}`);
        const loadBtn = document.getElementById(`sr-dev-load-${moduleKey}`);
        const clearBtn = document.getElementById(`sr-dev-clear-${moduleKey}`);
        if (saveBtn) saveBtn.onclick = () => saveAnalysisDevCache(moduleKey);
        if (loadBtn) loadBtn.onclick = () => loadAnalysisDevCache(moduleKey);
        if (clearBtn) clearBtn.onclick = () => clearAnalysisDevCache(moduleKey);
    }

    function injectAnalysisDevControls() {
        if (!ENABLE_ANALYSIS_DEV_MODE) return;
        const modules = [
            { key: 'outbound', anchorId: 'sr-view-outbound' },
            { key: 'region', anchorId: 'sr-view-region' },
            { key: 'inventory', anchorId: 'sr-view-inventory' },
            { key: 'skuInventory', anchorId: 'sr-view-sku-inventory' },
            { key: 'skuSales', anchorId: 'sr-view-sku-sales' }
        ];

        modules.forEach(({ key, anchorId }) => {
            const view = document.getElementById(anchorId);
            if (!view || view.querySelector(`[data-dev-module="${key}"]`)) return;
            const firstControl = view.querySelector('.sr-controls');
            if (!firstControl) return;
            const panel = document.createElement('div');
            panel.className = 'sr-controls sr-dev-controls';
            panel.dataset.devModule = key;
            panel.style.cssText = 'background:#fff1f0; padding:10px 12px; border-radius:6px; border:1px dashed #ff7875; margin-top:-4px; margin-bottom:15px; flex-wrap:wrap;';
            panel.innerHTML = `
                <span style="font-size:13px; font-weight:bold; color:#cf1322;">开发模式</span>
                <span style="font-size:12px; color:#8c8c8c;">保存当前分析结果到本地，便于反复调试，无需重新抓取。</span>
                <button class="sr-btn sr-btn-danger" id="sr-dev-save-${key}" type="button">保存抓取数据</button>
                <button class="sr-btn sr-btn-info" id="sr-dev-load-${key}" type="button">加载抓取数据</button>
                <button class="sr-btn sr-btn-warning" id="sr-dev-clear-${key}" type="button">清空抓取数据</button>
            `;
            firstControl.insertAdjacentElement('afterend', panel);
        });
    }

    function syncPageJumpControl(inputId, buttonId, totalPages, currentPage, disabled) {
        const input = document.getElementById(inputId);
        const button = document.getElementById(buttonId);
        const safeTotalPages = Math.max(1, Number(totalPages) || 1);
        const safeCurrentPage = Math.min(Math.max(1, Number(currentPage) || 1), safeTotalPages);
        const isDisabled = !!disabled;
        if (input) {
            input.dataset.totalPages = String(safeTotalPages);
            input.dataset.currentPage = String(safeCurrentPage);
            input.placeholder = `1-${safeTotalPages}`;
            input.value = '';
            input.disabled = isDisabled;
            input.title = isDisabled ? '暂无可跳转页码' : `输入 1-${safeTotalPages} 的页码`;
        }
        if (button) {
            button.disabled = isDisabled;
        }
    }

    function jumpToTablePage(tableState, inputId, renderFn) {
        const input = document.getElementById(inputId);
        if (!input) return;
        const totalPages = Math.max(1, Number(input.dataset.totalPages || 1));
        const parsedPage = Number(input.value);
        if (!Number.isFinite(parsedPage)) {
            input.focus();
            input.select();
            return;
        }
        tableState.page = Math.min(Math.max(1, Math.floor(parsedPage)), totalPages);
        renderFn();
        input.value = '';
    }

    function bindTablePageJump(inputId, buttonId, tableState, renderFn) {
        const input = document.getElementById(inputId);
        const button = document.getElementById(buttonId);
        if (!input || !button) return;
        button.onclick = () => jumpToTablePage(tableState, inputId, renderFn);
        input.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                jumpToTablePage(tableState, inputId, renderFn);
            }
        });
    }

    function exportInventoryAnalysisExcel() {
        if (!window.XLSX) return alert('Excel组件加载中...');
        if (!inventoryReportData) return alert('请先完成按客户库存分析！');

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inventoryReportData.detailRows || []), '客户分仓明细');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inventoryReportData.customerRows || []), '按客户汇总');

        const summaryRows = [{
            '开始日期': inventoryReportData.startStr,
            '结束日期': inventoryReportData.endStr,
            '刷新日期': inventoryReportData.refreshTime || '',
            '客户数': inventoryReportData.customerCount || 0,
            '客户分仓组合数': inventoryReportData.detailCount || 0,
            '总仓库数': inventoryReportData.totalWarehouseCount || 0,
            '期初库存': inventoryReportData.totalPreStockQty || 0,
            '期末库存': inventoryReportData.totalCloseStockQty || 0,
            '总周转率': inventoryReportData.totalTurnoverRate || 0,
            '总周转天数': inventoryReportData.totalTurnoverDays || 0,
            '总库存售罄率': inventoryReportData.totalStockSaleRate || 0
        }];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), '汇总指标');
        XLSX.writeFile(wb, `按客户库存分析_${inventoryReportData.startStr}_至_${inventoryReportData.endStr}.xlsx`);
    }

    function exportSkuInventoryAnalysisExcel() {
        if (!window.XLSX) return alert('Excel组件加载中...');
        if (!skuInventoryReportData) return alert('请先完成按sku库存分析！');

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skuInventoryReportData.detailRows || []), '客户SKU明细');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skuInventoryReportData.customerRows || []), '按客户汇总');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skuInventoryReportData.skuRows || []), '按SKU汇总');

        const summaryRows = [{
            '开始日期': skuInventoryReportData.startStr,
            '结束日期': skuInventoryReportData.endStr,
            '刷新日期': skuInventoryReportData.refreshTime || '',
            '客户数': skuInventoryReportData.customerCount || 0,
            'SKU数': skuInventoryReportData.skuCount || 0,
            '明细组合数': skuInventoryReportData.detailCount || 0,
            '期初库存': skuInventoryReportData.totalPreStockQty || 0,
            '期末库存': skuInventoryReportData.totalCloseStockQty || 0,
            '总周转率': skuInventoryReportData.totalTurnoverRate || 0,
            '总周转天数': skuInventoryReportData.totalTurnoverDays || 0,
            '总库存售罄率': skuInventoryReportData.totalStockSaleRate || 0
        }];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), '汇总指标');
        XLSX.writeFile(wb, `按SKU库存分析_${skuInventoryReportData.startStr}_至_${skuInventoryReportData.endStr}.xlsx`);
    }

    function exportSkuSalesAnalysisExcel() {
        if (!window.XLSX) return alert('Excel组件加载中...');
        if (!skuSalesReportData) return alert('请先完成sku销量分析！');

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skuSalesReportData.detailRows || []), '客户SKU明细');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skuSalesReportData.customerRows || []), '按客户汇总');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skuSalesReportData.skuRows || []), '按SKU汇总');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skuSalesReportData.dailyRows || []), '按日销量趋势');

        const summaryRows = [{
            '开始日期': skuSalesReportData.startStr,
            '结束日期': skuSalesReportData.endStr,
            '客户数': skuSalesReportData.customerCount || 0,
            'SKU数': skuSalesReportData.skuCount || 0,
            '客户SKU组合数': skuSalesReportData.detailCount || 0,
            '订单商品行数': skuSalesReportData.orderCount || 0,
            '总销量': skuSalesReportData.totalQty || 0,
            '销量最高客户': skuSalesReportData.topCustomer?.customerName || '',
            '最高客户销量': skuSalesReportData.topCustomer?.totalQty || 0,
            '销量最高SKU': skuSalesReportData.topSku?.sku || '',
            '最高SKU销量': skuSalesReportData.topSku?.totalQty || 0
        }];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), '汇总指标');
        XLSX.writeFile(wb, `SKU销量分析_${skuSalesReportData.startStr}_至_${skuSalesReportData.endStr}.xlsx`);
    }

    function getReportCustomerKey(code, name) {
        const safeCode = String(code || '').trim();
        const safeName = String(name || '').trim();
        return `${safeCode}__${safeName}`;
    }

    function getSalesReportSourceDateRange() {
        return {
            region: regionReportData ? `${regionReportData.startStr} 至 ${regionReportData.endStr}` : '-',
            inventory: inventoryReportData ? `${inventoryReportData.startStr} 至 ${inventoryReportData.endStr}` : '-',
            skuInventory: skuInventoryReportData ? `${skuInventoryReportData.startStr} 至 ${skuInventoryReportData.endStr}` : '-',
            skuSales: skuSalesReportData ? `${skuSalesReportData.startStr} 至 ${skuSalesReportData.endStr}` : '-',
            outboundEfficiency: outboundReportData?.rows?.length ? `${outboundReportData.startDate} 至 ${outboundReportData.endDate}` : '-'
        };
    }

    const SALES_REPORT_MODULE_META = {
        region: { label: '订单分布分析', tagLabel: '地区周期' },
        inventory: { label: '按客户库存分析', tagLabel: '客户库存周期' },
        skuInventory: { label: '按SKU库存分析', tagLabel: 'SKU库存周期' },
        skuSales: { label: 'SKU销量分析', tagLabel: 'SKU销量周期' },
        outboundEfficiency: { label: '出库效率', tagLabel: '出库效率周期' }
    };

    function getSalesReportReadyModules() {
        return {
            region: !!regionReportData,
            inventory: !!inventoryReportData,
            skuInventory: !!skuInventoryReportData,
            skuSales: !!skuSalesReportData,
            outboundEfficiency: !!outboundReportData?.rows?.length
        };
    }

    function getSalesReportModuleKeysByState(state, expected = true) {
        return Object.keys(SALES_REPORT_MODULE_META).filter((key) => Boolean(state?.[key]) === expected);
    }

    function getSalesReportModuleLabels(moduleKeys) {
        return (moduleKeys || []).map((key) => SALES_REPORT_MODULE_META[key]?.label || key);
    }

    function getSalesReportDateRangeTags(moduleKeys, ranges = getSalesReportSourceDateRange()) {
        const rangeItems = (moduleKeys || [])
            .map((key) => ({
                key,
                label: SALES_REPORT_MODULE_META[key]?.tagLabel || key,
                range: ranges?.[key] || '-'
            }))
            .filter((item) => item.range && item.range !== '-');
        const uniqueRanges = Array.from(new Set(rangeItems.map((item) => item.range)));
        if (rangeItems.length > 1 && uniqueRanges.length === 1) {
            return [`统计周期：${uniqueRanges[0]}`];
        }
        return rangeItems.map((item) => `${item.label}：${item.range}`);
    }

    function buildSalesReportReadyDateRangeText(moduleKeys, ranges = getSalesReportSourceDateRange()) {
        const rangeTags = getSalesReportDateRangeTags(moduleKeys, ranges);
        if (rangeTags.length === 0) return '';
        return `数据周期：${rangeTags.join('；')}。`;
    }

    function getCustomerNameInfoKey(code, name) {
        const safeCode = String(code || '').trim();
        const safeName = String(name || '').trim();
        return `${safeCode}__${safeName}`;
    }

    function buildCustomerReportTitle(customerName, companyName) {
        const displayName = buildCustomerDisplayName(customerName, companyName);
        return `${displayName} 销售报告`;
    }

    function buildCustomerDisplayName(customerName, companyName) {
        const safeCustomerName = normalizeSalesText(customerName, '-');
        const safeCompanyName = normalizeSalesText(companyName, '');
        return safeCompanyName ? `${safeCustomerName}（${safeCompanyName}）` : safeCustomerName;
    }

    function getCustomerNameInfo(code, name) {
        if (!salesReportCustomerNameMap) return null;
        const safeCode = String(code || '').trim();
        const safeName = String(name || '').trim();
        return salesReportCustomerNameMap.byKey.get(getCustomerNameInfoKey(safeCode, safeName))
            || (safeCode ? salesReportCustomerNameMap.byCode.get(safeCode) : null)
            || (safeName ? salesReportCustomerNameMap.byName.get(safeName) : null)
            || null;
    }

    function getSalesReportDisplayName(code, name) {
        const info = getCustomerNameInfo(code, name);
        return buildCustomerDisplayName(name, info?.companyName || '');
    }

    function buildSalesReportCustomerNameMap(records) {
        const byKey = new Map();
        const byCode = new Map();
        const byName = new Map();
        (records || []).forEach((record) => {
            const customerCode = normalizeSalesText(record.customerCode, '');
            const customerName = normalizeSalesText(record.customerName, '');
            const companyName = normalizeSalesText(record.companyName, '');
            if (!customerName && !customerCode) return;
            const info = { customerCode, customerName, companyName };
            byKey.set(getCustomerNameInfoKey(customerCode, customerName), info);
            if (customerCode) byCode.set(customerCode, info);
            if (customerName) byName.set(customerName, info);
        });
        return { byKey, byCode, byName, count: records.length, loadedAt: new Date().toISOString() };
    }

    function updateSalesReportCustomerNameState() {
        const statusEl = document.getElementById('sr-sales-report-cn-name-status');
        const checkbox = document.getElementById('sr-sales-report-use-cn-name');
        if (checkbox) checkbox.disabled = !salesReportCustomerNameMap?.count;
        if (!statusEl) return;
        if (salesReportCustomerNameMap?.count) {
            const timeText = salesReportCustomerNameMap.loadedAt
                ? `，获取时间 ${new Date(salesReportCustomerNameMap.loadedAt).toLocaleString('zh-CN')}`
                : '';
            statusEl.innerText = `已获取 ${salesReportCustomerNameMap.count} 个客户中文名称${timeText}`;
        } else {
            statusEl.innerText = '未获取客户中文名称，标题默认使用客户名称。';
        }
    }

    function buildSalesReportCustomerOptions() {
        const readyModules = getSalesReportReadyModules();
        if (!Object.values(readyModules).some(Boolean)) return [];

        const customerMap = new Map();
        const registerCustomer = (code, name, source) => {
            const key = getReportCustomerKey(code, name);
            if (!customerMap.has(key)) {
                customerMap.set(key, {
                    key,
                    customerCode: code || '',
                    customerName: name || '',
                    sources: new Set()
                });
            }
            customerMap.get(key).sources.add(source);
        };

        if (regionReportData) (regionReportData.customerRows || []).forEach((row) => registerCustomer(row["客户编码"], row["客户名称"], 'region'));
        if (inventoryReportData) (inventoryReportData.customerRows || []).forEach((row) => registerCustomer(row["客户编码"], row["客户名称"], 'inventory'));
        if (skuInventoryReportData) (skuInventoryReportData.customerRows || []).forEach((row) => registerCustomer(row.customerCode, row.customerName, 'skuInventory'));
        if (skuSalesReportData) (skuSalesReportData.customerRows || []).forEach((row) => registerCustomer(row.customerCode, row.customerName, 'skuSales'));
        if (outboundReportData?.rows?.length && customerMap.size === 0) {
            registerCustomer('', '全局出库效率', 'outboundEfficiency');
        }

        return Array.from(customerMap.values())
            .sort((a, b) =>
                a.customerName.localeCompare(b.customerName, 'zh-Hans-CN') ||
                a.customerCode.localeCompare(b.customerCode, 'zh-Hans-CN')
            );
    }

    function setSalesReportActionState(disabled) {
        const exportBtn = document.getElementById('sr-sales-report-export-btn');
        const exportPdfBtn = document.getElementById('sr-sales-report-export-pdf-btn');
        const exportAllBtn = document.getElementById('sr-sales-report-export-all-btn');
        const exportAllPdfBtn = document.getElementById('sr-sales-report-export-all-pdf-btn');
        if (exportBtn) exportBtn.disabled = disabled;
        if (exportPdfBtn) exportPdfBtn.disabled = disabled;
        if (exportAllBtn) exportAllBtn.disabled = disabled;
        if (exportAllPdfBtn) exportAllPdfBtn.disabled = disabled;
    }

    function updateSalesReportReadyState() {
        const readyEl = document.getElementById('sr-sales-report-ready');
        const customerSelect = document.getElementById('sr-sales-report-customer');
        if (!readyEl || !customerSelect) return;
        const previousValue = customerSelect.value || '';

        const readyModules = getSalesReportReadyModules();
        const readyModuleKeys = getSalesReportModuleKeysByState(readyModules, true);
        const missingModuleKeys = getSalesReportModuleKeysByState(readyModules, false);

        customerSelect.innerHTML = '<option value="">请选择客户</option>';

        if (readyModuleKeys.length === 0) {
            readyEl.innerText = '请先完成至少一个分析模块，再生成销售报告。客户若缺少某模块数据，报告会自动跳过该模块。';
            setSalesReportActionState(true);
            return;
        }

        const customers = buildSalesReportCustomerOptions();
        if (customers.length === 0) {
            readyEl.innerText = '当前已完成的分析模块中还没有可用于生成报告的客户数据。';
            setSalesReportActionState(true);
            return;
        }

        customers.forEach((customer) => {
            const option = document.createElement('option');
            option.value = customer.key;
            option.textContent = document.getElementById('sr-sales-report-use-cn-name')?.checked
                ? getSalesReportDisplayName(customer.customerCode, customer.customerName)
                : customer.customerName;
            customerSelect.appendChild(option);
        });
        if (previousValue && customers.some((customer) => customer.key === previousValue)) {
            customerSelect.value = previousValue;
        }

        const ranges = getSalesReportSourceDateRange();
        const readyText = getSalesReportModuleLabels(readyModuleKeys).join('、');
        const missingText = missingModuleKeys.length > 0 ? `；未接入模块：${getSalesReportModuleLabels(missingModuleKeys).join('、')}` : '';
        readyEl.innerText = `已就绪，共 ${customers.length} 个客户可生成报告。已接入模块：${readyText}${missingText}。客户若缺少某模块数据，报告会自动跳过对应客户模块；出库效率为全局仓库维度模块。${buildSalesReportReadyDateRangeText(readyModuleKeys, ranges)}`;
        setSalesReportActionState(false);
    }

    function getSalesReportCustomerBundle(customerKey) {
        if (!customerKey) return null;

        const [customerCode = '', customerName = ''] = customerKey.split('__');
        const regionCustomerRow = regionReportData
            ? (regionReportData.customerRows || []).find((row) => getReportCustomerKey(row["客户编码"], row["客户名称"]) === customerKey)
            : null;
        const regionDetailRows = regionReportData
            ? (regionReportData.detailRows || []).filter((row) => getReportCustomerKey(row["客户编码"], row["客户名称"]) === customerKey)
            : [];
        const inventoryCustomerRow = inventoryReportData
            ? (inventoryReportData.customerRows || []).find((row) => getReportCustomerKey(row["客户编码"], row["客户名称"]) === customerKey)
            : null;
        const inventoryDetailRows = inventoryReportData
            ? (inventoryReportData.detailRows || []).filter((row) => getReportCustomerKey(row["客户编码"], row["客户名称"]) === customerKey)
            : [];
        const skuInventoryCustomerRow = skuInventoryReportData
            ? (skuInventoryReportData.customerRows || []).find((row) => getReportCustomerKey(row.customerCode, row.customerName) === customerKey)
            : null;
        const skuInventoryDetailRows = skuInventoryReportData
            ? (skuInventoryReportData.detailRows || []).filter((row) => getReportCustomerKey(row.customerCode, row.customerName) === customerKey)
            : [];
        const skuSalesCustomerRow = skuSalesReportData
            ? (skuSalesReportData.customerRows || []).find((row) => getReportCustomerKey(row.customerCode, row.customerName) === customerKey)
            : null;
        const skuSalesDetailRows = skuSalesReportData
            ? (skuSalesReportData.detailRows || []).filter((row) => getReportCustomerKey(row.customerCode, row.customerName) === customerKey)
            : [];
        const skuSalesChart = skuSalesReportData
            ? (skuSalesReportData.customerCharts || []).find((row) => getReportCustomerKey(row.customerCode, row.customerName) === customerKey)
            : null;

        const availableModules = {
            region: !!regionCustomerRow,
            inventory: !!inventoryCustomerRow,
            skuInventory: !!skuInventoryCustomerRow,
            skuSales: !!skuSalesCustomerRow,
            outboundEfficiency: !!outboundReportData?.rows?.length
        };
        const availableModuleKeys = getSalesReportModuleKeysByState(availableModules, true);
        if (availableModuleKeys.length === 0) return null;
        const missingModuleKeys = getSalesReportModuleKeysByState(availableModules, false);
        const resolvedCustomerName = String(
            regionCustomerRow?.["客户名称"]
            || inventoryCustomerRow?.["客户名称"]
            || skuInventoryCustomerRow?.customerName
            || skuSalesCustomerRow?.customerName
            || customerName
            || ''
        ).trim() || '-';
        const resolvedCustomerCode = String(
            regionCustomerRow?.["客户编码"]
            || inventoryCustomerRow?.["客户编码"]
            || skuInventoryCustomerRow?.customerCode
            || skuSalesCustomerRow?.customerCode
            || customerCode
            || ''
        ).trim();
        const customerNameInfo = document.getElementById('sr-sales-report-use-cn-name')?.checked
            ? getCustomerNameInfo(resolvedCustomerCode, resolvedCustomerName)
            : null;
        const customerCompanyName = customerNameInfo?.companyName || '';
        const reportTitle = buildCustomerReportTitle(resolvedCustomerName, customerCompanyName);

        return {
            customerKey,
            customerCode: resolvedCustomerCode,
            customerName: resolvedCustomerName,
            customerCompanyName,
            reportTitle,
            regionCustomerRow,
            regionDetailRows,
            inventoryCustomerRow,
            inventoryDetailRows,
            skuInventoryCustomerRow,
            skuInventoryDetailRows,
            skuSalesCustomerRow,
            skuSalesDetailRows,
            skuSalesChart,
            outboundEfficiencyRows: sortOutboundRowsByWarehouseName(outboundReportData?.rows || []),
            outboundEfficiencySummary: outboundReportData?.summary || null,
            availableModules,
            availableModuleKeys,
            missingModuleKeys,
            dateRanges: getSalesReportSourceDateRange()
        };
    }

    function getTopRegionForCustomer(regionCustomerRow) {
        if (!regionCustomerRow) return { region: '-', value: 0 };
        return REGION_COLUMNS
            .map((region) => ({ region, value: Number(regionCustomerRow?.[region] || 0) }))
            .sort((a, b) => b.value - a.value || a.region.localeCompare(b.region, 'zh-Hans-CN'))[0] || { region: '-', value: 0 };
    }

    function buildSalesReportNarratives(bundle) {
        const topRegion = getTopRegionForCustomer(bundle.regionCustomerRow);
        const topInventorySku = (bundle.skuInventoryDetailRows || [])
            .slice()
            .sort((a, b) => Number(b.closeStockQty || 0) - Number(a.closeStockQty || 0) || String(a.sku || '').localeCompare(String(b.sku || ''), 'zh-Hans-CN'))[0] || null;
        const topSalesSku = (bundle.skuSalesDetailRows || [])
            .slice()
            .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0) || String(a.sku || '').localeCompare(String(b.sku || ''), 'zh-Hans-CN'))[0] || null;
        const availableLabels = getSalesReportModuleLabels(bundle.availableModuleKeys);
        const missingLabels = getSalesReportModuleLabels(bundle.missingModuleKeys);
        const outboundSummary = bundle.outboundEfficiencySummary || {};
        const overviewParts = [
            `本报告围绕 ${bundle.customerName} 的订单分布、库存结构、SKU动销及出库效率进行汇总分析，当前包含：${availableLabels.join('、')}。`
        ];
        if (bundle.skuSalesCustomerRow) {
            overviewParts.push(`统计周期内共有 ${bundle.skuSalesCustomerRow.orderCount} 个订单商品行，累计销量 ${formatChartNumber(bundle.skuSalesCustomerRow.totalQty)}，覆盖 ${bundle.skuSalesCustomerRow.skuCount} 个SKU。`);
        } else if (bundle.inventoryCustomerRow) {
            overviewParts.push(`当前客户期末库存 ${formatChartNumber(bundle.inventoryCustomerRow["期末库存"])}，库存周转率 ${formatChartNumber(bundle.inventoryCustomerRow["库存周转率"])}%。`);
        } else if (bundle.regionCustomerRow) {
            overviewParts.push(`当前客户订单量 ${formatChartNumber(bundle.regionCustomerRow["总订单量"] || 0)}，主力地区为 ${topRegion.region}。`);
        } else if (bundle.skuInventoryCustomerRow) {
            overviewParts.push(`当前客户库存SKU数 ${formatChartNumber(bundle.skuInventoryCustomerRow.skuCount || 0)}。`);
        }
        if (missingLabels.length > 0) {
            overviewParts.push(`本期暂无 ${missingLabels.join('、')} 数据，因此正文未展示对应分析。`);
        }

        return {
            overview: overviewParts.join(' '),
            region: bundle.regionCustomerRow
                ? `订单区域以 ${topRegion.region} 为主，共 ${topRegion.value} 单；邮编信息不完整或详情暂不可用的订单未纳入地区统计。`
                : '',
            inventory: bundle.inventoryCustomerRow
                ? `客户库存汇总显示期末库存 ${formatChartNumber(bundle.inventoryCustomerRow["期末库存"])}，库存周转率 ${formatChartNumber(bundle.inventoryCustomerRow["库存周转率"])}%，库存周转天数 ${formatChartNumber(bundle.inventoryCustomerRow["库存周转天数"])}。`
                : '',
            skuInventory: topInventorySku
                ? `SKU库存结构中，当前库存最高的SKU为 ${topInventorySku.sku}，期末库存 ${formatChartNumber(topInventorySku.closeStockQty)}。`
                : '',
            skuSales: topSalesSku
                ? `SKU销量结构中，销量最高SKU为 ${topSalesSku.sku}，销量 ${formatChartNumber(topSalesSku.qty)}。`
                : '',
            outboundEfficiency: bundle.availableModules?.outboundEfficiency
                ? `出库效率基于全局仓库维度数据，覆盖 ${formatChartNumber(outboundSummary.warehouseCount || 0)} 个仓库；平均 24H / 48H / 72H 出库发货率分别为 ${formatChartNumber(outboundSummary.avg24 || 0)}%、${formatChartNumber(outboundSummary.avg48 || 0)}%、${formatChartNumber(outboundSummary.avg72 || 0)}%。`
                : '',
            missingModules: missingLabels
        };
    }

    function createDetachedReportChart(width = 880, height = 320) {
        const el = document.createElement('div');
        el.style.cssText = `position:fixed; left:-99999px; top:-99999px; width:${width}px; height:${height}px; background:#fff; z-index:-1;`;
        document.body.appendChild(el);
        return el;
    }

    function renderChartDataUrl(option, width = 880, height = 320) {
        if (!window.echarts) return '';
        const el = createDetachedReportChart(width, height);
        const chart = echarts.init(el, null, { renderer: 'canvas' });
        chart.setOption(option, true);
        chart.resize({ width, height });
        const dataUrl = chart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: '#ffffff' });
        chart.dispose();
        el.remove();
        return dataUrl;
    }

    function sortReportItemsDesc(list, getValue, getName) {
        return (Array.isArray(list) ? list : [])
            .slice()
            .sort((a, b) => {
                const diff = Number(getValue(b) || 0) - Number(getValue(a) || 0);
                if (diff !== 0) return diff;
                return String(getName(a) || '').localeCompare(String(getName(b) || ''), 'zh-Hans-CN');
            });
    }

    function buildTopNLabel(baseText, count) {
        return `${baseText} Top ${Math.max(0, Number(count) || 0)}`;
    }

    function buildConditionalTopLabel(baseText, count, topCount = 5) {
        const safeTopCount = Math.max(1, Number(topCount) || 1);
        return Number(count || 0) >= safeTopCount ? `${baseText} Top ${safeTopCount}` : baseText;
    }

    function buildBarValueLabel(suffix = '') {
        return {
            show: true,
            position: 'top',
            formatter: ({ value }) => value == null ? '' : `${formatChartNumber(value)}${suffix}`
        };
    }

    function buildSalesReportImageCard(title, dataUrl, description = '') {
        if (!dataUrl) return '';
        return `
            <div class="chart-card">
                <div class="chart-card-head">
                    <div class="chart-card-title">${escapeHtml(title)}</div>
                    ${description ? `<div class="chart-card-desc">${escapeHtml(description)}</div>` : ''}
                </div>
                <div class="chart"><img src="${dataUrl}" alt="${escapeHtml(title)}" /></div>
            </div>
        `;
    }

    function buildReportPdfContainer(html, mountToBody = false) {
        const wrapper = document.createElement('div');
        wrapper.style.cssText = 'position:fixed; left:0; top:-20000px; width:1120px; background:#fff; z-index:1; pointer-events:none; overflow:visible;';
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const styleNodes = Array.from(doc.querySelectorAll('style'));
        styleNodes.forEach((node) => {
            const styleEl = document.createElement('style');
            styleEl.textContent = node.textContent || '';
            wrapper.appendChild(styleEl);
        });
        const pageNode = doc.body?.firstElementChild;
        if (pageNode) {
            wrapper.insertAdjacentHTML('beforeend', pageNode.outerHTML);
        } else {
            wrapper.insertAdjacentHTML('beforeend', html);
        }
        if (mountToBody) document.body.appendChild(wrapper);
        return wrapper;
    }

    async function waitForSalesReportContainerReady(container) {
        const images = Array.from(container.querySelectorAll('img'));
        await Promise.all(images.map((img) => new Promise((resolve) => {
            if (img.complete && img.naturalWidth > 0) return resolve();
            const done = () => {
                img.removeEventListener('load', done);
                img.removeEventListener('error', done);
                resolve();
            };
            img.addEventListener('load', done, { once: true });
            img.addEventListener('error', done, { once: true });
        })));
        if (document.fonts && document.fonts.ready) {
            try {
                await document.fonts.ready;
            } catch (error) {
                logger.warn('等待字体加载失败，继续导出 PDF', error);
            }
        }
        container.getBoundingClientRect();
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }

    function getSalesReportPdfOptions() {
        return {
            margin: 8,
            imageType: 'jpeg',
            imageQuality: 0.98,
            html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff', windowWidth: 1120, scrollX: 0, scrollY: 0, logging: false },
            pdf: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };
    }

    function getSalesReportPdfRuntime() {
        const globalScope = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
        const html2canvasRef = globalScope.html2canvas || (typeof html2canvas !== 'undefined' ? html2canvas : null);
        const jspdfNamespace = globalScope.jspdf || (typeof jspdf !== 'undefined' ? jspdf : null);
        const jsPDFRef = jspdfNamespace?.jsPDF || globalScope.jsPDF || (typeof jsPDF !== 'undefined' ? jsPDF : null);
        return { html2canvas: html2canvasRef, jsPDF: jsPDFRef };
    }

    async function renderSalesReportBlockCanvas(element, options, runtime) {
        return runtime.html2canvas(element, options.html2canvas);
    }

    function getSalesReportPdfBlocks(page) {
        const pageChildren = Array.from(page.children).filter((node) => node instanceof HTMLElement);
        if (pageChildren.length === 1 && pageChildren[0].classList.contains('dense-shell')) {
            return Array.from(pageChildren[0].children).filter((node) => node instanceof HTMLElement);
        }
        return pageChildren;
    }

    function groupSalesReportGridRows(element) {
        const rect = element.getBoundingClientRect();
        const children = Array.from(element.children)
            .filter((node) => node instanceof HTMLElement)
            .filter((node) => {
                const childRect = node.getBoundingClientRect();
                return childRect.width >= 2 && childRect.height >= 2;
            })
            .sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);

        const rows = [];
        const tolerance = 4;
        children.forEach((child) => {
            const childRect = child.getBoundingClientRect();
            const top = childRect.top - rect.top;
            const bottom = childRect.bottom - rect.top;
            const lastRow = rows[rows.length - 1];
            if (lastRow && Math.abs(lastRow.anchorTop - top) <= tolerance) {
                lastRow.top = Math.min(lastRow.top, top);
                lastRow.bottom = Math.max(lastRow.bottom, bottom);
            } else {
                rows.push({ top, bottom, anchorTop: top });
            }
        });
        return rows;
    }

    async function renderSalesReportGridRowCanvases(element, options, runtime) {
        const rows = groupSalesReportGridRows(element);
        if (rows.length === 0) {
            const fullCanvas = await renderSalesReportBlockCanvas(element, options, runtime);
            return fullCanvas ? [fullCanvas] : [];
        }

        const canvas = await renderSalesReportBlockCanvas(element, options, runtime);
        if (!canvas || !canvas.width || !canvas.height) return [];

        const rect = element.getBoundingClientRect();
        const scaleY = rect.height > 0 ? canvas.height / rect.height : 1;
        const bleedPx = Math.max(4, Math.round(6 * scaleY));

        return rows.map((row, index) => {
            const isFirst = index === 0;
            const isLast = index === rows.length - 1;
            const start = Math.max(0, Math.floor(row.top * scaleY) - (isFirst ? 0 : bleedPx));
            const end = Math.min(canvas.height, Math.ceil(row.bottom * scaleY) + (isLast ? 0 : bleedPx));
            const height = Math.max(1, end - start);
            const rowCanvas = document.createElement('canvas');
            rowCanvas.width = canvas.width;
            rowCanvas.height = height;
            const rowCtx = rowCanvas.getContext('2d');
            if (!rowCtx) return null;
            rowCtx.fillStyle = '#ffffff';
            rowCtx.fillRect(0, 0, rowCanvas.width, rowCanvas.height);
            rowCtx.drawImage(
                canvas,
                0, start, canvas.width, height,
                0, 0, rowCanvas.width, rowCanvas.height
            );
            return rowCanvas;
        }).filter(Boolean);
    }

    function appendSalesReportCanvasToPdf(pdf, canvas, options, state, blockGap = 4) {
        if (!canvas || !canvas.width || !canvas.height) return;
        const { margin, imageType, imageQuality } = options;
        const pageHeight = state.pageHeight;
        const usableWidth = state.usableWidth;
        const usableHeight = state.usableHeight;
        const fullRenderHeightMm = (canvas.height * usableWidth) / canvas.width;

        if (state.wroteAnyContent && fullRenderHeightMm <= usableHeight && state.currentY + fullRenderHeightMm > pageHeight - margin + 0.01) {
            pdf.addPage();
            state.currentY = margin;
        }

        const maxSliceHeightPx = Math.max(1, Math.floor((usableHeight * canvas.width) / usableWidth));
        let offsetY = 0;

        while (offsetY < canvas.height) {
            const sliceHeightPx = Math.min(maxSliceHeightPx, canvas.height - offsetY);
            const sliceCanvas = document.createElement('canvas');
            sliceCanvas.width = canvas.width;
            sliceCanvas.height = sliceHeightPx;
            const sliceCtx = sliceCanvas.getContext('2d');
            if (!sliceCtx) throw new Error('PDF切片画布创建失败');
            sliceCtx.fillStyle = '#ffffff';
            sliceCtx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
            sliceCtx.drawImage(
                canvas,
                0, offsetY, canvas.width, sliceHeightPx,
                0, 0, sliceCanvas.width, sliceCanvas.height
            );

            const renderHeightMm = (sliceHeightPx * usableWidth) / canvas.width;
            if (state.wroteAnyContent && state.currentY + renderHeightMm > pageHeight - margin + 0.01) {
                pdf.addPage();
                state.currentY = margin;
            }

            const imageData = sliceCanvas.toDataURL(`image/${imageType}`, imageQuality);
            pdf.addImage(imageData, imageType.toUpperCase(), margin, state.currentY, usableWidth, renderHeightMm, undefined, 'FAST');
            state.wroteAnyContent = true;
            state.currentY += renderHeightMm;
            offsetY += sliceHeightPx;

            if (offsetY < canvas.height) {
                pdf.addPage();
                state.currentY = margin;
            } else {
                state.currentY += blockGap;
            }
        }
    }

    async function generateSalesReportPdfBlob(html, filename) {
        const runtime = getSalesReportPdfRuntime();
        if (!runtime.html2canvas || !runtime.jsPDF) throw new Error('PDF组件加载失败，请刷新页面后重试');
        const container = buildReportPdfContainer(html, true);
        try {
            await waitForSalesReportContainerReady(container);
            const page = container.querySelector('.page') || container;
            const blocks = getSalesReportPdfBlocks(page);
            const renderBlocks = blocks.length > 0 ? blocks : [page];
            const options = getSalesReportPdfOptions();
            const pdf = new runtime.jsPDF(options.pdf);
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();
            const margin = options.margin;
            const usableWidth = pageWidth - margin * 2;
            const usableHeight = pageHeight - margin * 2;
            const pdfState = {
                currentY: margin,
                wroteAnyContent: false,
                pageHeight,
                usableWidth,
                usableHeight
            };

            for (const block of renderBlocks) {
                if (!(block instanceof HTMLElement)) continue;
                const rect = block.getBoundingClientRect();
                if (rect.width < 2 || rect.height < 2) continue;
                if (block.classList.contains('page-break') && pdfState.wroteAnyContent && pdfState.currentY > margin) {
                    pdf.addPage();
                    pdfState.currentY = margin;
                }

                const blockGap = (block.classList.contains('footer') || block.classList.contains('dense-footer')) ? 0 : 4;
                if (block.classList.contains('dense-grid') || block.classList.contains('dense-grid-compact')) {
                    const rowCanvases = await renderSalesReportGridRowCanvases(block, options, runtime);
                    for (const rowCanvas of rowCanvases) {
                        appendSalesReportCanvasToPdf(pdf, rowCanvas, options, pdfState, blockGap);
                    }
                } else {
                    const canvas = await renderSalesReportBlockCanvas(block, options, runtime);
                    appendSalesReportCanvasToPdf(pdf, canvas, options, pdfState, blockGap);
                }
            }

            if (!pdfState.wroteAnyContent) {
                throw new Error(`PDF内容渲染失败：${filename} 未生成任何有效页面`);
            }

            return pdf.output('blob');
        } finally {
            container.remove();
        }
    }

    function buildSalesReportImages(bundle) {
        const hasRegion = !!bundle.regionCustomerRow;
        const hasInventory = !!bundle.inventoryCustomerRow;
        const hasSkuInventory = !!bundle.skuInventoryCustomerRow;
        const hasSkuSales = !!bundle.skuSalesCustomerRow;
        const hasOutboundEfficiency = !!bundle.availableModules?.outboundEfficiency;

        const topRegionRows = hasRegion
            ? sortReportItemsDesc(
                REGION_COLUMNS.map((region) => ({ region, value: Number(bundle.regionCustomerRow?.[region] || 0) })),
                (item) => item.value,
                (item) => item.region
            ).filter((item) => item.value > 0).slice(0, 5)
            : [];
        const regionWarehouseRows = hasRegion
            ? sortReportItemsDesc(
                bundle.regionDetailRows || [],
                (row) => row["总订单量"],
                (row) => row["发货仓库"]
            ).filter((row) => Number(row["总订单量"] || 0) > 0).slice(0, 6)
            : [];

        const inventoryWarehouseByStockRows = hasInventory
            ? sortReportItemsDesc(
                bundle.inventoryDetailRows || [],
                (row) => row["期末库存"],
                (row) => row["发货仓库"]
            ).slice(0, 8)
            : [];
        const inventoryWarehouseByDaysRows = hasInventory
            ? sortReportItemsDesc(
                (bundle.inventoryDetailRows || []).filter((row) => Number(row["库存周转天数"] || 0) > 0),
                (row) => row["库存周转天数"],
                (row) => row["发货仓库"]
            ).slice(0, 8)
            : [];
        const inventoryWarehouseByRateRows = hasInventory
            ? sortReportItemsDesc(
                (bundle.inventoryDetailRows || []).filter((row) => Number(row["库存周转率"] || 0) > 0),
                (row) => row["库存周转率"],
                (row) => row["发货仓库"]
            ).slice(0, 8)
            : [];
        const inventoryWarehouseShareRows = inventoryWarehouseByStockRows.filter((row) => Number(row["期末库存"] || 0) > 0);

        const skuInventoryByStockRows = hasSkuInventory
            ? sortReportItemsDesc(
                bundle.skuInventoryDetailRows || [],
                (row) => row.closeStockQty,
                (row) => row.sku
            ).slice(0, 5)
            : [];
        const skuInventoryByDaysRows = hasSkuInventory
            ? sortReportItemsDesc(
                (bundle.skuInventoryDetailRows || []).filter((row) => Number(row.stockTurnoverDays || 0) > 0),
                (row) => row.stockTurnoverDays,
                (row) => row.sku
            ).slice(0, 5)
            : [];

        const skuSalesRows = hasSkuSales
            ? sortReportItemsDesc(
                (bundle.skuSalesDetailRows || [])
                    .map((row) => ({ name: row.sku, value: Number(row.qty || 0) })),
                (row) => row.value,
                (row) => row.name
            ).slice(0, 5)
            : [];
        const skuSalesShareRows = hasSkuSales
            ? (() => {
                const rankedRows = sortReportItemsDesc(
                    (bundle.skuSalesDetailRows || [])
                        .map((row) => ({ name: row.sku, value: Number(row.qty || 0) }))
                        .filter((row) => row.value > 0),
                    (row) => row.value,
                    (row) => row.name
                );
                const topRows = rankedRows.slice(0, 5);
                const otherValue = rankedRows.slice(5).reduce((sum, row) => sum + row.value, 0);
                return otherValue > 0
                    ? topRows.concat({ name: '其他SKU', value: otherValue })
                    : topRows;
            })()
            : [];
        const skuSalesDetailRows = hasSkuSales
            ? sortReportItemsDesc(
                bundle.skuSalesDetailRows || [],
                (row) => row.qty,
                (row) => row.sku
            ).slice(0, 5)
            : [];
        const outboundSummary = bundle.outboundEfficiencySummary || {};
        const outboundRows = hasOutboundEfficiency
            ? sortReportItemsDesc(
                bundle.outboundEfficiencyRows || [],
                (row) => row.rate24,
                (row) => row.whName
            )
            : [];
        const outboundTopRows = outboundRows.slice(0, 12);

        const regionWarehousePieRows = regionWarehouseRows.map((row) => ({
            warehouseName: row["发货仓库"] || '未知仓库',
            total: Number(row["总订单量"] || 0),
            data: sortReportItemsDesc(
                REGION_COLUMNS.map((region) => ({ name: region, value: Number(row[region] || 0) })),
                (item) => item.value,
                (item) => item.name
            ).filter((item) => item.value > 0)
        })).filter((item) => item.data.length > 0);

        const chartImages = {
            regionTop: hasRegion && topRegionRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} ${buildConditionalTopLabel('订单地区分布', topRegionRows.length, 5)}`, left: 'center' },
                tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                grid: { left: '4%', right: '4%', bottom: 48, top: 58, containLabel: true },
                xAxis: { type: 'category', data: topRegionRows.map((row) => row.region), axisLabel: { interval: 0, rotate: 16 } },
                yAxis: { type: 'value', name: '订单量' },
                series: [{ type: 'bar', data: topRegionRows.map((row) => row.value), itemStyle: { color: '#fa8c16' }, barMaxWidth: 42, label: buildBarValueLabel() }]
            }) : '',
            inventoryShare: hasInventory && inventoryWarehouseShareRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} 分仓期末库存占比`, left: 'center' },
                tooltip: { trigger: 'item', formatter: '{b}<br/>期末库存: {c}<br/>占比: {d}%' },
                legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 8 },
                series: [{
                    type: 'pie',
                    radius: ['34%', '66%'],
                    center: ['36%', '56%'],
                    data: inventoryWarehouseShareRows.map((row) => ({ name: row["发货仓库"], value: Number(row["期末库存"] || 0) })),
                    label: { formatter: '{b}\n{d}%' }
                }]
            }) : '',
            inventoryStock: hasInventory && inventoryWarehouseByStockRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} 分仓期末库存`, left: 'center' },
                tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                grid: { left: '4%', right: '4%', bottom: 68, top: 58, containLabel: true },
                xAxis: { type: 'category', data: inventoryWarehouseByStockRows.map((row) => row["发货仓库"]), axisLabel: { interval: 0, rotate: 18 } },
                yAxis: { type: 'value', name: '期末库存' },
                series: [{ type: 'bar', data: inventoryWarehouseByStockRows.map((row) => Number(row["期末库存"] || 0)), itemStyle: { color: '#52c41a' }, barMaxWidth: 40, label: buildBarValueLabel() }]
            }) : '',
            inventoryDays: hasInventory && inventoryWarehouseByDaysRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} 分仓库存周转天数`, left: 'center' },
                tooltip: { trigger: 'item', formatter: '{b}<br/>周转天数: {c}<br/>占比: {d}%' },
                legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 8 },
                series: [{
                    type: 'pie',
                    radius: ['34%', '66%'],
                    center: ['36%', '56%'],
                    data: inventoryWarehouseByDaysRows.map((row) => ({ name: row["发货仓库"], value: Number(row["库存周转天数"] || 0) })),
                    label: { formatter: '{b}\n{d}%' }
                }]
            }) : '',
            inventoryRateShare: hasInventory && inventoryWarehouseByRateRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} 分仓库存周转率占比`, left: 'center' },
                tooltip: { trigger: 'item', formatter: '{b}<br/>库存周转率: {c}%<br/>占比: {d}%' },
                legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 8 },
                series: [{
                    type: 'pie',
                    radius: ['34%', '66%'],
                    center: ['36%', '56%'],
                    data: inventoryWarehouseByRateRows.map((row) => ({ name: row["发货仓库"], value: Number(row["库存周转率"] || 0) })),
                    label: { formatter: '{b}\n{c}%' }
                }]
            }) : '',
            inventoryRate: hasInventory && inventoryWarehouseByRateRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} 分仓库存周转率`, left: 'center' },
                tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                grid: { left: '4%', right: '4%', bottom: 68, top: 58, containLabel: true },
                xAxis: { type: 'category', data: inventoryWarehouseByRateRows.map((row) => row["发货仓库"]), axisLabel: { interval: 0, rotate: 18 } },
                yAxis: { type: 'value', name: '周转率(%)' },
                series: [{ type: 'bar', data: inventoryWarehouseByRateRows.map((row) => Number(row["库存周转率"] || 0)), itemStyle: { color: '#389e0d' }, barMaxWidth: 40, label: buildBarValueLabel('%') }]
            }) : '',
            skuInventoryStock: hasSkuInventory && skuInventoryByStockRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} ${buildConditionalTopLabel('SKU库存对比', skuInventoryByStockRows.length, 5)}`, left: 'center' },
                tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                legend: { top: 28 },
                grid: { left: '4%', right: '4%', bottom: 72, top: 78, containLabel: true },
                xAxis: { type: 'category', data: skuInventoryByStockRows.map((row) => row.sku), axisLabel: { interval: 0, rotate: 20 } },
                yAxis: { type: 'value', name: '库存' },
                series: [
                    { name: '期初库存', type: 'bar', data: skuInventoryByStockRows.map((row) => Number(row.preStockQty || 0)), itemStyle: { color: '#69c0ff' }, barMaxWidth: 28, label: buildBarValueLabel() },
                    { name: '期末库存', type: 'bar', data: skuInventoryByStockRows.map((row) => Number(row.closeStockQty || 0)), itemStyle: { color: '#722ed1' }, barMaxWidth: 28, label: buildBarValueLabel() }
                ]
            }) : '',
            skuInventoryDays: hasSkuInventory && skuInventoryByDaysRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} ${buildConditionalTopLabel('SKU库存周转天数', skuInventoryByDaysRows.length, 5)}`, left: 'center' },
                tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                grid: { left: '4%', right: '4%', bottom: 72, top: 58, containLabel: true },
                xAxis: { type: 'category', data: skuInventoryByDaysRows.map((row) => row.sku), axisLabel: { interval: 0, rotate: 20 } },
                yAxis: { type: 'value', name: '周转天数' },
                series: [{ type: 'bar', data: skuInventoryByDaysRows.map((row) => Number(row.stockTurnoverDays || 0)), itemStyle: { color: '#9254de' }, barMaxWidth: 38, label: buildBarValueLabel() }]
            }) : '',
            skuSalesTop: hasSkuSales && skuSalesRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} ${buildConditionalTopLabel('SKU销量', skuSalesRows.length, 5)}`, left: 'center' },
                tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                grid: { left: '4%', right: '4%', bottom: 68, top: 58, containLabel: true },
                xAxis: { type: 'category', data: skuSalesRows.map((row) => row.name), axisLabel: { interval: 0, rotate: 18 } },
                yAxis: { type: 'value', name: '销量' },
                series: [{ type: 'bar', data: skuSalesRows.map((row) => row.value), itemStyle: { color: '#13c2c2' }, barMaxWidth: 40, label: buildBarValueLabel() }]
            }) : '',
            skuSalesShare: hasSkuSales && skuSalesShareRows.length > 0 ? renderChartDataUrl({
                title: { text: `${bundle.customerName} ${buildConditionalTopLabel('SKU销量占比', skuSalesRows.length, 5)}`, left: 'center' },
                tooltip: { trigger: 'item', formatter: '{b}<br/>销量: {c}<br/>占比: {d}%' },
                legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 8 },
                series: [{
                    type: 'pie',
                    radius: ['34%', '66%'],
                    center: ['36%', '56%'],
                    data: skuSalesShareRows.map((row) => ({ name: row.name, value: row.value })),
                    label: { formatter: '{b}\n{d}%' }
                }]
            }) : '',
            outboundEfficiencyTrend: hasOutboundEfficiency && outboundRows.length > 0 ? renderChartDataUrl({
                title: { text: '整体平均出库效率', left: 'center' },
                tooltip: { trigger: 'axis', valueFormatter: (value) => `${Number(value || 0).toFixed(2)}%` },
                grid: { left: '4%', right: '4%', bottom: 48, top: 58, containLabel: true },
                xAxis: { type: 'category', data: ['24H', '48H', '72H'] },
                yAxis: { type: 'value', name: '出库发货率(%)', min: 0, max: 100 },
                series: [{
                    name: '平均出库发货率',
                    type: 'line',
                    data: [outboundSummary.avg24 || 0, outboundSummary.avg48 || 0, outboundSummary.avg72 || 0],
                    itemStyle: { color: '#faad14' },
                    lineStyle: { width: 3 },
                    symbolSize: 8,
                    label: { show: true, position: 'top', formatter: ({ value }) => `${Number(value || 0).toFixed(2)}%` }
                }]
            }) : '',
            outboundEfficiencyWarehouse: hasOutboundEfficiency && outboundTopRows.length > 0 ? renderChartDataUrl({
                title: { text: outboundTopRows.length >= 12 ? 'Top 12 仓库出库效率' : '仓库出库效率', left: 'center' },
                tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: (value) => `${Number(value || 0).toFixed(2)}%` },
                legend: { top: 28 },
                grid: { left: '4%', right: '4%', bottom: 72, top: 78, containLabel: true },
                xAxis: { type: 'category', data: outboundTopRows.map((row) => row.whName), axisLabel: { interval: 0, rotate: 18 } },
                yAxis: { type: 'value', name: '出库发货率(%)', min: 0, max: 100 },
                series: [
                    { name: '24H', type: 'bar', data: outboundTopRows.map((row) => Number(row.rate24 || 0)), itemStyle: { color: '#faad14' }, barMaxWidth: 24, label: buildBarValueLabel('%') },
                    { name: '48H', type: 'bar', data: outboundTopRows.map((row) => Number(row.rate48 || 0)), itemStyle: { color: '#52c41a' }, barMaxWidth: 24, label: buildBarValueLabel('%') },
                    { name: '72H', type: 'bar', data: outboundTopRows.map((row) => Number(row.rate72 || 0)), itemStyle: { color: '#1890ff' }, barMaxWidth: 24, label: buildBarValueLabel('%') }
                ]
            }) : ''
        };

        return {
            region: chartImages.regionTop,
            inventory: chartImages.inventoryStock,
            skuInventory: chartImages.skuInventoryStock,
            skuSales: chartImages.skuSalesTop,
            outboundEfficiency: chartImages.outboundEfficiencyTrend,
            groups: {
                region: [
                    ...regionWarehousePieRows.map((item) => ({
                        key: `regionShareByWarehouse_${sanitizeFileName(item.warehouseName, 'warehouse')}`,
                        title: `${item.warehouseName} 分仓订单地区占比`,
                        description: `当前分仓总订单量 ${formatChartNumber(item.total)}，地区占比按订单量从大到小展示。`,
                        dataUrl: item.data.length > 0 ? renderChartDataUrl({
                            title: { text: `${item.warehouseName} 分仓订单地区占比`, left: 'center' },
                            tooltip: { trigger: 'item', formatter: '{b}<br/>订单量: {c} 单<br/>占比: {d}%' },
                            legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 8 },
                            series: [{
                                name: item.warehouseName,
                                type: 'pie',
                                radius: ['34%', '66%'],
                                center: ['36%', '56%'],
                                data: item.data,
                                label: { formatter: '{b}\n{d}%' }
                            }]
                        }, 880, 320) : ''
                    }))
                ],
                inventory: [
                    {
                        key: 'inventoryShare',
                        title: '分仓期末库存占比',
                        description: '按分仓期末库存从大到小排序后展示各仓库存占比结构。',
                        dataUrl: chartImages.inventoryShare
                    },
                    {
                        key: 'inventoryRateShare',
                        title: '分仓库存周转率占比',
                        description: '按分仓库存周转率从大到小排序后展示各仓周转率结构。',
                        dataUrl: chartImages.inventoryRateShare
                    },
                    {
                        key: 'inventoryDays',
                        title: '分仓库存周转天数',
                        description: '按库存周转天数从大到小排序。',
                        dataUrl: chartImages.inventoryDays
                    }
                ],
                skuInventory: [
                    {
                        key: 'skuInventoryStock',
                        title: buildConditionalTopLabel('SKU库存对比', skuInventoryByStockRows.length, 5),
                        description: `按 SKU 期末库存从大到小排序，当前展示 ${skuInventoryByStockRows.length >= 5 ? 'Top 5' : '全部已统计 SKU'}，并同时展示期初库存与期末库存。`,
                        dataUrl: chartImages.skuInventoryStock
                    },
                    {
                        key: 'skuInventoryDays',
                        title: buildConditionalTopLabel('SKU库存周转天数', skuInventoryByDaysRows.length, 5),
                        description: `按 SKU 周转天数从大到小排序，当前展示 ${skuInventoryByDaysRows.length >= 5 ? 'Top 5' : '全部已统计 SKU'}。`,
                        dataUrl: chartImages.skuInventoryDays
                    }
                ],
                skuSales: [
                    {
                        key: 'skuSalesShare',
                        title: buildConditionalTopLabel('SKU销量占比', skuSalesRows.length, 5),
                        description: `按全部 SKU 销量计算占比，超过 5 个 SKU 时将 Top 5 之外的销量合并为“其他SKU”。`,
                        dataUrl: chartImages.skuSalesShare
                    }
                ],
                outboundEfficiency: [
                    {
                        key: 'outboundEfficiencyTrend',
                        title: '整体平均出库效率',
                        description: '展示本次出库发货率统计周期内 24H、48H、72H 的全仓平均表现。',
                        dataUrl: chartImages.outboundEfficiencyTrend
                    },
                ]
            },
            meta: {
                topRegionRows,
                regionWarehousePieRows,
                inventoryWarehouseByStockRows,
                inventoryWarehouseByRateRows,
                inventoryWarehouseByDaysRows,
                inventoryWarehouseShareRows,
                skuInventoryByStockRows,
                skuInventoryByDaysRows,
                skuSalesRows,
                skuSalesShareRows,
                skuSalesDetailRows,
                outboundRows,
                outboundTopRows
            }
        };
    }

    function buildSalesReportMetrics(bundle) {
        const metrics = [];
        if (bundle.regionCustomerRow) metrics.push({ label: '订单量', value: formatChartNumber(bundle.regionCustomerRow["总订单量"] || 0) });
        if (bundle.inventoryCustomerRow) metrics.push({ label: '期末库存', value: formatChartNumber(bundle.inventoryCustomerRow["期末库存"] || 0) });
        if (bundle.skuInventoryCustomerRow) metrics.push({ label: '库存SKU数', value: formatChartNumber(bundle.skuInventoryCustomerRow.skuCount || 0) });
        if (bundle.skuSalesCustomerRow) metrics.push({ label: '销量', value: formatChartNumber(bundle.skuSalesCustomerRow.totalQty || 0) });
        if (bundle.availableModules?.outboundEfficiency) metrics.push({ label: '24H出库率', value: `${formatChartNumber(bundle.outboundEfficiencySummary?.avg24 || 0)}%` });
        return metrics;
    }

    function buildSalesReportTagItems(bundle) {
        return getSalesReportDateRangeTags(bundle.availableModuleKeys, bundle.dateRanges);
    }

    function calcReportPercent(part, total) {
        const safeTotal = Number(total || 0);
        if (!safeTotal) return 0;
        return Number((Number(part || 0) / safeTotal * 100).toFixed(2));
    }

    function getReportWarehouseName(row) {
        return row?.["发货仓库"] || row?.whName || '-';
    }

    function getInventoryChangeText(preStock, closeStock) {
        const diff = Number(closeStock || 0) - Number(preStock || 0);
        if (diff === 0) return '与期初库存持平';
        const direction = diff > 0 ? '较期初增加' : '较期初减少';
        return `${direction} ${formatChartNumber(Math.abs(diff))}`;
    }

    function buildSalesReportTextBlock(items) {
        const rows = (items || []).filter(Boolean);
        if (rows.length === 0) return '';
        return rows.map((text) => `<p>${escapeHtml(text)}</p>`).join('');
    }

    function buildSalesReportAdviceBlock(moduleKey, options = {}) {
        const text = normalizeAdviceText(salesReportAdviceState?.[moduleKey] || '');
        if (!text) return '';
        const title = options.title || '建议';
        const className = options.className || 'notice';
        return `<div class="${className}"><strong>${escapeHtml(title)}：</strong>${escapeHtml(text)}</div>`;
    }

    function getSalesReportModuleOrder(bundle) {
        return [
            bundle?.regionCustomerRow ? { key: 'region', title: '订单分布分析' } : null,
            bundle?.inventoryCustomerRow ? { key: 'inventory', title: '按客户库存分析' } : null,
            bundle?.skuInventoryCustomerRow ? { key: 'skuInventory', title: '按SKU库存分析' } : null,
            bundle?.skuSalesCustomerRow ? { key: 'skuSales', title: 'SKU销量分析' } : null,
            bundle?.availableModules?.outboundEfficiency ? { key: 'outboundEfficiency', title: '出库效率' } : null
        ].filter(Boolean);
    }

    function buildSalesReportModuleTitle(bundle, moduleKey, fallbackTitle = '') {
        const modules = getSalesReportModuleOrder(bundle);
        const index = modules.findIndex((item) => item.key === moduleKey);
        return modules[index]?.title || fallbackTitle || moduleKey;
    }

    function buildSalesReportFormalDescriptions(bundle, images) {
        const meta = images?.meta || {};
        const descriptions = {
            region: [],
            inventory: [],
            skuInventory: [],
            skuSales: [],
            outboundEfficiency: []
        };

        if (bundle.regionCustomerRow) {
            const totalOrders = Number(bundle.regionCustomerRow["总订单量"] || 0);
            const topRegion = getTopRegionForCustomer(bundle.regionCustomerRow);
            const topRegionShare = calcReportPercent(topRegion.value, totalOrders);
            const regionRankText = (meta.topRegionRows || [])
                .slice(0, 3)
                .map((row) => `${row.region} ${formatChartNumber(row.value)} 单`)
                .join('、');
            const topWarehouse = sortReportItemsDesc(
                bundle.regionDetailRows || [],
                (row) => row["总订单量"],
                (row) => row["发货仓库"]
            )[0];
            const warehouseCount = new Set((bundle.regionDetailRows || []).map((row) => row["发货仓库"]).filter(Boolean)).size || (bundle.regionDetailRows || []).length;
            descriptions.region.push(`本周期共纳入 ${formatChartNumber(totalOrders)} 单有效订单进行地区分析，覆盖 ${formatChartNumber(warehouseCount)} 个发货仓库。${topRegion.value > 0 ? `订单量最高的地区为 ${topRegion.region}，共 ${formatChartNumber(topRegion.value)} 单，占当前客户有效订单量的 ${formatChartNumber(topRegionShare)}%。` : '当前周期暂无明显的地区集中表现。'}`);
            if (regionRankText) {
                descriptions.region.push(`主要订单地区依次为：${regionRankText}。该结构反映当前客户订单目的地的集中度与区域分布情况。`);
            }
            if (topWarehouse) {
                descriptions.region.push(`从发货仓库维度看，${topWarehouse["发货仓库"] || '未知仓库'} 的订单量最高，共 ${formatChartNumber(topWarehouse["总订单量"] || 0)} 单。该仓库是本周期订单贡献最高的发货仓库。`);
            }
        }

        if (bundle.inventoryCustomerRow) {
            const row = bundle.inventoryCustomerRow;
            const preStock = Number(row["期初库存"] || 0);
            const closeStock = Number(row["期末库存"] || 0);
            const turnoverRate = Number(row["库存周转率"] || 0);
            const turnoverDays = Number(row["库存周转天数"] || 0);
            const warehouseCount = Number(row["分仓数"] || (bundle.inventoryDetailRows || []).length || 0);
            const topStockWarehouse = (meta.inventoryWarehouseByStockRows || [])[0];
            const slowWarehouse = (meta.inventoryWarehouseByDaysRows || [])[0];
            const fastWarehouse = (meta.inventoryWarehouseByRateRows || [])[0];
            descriptions.inventory.push(`本周期客户库存覆盖 ${formatChartNumber(warehouseCount)} 个分仓，期初库存 ${formatChartNumber(preStock)}，期末库存 ${formatChartNumber(closeStock)}，${getInventoryChangeText(preStock, closeStock)}。整体库存周转率为 ${formatChartNumber(turnoverRate)}%，库存周转天数为 ${formatChartNumber(turnoverDays)} 天。`);
            if (topStockWarehouse) {
                const share = calcReportPercent(topStockWarehouse["期末库存"], closeStock);
                descriptions.inventory.push(`期末库存最高的分仓为 ${getReportWarehouseName(topStockWarehouse)}，期末库存 ${formatChartNumber(topStockWarehouse["期末库存"] || 0)}，占客户期末库存的 ${formatChartNumber(share)}%。`);
            }
            if (slowWarehouse || fastWarehouse) {
                const slowText = slowWarehouse ? `${getReportWarehouseName(slowWarehouse)} 周转天数 ${formatChartNumber(slowWarehouse["库存周转天数"] || 0)} 天` : '';
                const fastText = fastWarehouse ? `${getReportWarehouseName(fastWarehouse)} 周转率 ${formatChartNumber(fastWarehouse["库存周转率"] || 0)}%` : '';
                descriptions.inventory.push(`周转表现方面，${[slowText, fastText].filter(Boolean).join('；')}。`);
            }
        }

        if (bundle.skuInventoryCustomerRow) {
            const rows = bundle.skuInventoryDetailRows || [];
            const skuCount = Number(bundle.skuInventoryCustomerRow.skuCount || rows.length || 0);
            const totalPreStock = rows.reduce((sum, row) => sum + Number(row.preStockQty || 0), 0);
            const totalCloseStock = Number(bundle.skuInventoryCustomerRow.totalCloseStockQty || rows.reduce((sum, row) => sum + Number(row.closeStockQty || 0), 0));
            const topSku = (meta.skuInventoryByStockRows || [])[0];
            const slowSku = (meta.skuInventoryByDaysRows || [])[0];
            descriptions.skuInventory.push(`SKU库存分析覆盖 ${formatChartNumber(skuCount)} 个 SKU，期初库存合计 ${formatChartNumber(totalPreStock)}，期末库存合计 ${formatChartNumber(totalCloseStock)}，${getInventoryChangeText(totalPreStock, totalCloseStock)}。`);
            if (topSku) {
                const share = calcReportPercent(topSku.closeStockQty, totalCloseStock);
                descriptions.skuInventory.push(`期末库存最高的 SKU 为 ${topSku.sku}，期末库存 ${formatChartNumber(topSku.closeStockQty)}，占当前 SKU 期末库存总量的 ${formatChartNumber(share)}%。`);
            }
            if (slowSku) {
                descriptions.skuInventory.push(`库存周转天数最高的 SKU 为 ${slowSku.sku}，周转天数 ${formatChartNumber(slowSku.stockTurnoverDays)} 天。`);
            }
        }

        if (bundle.skuSalesCustomerRow) {
            const row = bundle.skuSalesCustomerRow;
            const totalQty = Number(row.totalQty || 0);
            const skuCount = Number(row.skuCount || 0);
            const orderCount = Number(row.orderCount || 0);
            const topSku = (meta.skuSalesDetailRows || [])[0];
            const topFiveQty = (meta.skuSalesRows || []).reduce((sum, item) => sum + Number(item.value || 0), 0);
            const otherSku = (meta.skuSalesShareRows || []).find((item) => item.name === '其他SKU');
            descriptions.skuSales.push(`本周期 SKU 销量分析覆盖 ${formatChartNumber(skuCount)} 个 SKU、${formatChartNumber(orderCount)} 个订单商品行，累计销量 ${formatChartNumber(totalQty)}。`);
            if (topSku) {
                const share = calcReportPercent(topSku.qty, totalQty);
                descriptions.skuSales.push(`销量最高的 SKU 为 ${topSku.sku}，销量 ${formatChartNumber(topSku.qty)}，占客户总销量的 ${formatChartNumber(share)}%。该 SKU 是本周期主要销量来源。`);
            }
            if (topFiveQty > 0) {
                const topFiveShare = calcReportPercent(topFiveQty, totalQty);
                const otherText = otherSku ? `，Top 5 之外的“其他SKU”合计销量 ${formatChartNumber(otherSku.value)}，占比 ${formatChartNumber(calcReportPercent(otherSku.value, totalQty))}%` : '';
                descriptions.skuSales.push(`Top 5 SKU 合计销量 ${formatChartNumber(topFiveQty)}，占总销量的 ${formatChartNumber(topFiveShare)}%${otherText}。`);
            }
        }

        if (bundle.availableModules?.outboundEfficiency) {
            const summary = bundle.outboundEfficiencySummary || {};
            const rows = meta.outboundRows || bundle.outboundEfficiencyRows || [];
            const bestWarehouse = rows[0];
            const weakWarehouse = rows.length > 1 ? rows[rows.length - 1] : null;
            const skipWeekendText = (bundle.outboundEfficiencyRows || []).some((row) => row.skipWeekends) ? '已排除周末' : '包含周末';
            descriptions.outboundEfficiency.push(`出库效率为全仓口径统计，覆盖 ${formatChartNumber(summary.warehouseCount || 0)} 个仓库，统计周期 ${skipWeekendText}。平均 24H、48H、72H 出库发货率分别为 ${formatChartNumber(summary.avg24 || 0)}%、${formatChartNumber(summary.avg48 || 0)}%、${formatChartNumber(summary.avg72 || 0)}%。`);
            if (bestWarehouse) {
                descriptions.outboundEfficiency.push(`24H 出库发货率最高的仓库为 ${bestWarehouse.whName}，24H 出库率 ${formatChartNumber(bestWarehouse.rate24)}%，48H 出库率 ${formatChartNumber(bestWarehouse.rate48)}%，72H 出库率 ${formatChartNumber(bestWarehouse.rate72)}%。`);
            }
            if (weakWarehouse) {
                const gap = Number(bestWarehouse?.rate24 || 0) - Number(weakWarehouse.rate24 || 0);
                descriptions.outboundEfficiency.push(`24H 出库发货率最低的仓库为 ${weakWarehouse.whName}，24H 出库率 ${formatChartNumber(weakWarehouse.rate24)}%，与最高仓库相差 ${formatChartNumber(gap)} 个百分点。`);
            }
        }

        return descriptions;
    }

    function buildSalesReportRadarConfig(bundle) {
        const indicator = [];
        const value = [];
        if (bundle.regionCustomerRow) {
            const metricValue = Number(bundle.regionCustomerRow["总订单量"] || 0);
            indicator.push({ name: '订单量', max: Math.max(1, metricValue * 1.2) });
            value.push(metricValue);
        }
        if (bundle.inventoryCustomerRow) {
            const metricValue = Number(bundle.inventoryCustomerRow["期末库存"] || 0);
            indicator.push({ name: '期末库存', max: Math.max(1, metricValue * 1.2) });
            value.push(metricValue);
        }
        if (bundle.skuInventoryCustomerRow) {
            const metricValue = Number(bundle.skuInventoryCustomerRow.skuCount || 0);
            indicator.push({ name: '库存SKU数', max: Math.max(1, metricValue * 1.2) });
            value.push(metricValue);
        }
        if (bundle.skuSalesCustomerRow) {
            const metricValue = Number(bundle.skuSalesCustomerRow.totalQty || 0);
            indicator.push({ name: '销量', max: Math.max(1, metricValue * 1.2) });
            value.push(metricValue);
        }
        return { indicator, value };
    }

    function buildSalesReportSectionHtml(title, narrative, imageItems, options = {}) {
        const validImageItems = (imageItems || []).filter((item) => item && item.dataUrl);
        const cards = validImageItems.map((item) =>
            buildSalesReportImageCard(item.title, item.dataUrl, item.description)
        ).join('');
        const tableHtml = options.tableHtml || '';
        const adviceHtml = options.adviceHtml || '';
        const sectionClass = options.pageBreak ? 'section page-break' : 'section';
        const chartGridClass = validImageItems.length <= 1 ? 'chart-grid chart-grid-single' : 'chart-grid';
        if (!narrative && !cards && !tableHtml && !adviceHtml) return '';
        return `
            <div class="${sectionClass}">
                <h2>${escapeHtml(title)}</h2>
                ${narrative || ''}
                ${adviceHtml}
                ${cards ? `<div class="${chartGridClass}">${cards}</div>` : ''}
                ${tableHtml}
            </div>
        `;
    }

    function getSalesReportOptions() {
        return {
            showChineseSkuName: Boolean(document.getElementById('sr-sales-report-show-cn-sku')?.checked),
            useChineseCustomerName: Boolean(document.getElementById('sr-sales-report-use-cn-name')?.checked),
            denseMode: Boolean(document.getElementById('sr-sales-report-dense-mode')?.checked)
        };
    }

    function isSameSalesReportOptions(left, right) {
        return Boolean(left?.showChineseSkuName) === Boolean(right?.showChineseSkuName)
            && Boolean(left?.useChineseCustomerName) === Boolean(right?.useChineseCustomerName)
            && Boolean(left?.denseMode) === Boolean(right?.denseMode);
    }

    function getSalesReportFileBase(bundle) {
        return sanitizeFileName(bundle.reportTitle || `${bundle.customerName} 销售报告`);
    }

    function refreshSalesReportPreviewByOptions() {
        salesReportSnapshot = null;
        updateSalesReportReadyState();
        const customerKey = document.getElementById('sr-sales-report-customer')?.value || '';
        const contentEl = document.getElementById('sr-sales-report-preview-content');
        if (customerKey && contentEl && contentEl.style.display !== 'none') {
            previewSalesReport();
        }
    }

    function renderSalesReportPreview(bundle, images, narratives) {
        const emptyEl = document.getElementById('sr-sales-report-preview-empty');
        const contentEl = document.getElementById('sr-sales-report-preview-content');
        const titleEl = document.getElementById('sr-sales-report-preview-title');
        const metaEl = document.getElementById('sr-sales-report-preview-meta');
        const frameEl = document.getElementById('sr-sales-report-preview-frame');
        if (!emptyEl || !contentEl || !titleEl || !metaEl || !frameEl) return;

        emptyEl.style.display = 'none';
        contentEl.style.display = 'block';
        titleEl.innerText = `${bundle.reportTitle || `${bundle.customerName} 销售报告`} HTML预览`;
        metaEl.innerText = `已接入模块：${getSalesReportModuleLabels(bundle.availableModuleKeys).join('、')}｜缺失模块会自动跳过导出`;

        const options = getSalesReportOptions();
        const html = options.denseMode
            ? buildSalesReportDenseHtml(bundle, images, narratives, options)
            : buildSalesReportHtml(bundle, images, narratives, options);
        frameEl.srcdoc = html;
        salesReportSnapshot = { bundle, images, narratives, options, html };
    }

    function buildSalesReportHtml(bundle, images, narratives, options = {}) {
        const topRegion = bundle.regionCustomerRow ? getTopRegionForCustomer(bundle.regionCustomerRow) : { region: '-', value: 0 };
        const topInventoryRows = bundle.skuInventoryCustomerRow ? (bundle.skuInventoryDetailRows || [])
            .slice()
            .sort((a, b) => Number(b.closeStockQty || 0) - Number(a.closeStockQty || 0) || String(a.sku || '').localeCompare(String(b.sku || ''), 'zh-Hans-CN'))
            .slice(0, 5) : [];
        const topSalesRows = bundle.skuSalesCustomerRow ? (bundle.skuSalesDetailRows || [])
            .slice()
            .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0) || String(a.sku || '').localeCompare(String(b.sku || ''), 'zh-Hans-CN'))
            .slice(0, 5) : [];
        const outboundRows = bundle.availableModules?.outboundEfficiency
            ? sortOutboundRowsByWarehouseName(bundle.outboundEfficiencyRows || []).slice(0, 12)
            : [];
        const skuInventoryTableTitle = buildConditionalTopLabel('SKU库存明细', topInventoryRows.length, 5);
        const skuSalesTableTitle = buildConditionalTopLabel('SKU销量明细', topSalesRows.length, 5);
        const outboundTableTitle = outboundRows.length >= 12 ? '出库效率明细 Top 12' : '出库效率明细';
        const metrics = buildSalesReportMetrics(bundle);
        const imageGroups = images.groups || {};
        const tagItems = buildSalesReportTagItems(bundle);
        const formalDescriptions = buildSalesReportFormalDescriptions(bundle, images);
        const missingModuleNote = (narratives.missingModules || []).length > 0
            ? `<div class="notice">本期暂无以下分析数据：${escapeHtml(narratives.missingModules.join('、'))}，正文仅展示已有数据内容。</div>`
            : '';
        const showChineseSkuName = Boolean(options.showChineseSkuName);
        const productNameHeader = showChineseSkuName ? '<th>产品名称</th>' : '';
        const productNameCell = (row) => showChineseSkuName ? `<td>${escapeHtml(row.productName)}</td>` : '';
        const skuInventoryTableHtml = topInventoryRows.length > 0 ? `
        <div class="table-title">${escapeHtml(skuInventoryTableTitle)}</div>
        <table class="table">
            <thead><tr><th>SKU</th>${productNameHeader}<th>期初库存</th><th>期末库存</th><th>库存周转率</th><th>库存周转天数</th></tr></thead>
            <tbody>
                ${topInventoryRows.map((row) => `<tr><td>${escapeHtml(row.sku)}</td>${productNameCell(row)}<td>${escapeHtml(formatChartNumber(row.preStockQty))}</td><td>${escapeHtml(formatChartNumber(row.closeStockQty))}</td><td>${escapeHtml(formatChartNumber(row.stockTurnoverRate))}%</td><td>${escapeHtml(formatChartNumber(row.stockTurnoverDays))}</td></tr>`).join('')}
            </tbody>
        </table>` : '';
        const skuSalesTableHtml = topSalesRows.length > 0 ? `
        <div class="table-title">${escapeHtml(skuSalesTableTitle)}</div>
        <table class="table">
            <thead><tr><th>SKU</th>${productNameHeader}<th>销量</th><th>客户总销量</th><th>订单商品行数</th></tr></thead>
            <tbody>
                ${topSalesRows.map((row) => `<tr><td>${escapeHtml(row.sku)}</td>${productNameCell(row)}<td>${escapeHtml(formatChartNumber(row.qty))}</td><td>${escapeHtml(formatChartNumber(row.customerTotalQty))}</td><td>${escapeHtml(formatChartNumber(row.customerOrderCount))}</td></tr>`).join('')}
            </tbody>
        </table>` : '';
        const outboundEfficiencyTableHtml = outboundRows.length > 0 ? `
        <div class="table-title">${escapeHtml(outboundTableTitle)}</div>
        <table class="table">
            <thead><tr><th>仓库名称</th><th>仓库代码</th><th>24H 出库发货率</th><th>48H 出库发货率</th><th>72H 出库发货率</th><th>纳入日期数</th></tr></thead>
            <tbody>
                ${outboundRows.map((row) => `<tr><td>${escapeHtml(row.whName)}</td><td>${escapeHtml(row.whCode)}</td><td>${escapeHtml(formatOutboundRate(row.rate24))}</td><td>${escapeHtml(formatOutboundRate(row.rate48))}</td><td>${escapeHtml(formatOutboundRate(row.rate72))}</td><td>${escapeHtml(formatChartNumber(row.includedDays || 0))}</td></tr>`).join('')}
            </tbody>
        </table>` : '';
        const reportTitle = bundle.reportTitle || `${bundle.customerName} 销售报告`;

        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>${escapeHtml(reportTitle)}</title>
<style>
body { margin:0; background:#f5f7fa; color:#1f1f1f; font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif; }
.page { width:1120px; margin:0 auto; padding:28px 24px 40px; }
.hero { background:linear-gradient(135deg, #fff7e6, #ffffff 42%, #e6fffb); border:1px solid #f0f0f0; border-radius:18px; padding:28px; box-shadow:0 10px 30px rgba(0,0,0,0.05); }
.title { font-size:30px; font-weight:800; margin:0 0 8px; }
.sub { color:#666; font-size:14px; line-height:1.8; }
.metrics { display:grid; grid-template-columns:repeat(var(--metric-count, 5), minmax(0, 1fr)); gap:10px; margin:22px 0 4px; }
.metric { min-width:0; background:#ffffff; border:1px solid #eef0f2; border-radius:12px; padding:14px 12px; }
.metric-label { color:#888; font-size:12px; margin-bottom:8px; white-space:nowrap; }
.metric-value { font-size:24px; font-weight:700; color:#111; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.section { margin-top:18px; background:#fff; border:1px solid #f0f0f0; border-radius:18px; padding:24px; box-shadow:0 8px 24px rgba(0,0,0,0.04); }
.section h2 { margin:0 0 8px; font-size:20px; }
.section p { margin:8px 0 0; color:#555; line-height:1.9; font-size:14px; }
.tag-row { display:flex; flex-wrap:wrap; gap:8px; margin-top:14px; }
.tag { background:#f0f5ff; color:#1d39c4; border:1px solid #adc6ff; border-radius:999px; padding:4px 10px; font-size:12px; }
.chart { margin-top:16px; border:1px solid #f0f0f0; border-radius:14px; overflow:hidden; background:#fff; }
.chart img { display:block; width:100%; height:auto; }
.chart-grid { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:16px; margin-top:16px; }
.chart-grid.chart-grid-single { grid-template-columns:minmax(0, 1fr); }
.chart-card { border:1px solid #eef0f2; border-radius:16px; background:#fff; overflow:hidden; }
.chart-card-head { padding:16px 16px 0; }
.chart-card-title { font-size:16px; font-weight:700; color:#111; }
.chart-card-desc { margin-top:6px; color:#666; font-size:13px; line-height:1.75; min-height:44px; }
.notice { margin-top:14px; padding:12px 14px; border-radius:12px; background:#fff7e6; border:1px solid #ffd591; color:#ad6800; font-size:13px; line-height:1.8; }
.table-title { margin-top:18px; font-size:15px; font-weight:700; color:#111; }
.table { width:100%; border-collapse:collapse; margin-top:14px; font-size:13px; }
.table th, .table td { border:1px solid #f0f0f0; padding:8px 10px; text-align:left; }
.table th { background:#fafafa; }
.page-break { page-break-before: always; }
.footer { color:#888; font-size:12px; text-align:center; margin-top:18px; }
@media print {
    .page { width:auto; padding:0; }
    .section, .hero, .chart-card { break-inside: avoid; page-break-inside: avoid; }
}
</style>
</head>
<body>
<div class="page">
    <div class="hero">
        <div class="title">${escapeHtml(reportTitle)}</div>
        <div class="sub">${escapeHtml(narratives.overview)}</div>
        <div class="tag-row">
            ${tagItems.map((text) => `<span class="tag">${escapeHtml(text)}</span>`).join('')}
        </div>
        <div class="metrics" style="--metric-count:${Math.max(1, Math.min(metrics.length, 5))}">
            ${metrics.map((item) => `<div class="metric"><div class="metric-label">${escapeHtml(item.label)}</div><div class="metric-value">${escapeHtml(item.value)}</div></div>`).join('')}
        </div>
        ${missingModuleNote}
    </div>

    ${bundle.regionCustomerRow ? buildSalesReportSectionHtml(
        buildSalesReportModuleTitle(bundle, 'region', '订单分布分析'),
        buildSalesReportTextBlock(formalDescriptions.region),
        imageGroups.region || [],
        { adviceHtml: buildSalesReportAdviceBlock('region') }
    ) : ''}

    ${bundle.inventoryCustomerRow ? buildSalesReportSectionHtml(
        buildSalesReportModuleTitle(bundle, 'inventory', '按客户库存分析'),
        buildSalesReportTextBlock(formalDescriptions.inventory),
        imageGroups.inventory || [],
        { adviceHtml: buildSalesReportAdviceBlock('inventory') }
    ) : ''}

    ${bundle.skuInventoryCustomerRow ? buildSalesReportSectionHtml(
        buildSalesReportModuleTitle(bundle, 'skuInventory', '按SKU库存分析'),
        buildSalesReportTextBlock(formalDescriptions.skuInventory),
        imageGroups.skuInventory || [],
        { pageBreak: !!bundle.regionCustomerRow || !!bundle.inventoryCustomerRow, tableHtml: skuInventoryTableHtml, adviceHtml: buildSalesReportAdviceBlock('skuInventory') }
    ) : ''}

    ${bundle.skuSalesCustomerRow ? buildSalesReportSectionHtml(
        buildSalesReportModuleTitle(bundle, 'skuSales', 'SKU销量分析'),
        buildSalesReportTextBlock(formalDescriptions.skuSales),
        imageGroups.skuSales || [],
        { tableHtml: skuSalesTableHtml, adviceHtml: buildSalesReportAdviceBlock('skuSales') }
    ) : ''}

    ${bundle.availableModules?.outboundEfficiency ? buildSalesReportSectionHtml(
        buildSalesReportModuleTitle(bundle, 'outboundEfficiency', '出库效率'),
        buildSalesReportTextBlock(formalDescriptions.outboundEfficiency),
        imageGroups.outboundEfficiency || [],
        { tableHtml: outboundEfficiencyTableHtml, adviceHtml: buildSalesReportAdviceBlock('outboundEfficiency') }
    ) : ''}

    <div class="footer">报告生成时间：${escapeHtml(new Date().toLocaleString('zh-CN'))}</div>
</div>
</body>
</html>`;
    }

    function buildDenseKpiCards(bundle) {
        const cards = [];
        if (bundle.regionCustomerRow) cards.push({ label: '订单量', value: formatChartNumber(bundle.regionCustomerRow["总订单量"] || 0) });
        if (bundle.inventoryCustomerRow) cards.push({ label: '期末库存', value: formatChartNumber(bundle.inventoryCustomerRow["期末库存"] || 0) });
        if (bundle.skuInventoryCustomerRow) cards.push({ label: '库存SKU数', value: formatChartNumber(bundle.skuInventoryCustomerRow.skuCount || 0) });
        if (bundle.skuSalesCustomerRow) cards.push({ label: '总销量', value: formatChartNumber(bundle.skuSalesCustomerRow.totalQty || 0) });
        return cards.slice(0, 4);
    }

    function buildSalesReportDenseHtml(bundle, images, narratives, options = {}) {
        const reportTitle = bundle.reportTitle || `${bundle.customerName} 销售报告`;
        const tagItems = buildSalesReportTagItems(bundle);
        const kpis = buildDenseKpiCards(bundle);
        const imageGroups = images.groups || {};
        const formalDescriptions = buildSalesReportFormalDescriptions(bundle, images);
        const outboundRows = bundle.availableModules?.outboundEfficiency
            ? sortOutboundRowsByWarehouseName(bundle.outboundEfficiencyRows || []).slice(0, 6)
            : [];
        const topInventoryRows = bundle.skuInventoryCustomerRow ? (bundle.skuInventoryDetailRows || [])
            .slice()
            .sort((a, b) => Number(b.closeStockQty || 0) - Number(a.closeStockQty || 0) || String(a.sku || '').localeCompare(String(b.sku || ''), 'zh-Hans-CN'))
            .slice(0, 3) : [];
        const topSalesRows = bundle.skuSalesCustomerRow ? (bundle.skuSalesDetailRows || [])
            .slice()
            .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0) || String(a.sku || '').localeCompare(String(b.sku || ''), 'zh-Hans-CN'))
            .slice(0, 3) : [];
        const showChineseSkuName = Boolean(options.showChineseSkuName);
        const productHeader = showChineseSkuName ? '<th>产品名称</th>' : '';
        const productCell = (row) => showChineseSkuName ? `<td>${escapeHtml(row.productName)}</td>` : '';

        const moduleCards = [
            bundle.regionCustomerRow ? {
                title: buildSalesReportModuleTitle(bundle, 'region', '订单分布分析'),
                advice: buildSalesReportAdviceBlock('region', { className: 'dense-advice' }),
                narrative: buildSalesReportTextBlock(formalDescriptions.region.slice(0, 2)),
                charts: imageGroups.region || [],
                compact: true
            } : null,
            bundle.inventoryCustomerRow ? {
                title: buildSalesReportModuleTitle(bundle, 'inventory', '按客户库存分析'),
                advice: buildSalesReportAdviceBlock('inventory', { className: 'dense-advice' }),
                narrative: buildSalesReportTextBlock(formalDescriptions.inventory.slice(0, 2)),
                charts: imageGroups.inventory || [],
                compact: true
            } : null,
            bundle.skuInventoryCustomerRow ? {
                title: buildSalesReportModuleTitle(bundle, 'skuInventory', '按SKU库存分析'),
                advice: buildSalesReportAdviceBlock('skuInventory', { className: 'dense-advice' }),
                narrative: buildSalesReportTextBlock(formalDescriptions.skuInventory.slice(0, 2)),
                charts: imageGroups.skuInventory || [],
                table: topInventoryRows.length ? `
                    <table class="dense-table">
                        <thead><tr><th>SKU</th>${productHeader}<th>期末库存</th><th>周转天数</th></tr></thead>
                        <tbody>${topInventoryRows.map((row) => `<tr><td>${escapeHtml(row.sku)}</td>${productCell(row)}<td>${escapeHtml(formatChartNumber(row.closeStockQty))}</td><td>${escapeHtml(formatChartNumber(row.stockTurnoverDays))}</td></tr>`).join('')}</tbody>
                    </table>
                ` : ''
            } : null,
            bundle.skuSalesCustomerRow ? {
                title: buildSalesReportModuleTitle(bundle, 'skuSales', 'SKU销量分析'),
                advice: buildSalesReportAdviceBlock('skuSales', { className: 'dense-advice' }),
                narrative: buildSalesReportTextBlock(formalDescriptions.skuSales.slice(0, 2)),
                charts: imageGroups.skuSales || [],
                compact: true,
                table: topSalesRows.length ? `
                    <table class="dense-table">
                        <thead><tr><th>SKU</th>${productHeader}<th>销量</th><th>订单行数</th></tr></thead>
                        <tbody>${topSalesRows.map((row) => `<tr><td>${escapeHtml(row.sku)}</td>${productCell(row)}<td>${escapeHtml(formatChartNumber(row.qty))}</td><td>${escapeHtml(formatChartNumber(row.customerOrderCount))}</td></tr>`).join('')}</tbody>
                    </table>
                ` : ''
            } : null,
            bundle.availableModules?.outboundEfficiency ? {
                title: buildSalesReportModuleTitle(bundle, 'outboundEfficiency', '出库效率'),
                advice: buildSalesReportAdviceBlock('outboundEfficiency', { className: 'dense-advice' }),
                narrative: buildSalesReportTextBlock(formalDescriptions.outboundEfficiency.slice(0, 2)),
                charts: imageGroups.outboundEfficiency || [],
                compact: true,
                table: outboundRows.length ? `
                    <table class="dense-table">
                        <thead><tr><th>仓库</th><th>24H</th><th>48H</th><th>72H</th></tr></thead>
                        <tbody>${outboundRows.map((row) => `<tr><td>${escapeHtml(row.whName)}</td><td>${escapeHtml(formatOutboundRate(row.rate24))}</td><td>${escapeHtml(formatOutboundRate(row.rate48))}</td><td>${escapeHtml(formatOutboundRate(row.rate72))}</td></tr>`).join('')}</tbody>
                    </table>
                ` : ''
            } : null
        ].filter(Boolean);

        const compactModuleCards = moduleCards.filter((item) => item.compact);
        const regularModuleCards = moduleCards.filter((item) => !item.compact);

        const chartCardHtml = (item) => {
            const chart = (item.charts || []).find((entry) => entry && entry.dataUrl);
            const extraChart = (item.charts || []).slice(1, 2).find((entry) => entry && entry.dataUrl);
            const moduleClass = item.compact ? 'dense-module compact' : 'dense-module';
            return `
                <section class="${moduleClass}">
                    <div class="dense-module-header">
                        <div>
                            <h2>${escapeHtml(item.title)}</h2>
                            ${item.advice || ''}
                        </div>
                    </div>
                    <div class="dense-module-body">
                        <div class="dense-copy">${item.narrative || ''}</div>
                        ${chart ? `<div class="dense-chart-shell"><img src="${chart.dataUrl}" alt="${escapeHtml(chart.title)}" /></div>` : ''}
                        ${extraChart ? `<div class="dense-chart-shell secondary"><img src="${extraChart.dataUrl}" alt="${escapeHtml(extraChart.title)}" /></div>` : ''}
                        ${item.table || ''}
                    </div>
                </section>
            `;
        };

        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>${escapeHtml(reportTitle)}</title>
<style>
body { margin:0; background:#eef3f8; color:#0f172a; font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif; }
.page { width:1120px; margin:0 auto; padding:26px 26px 30px; }
.dense-shell { background:linear-gradient(180deg, #ffffff, #f8fbff); border:1px solid #dbe4ee; border-radius:28px; padding:24px 24px 20px; box-shadow:0 20px 60px rgba(15, 23, 42, 0.08); }
.dense-hero { display:grid; grid-template-columns:1.3fr 0.9fr; gap:20px; align-items:stretch; }
.dense-title { font-size:30px; line-height:1.18; margin:0 0 10px; font-weight:800; letter-spacing:0.01em; }
.dense-sub { font-size:14px; line-height:1.95; color:#475569; margin:0; }
.dense-tag-row { display:flex; flex-wrap:wrap; gap:10px; margin-top:14px; }
.dense-tag { padding:6px 12px; border-radius:999px; font-size:12px; border:1px solid #bfdbfe; background:#eff6ff; color:#1d4ed8; }
.dense-kpis { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:12px; }
.dense-kpi { border:1px solid #e2e8f0; border-radius:20px; padding:16px; background:linear-gradient(180deg, #fffdf9, #ffffff); min-height:82px; }
.dense-kpi-label { color:#64748b; font-size:13px; }
.dense-kpi-value { margin-top:10px; font-size:26px; font-weight:800; color:#0f172a; }
.dense-grid { display:grid; grid-template-columns:minmax(0, 1fr); gap:18px; margin-top:20px; }
.dense-grid-compact { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:14px; margin-top:20px; }
.dense-grid-compact.single { grid-template-columns:minmax(0, 1fr); }
.dense-module { border:1px solid #dfe7ef; border-radius:22px; background:#ffffff; overflow:hidden; box-shadow:0 12px 32px rgba(15, 23, 42, 0.05); }
.dense-module-header { padding:18px 18px 12px; background:linear-gradient(180deg, #ffffff, #f8fbff); border-bottom:1px solid #edf2f7; }
.dense-module-header h2 { margin:0; font-size:18px; color:#0f172a; }
.dense-module-body { padding:16px 18px 18px; }
.dense-copy p { margin:0 0 10px; color:#475569; font-size:13px; line-height:1.85; }
.dense-chart-shell { margin-top:12px; border:1px solid #e5edf5; border-radius:18px; overflow:hidden; background:#fff; }
.dense-chart-shell.secondary { margin-top:12px; }
.dense-chart-shell img { display:block; width:100%; height:auto; }
.dense-module.compact { border-radius:18px; }
.dense-module.compact .dense-module-header { padding:14px 14px 10px; }
.dense-module.compact .dense-module-header h2 { font-size:16px; }
.dense-module.compact .dense-module-body { padding:12px 14px 14px; }
.dense-module.compact .dense-copy p { margin:0 0 8px; font-size:12px; line-height:1.72; }
.dense-module.compact .dense-chart-shell { margin-top:8px; border-radius:14px; }
.dense-module.compact .dense-chart-shell.secondary { margin-top:8px; }
.dense-module.compact .dense-advice { margin-top:8px; padding:8px 10px; font-size:11px; line-height:1.6; }
.dense-table { width:100%; border-collapse:collapse; margin-top:12px; font-size:12px; }
.dense-table th, .dense-table td { border:1px solid #e5edf5; padding:8px 10px; text-align:left; }
.dense-table th { background:#f8fafc; color:#334155; }
.dense-advice { margin-top:10px; padding:10px 12px; border-radius:12px; background:#fff7ed; border:1px solid #fdba74; color:#9a3412; font-size:12px; line-height:1.75; }
.dense-footer { margin-top:16px; color:#64748b; font-size:12px; text-align:right; }
@media print {
    .page { padding:0; }
    .dense-shell, .dense-module, .dense-kpi { break-inside: avoid; page-break-inside: avoid; }
}
</style>
</head>
<body>
<div class="page">
    <div class="dense-shell">
        <section class="dense-hero">
            <div>
                <h1 class="dense-title">${escapeHtml(reportTitle)}</h1>
                <p class="dense-sub">${escapeHtml(narratives.overview)}</p>
                <div class="dense-tag-row">${tagItems.map((text) => `<span class="dense-tag">${escapeHtml(text)}</span>`).join('')}</div>
            </div>
            <div class="dense-kpis">
                ${kpis.map((item) => `<div class="dense-kpi"><div class="dense-kpi-label">${escapeHtml(item.label)}</div><div class="dense-kpi-value">${escapeHtml(item.value)}</div></div>`).join('')}
            </div>
        </section>
        ${compactModuleCards.length > 0 ? `
        <section class="dense-grid-compact${compactModuleCards.length === 1 ? ' single' : ''}">
            ${compactModuleCards.map((item) => chartCardHtml(item)).join('')}
        </section>` : ''}
        ${regularModuleCards.length > 0 ? `
        <section class="dense-grid">
            ${regularModuleCards.map((item) => chartCardHtml(item)).join('')}
        </section>` : ''}
        <div class="dense-footer">报告生成时间：${escapeHtml(new Date().toLocaleString('zh-CN'))}</div>
    </div>
</div>
</body>
</html>`;
    }

    function refreshSalesReportCustomers() {
        updateSalesReportReadyState();
        updateSalesReportCustomerNameState();
        salesReportSnapshot = null;
        const emptyEl = document.getElementById('sr-sales-report-preview-empty');
        const contentEl = document.getElementById('sr-sales-report-preview-content');
        const frameEl = document.getElementById('sr-sales-report-preview-frame');
        if (emptyEl) emptyEl.style.display = 'flex';
        if (contentEl) contentEl.style.display = 'none';
        if (frameEl) frameEl.srcdoc = '';
    }

    let outboundEditorDraftRows = [];
    let outboundEditorVisibleColumns = ['whName', 'whCode', 'rate24', 'rate48', 'rate72', 'includedDays'];

    const OUTBOUND_EDITOR_COLUMN_OPTIONS = [
        { key: 'whName', label: '仓库名称' },
        { key: 'whCode', label: '仓库代码' },
        { key: 'rate24', label: '24H 出库率 (%)' },
        { key: 'rate48', label: '48H 出库率 (%)' },
        { key: 'rate72', label: '72H 出库率 (%)' },
        { key: 'includedDays', label: '纳入天数' },
        { key: 'skipWeekends', label: '排除周末' },
        { key: 'startDate', label: '开始日期' },
        { key: 'endDate', label: '结束日期' }
    ];

    const SALES_REPORT_ADVICE_MODULES = [
        { key: 'region', label: '1. 订单分布分析' },
        { key: 'inventory', label: '2. 按客户库存分析' },
        { key: 'skuInventory', label: '3. 按SKU库存分析' },
        { key: 'skuSales', label: '4. SKU销量分析' },
        { key: 'outboundEfficiency', label: '5. 出库效率' }
    ];

    function syncSalesReportPreviewIfVisible() {
        const customerKey = document.getElementById('sr-sales-report-customer')?.value || '';
        const contentEl = document.getElementById('sr-sales-report-preview-content');
        if (customerKey && contentEl && contentEl.style.display !== 'none') {
            previewSalesReport();
        }
    }

    function sortOutboundRowsByWarehouseName(rows = []) {
        return cloneOutboundRows(rows).sort((left, right) =>
            String(left.whName || '').localeCompare(String(right.whName || ''), 'zh-Hans-CN')
            || String(left.whCode || '').localeCompare(String(right.whCode || ''), 'zh-Hans-CN')
        );
    }

    function formatOutboundEditorCellValue(row, field) {
        if (field === 'skipWeekends') return row.skipWeekends ? 'true' : 'false';
        if (['rate24', 'rate48', 'rate72'].includes(field)) return Number(row[field] || 0).toFixed(2);
        return String(row?.[field] ?? '');
    }

    function renderOutboundEditorColumnOptions() {
        const container = document.getElementById('sr-outbound-editor-columns');
        if (!container) return;
        container.innerHTML = OUTBOUND_EDITOR_COLUMN_OPTIONS.map((item) => `
            <label class="sr-editor-chip">
                <input type="checkbox" data-column="${item.key}" ${outboundEditorVisibleColumns.includes(item.key) ? 'checked' : ''}>
                <span>${escapeHtml(item.label)}</span>
            </label>
        `).join('');
        container.querySelectorAll('input[data-column]').forEach((input) => {
            input.addEventListener('change', () => {
                const checked = Array.from(container.querySelectorAll('input[data-column]:checked')).map((node) => node.dataset.column);
                if (checked.length === 0) {
                    input.checked = true;
                    alert('至少保留一列。');
                    return;
                }
                outboundEditorVisibleColumns = checked;
                renderOutboundEditorTable();
            });
        });
    }

    function renderOutboundEditorTable() {
        const tbody = document.querySelector('#sr-outbound-editor-table tbody');
        const metaEl = document.getElementById('sr-outbound-editor-meta');
        const saveBtn = document.getElementById('sr-outbound-editor-save');
        const thead = document.querySelector('#sr-outbound-editor-table thead');
        if (!tbody) return;

        const rows = outboundEditorDraftRows || [];
        const visibleColumns = OUTBOUND_EDITOR_COLUMN_OPTIONS.filter((item) => outboundEditorVisibleColumns.includes(item.key));
        if (thead) {
            thead.innerHTML = `<tr>${visibleColumns.map((item) => `<th>${escapeHtml(item.label)}</th>`).join('')}<th>操作</th></tr>`;
        }
        tbody.innerHTML = '';
        if (rows.length === 0) {
            const colSpan = visibleColumns.length + 1;
            tbody.innerHTML = `<tr><td colspan="${colSpan}" style="color:#999; padding:24px 0; text-align:center;">请先抓取出库效率数据，再进入编辑。</td></tr>`;
            if (metaEl) metaEl.innerText = '当前没有可编辑的出库效率数据。';
            if (saveBtn) saveBtn.disabled = true;
            return;
        }

        rows.forEach((row, index) => {
            const tr = document.createElement('tr');
            const cells = visibleColumns.map((column) => {
                const value = formatOutboundEditorCellValue(row, column.key);
                const align = column.key === 'whName' ? 'left' : 'right';
                return `<td><input class="sr-editor-input" style="text-align:${align};" data-field="${column.key}" data-index="${index}" value="${escapeAttribute(value)}" /></td>`;
            }).join('');
            tr.innerHTML = `${cells}<td><button class="sr-btn sr-btn-danger sr-editor-row-delete" data-index="${index}" type="button">删除</button></td>`;
            tbody.appendChild(tr);
        });

        tbody.querySelectorAll('.sr-editor-row-delete').forEach((button) => {
            button.addEventListener('click', () => {
                const index = Number(button.dataset.index);
                if (!Number.isInteger(index)) return;
                outboundEditorDraftRows.splice(index, 1);
                renderOutboundEditorTable();
            });
        });

        if (metaEl) {
            metaEl.innerText = `共 ${rows.length} 个仓库，修改后会同步覆盖销售报告中的“出库效率”模块。`;
        }
        if (saveBtn) saveBtn.disabled = false;
    }

    function openOutboundEditorModal() {
        const mask = document.getElementById('sr-outbound-editor-modal-mask');
        if (!mask) return;
        if (!outboundReportData?.rows?.length) {
            alert('请先抓取出库效率数据，再进行修改。');
            return;
        }
        outboundEditorDraftRows = sortOutboundRowsByWarehouseName(outboundReportData?.rows || []);
        outboundEditorVisibleColumns = ['whName', 'whCode', 'rate24', 'rate48', 'rate72', 'includedDays'];
        renderOutboundEditorColumnOptions();
        renderOutboundEditorTable();
        mask.classList.add('show');
    }

    function closeOutboundEditorModal() {
        const mask = document.getElementById('sr-outbound-editor-modal-mask');
        if (!mask) return;
        mask.classList.remove('show');
    }

    function resetOutboundEditorDraft() {
        outboundEditorDraftRows = sortOutboundRowsByWarehouseName(outboundReportData?.rows || []);
        outboundEditorVisibleColumns = ['whName', 'whCode', 'rate24', 'rate48', 'rate72', 'includedDays'];
        renderOutboundEditorColumnOptions();
        renderOutboundEditorTable();
    }

    function addOutboundEditorRow() {
        outboundEditorDraftRows.push({
            whName: `新仓库${outboundEditorDraftRows.length + 1}`,
            whCode: '',
            rate24: 0,
            rate48: 0,
            rate72: 0,
            includedDays: outboundReportData?.rows?.[0]?.includedDays || 0,
            skipWeekends: Boolean(outboundReportData?.skipWeekends),
            startDate: outboundReportData?.startDate || '',
            endDate: outboundReportData?.endDate || ''
        });
        outboundEditorDraftRows = sortOutboundRowsByWarehouseName(outboundEditorDraftRows);
        renderOutboundEditorTable();
    }

    function saveOutboundEditorChanges() {
        if (!outboundEditorDraftRows.length) {
            alert('请先抓取出库效率数据。');
            return;
        }

        const inputs = Array.from(document.querySelectorAll('#sr-outbound-editor-table .sr-editor-input'));
        const nextRows = cloneOutboundRows(outboundEditorDraftRows);
        try {
            inputs.forEach((input) => {
                const index = Number(input.dataset.index);
                const field = input.dataset.field;
                if (!Number.isInteger(index) || !nextRows[index] || !field) return;
                if (['rate24', 'rate48', 'rate72', 'includedDays'].includes(field)) {
                    nextRows[index][field] = normalizeEditableRateValue(input.value);
                } else if (field === 'skipWeekends') {
                    const text = String(input.value || '').trim().toLowerCase();
                    nextRows[index][field] = text === 'true' || text === '1' || text === '是';
                } else {
                    nextRows[index][field] = String(input.value ?? '').trim();
                }
            });
        } catch (error) {
            alert(error.message);
            return;
        }

        const sanitizedRows = sortOutboundRowsByWarehouseName(nextRows);

        outboundReportData = {
            ...outboundReportData,
            rows: sanitizedRows,
            summary: buildOutboundSummary(sanitizedRows)
        };
        outboundEditorDraftRows = cloneOutboundRows(sanitizedRows);
        renderOutboundRateView();
        updateSalesReportReadyState();
        salesReportSnapshot = null;
        syncSalesReportPreviewIfVisible();
        const statusEl = document.getElementById('sr-status');
        if (statusEl) statusEl.innerText = `已保存出库效率修改，共更新 ${sanitizedRows.length} 个仓库的报表数据`;
        closeOutboundEditorModal();
    }

    function openSalesReportAdviceModal() {
        const mask = document.getElementById('sr-sales-report-advice-modal-mask');
        if (!mask) return;
        const customerKey = document.getElementById('sr-sales-report-customer')?.value || '';
        const bundle = customerKey ? getSalesReportCustomerBundle(customerKey) : null;
        SALES_REPORT_ADVICE_MODULES.forEach((module) => {
            const textarea = document.getElementById(`sr-sales-report-advice-${module.key}`);
            const titleEl = document.querySelector(`[data-advice-title="${module.key}"]`);
            if (titleEl) {
                titleEl.textContent = bundle
                    ? buildSalesReportModuleTitle(bundle, module.key, module.label.replace(/^\d+\.\s*/, ''))
                    : module.label;
            }
            if (textarea) textarea.value = salesReportAdviceState[module.key] || '';
        });
        mask.classList.add('show');
    }

    function closeSalesReportAdviceModal() {
        const mask = document.getElementById('sr-sales-report-advice-modal-mask');
        if (mask) mask.classList.remove('show');
    }

    function saveSalesReportAdvice() {
        const nextState = { ...salesReportAdviceState };
        SALES_REPORT_ADVICE_MODULES.forEach((module) => {
            const textarea = document.getElementById(`sr-sales-report-advice-${module.key}`);
            nextState[module.key] = normalizeAdviceText(textarea?.value || '');
        });
        salesReportAdviceState = nextState;
        salesReportSnapshot = null;
        syncSalesReportPreviewIfVisible();
        const statusEl = document.getElementById('sr-status');
        if (statusEl) statusEl.innerText = '已保存销售报告意见建议，后续预览和导出会自动带入';
        closeSalesReportAdviceModal();
    }

    async function fetchSalesReportCustomerNames() {
        const btn = document.getElementById('sr-sales-report-fetch-cn-name');
        const checkbox = document.getElementById('sr-sales-report-use-cn-name');
        const statusEl = document.getElementById('sr-status');
        if (btn) {
            btn.disabled = true;
            btn.innerText = '获取中...';
        }
        try {
            const records = await fetchAllCustomerNames((current, pages, count) => {
                if (statusEl) statusEl.innerText = `正在获取客户中文名称... 第 ${current}/${pages} 页，已获取 ${count} 条`;
            });
            salesReportCustomerNameMap = buildSalesReportCustomerNameMap(records);
            if (checkbox) checkbox.checked = true;
            updateSalesReportCustomerNameState();
            refreshSalesReportPreviewByOptions();
            if (statusEl) statusEl.innerText = `已获取 ${records.length} 个客户中文名称，销售报告标题将使用客户名称（公司名称）格式`;
        } catch (error) {
            logger.error('获取客户中文名称失败', error);
            alert(`获取客户中文名称失败: ${error.message}`);
            updateSalesReportCustomerNameState();
            if (statusEl) statusEl.innerText = `获取客户中文名称失败: ${error.message}`;
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerText = '获取客户公司名称';
            }
        }
    }

    function previewSalesReport() {
        const customerKey = document.getElementById('sr-sales-report-customer')?.value || '';
        if (!customerKey) return alert('请先选择客户');
        const bundle = getSalesReportCustomerBundle(customerKey);
        if (!bundle) return alert('当前客户在已完成的分析模块中没有可用于生成报告的数据');
        const narratives = buildSalesReportNarratives(bundle);
        const images = buildSalesReportImages(bundle);
        renderSalesReportPreview(bundle, images, narratives);
        const statusEl = document.getElementById('sr-status');
        if (statusEl) statusEl.innerText = `已生成 ${bundle.reportTitle || bundle.customerName} 预览，已使用模块：${getSalesReportModuleLabels(bundle.availableModuleKeys).join('、')}`;
    }

    async function exportCurrentSalesReport() {
        const customerKey = document.getElementById('sr-sales-report-customer')?.value || '';
        if (!customerKey) return alert('请先选择客户');
        const options = getSalesReportOptions();
        const canReuseSnapshot = salesReportSnapshot?.bundle?.customerKey === customerKey
            && isSameSalesReportOptions(salesReportSnapshot?.options, options);
        const bundle = canReuseSnapshot
            ? salesReportSnapshot.bundle
            : getSalesReportCustomerBundle(customerKey);
        if (!bundle) return alert('当前客户在已完成的分析模块中没有可用于导出的报告数据');
        const narratives = canReuseSnapshot ? salesReportSnapshot.narratives : buildSalesReportNarratives(bundle);
        const images = canReuseSnapshot ? salesReportSnapshot.images : buildSalesReportImages(bundle);
        const html = options.denseMode
            ? buildSalesReportDenseHtml(bundle, images, narratives, options)
            : buildSalesReportHtml(bundle, images, narratives, options);
        downloadBlob(new Blob([html], { type: 'text/html;charset=utf-8' }), `${getSalesReportFileBase(bundle)}.html`);
        const statusEl = document.getElementById('sr-status');
        if (statusEl) statusEl.innerText = `已导出 ${bundle.reportTitle || bundle.customerName}，已使用模块：${getSalesReportModuleLabels(bundle.availableModuleKeys).join('、')}`;
    }

    async function exportCurrentSalesReportPdf() {
        const customerKey = document.getElementById('sr-sales-report-customer')?.value || '';
        if (!customerKey) return alert('请先选择客户');
        const options = getSalesReportOptions();
        const canReuseSnapshot = salesReportSnapshot?.bundle?.customerKey === customerKey
            && isSameSalesReportOptions(salesReportSnapshot?.options, options);
        const bundle = canReuseSnapshot
            ? salesReportSnapshot.bundle
            : getSalesReportCustomerBundle(customerKey);
        if (!bundle) return alert('当前客户在已完成的分析模块中没有可用于导出的 PDF 报告数据');

        const btn = document.getElementById('sr-sales-report-export-pdf-btn');
        const statusEl = document.getElementById('sr-status');
        if (btn) btn.disabled = true;

        try {
            const narratives = canReuseSnapshot ? salesReportSnapshot.narratives : buildSalesReportNarratives(bundle);
            const images = canReuseSnapshot ? salesReportSnapshot.images : buildSalesReportImages(bundle);
            const html = options.denseMode
                ? buildSalesReportDenseHtml(bundle, images, narratives, options)
                : buildSalesReportHtml(bundle, images, narratives, options);
            const filename = `${getSalesReportFileBase(bundle)}.pdf`;
            const blob = await generateSalesReportPdfBlob(html, filename);
            downloadBlob(blob, filename);
            if (statusEl) statusEl.innerText = `已导出 ${bundle.reportTitle || bundle.customerName} PDF，已使用模块：${getSalesReportModuleLabels(bundle.availableModuleKeys).join('、')}`;
        } catch (error) {
            logger.error('导出 PDF 销售报告失败', error);
            alert(`导出 PDF 销售报告失败: ${error.message}`);
            if (statusEl) statusEl.innerText = `导出 PDF 销售报告失败: ${error.message}`;
        } finally {
            if (btn) btn.disabled = false;
            updateSalesReportReadyState();
        }
    }

    async function exportAllSalesReports() {
        if (!window.JSZip) return alert('ZIP组件加载中...');
        const customers = buildSalesReportCustomerOptions();
        if (customers.length === 0) return alert('暂无可导出的客户报告');

        const btn = document.getElementById('sr-sales-report-export-all-btn');
        const exportPdfBtn = document.getElementById('sr-sales-report-export-pdf-btn');
        const exportAllPdfBtn = document.getElementById('sr-sales-report-export-all-pdf-btn');
        const previewBtn = document.getElementById('sr-sales-report-preview-btn');
        const exportBtn = document.getElementById('sr-sales-report-export-btn');
        const refreshBtn = document.getElementById('sr-sales-report-refresh');
        const statusEl = document.getElementById('sr-status');
        const zip = new JSZip();
        const options = getSalesReportOptions();

        btn.disabled = true;
        if (exportPdfBtn) exportPdfBtn.disabled = true;
        if (exportAllPdfBtn) exportAllPdfBtn.disabled = true;
        previewBtn.disabled = true;
        exportBtn.disabled = true;
        refreshBtn.disabled = true;

        try {
            for (let index = 0; index < customers.length; index++) {
                const customer = customers[index];
                const bundle = getSalesReportCustomerBundle(customer.key);
                if (!bundle) continue;
                const narratives = buildSalesReportNarratives(bundle);
                const images = buildSalesReportImages(bundle);
                const html = options.denseMode
                    ? buildSalesReportDenseHtml(bundle, images, narratives, options)
                    : buildSalesReportHtml(bundle, images, narratives, options);
                zip.file(`${getSalesReportFileBase(bundle)}.html`, html);
                if (statusEl) statusEl.innerText = `正在生成销售报告 ${index + 1}/${customers.length}：${bundle.reportTitle || bundle.customerName}`;
                await sleep(20);
            }

            const blob = await zip.generateAsync({ type: 'blob' });
            downloadBlob(blob, `销售报告_${formatDateStandard(new Date())}.zip`);
            if (statusEl) statusEl.innerText = `已打包导出 ${customers.length} 份销售报告`;
        } catch (error) {
            logger.error('批量导出销售报告失败', error);
            alert(`批量导出销售报告失败: ${error.message}`);
            if (statusEl) statusEl.innerText = `批量导出销售报告失败: ${error.message}`;
        } finally {
            btn.disabled = false;
            previewBtn.disabled = false;
            refreshBtn.disabled = false;
            updateSalesReportReadyState();
        }
    }

    async function exportAllSalesReportsPdf() {
        if (!window.JSZip) return alert('ZIP组件加载中...');
        const runtime = getSalesReportPdfRuntime();
        if (!runtime.html2canvas || !runtime.jsPDF) return alert('PDF组件加载失败，请刷新页面后重试');
        const customers = buildSalesReportCustomerOptions();
        if (customers.length === 0) return alert('暂无可导出的客户报告');

        const btn = document.getElementById('sr-sales-report-export-all-pdf-btn');
        const exportHtmlBtn = document.getElementById('sr-sales-report-export-btn');
        const exportSinglePdfBtn = document.getElementById('sr-sales-report-export-pdf-btn');
        const exportAllHtmlBtn = document.getElementById('sr-sales-report-export-all-btn');
        const previewBtn = document.getElementById('sr-sales-report-preview-btn');
        const refreshBtn = document.getElementById('sr-sales-report-refresh');
        const statusEl = document.getElementById('sr-status');
        const zip = new JSZip();
        const options = getSalesReportOptions();

        if (btn) btn.disabled = true;
        if (exportHtmlBtn) exportHtmlBtn.disabled = true;
        if (exportSinglePdfBtn) exportSinglePdfBtn.disabled = true;
        if (exportAllHtmlBtn) exportAllHtmlBtn.disabled = true;
        if (previewBtn) previewBtn.disabled = true;
        if (refreshBtn) refreshBtn.disabled = true;

        try {
            for (let index = 0; index < customers.length; index++) {
                const customer = customers[index];
                const bundle = getSalesReportCustomerBundle(customer.key);
                if (!bundle) continue;
                const narratives = buildSalesReportNarratives(bundle);
                const images = buildSalesReportImages(bundle);
                const html = options.denseMode
                    ? buildSalesReportDenseHtml(bundle, images, narratives, options)
                    : buildSalesReportHtml(bundle, images, narratives, options);
                const fileBase = getSalesReportFileBase(bundle);
                const pdfBlob = await generateSalesReportPdfBlob(html, `${fileBase}.pdf`);
                zip.file(`${fileBase}.pdf`, pdfBlob);
                if (statusEl) statusEl.innerText = `正在生成 PDF 销售报告 ${index + 1}/${customers.length}：${bundle.reportTitle || bundle.customerName}`;
                await sleep(60);
            }

            const blob = await zip.generateAsync({ type: 'blob' });
            downloadBlob(blob, `销售报告_PDF_${formatDateStandard(new Date())}.zip`);
            if (statusEl) statusEl.innerText = `已打包导出 ${customers.length} 份 PDF 销售报告`;
        } catch (error) {
            logger.error('批量导出 PDF 销售报告失败', error);
            alert(`批量导出 PDF 销售报告失败: ${error.message}`);
            if (statusEl) statusEl.innerText = `批量导出 PDF 销售报告失败: ${error.message}`;
        } finally {
            if (btn) btn.disabled = false;
            if (previewBtn) previewBtn.disabled = false;
            if (refreshBtn) refreshBtn.disabled = false;
            updateSalesReportReadyState();
        }
    }

    function buildSkuSalesReport(records, startStr, endStr) {
        const customerMap = {};
        const skuMap = {};
        const customerSkuMap = {};
        const dailyTotalsMap = {};
        const detailRows = [];
        const customerOrderSet = new Set();
        let totalQty = 0;

        records.forEach((record) => {
            const customerName = normalizeSalesText(record.customerName, '未知客户');
            const customerCode = normalizeSalesText(record.customerCode, '');
            const orderNo = normalizeSalesText(record.deliveryNo || record.sourceNo || record.referOrderNo || record.platformOrderNo, '');
            const dateKey = normalizeSalesText((record.createTime || record.orderCreateTime || '').slice(0, 10), '未知日期');
            const productList = Array.isArray(record.productList) ? record.productList : [];

            if (orderNo) {
                customerOrderSet.add(`${customerCode}__${customerName}__${orderNo}`);
            }

            productList.forEach((item) => {
                const sku = normalizeSalesText(item.productSku, '-');
                const productName = normalizeSalesText(item.productName || item.productSku, sku);
                const qty = Number(item.qty || 0);
                if (!sku || qty <= 0) return;

                const customerKey = `${customerCode}__${customerName}`;
                const skuKey = `${sku}__${productName}`;
                const customerSkuKey = `${customerKey}__${skuKey}`;

                totalQty += qty;
                dailyTotalsMap[dateKey] = (dailyTotalsMap[dateKey] || 0) + qty;

                if (!customerMap[customerKey]) {
                    customerMap[customerKey] = {
                        customerCode,
                        customerName,
                        totalQty: 0,
                        skuSet: new Set(),
                        orderSet: new Set()
                    };
                }
                customerMap[customerKey].totalQty += qty;
                customerMap[customerKey].skuSet.add(skuKey);
                if (orderNo) customerMap[customerKey].orderSet.add(orderNo);

                if (!skuMap[skuKey]) {
                    skuMap[skuKey] = {
                        sku,
                        productName,
                        totalQty: 0,
                        customerSet: new Set()
                    };
                }
                skuMap[skuKey].totalQty += qty;
                skuMap[skuKey].customerSet.add(customerKey);

                if (!customerSkuMap[customerSkuKey]) {
                    customerSkuMap[customerSkuKey] = {
                        customerCode,
                        customerName,
                        sku,
                        productName,
                        qty: 0
                    };
                }
                customerSkuMap[customerSkuKey].qty += qty;
            });
        });

        const customerRows = Object.values(customerMap)
            .map((item) => ({
                customerCode: item.customerCode,
                customerName: item.customerName,
                totalQty: item.totalQty,
                skuCount: item.skuSet.size,
                orderCount: item.orderSet.size
            }))
            .sort((a, b) => a.customerName.localeCompare(b.customerName, 'zh-Hans-CN') || (a.customerCode || '').localeCompare(b.customerCode || ''));

        const skuRows = Object.values(skuMap)
            .map((item) => ({
                sku: item.sku,
                productName: item.productName,
                totalQty: item.totalQty,
                customerCount: item.customerSet.size
            }))
            .sort((a, b) => b.totalQty - a.totalQty || b.customerCount - a.customerCount || a.sku.localeCompare(b.sku, 'zh-Hans-CN'));

        Object.values(customerSkuMap).forEach((item) => {
            const customerMeta = customerMap[`${item.customerCode}__${item.customerName}`];
            detailRows.push({
                customerCode: item.customerCode,
                customerName: item.customerName,
                sku: item.sku,
                productName: item.productName,
                qty: item.qty,
                customerTotalQty: customerMeta ? customerMeta.totalQty : item.qty,
                customerOrderCount: customerMeta ? customerMeta.orderSet.size : 0
            });
        });

        detailRows.sort((a, b) =>
            a.customerName.localeCompare(b.customerName, 'zh-Hans-CN')
            || (a.customerCode || '').localeCompare(b.customerCode || '')
            || b.qty - a.qty
            || a.sku.localeCompare(b.sku, 'zh-Hans-CN')
        );

        const customerCharts = customerRows.map((customerRow) => {
            const rows = detailRows
                .filter((row) => row.customerCode === customerRow.customerCode && row.customerName === customerRow.customerName)
                .sort((a, b) => b.qty - a.qty || a.sku.localeCompare(b.sku, 'zh-Hans-CN'));
            return {
                ...customerRow,
                topSkus: rows.slice(0, 4),
                otherQty: rows.slice(4).reduce((sum, row) => sum + Number(row.qty || 0), 0)
            };
        });

        const dailyRows = Object.keys(dailyTotalsMap)
            .sort((a, b) => a.localeCompare(b))
            .map((date) => ({ date, qty: dailyTotalsMap[date] }));

        const topCustomer = customerRows[0] || null;
        const topSku = skuRows[0] || null;

        return {
            startStr,
            endStr,
            totalQty,
            orderCount: customerOrderSet.size,
            customerCount: customerRows.length,
            skuCount: skuRows.length,
            detailCount: detailRows.length,
            topCustomer,
            topSku,
            customerRows,
            skuRows,
            detailRows,
            customerCharts,
            dailyRows
        };
    }

    function getSkuSalesDimensionMeta() {
        const dim = document.getElementById('sr-sku-sales-dim')?.value || 'detail';
        if (!skuSalesReportData) return { dim, rows: [] };
        if (dim === 'customer') return { dim, rows: skuSalesReportData.customerRows || [] };
        if (dim === 'sku') return { dim, rows: skuSalesReportData.skuRows || [] };
        return { dim, rows: skuSalesReportData.detailRows || [] };
    }

    function getSkuSalesRowLabel(row, dim) {
        if (dim === 'customer') return row.customerName || '-';
        if (dim === 'sku') return row.sku || '-';
        return `${row.customerName || '-'} / ${row.sku || '-'}`;
    }

    function renderSkuSalesCustomerCharts() {
        const container = document.getElementById('sr-sku-sales-customer-charts');
        if (!container) return;

        container.innerHTML = '';
        charts.skuSalesCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.skuSalesCustomerCharts = [];

        if (!window.echarts || !skuSalesReportData) return;

        const groups = (skuSalesReportData.customerCharts || []).filter((item) => (item.topSkus || []).length > 0);
        if (groups.length === 0) {
            container.innerHTML = '<div style="color:#999; padding:12px 4px;">暂无客户SKU销量图表数据</div>';
            return;
        }

        const renderQueue = [];
        groups.forEach((group, index) => {
            const card = document.createElement('div');
            card.className = 'sr-inventory-card';
            const chartId = `sr-sku-sales-customer-bar-${index}`;
            card.innerHTML = `
                <div class="sr-inventory-card-title">${group.customerName || '-'}</div>
                <div class="sr-inventory-card-meta">订单数 ${group.orderCount || 0} ｜ SKU数 ${group.skuCount || 0} ｜ 总销量 ${formatChartNumber(group.totalQty)}</div>
                <div id="${chartId}" class="sr-inventory-card-box"></div>
            `;
            container.appendChild(card);
            renderQueue.push({ chartId, group });
        });

        requestAnimationFrame(() => {
            renderQueue.forEach(({ chartId, group }) => {
                const chartEl = document.getElementById(chartId);
                if (!chartEl) return;

                const chart = echarts.init(chartEl);
                const chartRows = (group.topSkus || [])
                    .map((item) => ({ name: item.sku, value: Number(item.qty || 0) }))
                    .concat(Number(group.otherQty || 0) > 0 ? [{ name: '其他SKU', value: Number(group.otherQty || 0) }] : [])
                    .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name, 'zh-Hans-CN'));
                const labels = chartRows.map((item) => item.name);
                const values = chartRows.map((item) => item.value);

                chart.setOption({
                    title: { text: 'SKU销量前四', left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                    grid: { left: '6%', right: '4%', bottom: 54, top: 42, containLabel: true },
                    xAxis: {
                        type: 'category',
                        data: labels,
                        axisLabel: { interval: 0, rotate: 20 }
                    },
                    yAxis: { type: 'value', name: '销量' },
                    series: [{
                        name: '销量',
                        type: 'bar',
                        data: values,
                        barMaxWidth: 42,
                        itemStyle: { color: '#5b8ff9' },
                        label: {
                            show: true,
                            position: 'top',
                            formatter: ({ value }) => formatChartNumber(value)
                        }
                    }]
                }, true);
                chart.resize();
                charts.skuSalesCustomerCharts.push(chart);
            });
        });
    }

    function renderSkuSalesAnalysisTable() {
        const thead = document.querySelector('#sr-sku-sales-table thead');
        const tbody = document.querySelector('#sr-sku-sales-table tbody');
        if (!thead || !tbody) return;

        const { dim, rows } = getSkuSalesDimensionMeta();
        const pagination = paginateRows(rows, skuSalesTableState.page, skuSalesTableState.pageSize);
        skuSalesTableState.page = pagination.currentPage;

        if (dim === 'customer') {
            thead.innerHTML = '<tr><th>客户名称</th><th>总销量</th><th>SKU数</th><th>订单数</th></tr>';
        } else if (dim === 'sku') {
            thead.innerHTML = '<tr><th>SKU</th><th>产品名称</th><th>总销量</th><th>客户数</th></tr>';
        } else {
            thead.innerHTML = '<tr><th>客户名称</th><th>SKU</th><th>产品名称</th><th>销量</th><th>客户总销量</th><th>订单数</th></tr>';
        }

        tbody.innerHTML = '';
        if (!skuSalesReportData || rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="${dim === 'detail' ? 6 : 4}" style="color:#999; padding:30px 0;">暂无SKU销量分析结果</td></tr>`;
            document.getElementById('sr-sku-sales-page-info').innerText = '共 0 条';
            document.getElementById('sr-sku-sales-prev').disabled = true;
            document.getElementById('sr-sku-sales-next').disabled = true;
            syncPageJumpControl('sr-sku-sales-page-jump', 'sr-sku-sales-page-go', 1, 1, true);
            return;
        }

        pagination.pageRows.forEach((row) => {
            const tr = document.createElement('tr');
            if (dim === 'customer') {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row.customerName}</td><td>${formatChartNumber(row.totalQty)}</td><td>${row.skuCount}</td><td>${row.orderCount}</td>`;
            } else if (dim === 'sku') {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row.sku}</td><td style="text-align:left;">${row.productName}</td><td>${formatChartNumber(row.totalQty)}</td><td>${row.customerCount}</td>`;
            } else {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row.customerName}</td><td>${row.sku}</td><td style="text-align:left;">${row.productName}</td><td>${formatChartNumber(row.qty)}</td><td>${formatChartNumber(row.customerTotalQty)}</td><td>${row.customerOrderCount}</td>`;
            }
            tbody.appendChild(tr);
        });

        document.getElementById('sr-sku-sales-page-info').innerText = `第 ${pagination.currentPage}/${pagination.totalPages} 页，共 ${pagination.total} 条`;
        document.getElementById('sr-sku-sales-prev').disabled = pagination.currentPage <= 1;
        document.getElementById('sr-sku-sales-next').disabled = pagination.currentPage >= pagination.totalPages;
        syncPageJumpControl('sr-sku-sales-page-jump', 'sr-sku-sales-page-go', pagination.totalPages, pagination.currentPage, false);
    }

    function renderSkuSalesAnalysisCharts() {
        if (!window.echarts || !skuSalesReportData) return;

        const customerRows = (skuSalesReportData.customerRows || [])
            .slice()
            .sort((a, b) => Number(b.totalQty || 0) - Number(a.totalQty || 0) || a.customerName.localeCompare(b.customerName, 'zh-Hans-CN'))
            .slice(0, 12);
        const skuRows = (skuSalesReportData.skuRows || [])
            .slice()
            .sort((a, b) => Number(b.totalQty || 0) - Number(a.totalQty || 0) || a.sku.localeCompare(b.sku, 'zh-Hans-CN'))
            .slice(0, 12);
        const dailyRows = skuSalesReportData.dailyRows || [];

        if (!charts.skuSalesCustomerBar) charts.skuSalesCustomerBar = echarts.init(document.getElementById('chart-sku-sales-customer'));
        charts.skuSalesCustomerBar.setOption({
            title: { text: '客户销量排行', left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            grid: { left: '4%', right: '4%', bottom: 68, top: 48, containLabel: true },
            xAxis: { type: 'category', data: customerRows.map((row) => row.customerName), axisLabel: { interval: 0, rotate: 22 } },
            yAxis: { type: 'value', name: '销量' },
            series: [{
                name: '销量',
                type: 'bar',
                barMaxWidth: 38,
                itemStyle: { color: '#36cfc9' },
                data: customerRows.map((row) => Number(row.totalQty || 0))
            }]
        }, true);

        if (!charts.skuSalesSkuBar) charts.skuSalesSkuBar = echarts.init(document.getElementById('chart-sku-sales-sku'));
        charts.skuSalesSkuBar.setOption({
            title: { text: 'SKU销量排行', left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            grid: { left: '4%', right: '4%', bottom: 68, top: 48, containLabel: true },
            xAxis: { type: 'category', data: skuRows.map((row) => row.sku), axisLabel: { interval: 0, rotate: 22 } },
            yAxis: { type: 'value', name: '销量' },
            series: [{
                name: '销量',
                type: 'bar',
                barMaxWidth: 38,
                itemStyle: { color: '#ff9c6e' },
                data: skuRows.map((row) => Number(row.totalQty || 0))
            }]
        }, true);

        if (!charts.skuSalesTrend) charts.skuSalesTrend = echarts.init(document.getElementById('chart-sku-sales-trend'));
        charts.skuSalesTrend.setOption({
            title: { text: '周期销量趋势', left: 'center' },
            tooltip: { trigger: 'axis' },
            grid: { left: '4%', right: '4%', bottom: 30, top: 48, containLabel: true },
            xAxis: { type: 'category', data: dailyRows.map((row) => row.date) },
            yAxis: { type: 'value', name: '销量' },
            series: [{
                name: '销量',
                type: 'line',
                smooth: true,
                symbolSize: 8,
                lineStyle: { width: 3, color: '#597ef7' },
                itemStyle: { color: '#597ef7' },
                areaStyle: { color: 'rgba(89, 126, 247, 0.18)' },
                data: dailyRows.map((row) => Number(row.qty || 0))
            }]
        }, true);
    }

    function renderSkuSalesAnalysisView() {
        if (!skuSalesReportData) return;

        document.getElementById('ssb-customer-count').innerText = skuSalesReportData.customerCount;
        document.getElementById('ssb-sku-count').innerText = skuSalesReportData.skuCount;
        document.getElementById('ssb-detail-count').innerText = skuSalesReportData.detailCount;
        document.getElementById('ssb-total-qty').innerText = formatChartNumber(skuSalesReportData.totalQty);

        const topCustomerText = skuSalesReportData.topCustomer
            ? `${skuSalesReportData.topCustomer.customerName}（${formatChartNumber(skuSalesReportData.topCustomer.totalQty)}）`
            : '-';
        const topSkuText = skuSalesReportData.topSku
            ? `${skuSalesReportData.topSku.sku}（${formatChartNumber(skuSalesReportData.topSku.totalQty)}）`
            : '-';

        document.getElementById('sr-sku-sales-summary').innerText =
            `${skuSalesReportData.startStr} 至 ${skuSalesReportData.endStr}，共统计 ${skuSalesReportData.orderCount} 个订单商品行，销量最高客户：${topCustomerText}，销量最高SKU：${topSkuText}`;

        renderSkuSalesAnalysisCharts();
        renderSkuSalesCustomerCharts();
        renderSkuSalesAnalysisTable();
    }

    function resetSkuSalesAnalysisView() {
        document.getElementById('ssb-customer-count').innerText = '-';
        document.getElementById('ssb-sku-count').innerText = '-';
        document.getElementById('ssb-detail-count').innerText = '-';
        document.getElementById('ssb-total-qty').innerText = '-';
        document.getElementById('sr-sku-sales-summary').innerText = '请先点击【开始SKU销量分析】。';
        document.querySelector('#sr-sku-sales-table thead').innerHTML = '';
        document.querySelector('#sr-sku-sales-table tbody').innerHTML = '<tr><td colspan="6" style="color:#999; padding:30px 0;">暂无SKU销量分析结果</td></tr>';
        document.getElementById('sr-sku-sales-page-info').innerText = '共 0 条';
        syncPageJumpControl('sr-sku-sales-page-jump', 'sr-sku-sales-page-go', 1, 1, true);
        if (charts.skuSalesCustomerBar) charts.skuSalesCustomerBar.clear();
        if (charts.skuSalesSkuBar) charts.skuSalesSkuBar.clear();
        if (charts.skuSalesTrend) charts.skuSalesTrend.clear();
        const customerCharts = document.getElementById('sr-sku-sales-customer-charts');
        if (customerCharts) customerCharts.innerHTML = '';
        charts.skuSalesCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.skuSalesCustomerCharts = [];
    }

    async function computeSkuSalesAnalysisReport(startStr, endStr, skipWeekends, progress) {
        const startTime = `${startStr} 00:00:00`;
        const endTime = `${endStr} 23:59:59`;
        const records = [];
        let current = 1;
        let pages = 1;

        while (current <= pages) {
            if (progress) progress(current, pages, records.length);
            const data = await fetchOnlinePage(startTime, endTime, current);
            const pageRecords = (Array.isArray(data.records) ? data.records : []).filter((record) =>
                !skipWeekends || !isWeekendDate(record.createTime || record.orderCreateTime)
            );
            const pageSize = Number(data.size || 500);
            const pageTotal = Number(data.total || 0);
            pages = data.pages || (pageTotal > 0 ? Math.ceil(pageTotal / pageSize) : (pageRecords.length < pageSize ? current : current + 1));
            records.push(...pageRecords);
            current++;
            await sleep(150);
        }

        return buildSkuSalesReport(records, startStr, endStr);
    }

    async function startSkuSalesAnalysisProcess() {
        const btn = document.getElementById('sr-sku-sales-start-btn');
        const exportBtn = document.getElementById('sr-sku-sales-export');
        const startStr = document.getElementById('sr-sku-sales-start').value;
        const endStr = document.getElementById('sr-sku-sales-end').value;
        const skipWeekends = document.getElementById('sr-sku-sales-skip-weekends').checked;
        const statusEl = document.getElementById('sr-status');

        if (!startStr || !endStr) return alert('请先选择完整的SKU销量分析周期！');
        if (new Date(startStr) > new Date(endStr)) return alert('开始日期不能晚于结束日期！');

        btn.disabled = true;
        exportBtn.disabled = true;
        btn.innerText = 'SKU销量分析中...';
        skuSalesTableState.page = 1;
        resetSkuSalesAnalysisView();

        try {
            skuSalesReportData = await computeSkuSalesAnalysisReport(startStr, endStr, skipWeekends, (current, pages, count) => {
                btn.innerText = `SKU销量分析中... (${current}/${pages}页)`;
                statusEl.innerText = `正在拉取SKU销量分析数据... 已累计 ${count} 条订单`;
            });

            renderSkuSalesAnalysisView();
            exportBtn.disabled = false;
            statusEl.innerText = `✅ SKU销量分析完成，共汇总 ${skuSalesReportData.skuCount} 个SKU`;
            updateSalesReportReadyState();
        } catch (error) {
            skuSalesReportData = null;
            exportBtn.disabled = true;
            resetSkuSalesAnalysisView();
            statusEl.innerText = `❌ SKU销量分析失败: ${error.message}`;
            logger.error('SKU销量分析抓取失败', error);
            updateSalesReportReadyState();
        } finally {
            btn.disabled = false;
            btn.innerText = '开始SKU销量分析';
        }
    }

    function getSkuInventoryDimensionMeta() {
        const dim = document.getElementById('sr-sku-inventory-dim')?.value || 'detail';
        if (!skuInventoryReportData) return { dim, rows: [] };
        if (dim === 'customer') return { dim, rows: skuInventoryReportData.customerRows || [] };
        if (dim === 'sku') return { dim, rows: skuInventoryReportData.skuRows || [] };
        return { dim, rows: skuInventoryReportData.detailRows || [] };
    }

    function getSkuInventoryRowLabel(row, dim) {
        if (dim === 'customer') return row.customerName || '-';
        if (dim === 'sku') return row.sku || '-';
        return `${row.customerName || '-'} / ${row.sku || '-'}`;
    }

    const SKU_INVENTORY_CHART_LIMIT = 8;

    function renderSkuInventoryCustomerCharts() {
        const container = document.getElementById('sr-sku-inventory-customer-charts');
        if (!container) return;

        container.innerHTML = '';
        charts.skuInventoryCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.skuInventoryCustomerCharts = [];

        if (!window.echarts || !skuInventoryReportData) return;

        const detailGroupMap = {};
        (skuInventoryReportData.detailRows || []).forEach((row) => {
            const customerKey = `${row.customerCode || ''}__${row.customerName || ''}`;
            if (!detailGroupMap[customerKey]) detailGroupMap[customerKey] = [];
            detailGroupMap[customerKey].push(row);
        });

        const groups = (skuInventoryReportData.customerRows || [])
            .slice()
            .sort((a, b) => (a.customerName || '').localeCompare(b.customerName || '', 'zh-Hans-CN') || (a.customerCode || '').localeCompare(b.customerCode || ''))
            .map((customerRow) => {
                const customerKey = `${customerRow.customerCode || ''}__${customerRow.customerName || ''}`;
                const skuMap = {};
                (detailGroupMap[customerKey] || []).forEach((row) => {
                    const skuKey = row.sku || '';
                    const weight = row.closeStockQty > 0 ? row.closeStockQty : 1;
                    if (!skuMap[skuKey]) {
                        skuMap[skuKey] = { sku: skuKey, closeStockQty: 0, weightTotal: 0, rateWeighted: 0, daysWeighted: 0 };
                    }
                    skuMap[skuKey].closeStockQty += Number(row.closeStockQty || 0);
                    skuMap[skuKey].weightTotal += weight;
                    skuMap[skuKey].rateWeighted += row.stockTurnoverRate * weight;
                    skuMap[skuKey].daysWeighted += row.stockTurnoverDays * weight;
                });
                const rows = Object.values(skuMap)
                    .map((item) => ({
                        sku: item.sku,
                        closeStockQty: item.closeStockQty,
                        stockTurnoverRate: item.weightTotal === 0 ? 0 : Number((item.rateWeighted / item.weightTotal).toFixed(2)),
                        stockTurnoverDays: item.weightTotal === 0 ? 0 : Number((item.daysWeighted / item.weightTotal).toFixed(2))
                    }))
                    .sort((a, b) => b.closeStockQty - a.closeStockQty || b.stockTurnoverRate - a.stockTurnoverRate);
                return { customerRow, rows };
            })
            .filter((group) => group.rows.length > 0);

        if (groups.length === 0) {
            container.innerHTML = '<div style="color:#999; padding:12px 4px;">暂无客户图表数据</div>';
            return;
        }

        const renderQueue = [];
        groups.forEach(({ customerRow, rows }, index) => {
            const card = document.createElement('div');
            card.className = 'sr-inventory-card';
            const pieId = `sr-sku-inventory-customer-pie-${index}`;
            const barId = `sr-sku-inventory-customer-bar-${index}`;
            card.innerHTML = `
                <div class="sr-inventory-card-title">${customerRow.customerName || '-'}</div>
                <div class="sr-inventory-card-meta">SKU数 ${customerRow.skuCount || 0} ｜ 期末库存 ${customerRow.closeStockQty || 0}</div>
                <div class="sr-inventory-card-charts">
                    <div id="${pieId}" class="sr-inventory-card-box-half"></div>
                    <div id="${barId}" class="sr-inventory-card-box-half"></div>
                </div>
            `;
            container.appendChild(card);
            renderQueue.push({ pieId, barId, rows });
        });

        requestAnimationFrame(() => {
            renderQueue.forEach(({ pieId, barId, rows }) => {
                const pieEl = document.getElementById(pieId);
                const barEl = document.getElementById(barId);
                if (!pieEl || !barEl) return;

                const sortByMetricDesc = (list, field) => list
                    .slice()
                    .sort((a, b) => {
                        const aValue = Number(a[field] || 0);
                        const bValue = Number(b[field] || 0);
                        const aHasValue = aValue > 0 ? 1 : 0;
                        const bHasValue = bValue > 0 ? 1 : 0;
                        if (bHasValue !== aHasValue) return bHasValue - aHasValue;
                        return bValue - aValue || b.closeStockQty - a.closeStockQty || a.sku.localeCompare(b.sku, 'zh-Hans-CN');
                    });

                const turnoverRateRows = sortByMetricDesc(rows, 'stockTurnoverRate')
                    .filter((row) => Number(row.stockTurnoverRate || 0) > 0)
                    .slice(0, SKU_INVENTORY_CHART_LIMIT);
                const turnoverDaysRows = sortByMetricDesc(rows, 'stockTurnoverDays')
                    .filter((row) => Number(row.stockTurnoverDays || 0) > 0)
                    .slice(0, SKU_INVENTORY_CHART_LIMIT);

                const pieChart = echarts.init(pieEl);
                pieChart.setOption({
                    title: { text: `Top ${SKU_INVENTORY_CHART_LIMIT} SKU库存周转率`, left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'item', formatter: (params) => `${params.name}<br/>周转率: ${params.value}%<br/>占比: ${params.percent}%` },
                    legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 10 },
                    series: [{
                        type: 'pie',
                        radius: ['35%', '65%'],
                        center: ['38%', '50%'],
                        data: turnoverRateRows.map((row) => ({ name: row.sku, value: row.stockTurnoverRate })),
                        label: { formatter: '{b}\n{d}%' }
                    }]
                }, true);
                charts.skuInventoryCustomerCharts.push(pieChart);

                const barChart = echarts.init(barEl);
                barChart.setOption({
                    title: { text: `Top ${SKU_INVENTORY_CHART_LIMIT} SKU库存周转天数`, left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                    grid: { left: '8%', right: '6%', bottom: 72, top: 48, containLabel: true },
                    xAxis: { type: 'category', data: turnoverDaysRows.map((row) => row.sku), axisLabel: { interval: 0, rotate: 28, margin: 14 } },
                    yAxis: { type: 'value', name: '天数' },
                    series: [{
                        name: '库存周转天数',
                        type: 'bar',
                        data: turnoverDaysRows.map((row) => row.stockTurnoverDays),
                        itemStyle: { color: '#91cc75' },
                        barMaxWidth: 42,
                        label: { show: true, position: 'top', formatter: ({ value }) => value == null ? '' : value }
                    }]
                }, true);
                charts.skuInventoryCustomerCharts.push(barChart);
            });
        });
    }

    function renderSkuInventoryAnalysisTable() {
        const thead = document.querySelector('#sr-sku-inventory-table thead');
        const tbody = document.querySelector('#sr-sku-inventory-table tbody');
        const { dim, rows } = getSkuInventoryDimensionMeta();
        const pagination = paginateRows(rows, skuInventoryTableState.page, skuInventoryTableState.pageSize);
        skuInventoryTableState.page = pagination.currentPage;
        const pageRows = pagination.pageRows;

        if (dim === 'customer') {
            thead.innerHTML = '<tr><th>客户名称</th><th>SKU数</th><th>仓库数</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>';
        } else if (dim === 'sku') {
            thead.innerHTML = '<tr><th>SKU</th><th>产品名称</th><th>客户数</th><th>仓库数</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>';
        } else {
            thead.innerHTML = '<tr><th>客户名称</th><th>发货仓库</th><th>SKU</th><th>产品名称</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>';
        }

        tbody.innerHTML = '';
        if (!skuInventoryReportData || rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="${dim === 'customer' ? 9 : 10}" style="color:#999; padding:30px 0;">暂无SKU库存分析结果</td></tr>`;
            document.getElementById('sr-sku-inventory-page-info').innerText = '共 0 条';
            document.getElementById('sr-sku-inventory-prev').disabled = true;
            document.getElementById('sr-sku-inventory-next').disabled = true;
            return;
        }

        pageRows.forEach((row) => {
            const tr = document.createElement('tr');
            if (dim === 'customer') {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row.customerName}</td><td>${row.skuCount}</td><td>${row.warehouseCount}</td><td>${row.preStockQty}</td><td>${row.closeStockQty}</td><td>${row.outboundBookQty}</td><td>${row.stockTurnoverRate}%</td><td>${row.stockTurnoverDays}</td><td>${row.stockSaleRate}</td>`;
            } else if (dim === 'sku') {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row.sku}</td><td style="text-align:left;">${row.productName}</td><td>${row.customerCount}</td><td>${row.warehouseCount}</td><td>${row.preStockQty}</td><td>${row.closeStockQty}</td><td>${row.outboundBookQty}</td><td>${row.stockTurnoverRate}%</td><td>${row.stockTurnoverDays}</td><td>${row.stockSaleRate}</td>`;
            } else {
                tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${row.customerName}</td><td>${row.whName}</td><td>${row.sku}</td><td style="text-align:left;">${row.productName}</td><td>${row.preStockQty}</td><td>${row.closeStockQty}</td><td>${row.outboundBookQty}</td><td>${row.stockTurnoverRate}%</td><td>${row.stockTurnoverDays}</td><td>${row.stockSaleRate}</td>`;
            }
            tbody.appendChild(tr);
        });

        document.getElementById('sr-sku-inventory-page-info').innerText = `第 ${pagination.currentPage}/${pagination.totalPages} 页，共 ${pagination.total} 条`;
        document.getElementById('sr-sku-inventory-prev').disabled = pagination.currentPage <= 1;
        document.getElementById('sr-sku-inventory-next').disabled = pagination.currentPage >= pagination.totalPages;
        syncPageJumpControl('sr-sku-inventory-page-jump', 'sr-sku-inventory-page-go', pagination.totalPages, pagination.currentPage, false);
    }

    function renderSkuInventoryAnalysisCharts() {
        if (!window.echarts || !skuInventoryReportData) return;
        const { dim, rows } = getSkuInventoryDimensionMeta();
        const rankedRows = rows.slice().sort((a, b) => b.closeStockQty - a.closeStockQty || b.stockTurnoverRate - a.stockTurnoverRate);
        const stockRows = rankedRows.slice(0, SKU_INVENTORY_CHART_LIMIT);

        if (!charts.skuInventoryPie) charts.skuInventoryPie = echarts.init(document.getElementById('chart-sku-inventory-pie'));
        charts.skuInventoryPie.setOption({
            title: { text: dim === 'sku' ? `Top ${SKU_INVENTORY_CHART_LIMIT} SKU库存对比` : `Top ${SKU_INVENTORY_CHART_LIMIT} 库存对比`, left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            legend: { top: 28 },
            grid: { left: '3%', right: '3%', bottom: '10%', top: 72, containLabel: true },
            xAxis: { type: 'category', data: stockRows.map((row) => getSkuInventoryRowLabel(row, dim)), axisLabel: { interval: 0, rotate: 18 } },
            yAxis: { type: 'value', name: '库存' },
            series: [
                { name: '期初库存', type: 'bar', data: stockRows.map((row) => Number(row.preStockQty || 0)), itemStyle: { color: '#69c0ff' }, barMaxWidth: 28 },
                { name: '期末库存', type: 'bar', data: stockRows.map((row) => Number(row.closeStockQty || 0)), itemStyle: { color: '#722ed1' }, barMaxWidth: 28 }
            ]
        }, true);

        const barRows = rows
            .slice()
            .sort((a, b) => Number(b.stockTurnoverDays || 0) - Number(a.stockTurnoverDays || 0) || Number(b.closeStockQty || 0) - Number(a.closeStockQty || 0))
            .slice(0, SKU_INVENTORY_CHART_LIMIT);
        if (!charts.skuInventoryBar) charts.skuInventoryBar = echarts.init(document.getElementById('chart-sku-inventory-bar'));
        charts.skuInventoryBar.setOption({
            title: { text: dim === 'sku' ? `Top ${SKU_INVENTORY_CHART_LIMIT} SKU周转天数` : `Top ${SKU_INVENTORY_CHART_LIMIT} 周转天数`, left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            grid: { left: '3%', right: '3%', bottom: '10%', containLabel: true },
            xAxis: { type: 'category', data: barRows.map((row) => getSkuInventoryRowLabel(row, dim)), axisLabel: { interval: 0, rotate: 18 } },
            yAxis: { type: 'value', name: '天数' },
            series: [{ name: '库存周转天数', type: 'bar', data: barRows.map((row) => Number(row.stockTurnoverDays || 0)), itemStyle: { color: '#b37feb' } }]
        }, true);
    }

    function renderSkuInventoryAnalysisView() {
        if (!skuInventoryReportData) return;
        document.getElementById('sib-customer-count').innerText = skuInventoryReportData.customerCount;
        document.getElementById('sib-sku-count').innerText = skuInventoryReportData.skuCount;
        document.getElementById('sib-detail-count').innerText = skuInventoryReportData.detailCount;
        document.getElementById('sib-close-stock').innerText = skuInventoryReportData.totalCloseStockQty;

        const topSkuText = skuInventoryReportData.topSkuRow
            ? `${skuInventoryReportData.topSkuRow.sku}（${skuInventoryReportData.topSkuRow.closeStockQty}）`
            : '-';
        document.getElementById('sr-sku-inventory-summary').innerText =
            `${skuInventoryReportData.startStr} 至 ${skuInventoryReportData.endStr}，刷新日期 ${skuInventoryReportData.refreshTime || '-'}，总周转率 ${skuInventoryReportData.totalTurnoverRate}% ，总周转天数 ${skuInventoryReportData.totalTurnoverDays}，库存最高SKU：${topSkuText}`;

        renderSkuInventoryAnalysisCharts();
        renderSkuInventoryCustomerCharts();
        renderSkuInventoryAnalysisTable();
    }

    function resetSkuInventoryAnalysisView() {
        document.getElementById('sib-customer-count').innerText = '-';
        document.getElementById('sib-sku-count').innerText = '-';
        document.getElementById('sib-detail-count').innerText = '-';
        document.getElementById('sib-close-stock').innerText = '-';
        document.getElementById('sr-sku-inventory-summary').innerText = '请先点击【开始SKU库存分析】。';
        document.querySelector('#sr-sku-inventory-table thead').innerHTML = '';
        document.querySelector('#sr-sku-inventory-table tbody').innerHTML = '<tr><td colspan="10" style="color:#999; padding:30px 0;">暂无SKU库存分析结果</td></tr>';
        document.getElementById('sr-sku-inventory-page-info').innerText = '共 0 条';
        syncPageJumpControl('sr-sku-inventory-page-jump', 'sr-sku-inventory-page-go', 1, 1, true);
        if (charts.skuInventoryPie) charts.skuInventoryPie.clear();
        if (charts.skuInventoryBar) charts.skuInventoryBar.clear();
        const customerCharts = document.getElementById('sr-sku-inventory-customer-charts');
        if (customerCharts) customerCharts.innerHTML = '';
        charts.skuInventoryCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.skuInventoryCustomerCharts = [];
    }

    async function startSkuInventoryAnalysisProcess() {
        const btn = document.getElementById('sr-sku-inventory-start-btn');
        const exportBtn = document.getElementById('sr-sku-inventory-export');
        const startStr = document.getElementById('sr-sku-inventory-start').value;
        const endStr = document.getElementById('sr-sku-inventory-end').value;
        const skipWeekends = document.getElementById('sr-sku-inventory-skip-weekends').checked;
        const statusEl = document.getElementById('sr-status');

        if (!startStr || !endStr) return alert('请先选择完整的SKU库存分析周期！');
        if (new Date(startStr) > new Date(endStr)) return alert('开始日期不能晚于结束日期！');

        btn.disabled = true;
        exportBtn.disabled = true;
        btn.innerText = 'SKU库存分析中...';
        skuInventoryTableState.page = 1;
        resetSkuInventoryAnalysisView();

        try {
            const rawReport = await computeSkuInventoryAnalysisReport(startStr, endStr, skipWeekends, (current, pages, count) => {
                btn.innerText = `SKU库存分析中... (${current}/${pages}页)`;
                statusEl.innerText = `正在拉取SKU库存分析数据... 已累计 ${count} 条`;
            });
            skuInventoryReportData = normalizeSkuInventoryReport(rawReport);
            try {
                renderSkuInventoryAnalysisView();
                exportBtn.disabled = false;
                statusEl.innerText = `✅ SKU库存分析完成，共汇总 ${skuInventoryReportData.skuCount} 个SKU`;
                updateSalesReportReadyState();
            } catch (renderError) {
                exportBtn.disabled = true;
                reportSkuInventoryRenderError('渲染', renderError);
                updateSalesReportReadyState();
            }
        } catch (error) {
            skuInventoryReportData = null;
            exportBtn.disabled = true;
            resetSkuInventoryAnalysisView();
            statusEl.innerText = `❌ SKU库存分析失败: ${error.message}`;
            logger.error('SKU库存分析抓取失败', error);
            updateSalesReportReadyState();
        } finally {
            btn.disabled = false;
            btn.innerText = '开始SKU库存分析';
        }
    }

    function setupInventoryAnalysisUI() {
        const inventoryView = document.getElementById('sr-view-inventory');
        if (!inventoryView) return;

        const badges = inventoryView.querySelector('.sr-badges');
        if (badges) {
            badges.innerHTML = `
                <div class="sr-badge sr-badge-total"><span class="title">客户数</span><span class="num" id="ib-customer-count">-</span></div>
                <div class="sr-badge sr-badge-24"><span class="title">客户分仓组合数</span><span class="num" id="ib-detail-count">-</span></div>
                <div class="sr-badge sr-badge-48"><span class="title">期末库存</span><span class="num" id="ib-close-stock">-</span></div>
                <div class="sr-badge sr-badge-72"><span class="title">总周转率</span><span class="num" id="ib-turnover-rate">-</span></div>
            `;
        }

        const summaryControls = inventoryView.querySelectorAll('.sr-controls')[1];
        if (summaryControls) {
            summaryControls.innerHTML = `
                <select id="sr-inventory-dim" class="sr-select" style="width:220px;">
                    <option value="detail">按客户分仓查看</option>
                    <option value="customer">按客户汇总查看</option>
                </select>
                <div id="sr-inventory-summary" class="sr-inventory-summary">请先点击【开始库存分析】。</div>
            `;
        }

        const thead = inventoryView.querySelector('#sr-inventory-table thead');
        if (thead) thead.innerHTML = '';

        const chartRow = inventoryView.querySelector('.chart-row');
        if (chartRow && !inventoryView.querySelector('#sr-inventory-customer-charts')) {
            chartRow.insertAdjacentHTML('afterend', '<div id="sr-inventory-customer-charts" class="sr-inventory-cards"></div>');
        }
    }

    // 数据状态
    let finalReportData = {}, warehouseReportData = {}, customerReportData = {}, warehouseChannelReportData = {}, extendedData = {};
    let onlineRateDetailRows = [];
    let onlineRateDetailBuckets = { 'offline': [], '24h': [], '48h': [], '72h': [], 'over72h': [], 'cancelled': [] };
    let dateLabels = { '24h': '', '48h': '', '72h': '' };
    const charts = { bar: null, line: null, pie: null, rose: null, radar: null, outboundLine: null, outboundBar: null, regionBar: null, regionPie: null, regionDetailPies: [], inventoryPie: null, inventoryBar: null, inventoryCustomerCharts: [], skuInventoryPie: null, skuInventoryBar: null, skuInventoryCustomerCharts: [], skuSalesCustomerBar: null, skuSalesSkuBar: null, skuSalesTrend: null, skuSalesCustomerCharts: [], reportPreview: null };
    let outboundReportData = { rows: [], startDate: '', endDate: '', summary: { warehouseCount: 0, avg24: 0, avg48: 0, avg72: 0 } };
    let periodReportData = { channel: {}, warehouse: {}, customer: {}, summary: { total: 0, in24: 0, in48: 0, in72: 0 } };
    let regionReportData = null;
    let inventoryReportData = null;
    let skuInventoryReportData = null;
    let skuInventoryTableState = { page: 1, pageSize: 50 };
    let skuSalesReportData = null;
    let skuSalesTableState = { page: 1, pageSize: 50 };
    let salesReportCustomerNameMap = null;
    let salesReportSnapshot = null;
    let salesReportAdviceState = {
        region: '',
        inventory: '',
        skuInventory: '',
        skuSales: '',
        outboundEfficiency: ''
    };

    // ==========================================
    // 5. UI 界面构建
    // ==========================================
    function injectUI() {
        const style = document.createElement('style');
        style.innerHTML = `
            #sr-floating-btn { position: fixed; bottom: 30px; right: 30px; z-index: 99999; background: #1890ff; color: #fff; padding: 12px 20px; border-radius: 50px; box-shadow: 0 4px 12px rgba(24,144,255,0.4); cursor: grab; font-size: 14px; font-weight: bold; transition: background 0.3s, transform 0.3s; user-select: none; }
            #sr-floating-btn:hover { background: #40a9ff; transform: translateY(-2px); }
            #sr-panel { position: fixed; bottom: 80px; right: 30px; z-index: 99998; width: 980px; background: #fff; border-radius: 12px; box-shadow: 0 8px 30px rgba(0,0,0,0.25); display: none; flex-direction: column; overflow: hidden; font-family: sans-serif; border: 1px solid #e8e8e8; }
            #sr-panel.sr-fullscreen { top: 0; left: 0; right: 0; bottom: 0; width: 100vw; height: 100vh; border-radius: 0; border: none; box-shadow: none; }
            .sr-header { background: #fafafa; padding: 16px 20px; font-size: 16px; font-weight: bold; border-bottom: 1px solid #f0f0f0; display: flex; justify-content: space-between; align-items: center; color: #333; cursor: move; }
            #sr-panel.sr-fullscreen .sr-header { cursor: default; }
            .sr-header-actions { display:flex; align-items:center; gap:12px; }
            .sr-fullscreen-btn { cursor: pointer; color: #666; font-size: 14px; padding: 6px 10px; border-radius: 6px; border: 1px solid #d9d9d9; background:#fff; transition: all 0.3s; }
            .sr-fullscreen-btn:hover { color:#1890ff; border-color:#91caff; }
            .sr-close { cursor: pointer; color: #999; font-size: 20px; transition: color 0.3s; }
            .sr-close:hover { color: #ff4d4f; }
            .sr-nav-tabs { display: flex; background: #f5f5f5; border-bottom: 1px solid #d9d9d9;}
            .sr-nav-tab { padding: 10px 20px; cursor: pointer; font-size: 14px; color: #666; border-bottom: 2px solid transparent; transition: all 0.3s; }
            .sr-nav-tab.active { color: #1890ff; border-bottom: 2px solid #1890ff; background: #fff; font-weight: bold;}
            .sr-body { padding: 20px; max-height: 680px; overflow-y: auto; position: relative;}
            #sr-panel.sr-fullscreen .sr-body { max-height: calc(100vh - 70px); }
            .sr-view { display: none; }
            .sr-view.active { display: block; }
            .sr-controls { display: flex; gap: 10px; margin-bottom: 12px; align-items: center;}
            .sr-btn { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; color: #fff; transition: opacity 0.3s; font-weight: bold; }
            .sr-btn-primary { background: #1890ff; }
            .sr-btn-success { background: #52c41a; }
            .sr-btn-warning { background: #faad14; }
            .sr-btn-purple { background: #722ed1; }
            .sr-btn-info { background: #1677ff; }
            .sr-btn-danger { background: #cf1322; }
            .sr-select, .sr-input-date { background: #fff; color: #333; border: 1px solid #d9d9d9; padding: 7px 12px; border-radius: 6px; font-size: 13px; cursor:pointer;}
            .sr-btn:disabled { background: #f5f5f5; color:#b8b8b8; cursor: not-allowed; border-color:#d9d9d9;}
            .sr-btn:hover:not(:disabled) { opacity: 0.8; }
            #sr-status { font-size: 13px; color: #666; flex-grow: 1; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}
            .sr-status-bar { margin-bottom:12px; padding:8px 12px; background:#f8fafc; border:1px solid #eef2f7; border-radius:6px; }
            .sr-status-bar:empty { display:none; }
            .sr-table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }
            .sr-table th, .sr-table td { border: 1px solid #f0f0f0; padding: 8px; text-align: center; }
            .sr-table th { background: #fafafa; font-weight: bold; color: #333; white-space: pre-line; }
            .sr-table tbody tr:hover { background: #f5f5f5; }
            .echarts-box { width: 100%; height: 350px; margin-bottom: 30px; }
            .chart-row { display: flex; justify-content: space-between; gap: 20px; margin-bottom: 30px;}
            .echarts-box-half { width: 48%; height: 350px; }
            .chart-divider { border-top: 1px dashed #e8e8e8; margin: 30px 0; }
            .sr-badges { display: flex; gap: 15px; margin: 15px 0; }
            .sr-badge { flex: 1; padding: 15px; border-radius: 8px; text-align: center; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
            .sr-badge-total { background: linear-gradient(135deg, #8c8c8c, #bfbfbf); }
            .sr-badge-24 { background: linear-gradient(135deg, #1890ff, #69c0ff); }
            .sr-badge-48 { background: linear-gradient(135deg, #52c41a, #95de64); }
            .sr-badge-72 { background: linear-gradient(135deg, #faad14, #ffd666); }
            .sr-badge .title { font-size: 13px; opacity: 0.9; }
            .sr-badge .num { font-size: 28px; font-weight: bold; margin-top: 4px; display: block; }
            .sr-region-pies { display:grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap:20px; margin-top:18px; }
            .sr-region-pie-card { position: relative; overflow: hidden; border:1px solid #f0f0f0; border-radius:10px; padding:16px; background:#fff; }
            .sr-region-pie-title { font-size:13px; font-weight:bold; color:#333; margin-bottom:6px; }
            .sr-region-pie-meta { font-size:12px; color:#888; margin-bottom:10px; }
            .sr-region-pie-box { position: relative; overflow: hidden; width:100%; height:340px; }
            .sr-inventory-cards { display:grid; grid-template-columns: repeat(auto-fit, minmax(760px, 1fr)); gap:20px; margin-top:18px; }
            .sr-inventory-card { position: relative; overflow: hidden; border:1px solid #f0f0f0; border-radius:10px; padding:18px; background:#fff; }
            .sr-inventory-card-title { font-size:13px; font-weight:bold; color:#333; margin-bottom:6px; }
            .sr-inventory-card-meta { font-size:12px; color:#888; margin-bottom:12px; }
            .sr-inventory-card-charts { display:flex; flex-wrap:wrap; gap:20px; align-items:stretch; }
            .sr-inventory-card-box { position: relative; overflow: hidden; width:100%; height:320px; }
            .sr-inventory-card-box-half { position: relative; overflow: hidden; flex:1 1 420px; min-width:420px; height:420px; }
            .sr-inventory-summary { font-size:13px; color:#666; }
            .sr-table-pagination { display:flex; align-items:center; justify-content:flex-end; gap:10px; margin-top:14px; color:#666; font-size:13px; }
            .sr-table-pagination select { padding:6px 8px; border:1px solid #d9d9d9; border-radius:6px; background:#fff; }
            .sr-note-block { background:#fafafa; border:1px dashed #d9d9d9; border-radius:8px; padding:14px 16px; color:#555; line-height:1.8; font-size:13px; margin-bottom:14px; }
            .sr-note-title { font-weight:bold; color:#333; margin-bottom:6px; }
            .sr-note-list { margin: 6px 0 0 18px; padding: 0; }
            .sr-note-list li { margin: 2px 0; }
            .sr-link-btn { background:none; border:none; color:#1890ff; cursor:pointer; padding:0; font-size:13px; }
            .sr-link-btn:hover { text-decoration:underline; }
            .sr-modal-mask { position:fixed; inset:0; background:rgba(0,0,0,0.45); z-index:100000; display:none; align-items:center; justify-content:center; padding:20px; }
            .sr-modal-mask.show { display:flex; }
            .sr-modal { width:min(720px, 100%); max-height:min(80vh, 760px); background:#fff; border-radius:12px; box-shadow:0 12px 40px rgba(0,0,0,0.25); overflow:hidden; display:flex; flex-direction:column; }
            .sr-modal-header { display:flex; align-items:center; justify-content:space-between; padding:16px 18px; border-bottom:1px solid #f0f0f0; font-size:16px; font-weight:bold; color:#333; }
            .sr-modal-body { padding:16px 18px 18px; overflow:auto; color:#555; line-height:1.8; font-size:13px; }
            .sr-report-grid { display:grid; grid-template-columns: minmax(320px, 360px) minmax(0, 1fr); gap:20px; align-items:start; }
            .sr-report-side { display:flex; flex-direction:column; gap:16px; }
            .sr-report-card { border:1px solid #eceff3; border-radius:16px; background:linear-gradient(180deg, #ffffff, #fcfcfd); padding:18px; box-shadow:0 10px 30px rgba(15, 23, 42, 0.05); }
            .sr-report-card-title { font-size:15px; font-weight:700; color:#1f2937; margin-bottom:12px; }
            .sr-report-help { color:#667085; line-height:1.8; font-size:13px; }
            .sr-report-setting-stack { display:flex; flex-direction:column; gap:14px; }
            .sr-report-select-row { display:flex; flex-direction:column; gap:8px; }
            .sr-report-select-label { font-size:12px; font-weight:700; color:#6b7280; letter-spacing:0.02em; }
            .sr-report-option-row { display:flex; align-items:flex-start; gap:8px; padding:10px 12px; border:1px solid #e5e7eb; border-radius:12px; background:#f8fafc; color:#334155; cursor:pointer; }
            .sr-report-option-row input { margin-top:2px; }
            .sr-report-option-text { display:flex; flex-direction:column; gap:3px; line-height:1.45; }
            .sr-report-option-title { font-size:13px; font-weight:700; color:#1f2937; }
            .sr-report-option-desc { font-size:12px; color:#64748b; }
            .sr-report-actions { display:grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap:10px; }
            .sr-btn.sr-report-btn { min-height:42px; border-radius:12px; font-size:13px; font-weight:700; box-shadow:0 8px 18px rgba(15, 23, 42, 0.08); }
            .sr-btn.sr-report-btn:disabled { box-shadow:none; }
            .sr-btn.sr-btn-slate { background:linear-gradient(135deg, #475569, #64748b); }
            .sr-btn.sr-btn-teal { background:linear-gradient(135deg, #0f766e, #14b8a6); }
            .sr-btn.sr-btn-cyan { background:linear-gradient(135deg, #0891b2, #22d3ee); }
            .sr-btn.sr-btn-indigo { background:linear-gradient(135deg, #4f46e5, #818cf8); }
            .sr-btn.sr-btn-rose { background:linear-gradient(135deg, #e11d48, #fb7185); }
            .sr-btn.sr-btn-amber { background:linear-gradient(135deg, #d97706, #f59e0b); }
            .sr-report-preview { border:1px solid #e5e7eb; border-radius:16px; background:linear-gradient(180deg, #ffffff, #f8fafc); padding:16px; min-height:640px; box-shadow:0 12px 36px rgba(15, 23, 42, 0.06); }
            .sr-report-preview-empty { display:flex; align-items:center; justify-content:center; min-height:600px; text-align:center; color:#94a3b8; font-size:13px; line-height:1.9; border:1px dashed #cbd5e1; border-radius:14px; background:linear-gradient(180deg, #ffffff, #f8fafc); padding:24px; }
            .sr-report-preview-shell { border:1px solid #dbe4ee; border-radius:14px; overflow:hidden; background:#fff; box-shadow:0 12px 28px rgba(15, 23, 42, 0.08); }
            .sr-report-preview-toolbar { display:flex; align-items:center; justify-content:space-between; gap:12px; padding:12px 14px; border-bottom:1px solid #eef2f7; background:linear-gradient(180deg, #f8fafc, #f1f5f9); }
            .sr-report-preview-toolbar-left { display:flex; align-items:center; gap:10px; min-width:0; }
            .sr-report-preview-dots { display:flex; align-items:center; gap:6px; }
            .sr-report-preview-dot { width:10px; height:10px; border-radius:50%; background:#cbd5e1; }
            .sr-report-preview-dot.red { background:#fb7185; }
            .sr-report-preview-dot.yellow { background:#fbbf24; }
            .sr-report-preview-dot.green { background:#34d399; }
            .sr-report-preview-title { font-size:14px; font-weight:700; color:#111827; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
            .sr-report-preview-meta { font-size:12px; color:#64748b; text-align:right; }
            .sr-report-preview-frame { width:100%; height:940px; border:0; background:#fff; display:block; }
            .sr-editor-help { margin-bottom:12px; color:#64748b; line-height:1.75; }
            .sr-editor-meta { margin-bottom:12px; color:#475467; font-size:12px; }
            .sr-editor-toolbar { display:flex; flex-wrap:wrap; gap:10px; align-items:center; justify-content:space-between; margin-bottom:12px; }
            .sr-editor-chip-row { display:flex; flex-wrap:wrap; gap:8px; }
            .sr-editor-chip { display:inline-flex; align-items:center; gap:6px; padding:7px 10px; border:1px solid #dbe4ee; border-radius:999px; background:#f8fafc; color:#334155; font-size:12px; cursor:pointer; }
            .sr-editor-chip input { margin:0; }
            .sr-editor-table-wrap { border:1px solid #e5e7eb; border-radius:12px; overflow:auto; background:#fff; }
            .sr-editor-table { width:100%; border-collapse:collapse; font-size:13px; min-width:760px; }
            .sr-editor-table th, .sr-editor-table td { border:1px solid #eef2f7; padding:10px 12px; text-align:center; }
            .sr-editor-table th { background:#f8fafc; color:#334155; position:sticky; top:0; z-index:1; }
            .sr-editor-input { width:100%; min-width:88px; border:1px solid #d0d5dd; border-radius:8px; padding:7px 9px; font-size:13px; text-align:right; box-sizing:border-box; }
            .sr-editor-input:focus { outline:none; border-color:#3b82f6; box-shadow:0 0 0 3px rgba(59,130,246,0.12); }
            .sr-advice-grid { display:grid; grid-template-columns:1fr; gap:14px; }
            .sr-advice-card { border:1px solid #e5e7eb; border-radius:14px; background:#f8fafc; padding:14px; }
            .sr-advice-card-title { font-size:14px; font-weight:700; color:#111827; margin-bottom:8px; }
            .sr-advice-textarea { width:100%; min-height:92px; resize:vertical; border:1px solid #d0d5dd; border-radius:10px; padding:10px 12px; font-size:13px; line-height:1.75; box-sizing:border-box; font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif; }
            .sr-advice-textarea:focus { outline:none; border-color:#3b82f6; box-shadow:0 0 0 3px rgba(59,130,246,0.12); }
            .sr-modal-footer { display:flex; justify-content:flex-end; gap:10px; padding:14px 18px 18px; border-top:1px solid #eef2f7; background:#fff; }
            @media (max-width: 1180px) {
                .sr-report-grid { grid-template-columns: 1fr; }
                .sr-report-actions { grid-template-columns: 1fr; }
                .sr-report-preview-toolbar { flex-direction:column; align-items:flex-start; }
                .sr-report-preview-meta { text-align:left; }
            }
        `;
        document.head.appendChild(style);

        const html = `
            <div id="sr-floating-btn" title="按住拖拽，点击展开/收起">📊 物流统计控制台</div>
            <div id="sr-panel">
                <div class="sr-header"><span>物流数据看板</span><div class="sr-header-actions"><button class="sr-fullscreen-btn" id="sr-fullscreen-btn">全屏</button><span class="sr-close" id="sr-close-btn">×</span></div></div>
                <div class="sr-nav-tabs">
                    <div class="sr-nav-tab active" data-target="sr-view-table">📡 上网率</div>
                    <div class="sr-nav-tab" data-target="sr-view-period">📅 周期平均分析</div>
                    <div class="sr-nav-tab" data-target="sr-view-outbound">📦 出库发货率</div>
                    <div class="sr-nav-tab" data-target="sr-view-region">🗺️ 订单分布分析</div>
                    <div class="sr-nav-tab" data-target="sr-view-inventory">📦 按客户库存分析</div>
                    <div class="sr-nav-tab" data-target="sr-view-sku-inventory">🏷️ 按sku库存分析</div>
                    <div class="sr-nav-tab" data-target="sr-view-sku-sales">📊 sku销量分析</div>
                    <div class="sr-nav-tab" data-target="sr-view-sales-report">📝 销售报告</div>
                </div>
                <div class="sr-body">
                    <div id="sr-status" class="sr-status-bar" title=""></div>
                    <div id="sr-view-table" class="sr-view active">
                        <div class="sr-controls" style="background:#f9f9f9; padding:10px 12px; border-radius:6px; border:1px solid #eee;">
                            <button class="sr-btn sr-btn-primary" id="sr-start-btn">拉取上网率数据</button>
                            <span style="font-size:13px; font-weight:bold; color:#1677ff;">样本创建周期:</span>
                            <input type="date" id="sr-online-start" class="sr-input-date" title="开始日期"> <span style="font-size:12px;color:#666;">至</span>
                            <input type="date" id="sr-online-end" class="sr-input-date" title="结束日期">
                            <select id="sr-dimension-select" class="sr-select">
                                <option value="channel">按物流渠道查看</option>
                                <option value="warehouse">按发货仓库查看</option>
                                <option value="customer">按客户名称查看</option>
                            </select>
                            <label id="sr-warehouse-filter-wrap" style="display:none; align-items:center; gap:6px; font-size:13px; color:#333;">
                                <span>仓库筛选</span>
                                <select id="sr-warehouse-filter" class="sr-select">
                                    <option value="">全部仓库</option>
                                </select>
                            </label>
                            <label style="display:flex; align-items:center; font-size:13px; cursor:pointer; color:#d4380d; font-weight:bold;">
                                <input type="checkbox" id="sr-skip-weekends" checked style="margin-right:4px;"> 过滤周末
                            </label>
                            <button class="sr-btn sr-btn-success" id="sr-export-btn" disabled>导出 上网率 Excel</button>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>按所选创建日期范围内的订单作为固定样本，并按物流渠道、发货仓库或客户名称聚合展示；默认预置最近 3 个统计日样本。</div>
                            <div>24H、48H、72H 分别统计这同一批订单在创建后 24 / 48 / 72 小时内完成上网的比例，因此 48H 会包含 24H，72H 会包含 24H / 48H，整体呈递增趋势；若缺少上网时间但有出库时间（outboundTime），则按已上网兜底计入。</div>
                            <div>切换到“按物流渠道查看”时，可通过仓库筛选查看某个仓库内各渠道的上网率表现。</div>
                            <div>表格里的百分比为样本内在对应时效完成上网的订单数 ÷ 纳入统计的订单数，括号内显示具体单量；已取消（status=99）及上网时间超过 72H 的订单不计入总数。</div>
                            <div>订单明细会按首次命中的时效归类为 24H上网、48H上网、72H上网、已上网（超过72H）、已取消或未上网，可再按仓库和结果筛选查看。</div>
                        </div>
                        <table class="sr-table" id="sr-result-table">
                            <thead><tr><th id="th-dim-name" style="width: 25%">维度名称</th><th id="th-24h">24H上网率</th><th id="th-48h">48H上网率</th><th id="th-72h">72H上网率</th></tr></thead>
                            <tbody><tr><td colspan="4" style="color:#999; padding:30px 0;">暂无数据，请点击【拉取上网率数据】</td></tr></tbody>
                        </table>
                        <div class="chart-divider"></div>
                        <div id="chart-line" class="echarts-box" style="height:300px;"></div><div class="chart-divider"></div>
                        <div id="chart-bar" class="echarts-box"></div><div class="chart-divider"></div>
                        <div class="chart-row"><div id="chart-pie" class="echarts-box-half"></div><div id="chart-rose" class="echarts-box-half"></div></div><div class="chart-divider"></div>
                        <div id="chart-radar" class="echarts-box" style="height:400px;"></div>
                        <div class="sr-note-block" style="margin-top:15px;">
                            <div class="sr-note-title">订单明细</div>
                            <div style="margin-top:6px; color:#d46b08; font-size:13px;">说明：“已上网（超过72H）”和“已取消”的订单仅在明细中展示，不计入总数统计。</div>
                            <div class="sr-controls" style="margin-top:8px; margin-bottom:0; flex-wrap:wrap;">
                                <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#333;">
                                    <span>仓库筛选</span>
                                    <select id="sr-online-detail-warehouse" class="sr-select" style="width:220px;">
                                        <option value="">全部仓库</option>
                                    </select>
                                </label>
                                <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#333;">
                                    <span>结果筛选</span>
                                    <select id="sr-online-detail-status" class="sr-select" style="width:220px;">
                                        <option value="all">全部订单</option>
                                        <option value="offline">未上网</option>
                                        <option value="24h">24H上网</option>
                                        <option value="48h">48H上网</option>
                                        <option value="72h">72H上网</option>
                                        <option value="over72h">已上网（超过72H）</option>
                                        <option value="cancelled">已取消</option>
                                    </select>
                                </label>
                                <div id="sr-online-detail-summary" class="sr-inventory-summary">请先点击【拉取上网率数据】。</div>
                            </div>
                        </div>
                        <table class="sr-table" id="sr-online-detail-table">
                            <thead>
                                <tr>
                                    <th>订单号</th>
                                    <th>来源单号</th>
                                    <th>客户名称</th>
                                    <th>发货仓库</th>
                                    <th>物流渠道</th>
                                    <th>创建时间</th>
                                    <th>上网时间</th>
                                    <th>出库时间</th>
                                    <th>上网时长</th>
                                    <th>判定结果</th>
                                </tr>
                            </thead>
                            <tbody><tr><td colspan="10" style="color:#999; padding:30px 0;">暂无订单明细，请点击【拉取上网率数据】</td></tr></tbody>
                        </table>
                    </div>

                    <div id="sr-view-outbound" class="sr-view">
                        <div class="sr-controls" style="background:#fffbe6; padding:12px; border-radius:6px; border:1px solid #ffe58f; margin-bottom:15px;">
                            <span style="font-size:13px; font-weight:bold; color:#d46b08;">选择出库发货率周期:</span>
                            <input type="date" id="sr-outbound-start" class="sr-input-date" title="开始日期"> <span style="font-size:12px;color:#666;">至</span>
                            <input type="date" id="sr-outbound-end" class="sr-input-date" title="结束日期">
                            <label style="font-size:13px; cursor:pointer; color:#d4380d; font-weight:bold; margin-left:10px;">
                                <input type="checkbox" id="sr-outbound-skip" style="margin-right:4px;"> 排除周末
                            </label>
                            <button class="sr-btn sr-btn-warning" id="sr-outbound-btn" style="margin-left:auto;">抓取并展示</button>
                            <button class="sr-btn sr-btn-success" id="sr-outbound-export" disabled>导出 出库发货率 Excel</button>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>按所选周期逐仓库调用出库发货率接口，展示 24H、48H、72H 出库发货率。</div>
                            <div>勾选“排除周末”时，会按工作日逐日抓取并计算平均出库发货率。</div>
                            <div>图表与表格使用同一批抓取结果；导出 Excel 不会重复请求接口。</div>
                        </div>
                        <div class="sr-badges">
                            <div class="sr-badge sr-badge-total"><span class="title">仓库数</span><span class="num" id="ob-warehouse-count">-</span></div>
                            <div class="sr-badge sr-badge-24"><span class="title">24H 平均出库发货率</span><span class="num" id="ob-24">-</span></div>
                            <div class="sr-badge sr-badge-48"><span class="title">48H 平均出库发货率</span><span class="num" id="ob-48">-</span></div>
                            <div class="sr-badge sr-badge-72"><span class="title">72H 平均出库发货率</span><span class="num" id="ob-72">-</span></div>
                        </div>
                        <div class="chart-row" style="margin-top:10px;">
                            <div id="chart-outbound-line" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                            <div id="chart-outbound-bar" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                        </div>
                        <table class="sr-table" id="sr-outbound-table">
                            <thead>
                                <tr>
                                    <th>仓库名称</th>
                                    <th>仓库代码</th>
                                    <th>24H 出库发货率</th>
                                    <th>48H 出库发货率</th>
                                    <th>72H 出库发货率</th>
                                </tr>
                            </thead>
                            <tbody><tr><td colspan="5" style="color:#999; padding:30px 0;">请选择周期并点击【抓取并展示】</td></tr></tbody>
                        </table>
                    </div>

                    <!-- 周期分析模块 -->
                    <div id="sr-view-period" class="sr-view">
                        <div class="sr-controls" style="background:#f0f5ff; padding:12px; border-radius:6px; border:1px solid #adc6ff; margin-bottom:15px;">
                            <span style="font-size:13px; font-weight:bold; color:#096dd9;">选择分析周期:</span>
                            <input type="date" id="sr-period-start" class="sr-input-date"> <span style="font-size:12px;color:#666;">至</span>
                            <input type="date" id="sr-period-end" class="sr-input-date">
                            <label style="font-size:13px; cursor:pointer; color:#d4380d; font-weight:bold; margin-left:10px;">
                                <input type="checkbox" id="sr-period-skip" checked style="margin-right:4px;"> 排除周末
                            </label>
                            <button class="sr-btn sr-btn-purple" id="sr-period-btn" style="margin-left:auto;">🚀 开始计算周期平均</button>
                            <button class="sr-btn sr-btn-success" id="sr-period-export" disabled>导出 周期报表</button>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>按所选日期范围逐页抓取订单数据，再按物流渠道、发货仓库或客户名称聚合，统计整个周期内的平均表现。</div>
                            <div>勾选“排除周末”时，会直接跳过创建时间落在周六、周日的订单。</div>
                            <div>24H / 48H / 72H 平均达标率分别表示订单在创建后 24 / 48 / 72 小时内完成上网的比例；周期总单量不含已取消（status=99）及上网时间超过 72H 的订单。</div>
                        </div>

                        <div class="sr-badges">
                            <div class="sr-badge sr-badge-total"><span class="title">周期总发货单量</span><span class="num" id="pb-total">-</span></div>
                            <div class="sr-badge sr-badge-24"><span class="title">24H 综合平均达标率</span><span class="num" id="pb-24">-</span></div>
                            <div class="sr-badge sr-badge-48"><span class="title">48H 综合平均达标率</span><span class="num" id="pb-48">-</span></div>
                            <div class="sr-badge sr-badge-72"><span class="title">72H 综合平均达标率</span><span class="num" id="pb-72">-</span></div>
                        </div>

                        <select id="sr-period-dim" class="sr-select" style="margin-bottom:10px; width: 200px;">
                            <option value="channel">按物流渠道查看</option>
                            <option value="warehouse">按发货仓库查看</option>
                            <option value="customer">按客户名称查看</option>
                        </select>
                        <table class="sr-table" id="sr-period-table">
                            <thead>
                                <tr>
                                    <th style="width: 25%">维度名称</th>
                                    <th>周期总单量<br/><span style="font-size:11px; font-weight:normal;">(符合条件的订单)</span></th>
                                    <th>24H 平均达标率<br/><span style="font-size:11px; font-weight:normal;">(≤24h上网)</span></th>
                                    <th>48H 平均达标率<br/><span style="font-size:11px; font-weight:normal;">(≤48h上网)</span></th>
                                    <th>72H 平均达标率<br/><span style="font-size:11px; font-weight:normal;">(≤72h上网)</span></th>
                                </tr>
                            </thead>
                            <tbody><tr><td colspan="5" style="color:#999; padding:30px 0;">请设定时间范围并点击【开始计算周期平均】</td></tr></tbody>
                        </table>
                    </div>

                    <div id="sr-view-region" class="sr-view">
                        <div class="sr-controls" style="background:#fff7e6; padding:12px; border-radius:6px; border:1px solid #ffd591; margin-bottom:15px;">
                            <span style="font-size:13px; font-weight:bold; color:#d46b08;">选择地区统计周期:</span>
                            <input type="date" id="sr-region-start" class="sr-input-date"> <span style="font-size:12px;color:#666;">至</span>
                            <input type="date" id="sr-region-end" class="sr-input-date">
                            <label style="font-size:13px; cursor:pointer; color:#d4380d; font-weight:bold; margin-left:10px;">
                                <input type="checkbox" id="sr-region-skip" style="margin-right:4px;"> 排除周末
                            </label>
                            <button class="sr-btn sr-btn-warning" id="sr-region-start-btn" style="margin-left:auto;">开始统计并展示</button>
                            <button class="sr-btn sr-btn-success" id="sr-region-export" disabled>导出 订单分布分析</button>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>按客户 + 发货仓库汇总订单地区分布，地区依据地址信息接口返回的 postCode 邮编自动归类。</div>
                            <div>是否勾选“排除周末”会决定是否纳入创建时间落在周六、周日的订单；当前默认不排除周末。</div>
                            <div>导出文件包含：客户分仓地区分布、按客户地区汇总、邮编地区规则。</div>
                            <div>如果订单地址信息获取失败或邮编为空/格式异常，不纳入地区统计。</div>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">邮编匹配逻辑</div>
                            <div>系统会先提取邮编中的前 5 位数字，再按预设规则归类地区。</div>
                            <button class="sr-link-btn" id="sr-region-zip-rule-btn" type="button">查看邮编匹配说明</button>
                        </div>
                        <div class="sr-badges" style="margin-top:15px;">
                            <div class="sr-badge sr-badge-total"><span class="title">纳入统计单量</span><span class="num" id="rb-total">-</span></div>
                            <div class="sr-badge sr-badge-24"><span class="title">客户数</span><span class="num" id="rb-customers">-</span></div>
                            <div class="sr-badge sr-badge-48"><span class="title">客户分仓组合数</span><span class="num" id="rb-warehouses">-</span></div>
                            <div class="sr-badge sr-badge-72"><span class="title">未纳入统计单量</span><span class="num" id="rb-unknown">-</span></div>
                        </div>
                        <div class="sr-controls" style="margin-top:5px;">
                            <select id="sr-region-dim" class="sr-select" style="width:220px;">
                                <option value="warehouse">按客户分仓查看</option>
                                <option value="customer">按客户汇总查看</option>
                            </select>
                            <div id="sr-region-summary" style="font-size:13px; color:#666;">请先点击【开始统计并展示】。</div>
                        </div>
                        <div class="chart-row" style="margin-top:10px;">
                            <div id="chart-region-pie" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                            <div id="chart-region-bar" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                        </div>
                        <div id="sr-region-pies" class="sr-region-pies"></div>
                        <table class="sr-table" id="sr-region-table">
                            <thead></thead>
                            <tbody><tr><td style="color:#999; padding:30px 0;">暂无地区统计结果</td></tr></tbody>
                        </table>
                    </div>

                    <div id="sr-view-inventory" class="sr-view">
                        <div class="sr-controls" style="background:#f6ffed; padding:12px; border-radius:6px; border:1px solid #b7eb8f; margin-bottom:15px;">
                            <span style="font-size:13px; font-weight:bold; color:#389e0d;">选择库存分析周期:</span>
                            <input type="date" id="sr-inventory-start" class="sr-input-date"> <span style="font-size:12px;color:#666;">至</span>
                            <input type="date" id="sr-inventory-end" class="sr-input-date">
                            <label style="font-size:13px; cursor:pointer; color:#237804; font-weight:bold; margin-left:10px;">
                                <input type="checkbox" id="sr-inventory-skip-weekends" style="margin-right:4px;"> 排除周末
                            </label>
                            <button class="sr-btn sr-btn-success" id="sr-inventory-start-btn" style="margin-left:auto;">开始库存分析</button>
                            <button class="sr-btn sr-btn-success" id="sr-inventory-export" disabled>导出 库存分析</button>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>基于库存分析接口的客户分仓维度数据，按客户与发货仓库汇总期初库存、期末库存、出库预占、库存周转率、库存周转天数等指标。</div>
                            <div>勾选“排除周末”时，会剔除“statisticDate”落在周六、周日的记录；当前默认不排除周末。</div>
                            <div>“库存周转率”“库存周转天数”直接取接口返回分析结果；客户卡片图用于看每个客户下各分仓的库存结构和周转表现。</div>
                        </div>
                        <div class="sr-badges" style="margin-top:15px;">
                            <div class="sr-badge sr-badge-total"><span class="title">仓库数</span><span class="num" id="ib-warehouse-count">-</span></div>
                            <div class="sr-badge sr-badge-24"><span class="title">期末库存</span><span class="num" id="ib-close-stock">-</span></div>
                            <div class="sr-badge sr-badge-48"><span class="title">总周转率</span><span class="num" id="ib-turnover-rate">-</span></div>
                            <div class="sr-badge sr-badge-72"><span class="title">总周转天数</span><span class="num" id="ib-turnover-days">-</span></div>
                        </div>
                        <div class="sr-controls" style="margin-top:5px;">
                            <div id="sr-inventory-summary" class="sr-inventory-summary">请先点击【开始库存分析】。</div>
                        </div>
                        <div class="chart-row" style="margin-top:10px;">
                            <div id="chart-inventory-pie" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                            <div id="chart-inventory-bar" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                        </div>
                        <table class="sr-table" id="sr-inventory-table">
                            <thead>
                                <tr>
                                    <th>仓库名称</th>
                                    <th>客户数</th>
                                    <th>期初库存</th>
                                    <th>期末库存</th>
                                    <th>出库预占</th>
                                    <th>库存周转率</th>
                                    <th>库存周转天数</th>
                                    <th>库存销率</th>
                                </tr>
                            </thead>
                            <tbody><tr><td colspan="8" style="color:#999; padding:30px 0;">暂无库存分析结果</td></tr></tbody>
                        </table>
                    </div>

                    <div id="sr-view-sku-inventory" class="sr-view">
                        <div class="sr-controls" style="background:#f9f0ff; padding:12px; border-radius:6px; border:1px solid #d3adf7; margin-bottom:15px;">
                            <span style="font-size:13px; font-weight:bold; color:#722ed1;">选择SKU库存分析周期:</span>
                            <input type="date" id="sr-sku-inventory-start" class="sr-input-date"> <span style="font-size:12px;color:#666;">至</span>
                            <input type="date" id="sr-sku-inventory-end" class="sr-input-date">
                            <label style="font-size:13px; cursor:pointer; color:#531dab; font-weight:bold; margin-left:10px;">
                                <input type="checkbox" id="sr-sku-inventory-skip-weekends" style="margin-right:4px;"> 排除周末
                            </label>
                            <button class="sr-btn sr-btn-purple" id="sr-sku-inventory-start-btn" style="margin-left:auto;">开始SKU库存分析</button>
                            <button class="sr-btn sr-btn-success" id="sr-sku-inventory-export" disabled>导出 SKU库存分析</button>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>基于 SKU 聚合库存接口数据，支持按客户SKU明细、按客户汇总、按 SKU 汇总三种视角查看库存表现。</div>
                            <div>勾选“排除周末”时，会剔除“statisticDate”落在周六、周日的记录；当前默认不排除周末。</div>
                            <div>图表中的客户卡片会展示每个客户下 SKU 的库存周转率和周转天数，表格分页支持按页浏览与快捷跳转。</div>
                        </div>
                        <div class="sr-badges" style="margin-top:15px;">
                            <div class="sr-badge sr-badge-total"><span class="title">客户数</span><span class="num" id="sib-customer-count">-</span></div>
                            <div class="sr-badge sr-badge-24"><span class="title">SKU数</span><span class="num" id="sib-sku-count">-</span></div>
                            <div class="sr-badge sr-badge-48"><span class="title">明细组合数</span><span class="num" id="sib-detail-count">-</span></div>
                            <div class="sr-badge sr-badge-72"><span class="title">期末库存</span><span class="num" id="sib-close-stock">-</span></div>
                        </div>
                        <div class="sr-controls" style="margin-top:5px;">
                            <select id="sr-sku-inventory-dim" class="sr-select" style="width:220px;">
                                <option value="detail">按客户SKU明细查看</option>
                                <option value="customer">按客户汇总查看</option>
                                <option value="sku">按SKU汇总查看</option>
                            </select>
                            <div id="sr-sku-inventory-summary" class="sr-inventory-summary">请先点击【开始SKU库存分析】。</div>
                        </div>
                        <div class="chart-row" style="margin-top:10px;">
                            <div id="chart-sku-inventory-pie" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                            <div id="chart-sku-inventory-bar" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                        </div>
                        <div id="sr-sku-inventory-customer-charts" class="sr-inventory-cards"></div>
                        <table class="sr-table" id="sr-sku-inventory-table">
                            <thead></thead>
                            <tbody><tr><td colspan="10" style="color:#999; padding:30px 0;">暂无SKU库存分析结果</td></tr></tbody>
                        </table>
                        <div class="sr-table-pagination" id="sr-sku-inventory-pagination">
                            <span id="sr-sku-inventory-page-info">共 0 条</span>
                            <select id="sr-sku-inventory-page-size">
                                <option value="20">20/页</option>
                                <option value="50" selected>50/页</option>
                                <option value="100">100/页</option>
                            </select>
                            <span>跳至</span>
                            <input id="sr-sku-inventory-page-jump" class="sr-input-date" type="number" min="1" step="1" style="width:90px;" placeholder="页码">
                            <button class="sr-btn sr-btn-primary" id="sr-sku-inventory-page-go">跳转</button>
                            <button class="sr-btn sr-btn-primary" id="sr-sku-inventory-prev">上一页</button>
                            <button class="sr-btn sr-btn-primary" id="sr-sku-inventory-next">下一页</button>
                        </div>
                    </div>
                    <div id="sr-view-sku-sales" class="sr-view">
                        <div class="sr-controls" style="background:#e6fffb; padding:12px; border-radius:6px; border:1px solid #87e8de; margin-bottom:15px;">
                            <span style="font-size:13px; font-weight:bold; color:#08979c;">选择SKU销量分析周期:</span>
                            <input type="date" id="sr-sku-sales-start" class="sr-input-date"> <span style="font-size:12px;color:#666;">至</span>
                            <input type="date" id="sr-sku-sales-end" class="sr-input-date">
                            <label style="font-size:13px; cursor:pointer; color:#006d75; font-weight:bold; margin-left:10px;">
                                <input type="checkbox" id="sr-sku-sales-skip-weekends" style="margin-right:4px;"> 排除周末
                            </label>
                            <button class="sr-btn" id="sr-sku-sales-start-btn" style="margin-left:auto; background:#13c2c2;">开始SKU销量分析</button>
                            <button class="sr-btn sr-btn-success" id="sr-sku-sales-export" disabled>导出 SKU销量分析</button>
                        </div>
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>基于订单列表接口的“productList”明细统计销量，其中“productSku”作为 SKU，“qty”作为销量，按客户和 SKU 聚合。</div>
                            <div>勾选“排除周末”时，会剔除创建时间落在周六、周日的订单；当前默认不排除周末。</div>
                            <div>客户卡片按客户名称展示，每张卡片内部的“SKU销量前四”柱状图按销量从大到小排列；总览排行图同样按销量排序。</div>
                        </div>
                        <div class="sr-badges" style="margin-top:15px;">
                            <div class="sr-badge sr-badge-total"><span class="title">客户数</span><span class="num" id="ssb-customer-count">-</span></div>
                            <div class="sr-badge sr-badge-24"><span class="title">SKU数</span><span class="num" id="ssb-sku-count">-</span></div>
                            <div class="sr-badge sr-badge-48"><span class="title">客户SKU组合数</span><span class="num" id="ssb-detail-count">-</span></div>
                            <div class="sr-badge sr-badge-72"><span class="title">总销量</span><span class="num" id="ssb-total-qty">-</span></div>
                        </div>
                        <div class="sr-controls" style="margin-top:5px;">
                            <select id="sr-sku-sales-dim" class="sr-select" style="width:220px;">
                                <option value="detail">按客户SKU明细查看</option>
                                <option value="customer">按客户汇总查看</option>
                                <option value="sku">按SKU汇总查看</option>
                            </select>
                            <div id="sr-sku-sales-summary" class="sr-inventory-summary">请先点击【开始SKU销量分析】。</div>
                        </div>
                        <div class="chart-row" style="margin-top:10px;">
                            <div id="chart-sku-sales-customer" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                            <div id="chart-sku-sales-sku" class="echarts-box-half" style="border:1px solid #f0f0f0; border-radius:8px;"></div>
                        </div>
                        <div id="chart-sku-sales-trend" class="echarts-box" style="border:1px solid #f0f0f0; border-radius:8px; margin-bottom:20px;"></div>
                        <div id="sr-sku-sales-customer-charts" class="sr-inventory-cards"></div>
                        <table class="sr-table" id="sr-sku-sales-table">
                            <thead></thead>
                            <tbody><tr><td colspan="6" style="color:#999; padding:30px 0;">暂无SKU销量分析结果</td></tr></tbody>
                        </table>
                        <div class="sr-table-pagination" id="sr-sku-sales-pagination">
                            <span id="sr-sku-sales-page-info">共 0 条</span>
                            <select id="sr-sku-sales-page-size">
                                <option value="20">20/页</option>
                                <option value="50" selected>50/页</option>
                                <option value="100">100/页</option>
                            </select>
                            <span>跳至</span>
                            <input id="sr-sku-sales-page-jump" class="sr-input-date" type="number" min="1" step="1" style="width:90px;" placeholder="页码">
                            <button class="sr-btn sr-btn-primary" id="sr-sku-sales-page-go">跳转</button>
                            <button class="sr-btn sr-btn-primary" id="sr-sku-sales-prev">上一页</button>
                            <button class="sr-btn sr-btn-primary" id="sr-sku-sales-next">下一页</button>
                        </div>
                    </div>
                    <div id="sr-view-sales-report" class="sr-view">
                        <div class="sr-note-block">
                            <div class="sr-note-title">统计与分析说明</div>
                            <div>销售报告基于“订单分布分析、按客户库存分析、按sku库存分析、sku销量分析”这些模块的现有结果生成，不会重复请求接口。</div>
                            <div>单客户导出会输出一份包含图表图片、核心指标和文字分析说明的 HTML / PDF 报告；一键导出会按客户逐份生成并打包为 ZIP。</div>
                            <div>只要当前页面已完成至少一个相关分析模块，就可以生成报告；如果客户缺少某模块数据，报告会自动跳过该模块，仅导出已有数据。</div>
                        </div>
                        <div class="sr-report-grid">
                            <div class="sr-report-side">
                                <div class="sr-report-card">
                                    <div class="sr-report-card-title">导出设置</div>
                                    <div class="sr-report-setting-stack">
                                        <div class="sr-report-select-row">
                                            <div class="sr-report-select-label">选择客户</div>
                                            <select id="sr-sales-report-customer" class="sr-select">
                                                <option value="">请选择客户</option>
                                            </select>
                                        </div>
                                        <label class="sr-report-option-row" for="sr-sales-report-show-cn-sku">
                                            <input type="checkbox" id="sr-sales-report-show-cn-sku">
                                            <span class="sr-report-option-text">
                                                <span class="sr-report-option-title">启用中文SKU显示</span>
                                                <span class="sr-report-option-desc">在销售报告的SKU库存明细和SKU销量明细中显示产品名称列</span>
                                            </span>
                                        </label>
                                        <label class="sr-report-option-row" for="sr-sales-report-use-cn-name">
                                            <input type="checkbox" id="sr-sales-report-use-cn-name" disabled>
                                            <span class="sr-report-option-text">
                                                <span class="sr-report-option-title">启用客户中文名称</span>
                                                <span class="sr-report-option-desc">报告标题显示为“客户名称（公司名称）销售报告”</span>
                                            </span>
                                        </label>
                                        <label class="sr-report-option-row" for="sr-sales-report-dense-mode">
                                            <input type="checkbox" id="sr-sales-report-dense-mode">
                                            <span class="sr-report-option-text">
                                                <span class="sr-report-option-title">启用高密度版排版</span>
                                                <span class="sr-report-option-desc">勾选后，预览、HTML 导出、PDF 导出和批量导出统一使用高密度版模板</span>
                                            </span>
                                        </label>
                                        <div id="sr-sales-report-cn-name-status" class="sr-report-help">未获取客户中文名称，标题默认使用客户名称。</div>
                                        <div class="sr-report-actions">
                                            <button class="sr-btn sr-btn-slate sr-report-btn" id="sr-sales-report-refresh">刷新客户列表</button>
                                            <button class="sr-btn sr-btn-info sr-report-btn" id="sr-sales-report-fetch-cn-name">获取客户公司名称</button>
                                        </div>
                                        <div class="sr-report-actions">
                                            <button class="sr-btn sr-btn-amber sr-report-btn" id="sr-sales-report-edit-outbound">修改出库效率</button>
                                            <button class="sr-btn sr-btn-teal sr-report-btn" id="sr-sales-report-edit-advice">提出意见建议</button>
                                        </div>
                                        <div class="sr-report-actions">
                                            <button class="sr-btn sr-btn-warning sr-report-btn" id="sr-sales-report-preview-btn">生成HTML预览</button>
                                            <button class="sr-btn sr-btn-success sr-report-btn" id="sr-sales-report-export-btn" disabled>导出当前客户html</button>
                                        </div>
                                        <div class="sr-report-actions">
                                            <button class="sr-btn sr-btn-cyan sr-report-btn" id="sr-sales-report-export-pdf-btn" disabled>导出当前客户PDF</button>
                                            <button class="sr-btn sr-btn-indigo sr-report-btn" id="sr-sales-report-export-all-btn" disabled>一键导出全部html</button>
                                        </div>
                                        <div class="sr-report-actions">
                                            <button class="sr-btn sr-btn-rose sr-report-btn" id="sr-sales-report-export-all-pdf-btn" disabled>一键导出全部PDF</button>
                                        </div>
                                    </div>
                                </div>
                                <div class="sr-report-card">
                                    <div class="sr-report-card-title">数据就绪状态</div>
                                    <div id="sr-sales-report-ready" class="sr-report-help">请先完成四个分析模块，系统才能汇总客户报告。</div>
                                </div>
                            </div>
                            <div class="sr-report-preview">
                                <div id="sr-sales-report-preview-empty" class="sr-report-preview-empty">请先刷新客户列表并选择客户，然后点击“生成HTML预览”。</div>
                                <div id="sr-sales-report-preview-content" style="display:none;">
                                    <div class="sr-report-preview-shell">
                                        <div class="sr-report-preview-toolbar">
                                            <div class="sr-report-preview-toolbar-left">
                                                <div class="sr-report-preview-dots">
                                                    <span class="sr-report-preview-dot red"></span>
                                                    <span class="sr-report-preview-dot yellow"></span>
                                                    <span class="sr-report-preview-dot green"></span>
                                                </div>
                                                <div class="sr-report-preview-title" id="sr-sales-report-preview-title">销售报告HTML预览</div>
                                            </div>
                                            <div class="sr-report-preview-meta" id="sr-sales-report-preview-meta"></div>
                                        </div>
                                        <iframe id="sr-sales-report-preview-frame" class="sr-report-preview-frame" title="销售报告HTML预览"></iframe>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div id="sr-zip-rule-modal-mask" class="sr-modal-mask">
                <div class="sr-modal" role="dialog" aria-modal="true" aria-labelledby="sr-zip-rule-modal-title">
                    <div class="sr-modal-header">
                        <span id="sr-zip-rule-modal-title">邮编匹配说明</span>
                        <span class="sr-close" id="sr-zip-rule-modal-close">×</span>
                    </div>
                    <div class="sr-modal-body">
                        <div>系统会先提取订单邮编中的前 5 位数字，再按以下规则归类地区：</div>
                        <ul class="sr-note-list">${ZIP_REGION_RULES.map((rule) => `<li>${rule}</li>`).join('')}</ul>
                        <div style="margin-top:8px;">如果邮编为空、位数不足、格式异常，或者首位不是 0-9，则判定为“未知”并跳过统计。</div>
                    </div>
                </div>
            </div>
            <div id="sr-outbound-editor-modal-mask" class="sr-modal-mask">
                <div class="sr-modal" role="dialog" aria-modal="true" aria-labelledby="sr-outbound-editor-modal-title" style="width:min(980px, 100%);">
                    <div class="sr-modal-header">
                        <span id="sr-outbound-editor-modal-title">修改出库效率</span>
                        <span class="sr-close" id="sr-outbound-editor-close">×</span>
                    </div>
                    <div class="sr-modal-body">
                        <div class="sr-editor-help">这里编辑的是销售报告模块“5. 出库效率”使用的数据。保存后，报告预览、HTML 导出、PDF 导出和批量导出都会同步使用新数据。</div>
                        <div id="sr-outbound-editor-meta" class="sr-editor-meta">请先抓取出库效率数据，再进入编辑。</div>
                        <div class="sr-editor-toolbar">
                            <div id="sr-outbound-editor-columns" class="sr-editor-chip-row"></div>
                            <button class="sr-btn sr-btn-primary" id="sr-outbound-editor-add-row" type="button">新增一行</button>
                        </div>
                        <div class="sr-editor-table-wrap">
                            <table id="sr-outbound-editor-table" class="sr-editor-table">
                                <thead>
                                    <tr>
                                        <th>仓库名称</th>
                                        <th>仓库代码</th>
                                        <th>24H 出库率 (%)</th>
                                        <th>48H 出库率 (%)</th>
                                        <th>72H 出库率 (%)</th>
                                        <th>纳入天数</th>
                                    </tr>
                                </thead>
                                <tbody><tr><td colspan="6" style="color:#999; padding:24px 0; text-align:center;">请先抓取出库效率数据，再进入编辑。</td></tr></tbody>
                            </table>
                        </div>
                    </div>
                    <div class="sr-modal-footer">
                        <button class="sr-btn sr-btn-slate" id="sr-outbound-editor-reset" type="button">重置修改</button>
                        <button class="sr-btn sr-btn-warning" id="sr-outbound-editor-cancel" type="button">取消</button>
                        <button class="sr-btn sr-btn-success" id="sr-outbound-editor-save" type="button">保存</button>
                    </div>
                </div>
            </div>
            <div id="sr-sales-report-advice-modal-mask" class="sr-modal-mask">
                <div class="sr-modal" role="dialog" aria-modal="true" aria-labelledby="sr-sales-report-advice-modal-title" style="width:min(860px, 100%);">
                    <div class="sr-modal-header">
                        <span id="sr-sales-report-advice-modal-title">销售报告意见建议</span>
                        <span class="sr-close" id="sr-sales-report-advice-close">×</span>
                    </div>
                    <div class="sr-modal-body">
                        <div class="sr-editor-help">为每个模块单独填写意见建议。保存后，会在报告对应模块中追加展示这些建议内容。</div>
                        <div class="sr-advice-grid">
                            ${SALES_REPORT_ADVICE_MODULES.map((module) => `
                                <div class="sr-advice-card">
                                    <div class="sr-advice-card-title" data-advice-title="${module.key}">${module.label}</div>
                                    <textarea id="sr-sales-report-advice-${module.key}" class="sr-advice-textarea" placeholder="请输入该模块的意见建议，可为空"></textarea>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                    <div class="sr-modal-footer">
                        <button class="sr-btn sr-btn-warning" id="sr-sales-report-advice-cancel" type="button">取消</button>
                        <button class="sr-btn sr-btn-success" id="sr-sales-report-advice-save" type="button">保存建议</button>
                    </div>
                </div>
            </div>
        `;
        document.body.insertAdjacentHTML('beforeend', html);
        setupInventoryAnalysisUI();
        injectAnalysisDevControls();

        // 周期默认选中上周的周一至周五
        const lastWeek = getLastWeekMonFri();
        const defaultOnlineEnd = getTargetDate(1, true);
        const defaultOnlineStart = getTargetDate(3, true);
        document.getElementById('sr-outbound-start').value = lastWeek.start;
        document.getElementById('sr-outbound-end').value = lastWeek.end;
        document.getElementById('sr-period-start').value = lastWeek.start;
        document.getElementById('sr-period-end').value = lastWeek.end;
        document.getElementById('sr-online-start').value = defaultOnlineStart.dateStr;
        document.getElementById('sr-online-end').value = defaultOnlineEnd.dateStr;
        document.getElementById('sr-region-start').value = lastWeek.start;
        document.getElementById('sr-region-end').value = lastWeek.end;
        document.getElementById('sr-inventory-start').value = lastWeek.start;
        document.getElementById('sr-inventory-end').value = lastWeek.end;
        document.getElementById('sr-sku-inventory-start').value = lastWeek.start;
        document.getElementById('sr-sku-inventory-end').value = lastWeek.end;
        document.getElementById('sr-sku-sales-start').value = lastWeek.start;
        document.getElementById('sr-sku-sales-end').value = lastWeek.end;

        // 悬浮按钮拖拽逻辑
        const floatBtn = document.getElementById('sr-floating-btn');
        const panel = document.getElementById('sr-panel');
        const panelHeader = panel.querySelector('.sr-header');
        const fullscreenBtn = document.getElementById('sr-fullscreen-btn');
        const zipRuleModalMask = document.getElementById('sr-zip-rule-modal-mask');
        const outboundEditorModalMask = document.getElementById('sr-outbound-editor-modal-mask');
        const salesReportAdviceModalMask = document.getElementById('sr-sales-report-advice-modal-mask');
        let panelPositionState = null;
        let isDragging = false, startX, startY, initialLeft, initialTop;
        floatBtn.addEventListener('mousedown', (e) => {
            isDragging = false; startX = e.clientX; startY = e.clientY;
            const rect = floatBtn.getBoundingClientRect(); initialLeft = rect.left; initialTop = rect.top; floatBtn.style.transition = 'none';
            const onMouseMove = (moveEvent) => {
                if (Math.abs(moveEvent.clientX - startX) > 3 || Math.abs(moveEvent.clientY - startY) > 3) isDragging = true;
                if (isDragging) {
                    floatBtn.style.right = 'auto'; floatBtn.style.bottom = 'auto';
                    floatBtn.style.left = `${Math.max(0, Math.min(initialLeft + (moveEvent.clientX - startX), window.innerWidth - rect.width))}px`;
                    floatBtn.style.top = `${Math.max(0, Math.min(initialTop + (moveEvent.clientY - startY), window.innerHeight - rect.height))}px`;
                }
            };
            const onMouseUp = () => { document.removeEventListener('mousemove', onMouseMove); document.removeEventListener('mouseup', onMouseUp); floatBtn.style.transition = 'background 0.3s, transform 0.3s'; };
            document.addEventListener('mousemove', onMouseMove); document.addEventListener('mouseup', onMouseUp);
        });
        floatBtn.addEventListener('click', (e) => {
            if (isDragging) { e.preventDefault(); e.stopPropagation(); return; }
            panel.style.display = panel.style.display === 'flex' ? 'none' : 'flex';
            if (panel.style.display === 'flex') {
                if (!panel.classList.contains('sr-fullscreen')) {
                    fullscreenBtn.onclick();
                    return;
                }
                Object.values(charts).forEach(instance => {
                    if (Array.isArray(instance)) {
                        instance.forEach((chart) => { if (chart) chart.resize(); });
                    } else if (instance) {
                        instance.resize();
                    }
                });
            }
        });
        fullscreenBtn.onclick = () => {
            const enteringFullscreen = !panel.classList.contains('sr-fullscreen');
            if (enteringFullscreen) {
                panelPositionState = {
                    left: panel.style.left,
                    top: panel.style.top,
                    right: panel.style.right,
                    bottom: panel.style.bottom,
                    width: panel.style.width,
                    height: panel.style.height
                };
                panel.classList.add('sr-fullscreen');
                panel.style.left = '0';
                panel.style.top = '0';
                panel.style.right = '0';
                panel.style.bottom = '0';
                panel.style.width = '100vw';
                panel.style.height = '100vh';
            } else {
                panel.classList.remove('sr-fullscreen');
                panel.style.left = panelPositionState?.left || '';
                panel.style.top = panelPositionState?.top || '';
                panel.style.right = panelPositionState?.right || '';
                panel.style.bottom = panelPositionState?.bottom || '';
                panel.style.width = panelPositionState?.width || '';
                panel.style.height = panelPositionState?.height || '';
            }
            fullscreenBtn.innerText = panel.classList.contains('sr-fullscreen') ? '退出全屏' : '全屏';
            Object.values(charts).forEach(instance => {
                if (Array.isArray(instance)) {
                    instance.forEach((chart) => { if (chart) chart.resize(); });
                } else if (instance) {
                    instance.resize();
                }
            });
        };
        panelHeader.addEventListener('mousedown', (e) => {
            if (panel.classList.contains('sr-fullscreen')) return;
            if (e.target.closest('.sr-header-actions')) return;

            const rect = panel.getBoundingClientRect();
            const originX = e.clientX;
            const originY = e.clientY;
            const startLeft = rect.left;
            const startTop = rect.top;

            panel.style.left = `${startLeft}px`;
            panel.style.top = `${startTop}px`;
            panel.style.right = 'auto';
            panel.style.bottom = 'auto';

            const onMouseMove = (moveEvent) => {
                const nextLeft = Math.max(0, Math.min(startLeft + (moveEvent.clientX - originX), window.innerWidth - rect.width));
                const nextTop = Math.max(0, Math.min(startTop + (moveEvent.clientY - originY), window.innerHeight - 60));
                panel.style.left = `${nextLeft}px`;
                panel.style.top = `${nextTop}px`;
            };

            const onMouseUp = () => {
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            };

            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        });
        document.getElementById('sr-close-btn').onclick = () => panel.style.display = 'none';
        document.getElementById('sr-region-zip-rule-btn').onclick = () => zipRuleModalMask.classList.add('show');
        document.getElementById('sr-zip-rule-modal-close').onclick = () => zipRuleModalMask.classList.remove('show');
        zipRuleModalMask.addEventListener('click', (event) => {
            if (event.target === zipRuleModalMask) {
                zipRuleModalMask.classList.remove('show');
            }
        });
        document.getElementById('sr-outbound-editor-close').onclick = closeOutboundEditorModal;
        document.getElementById('sr-outbound-editor-cancel').onclick = closeOutboundEditorModal;
        document.getElementById('sr-outbound-editor-reset').onclick = resetOutboundEditorDraft;
        document.getElementById('sr-outbound-editor-save').onclick = saveOutboundEditorChanges;
        document.getElementById('sr-outbound-editor-add-row').onclick = addOutboundEditorRow;
        outboundEditorModalMask.addEventListener('click', (event) => {
            if (event.target === outboundEditorModalMask) {
                closeOutboundEditorModal();
            }
        });
        document.getElementById('sr-sales-report-advice-close').onclick = closeSalesReportAdviceModal;
        document.getElementById('sr-sales-report-advice-cancel').onclick = closeSalesReportAdviceModal;
        document.getElementById('sr-sales-report-advice-save').onclick = saveSalesReportAdvice;
        salesReportAdviceModalMask.addEventListener('click', (event) => {
            if (event.target === salesReportAdviceModalMask) {
                closeSalesReportAdviceModal();
            }
        });

        // 导航切换
        document.querySelectorAll('.sr-nav-tab').forEach(tab => {
            tab.onclick = function() {
                document.querySelectorAll('.sr-nav-tab').forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.sr-view').forEach(v => v.classList.remove('active'));
                this.classList.add('active'); document.getElementById(this.dataset.target).classList.add('active');

                if (this.dataset.target === 'sr-view-table' || this.dataset.target === 'sr-view-outbound' || this.dataset.target === 'sr-view-region' || this.dataset.target === 'sr-view-inventory' || this.dataset.target === 'sr-view-sku-inventory' || this.dataset.target === 'sr-view-sku-sales' || this.dataset.target === 'sr-view-sales-report') {
                    Object.values(charts).forEach(instance => {
                        if (Array.isArray(instance)) {
                            instance.forEach((chart) => { if (chart) chart.resize(); });
                        } else if (instance) {
                            instance.resize();
                        }
                    });
                }
            };
        });

        // 绑定事件
        document.getElementById('sr-start-btn').onclick = startOnlineRateProcess;
        document.getElementById('sr-export-btn').onclick = exportOnlineExcel;
        document.getElementById('sr-outbound-btn').onclick = startOutboundRateProcess;
        document.getElementById('sr-outbound-export').onclick = exportOutboundExcel;
        document.getElementById('sr-dimension-select').addEventListener('change', () => {
            syncOnlineRateFilterState();
            if (Object.keys(finalReportData).length > 0) { renderTable(); renderCharts(); renderOnlineRateDetailTable(); }
        });
        document.getElementById('sr-warehouse-filter').addEventListener('change', () => {
            if (Object.keys(finalReportData).length > 0) { renderTable(); renderCharts(); renderOnlineRateDetailTable(); }
        });
        document.getElementById('sr-online-detail-warehouse').addEventListener('change', renderOnlineRateDetailTable);
        document.getElementById('sr-online-detail-status').addEventListener('change', renderOnlineRateDetailTable);

        // 周期绑定事件
        document.getElementById('sr-period-btn').onclick = startPeriodProcess;
        document.getElementById('sr-period-export').onclick = exportPeriodExcel;
        document.getElementById('sr-region-start-btn').onclick = startRegionDistributionProcess;
        document.getElementById('sr-region-export').onclick = exportRegionDistributionExcel;
        document.getElementById('sr-inventory-start-btn').onclick = startInventoryAnalysisProcess;
        document.getElementById('sr-inventory-export').onclick = exportInventoryAnalysisExcel;
        document.getElementById('sr-sku-inventory-start-btn').onclick = startSkuInventoryAnalysisProcess;
        document.getElementById('sr-sku-inventory-export').onclick = exportSkuInventoryAnalysisExcel;
        document.getElementById('sr-sku-sales-start-btn').onclick = startSkuSalesAnalysisProcess;
        document.getElementById('sr-sku-sales-export').onclick = exportSkuSalesAnalysisExcel;
        document.getElementById('sr-sales-report-refresh').onclick = refreshSalesReportCustomers;
        document.getElementById('sr-sales-report-edit-outbound').onclick = openOutboundEditorModal;
        document.getElementById('sr-sales-report-edit-advice').onclick = openSalesReportAdviceModal;
        document.getElementById('sr-sales-report-preview-btn').onclick = previewSalesReport;
        document.getElementById('sr-sales-report-export-btn').onclick = exportCurrentSalesReport;
        document.getElementById('sr-sales-report-export-pdf-btn').onclick = exportCurrentSalesReportPdf;
        document.getElementById('sr-sales-report-export-all-btn').onclick = exportAllSalesReports;
        document.getElementById('sr-sales-report-export-all-pdf-btn').onclick = exportAllSalesReportsPdf;
        document.getElementById('sr-sales-report-fetch-cn-name').onclick = fetchSalesReportCustomerNames;
        document.getElementById('sr-sales-report-show-cn-sku').addEventListener('change', refreshSalesReportPreviewByOptions);
        document.getElementById('sr-sales-report-use-cn-name').addEventListener('change', refreshSalesReportPreviewByOptions);
        document.getElementById('sr-sales-report-dense-mode').addEventListener('change', refreshSalesReportPreviewByOptions);
        updateSalesReportCustomerNameState();
        syncOnlineRateFilterState();
        if (ENABLE_ANALYSIS_DEV_MODE) {
            bindAnalysisDevControls('outbound');
            bindAnalysisDevControls('region');
            bindAnalysisDevControls('inventory');
            bindAnalysisDevControls('skuInventory');
            bindAnalysisDevControls('skuSales');
        }
        document.getElementById('sr-sku-inventory-prev').onclick = () => {
            skuInventoryTableState.page = Math.max(1, skuInventoryTableState.page - 1);
            renderSkuInventoryAnalysisTable();
        };
        document.getElementById('sr-sku-inventory-next').onclick = () => {
            skuInventoryTableState.page += 1;
            renderSkuInventoryAnalysisTable();
        };
        document.getElementById('sr-sku-inventory-page-size').addEventListener('change', (e) => {
            skuInventoryTableState.pageSize = Number(e.target.value || 50);
            skuInventoryTableState.page = 1;
            if (skuInventoryReportData) renderSkuInventoryAnalysisTable();
        });
        document.getElementById('sr-sku-sales-prev').onclick = () => {
            skuSalesTableState.page = Math.max(1, skuSalesTableState.page - 1);
            renderSkuSalesAnalysisTable();
        };
        document.getElementById('sr-sku-sales-next').onclick = () => {
            skuSalesTableState.page += 1;
            renderSkuSalesAnalysisTable();
        };
        document.getElementById('sr-sku-sales-page-size').addEventListener('change', (e) => {
            skuSalesTableState.pageSize = Number(e.target.value || 50);
            skuSalesTableState.page = 1;
            if (skuSalesReportData) renderSkuSalesAnalysisTable();
        });
        bindTablePageJump('sr-sku-inventory-page-jump', 'sr-sku-inventory-page-go', skuInventoryTableState, renderSkuInventoryAnalysisTable);
        bindTablePageJump('sr-sku-sales-page-jump', 'sr-sku-sales-page-go', skuSalesTableState, renderSkuSalesAnalysisTable);
        document.getElementById('sr-period-dim').addEventListener('change', () => { if (periodReportData.summary.total > 0) renderPeriodTable(); });
        document.getElementById('sr-region-dim').addEventListener('change', () => { if (regionReportData) renderRegionDistributionView(); });
        document.getElementById('sr-inventory-dim').addEventListener('change', () => { if (inventoryReportData) renderInventoryAnalysisView(); });
        document.getElementById('sr-sku-inventory-dim').addEventListener('change', () => {
            skuInventoryTableState.page = 1;
            if (skuInventoryReportData) renderSkuInventoryAnalysisView();
        });
        document.getElementById('sr-sku-sales-dim').addEventListener('change', () => {
            skuSalesTableState.page = 1;
            if (skuSalesReportData) renderSkuSalesAnalysisView();
        });
        updateSalesReportReadyState();
    }

    // ==========================================
    // 6. 核心数据拉取及渲染
    // ==========================================
    function formatRate(total, online) { return total === 0 ? "-" : `${((online / total) * 100).toFixed(2)}% \n(${online}/${total})`; }

    function createOnlineRateStats() {
        return { '24h': { total: 0, online: 0 }, '48h': { total: 0, online: 0 }, '72h': { total: 0, online: 0 } };
    }

    function ensureOnlineRateDimension(obj, key) {
        if (!obj[key]) obj[key] = createOnlineRateStats();
        return obj[key];
    }

    function getOnlineRateSelectedWarehouse() {
        return document.getElementById('sr-warehouse-filter')?.value || '';
    }

    function getOnlineRateDimensionMeta() {
        const dimension = document.getElementById('sr-dimension-select').value;
        const selectedWarehouse = getOnlineRateSelectedWarehouse();

        if (dimension === 'warehouse') {
            return { dimension, selectedWarehouse: '', dataObj: warehouseReportData, dimName: '发货仓库', oppDataObj: finalReportData, oppDimName: '物流渠道' };
        }
        if (dimension === 'customer') {
            return { dimension, selectedWarehouse: '', dataObj: customerReportData, dimName: '客户名称', oppDataObj: finalReportData, oppDimName: '物流渠道' };
        }
        if (selectedWarehouse && warehouseChannelReportData[selectedWarehouse]) {
            return {
                dimension,
                selectedWarehouse,
                dataObj: warehouseChannelReportData[selectedWarehouse],
                dimName: `物流渠道（${selectedWarehouse}）`,
                oppDataObj: warehouseReportData,
                oppDimName: '发货仓库'
            };
        }
        return { dimension, selectedWarehouse: '', dataObj: finalReportData, dimName: '物流渠道', oppDataObj: warehouseReportData, oppDimName: '发货仓库' };
    }

    function updateOnlineWarehouseFilterOptions() {
        const select = document.getElementById('sr-warehouse-filter');
        if (!select) return;
        const previousValue = select.value || '';
        const warehouseNames = Object.keys(warehouseChannelReportData).sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
        select.innerHTML = '<option value="">全部仓库</option>';
        warehouseNames.forEach((warehouseName) => {
            const option = document.createElement('option');
            option.value = warehouseName;
            option.textContent = warehouseName;
            select.appendChild(option);
        });
        if (previousValue && warehouseNames.includes(previousValue)) {
            select.value = previousValue;
        }
    }

    function syncOnlineRateFilterState() {
        const wrap = document.getElementById('sr-warehouse-filter-wrap');
        const select = document.getElementById('sr-warehouse-filter');
        if (!wrap || !select) return;
        const dimension = document.getElementById('sr-dimension-select')?.value || 'channel';
        const isChannelView = dimension === 'channel';
        wrap.style.display = isChannelView ? 'flex' : 'none';
        select.disabled = !isChannelView || Object.keys(warehouseChannelReportData).length === 0;
        if (!isChannelView) select.value = '';
    }

    function getOnlineRateJudgeBucket(record, rate24, rate48, rate72) {
        if (isCancelledOrder(record)) return 'cancelled';
        if (rate24) return '24h';
        if (rate48) return '48h';
        if (rate72) return '72h';
        if (isOnlineOver72Hours(record)) return 'over72h';
        return 'offline';
    }

    function buildOnlineRateJudgeLabel(judgeBucket) {
        const labelMap = {
            all: '全部订单',
            offline: '未上网',
            '24h': '24H上网',
            '48h': '48H上网',
            '72h': '72H上网',
            over72h: '已上网（超过72H）',
            cancelled: '已取消'
        };
        return labelMap[judgeBucket] || '未上网';
    }

    function getOnlineRateDetailSelectedWarehouse() {
        return document.getElementById('sr-online-detail-warehouse')?.value || '';
    }

    function getOnlineRateDetailSelectedStatus() {
        return document.getElementById('sr-online-detail-status')?.value || 'all';
    }

    function buildOnlineRateDetailRow(record, rate24, rate48, rate72) {
        const judgeBucket = getOnlineRateJudgeBucket(record, rate24, rate48, rate72);
        const receiptTime = record?.receiptTime || '';
        const outboundTime = record?.outboundTime || '';
        const onlineTime = receiptTime || outboundTime || '';
        const detailRow = {
            orderNo: record?.deliveryNo || '',
            sourceNo: record?.sourceNo || record?.referOrderNo || record?.platformOrderNo || '',
            customerName: record?.customerName || '未知客户',
            whName: record?.whCodeName || '未知仓',
            channelName: record?.logisticsChannel || '未知渠道',
            createTime: record?.createTime || '',
            receiptTime,
            outboundTime,
            onlineDuration: formatOnlineDuration(getHoursDiff(record?.createTime, onlineTime)),
            judgeBucket,
            judgeLabel: buildOnlineRateJudgeLabel(judgeBucket),
            rate24,
            rate48,
            rate72
        };
        return detailRow;
    }

    function sortOnlineRateDetailRows(rows = []) {
        return rows.slice().sort((left, right) =>
            String(left.createTime || '').localeCompare(String(right.createTime || ''))
            || String(left.whName || '').localeCompare(String(right.whName || ''), 'zh-Hans-CN')
            || String(left.channelName || '').localeCompare(String(right.channelName || ''), 'zh-Hans-CN')
            || String(left.orderNo || '').localeCompare(String(right.orderNo || ''))
        );
    }

    function getOnlineRateFilteredDetailRows() {
        const selectedWarehouse = getOnlineRateDetailSelectedWarehouse();
        const selectedStatus = getOnlineRateDetailSelectedStatus();
        return onlineRateDetailRows.filter((row) => {
            if (selectedWarehouse && row.whName !== selectedWarehouse) return false;
            if (selectedStatus !== 'all' && row.judgeBucket !== selectedStatus) return false;
            return true;
        });
    }

    function updateOnlineDetailWarehouseOptions() {
        const select = document.getElementById('sr-online-detail-warehouse');
        if (!select) return;
        const previousValue = select.value || '';
        const warehouseNames = Array.from(new Set(onlineRateDetailRows.map((row) => row.whName).filter(Boolean)))
            .sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));

        select.innerHTML = '<option value="">全部仓库</option>';
        warehouseNames.forEach((warehouseName) => {
            const option = document.createElement('option');
            option.value = warehouseName;
            option.textContent = warehouseName;
            select.appendChild(option);
        });

        if (previousValue && warehouseNames.includes(previousValue)) {
            select.value = previousValue;
        }
    }

    function updateOnlineDetailStatusOptions() {
        const select = document.getElementById('sr-online-detail-status');
        if (!select) return;

        const selectedWarehouse = getOnlineRateDetailSelectedWarehouse();
        const scopeRows = selectedWarehouse
            ? onlineRateDetailRows.filter((row) => row.whName === selectedWarehouse)
            : onlineRateDetailRows;
        const counts = {
            all: scopeRows.length,
            offline: scopeRows.filter((row) => row.judgeBucket === 'offline').length,
            '24h': scopeRows.filter((row) => row.judgeBucket === '24h').length,
            '48h': scopeRows.filter((row) => row.judgeBucket === '48h').length,
            '72h': scopeRows.filter((row) => row.judgeBucket === '72h').length,
            over72h: scopeRows.filter((row) => row.judgeBucket === 'over72h').length,
            cancelled: scopeRows.filter((row) => row.judgeBucket === 'cancelled').length
        };

        Array.from(select.options).forEach((option) => {
            option.text = `${buildOnlineRateJudgeLabel(option.value)}（${counts[option.value] || 0}）`;
        });
    }

    function renderOnlineRateDetailTable() {
        const tbody = document.querySelector('#sr-online-detail-table tbody');
        const summary = document.getElementById('sr-online-detail-summary');
        if (!tbody || !summary) return;

        updateOnlineDetailStatusOptions();
        const rows = getOnlineRateFilteredDetailRows();
        const selectedWarehouse = getOnlineRateDetailSelectedWarehouse();
        const selectedStatus = getOnlineRateDetailSelectedStatus();
        const statusLabel = buildOnlineRateJudgeLabel(selectedStatus);

        summary.innerText = rows.length === 0
            ? `当前筛选下暂无订单明细${selectedWarehouse ? `（当前仓库：${selectedWarehouse}）` : ''}${selectedStatus !== 'all' ? `（当前结果：${statusLabel}）` : ''}`
            : `订单明细共 ${rows.length} 单${selectedWarehouse ? `（当前仓库：${selectedWarehouse}）` : ''}${selectedStatus !== 'all' ? `（当前结果：${statusLabel}）` : ''}`;

        tbody.innerHTML = '';
        if (rows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="10" style="color:#999; padding:30px 0;">当前条件下暂无订单明细</td></tr>';
            return;
        }

        rows.forEach((row) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${row.orderNo || '-'}</td>
                <td>${row.sourceNo || '-'}</td>
                <td>${row.customerName || '-'}</td>
                <td>${row.whName || '-'}</td>
                <td>${row.channelName || '-'}</td>
                <td>${row.createTime || '-'}</td>
                <td>${row.receiptTime || '-'}</td>
                <td>${row.outboundTime || '-'}</td>
                <td>${row.onlineDuration || '-'}</td>
                <td>${row.judgeLabel || '-'}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function renderTable() {
        const { dataObj = {}, dimName } = getOnlineRateDimensionMeta();

        document.getElementById('th-dim-name').innerText = dimName;
        document.getElementById('th-24h').innerHTML = `24H上网率 <br/><span style="font-size:11px;color:#1890ff">创建样本 ${dateLabels['24h']}</span>`;
        document.getElementById('th-48h').innerHTML = `48H上网率 <br/><span style="font-size:11px;color:#1890ff">创建样本 ${dateLabels['48h']}</span>`;
        document.getElementById('th-72h').innerHTML = `72H上网率 <br/><span style="font-size:11px;color:#1890ff">创建样本 ${dateLabels['72h']}</span>`;

        const tbody = document.querySelector('#sr-result-table tbody'); tbody.innerHTML = '';
        const keys = Object.keys(dataObj);
        if (keys.length === 0) { tbody.innerHTML = '<tr><td colspan="4">无符合条件的数据</td></tr>'; return; }
        keys.forEach(key => {
            const row = dataObj[key];
            const tr = document.createElement('tr');
            tr.innerHTML = `<td style="text-align:left; font-weight:bold;">${key}</td><td>${formatRate(row['24h'].total, row['24h'].online)}</td><td>${formatRate(row['48h'].total, row['48h'].online)}</td><td>${formatRate(row['72h'].total, row['72h'].online)}</td>`;
            tbody.appendChild(tr);
        });
    }

    function renderCharts() {
        if (!window.echarts || Object.keys(finalReportData).length === 0) return;
        const { dataObj: mainDataObj = {}, oppDataObj = finalReportData, dimName, oppDimName } = getOnlineRateDimensionMeta();

        const mainKeys = Object.keys(mainDataObj);
        if (mainKeys.length === 0) return;
        const getRate = (tot, onl) => tot === 0 ? 0 : Number(((onl / tot) * 100).toFixed(2));

        if (!charts.line) charts.line = echarts.init(document.getElementById('chart-line'));
        charts.line.setOption({
            title: { text: '固定样本 24H / 48H / 72H 上网率走势', left: 'center' }, tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
            legend: { data: ['样本总单量', '时效内上网量', '上网率'], top: 30 }, grid: { left: '3%', right: '3%', bottom: '5%', containLabel: true },
            xAxis: { type: 'category', data: ['24H', '48H', '72H'] }, yAxis: [ { type: 'value', name: '单量' }, { type: 'value', name: '上网率(%)', min: 0, max: 100 } ],
            series: [
                { name: '样本总单量', type: 'bar', data: [extendedData.trend['24h'].t, extendedData.trend['48h'].t, extendedData.trend['72h'].t], itemStyle:{color:'#5470c6'} },
                { name: '时效内上网量', type: 'bar', data: [extendedData.trend['24h'].o, extendedData.trend['48h'].o, extendedData.trend['72h'].o], itemStyle:{color:'#91cc75'} },
                { name: '上网率', type: 'line', yAxisIndex: 1, data: [ getRate(extendedData.trend['24h'].t, extendedData.trend['24h'].o), getRate(extendedData.trend['48h'].t, extendedData.trend['48h'].o), getRate(extendedData.trend['72h'].t, extendedData.trend['72h'].o) ], itemStyle:{color:'#fac858'}, lineStyle:{width:3}, symbolSize:8 }
            ]
        }, true);

        const barData = { c: [], r24: [], r48: [], r72: [] };
        mainKeys.forEach(k => { const r = mainDataObj[k]; barData.c.push(k); barData.r24.push(getRate(r['24h'].total, r['24h'].online)); barData.r48.push(getRate(r['48h'].total, r['48h'].online)); barData.r72.push(getRate(r['72h'].total, r['72h'].online)); });
        if (!charts.bar) charts.bar = echarts.init(document.getElementById('chart-bar'));
        charts.bar.setOption({
            title: { text: `各【${dimName}】24H/48H/72H上网率对比`, left: 'center' }, tooltip: { trigger: 'axis' }, legend: { data: ['24H上网率', '48H上网率', '72H上网率'], top: 30 }, grid: { left: '3%', right: '4%', bottom: '15%', containLabel: true },
            dataZoom: [ { type: 'slider', show: true, xAxisIndex: [0], start: 0, end: mainKeys.length > 6 ? Math.floor(600/mainKeys.length) : 100 } ], xAxis: { type: 'category', data: barData.c, axisLabel: { interval: 0, rotate: 15 } }, yAxis: { type: 'value', max: 100 },
            series: [ { name: '24H上网率', type: 'bar', data: barData.r24 }, { name: '48H上网率', type: 'bar', data: barData.r48 }, { name: '72H上网率', type: 'bar', data: barData.r72 } ]
        }, true);

        const pieData = mainKeys.map(k => ({ name: k, value: mainDataObj[k]['72h'].total || 0 })).filter(i => i.value > 0).sort((a,b) => b.value - a.value);
        if (!charts.pie) charts.pie = echarts.init(document.getElementById('chart-pie'));
        charts.pie.setOption({ title: { text: `【${dimName}】固定样本单量占比`, left: 'center' }, tooltip: { trigger: 'item' }, series: [{ type: 'pie', radius: ['35%', '65%'], center: ['50%', '55%'], data: pieData, label: { formatter: '{b}\n{c}单' } }] }, true);

        const roseData = Object.keys(oppDataObj).map(k => ({ name: k, value: oppDataObj[k]['72h'].total || 0 })).filter(i => i.value > 0).sort((a,b) => b.value - a.value);
        if (!charts.rose) charts.rose = echarts.init(document.getElementById('chart-rose'));
        charts.rose.setOption({ title: { text: `补充视角:【${oppDimName}】固定样本分布`, left: 'center' }, tooltip: { trigger: 'item' }, series: [{ type: 'pie', radius: [20, 100], center: ['50%', '55%'], roseType: 'area', data: roseData }] }, true);

        if (!charts.radar) charts.radar = echarts.init(document.getElementById('chart-radar'));
        const topItems = pieData.slice(0, 5);
        charts.radar.setOption({
            title: { text: `Top 5 【${dimName}】综合表现`, left: 'center' }, tooltip: { trigger: 'item' }, legend: { top: 30, data: topItems.map(c => c.name) },
            radar: { indicator: [ { name: '72H单量(票)', max: topItems.length ? topItems[0].value * 1.1 : 100 }, { name: '24H上网率(%)', max: 100 }, { name: '48H上网率(%)', max: 100 }, { name: '72H上网率(%)', max: 100 } ], center: ['50%', '55%'], radius: 100 },
            series: [{ type: 'radar', data: topItems.map(tc => { const r = mainDataObj[tc.name]; return { name: tc.name, value: [ tc.value, getRate(r['24h'].total, r['24h'].online), getRate(r['48h'].total, r['48h'].online), getRate(r['72h'].total, r['72h'].online) ] }; }), areaStyle: { opacity: 0.1 } }]
        }, true);
    }

    function appendOnlineRateStats(target, rate24, rate48, rate72) {
        ['24h', '48h', '72h'].forEach((key) => { target[key].total += 1; });
        if (rate24) target['24h'].online += 1;
        if (rate48) target['48h'].online += 1;
        if (rate72) target['72h'].online += 1;
    }

    async function fetchOnlineRateCohort(startStr, endStr, skipWeekends) {
        await forEachPeriodOrder(
            startStr,
            endStr,
            skipWeekends,
            (r) => {
                const channel = r.logisticsChannel || '未知渠道';
                const whName = r.whCodeName || '未知仓';
                const customer = r.customerName || '未知客户';
                const rate24 = isOnlineInSampleWindow(r, 24);
                const rate48 = isOnlineInSampleWindow(r, 48);
                const rate72 = isOnlineInSampleWindow(r, 72);
                const detailRow = buildOnlineRateDetailRow(r, rate24, rate48, rate72);

                onlineRateDetailRows.push(detailRow);
                onlineRateDetailBuckets[detailRow.judgeBucket].push(detailRow);

                if (shouldIncludeInOnlineRateStats(r)) {
                    appendOnlineRateStats(ensureOnlineRateDimension(finalReportData, channel), rate24, rate48, rate72);
                    appendOnlineRateStats(ensureOnlineRateDimension(warehouseReportData, whName), rate24, rate48, rate72);
                    appendOnlineRateStats(ensureOnlineRateDimension(customerReportData, customer), rate24, rate48, rate72);

                    if (!warehouseChannelReportData[whName]) warehouseChannelReportData[whName] = {};
                    appendOnlineRateStats(ensureOnlineRateDimension(warehouseChannelReportData[whName], channel), rate24, rate48, rate72);

                    ['24h', '48h', '72h'].forEach((key) => { extendedData.trend[key].t += 1; });
                    if (rate24) extendedData.trend['24h'].o += 1;
                    if (rate48) extendedData.trend['48h'].o += 1;
                    if (rate72) extendedData.trend['72h'].o += 1;
                }
            },
            (current, pages) => {
                document.getElementById('sr-status').innerText = `正在获取固定样本上网率数据... (${current}/${pages}页)`;
            }
        );

        onlineRateDetailRows = sortOnlineRateDetailRows(onlineRateDetailRows);
        onlineRateDetailBuckets = {
            offline: sortOnlineRateDetailRows(onlineRateDetailBuckets.offline),
            '24h': sortOnlineRateDetailRows(onlineRateDetailBuckets['24h']),
            '48h': sortOnlineRateDetailRows(onlineRateDetailBuckets['48h']),
            '72h': sortOnlineRateDetailRows(onlineRateDetailBuckets['72h']),
            over72h: sortOnlineRateDetailRows(onlineRateDetailBuckets.over72h),
            cancelled: sortOnlineRateDetailRows(onlineRateDetailBuckets.cancelled)
        };
    }

    async function startOnlineRateProcess() {
        const startBtn = document.getElementById('sr-start-btn');
        const exportBtn = document.getElementById('sr-export-btn');
        startBtn.disabled = true; exportBtn.disabled = true;

        finalReportData = {}; warehouseReportData = {}; customerReportData = {}; warehouseChannelReportData = {};
        onlineRateDetailRows = [];
        onlineRateDetailBuckets = { offline: [], '24h': [], '48h': [], '72h': [], over72h: [], cancelled: [] };
        extendedData = { trend: { '24h': {t:0,o:0}, '48h': {t:0,o:0}, '72h': {t:0,o:0} } };
        updateOnlineWarehouseFilterOptions();
        updateOnlineDetailWarehouseOptions();
        syncOnlineRateFilterState();
        document.getElementById('sr-online-detail-status').value = 'all';
        document.getElementById('sr-online-detail-warehouse').value = '';
        renderOnlineRateDetailTable();

        const skipWeekends = document.getElementById('sr-skip-weekends').checked;
        const sampleRange = getOnlineRateSampleRange(skipWeekends);
        const cohortWindow = buildOnlineRateWindow(sampleRange.startStr, sampleRange.endStr);
        dateLabels['24h'] = cohortWindow.label; dateLabels['48h'] = cohortWindow.label; dateLabels['72h'] = cohortWindow.label;

        try {
            if (new Date(sampleRange.startStr) > new Date(sampleRange.endStr)) {
                throw new Error('上网率样本开始日期不能晚于结束日期');
            }
            await fetchOnlineRateCohort(sampleRange.startStr, sampleRange.endStr, skipWeekends);
            updateOnlineWarehouseFilterOptions();
            updateOnlineDetailWarehouseOptions();
            syncOnlineRateFilterState();
            renderTable(); renderCharts(); renderOnlineRateDetailTable();
            document.getElementById('sr-status').innerText = '🎉 上网率数据渲染完成！';
            exportBtn.disabled = false;
        } catch (error) { document.getElementById('sr-status').innerText = `❌ 错误: ${error.message}`; }
        finally { startBtn.disabled = false; }
    }

    function exportOnlineExcel() {
        if (!window.XLSX) return alert('Excel组件加载中...');
        const wb = XLSX.utils.book_new();
        const genSheet = (obj, dimTitle) => Object.keys(obj).map(k => {
            const r = obj[k];
            return {
                [dimTitle]: k,
                [`24H上网率(创建样本${dateLabels['24h']})`]: r['24h'].total === 0 ? "-" : `${((r['24h'].online/r['24h'].total)*100).toFixed(2)}%`, "24H 详情": `${r['24h'].online} / ${r['24h'].total}`,
                [`48H上网率(创建样本${dateLabels['48h']})`]: r['48h'].total === 0 ? "-" : `${((r['48h'].online/r['48h'].total)*100).toFixed(2)}%`, "48H 详情": `${r['48h'].online} / ${r['48h'].total}`,
                [`72H上网率(创建样本${dateLabels['72h']})`]: r['72h'].total === 0 ? "-" : `${((r['72h'].online/r['72h'].total)*100).toFixed(2)}%`, "72H 详情": `${r['72h'].online} / ${r['72h'].total}`
            };
        });
        const genWarehouseChannelSheet = () => Object.keys(warehouseChannelReportData)
            .sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'))
            .flatMap((warehouseName) => Object.keys(warehouseChannelReportData[warehouseName] || {})
                .sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'))
                .map((channelName) => {
                    const row = warehouseChannelReportData[warehouseName][channelName];
                    return {
                        "发货仓库": warehouseName,
                        "物流渠道": channelName,
                        [`24H上网率(创建样本${dateLabels['24h']})`]: row['24h'].total === 0 ? "-" : `${((row['24h'].online / row['24h'].total) * 100).toFixed(2)}%`,
                        "24H 详情": `${row['24h'].online} / ${row['24h'].total}`,
                        [`48H上网率(创建样本${dateLabels['48h']})`]: row['48h'].total === 0 ? "-" : `${((row['48h'].online / row['48h'].total) * 100).toFixed(2)}%`,
                        "48H 详情": `${row['48h'].online} / ${row['48h'].total}`,
                        [`72H上网率(创建样本${dateLabels['72h']})`]: row['72h'].total === 0 ? "-" : `${((row['72h'].online / row['72h'].total) * 100).toFixed(2)}%`,
                        "72H 详情": `${row['72h'].online} / ${row['72h'].total}`
                    };
                }));
        const detailHeaders = ["订单号", "来源单号", "客户名称", "发货仓库", "物流渠道", "创建时间", "上网时间", "出库时间", "上网时长", "判定结果", "24H结果", "48H结果", "72H结果"];
        const genDetailSheet = (rows) => (rows || []).map((row) => ({
            "订单号": row.orderNo || '',
            "来源单号": row.sourceNo || '',
            "客户名称": row.customerName || '',
            "发货仓库": row.whName || '',
            "物流渠道": row.channelName || '',
            "创建时间": row.createTime || '',
            "上网时间": row.receiptTime || '',
            "出库时间": row.outboundTime || '',
            "上网时长": row.onlineDuration || '',
            "判定结果": row.judgeLabel || '',
            "24H结果": row.rate24 ? '已上网' : '未上网',
            "48H结果": row.rate48 ? '已上网' : '未上网',
            "72H结果": row.rate72 ? '已上网' : '未上网'
        }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genSheet(finalReportData, "渠道名称")), "按渠道上网率");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genSheet(warehouseReportData, "发货仓库")), "按仓库上网率");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genSheet(customerReportData, "客户名称")), "按客户上网率");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genWarehouseChannelSheet()), "按仓库渠道上网率");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genDetailSheet(onlineRateDetailRows), { header: detailHeaders }), "样本订单明细");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genDetailSheet(onlineRateDetailBuckets.offline), { header: detailHeaders }), "未上网明细");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genDetailSheet(onlineRateDetailBuckets['24h']), { header: detailHeaders }), "24H上网明细");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genDetailSheet(onlineRateDetailBuckets['48h']), { header: detailHeaders }), "48H上网明细");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genDetailSheet(onlineRateDetailBuckets['72h']), { header: detailHeaders }), "72H上网明细");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genDetailSheet(onlineRateDetailBuckets.over72h), { header: detailHeaders }), "超过72H已上网明细");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genDetailSheet(onlineRateDetailBuckets.cancelled), { header: detailHeaders }), "已取消明细");
        XLSX.writeFile(wb, `物流上网率统计_${formatDateStandard(new Date())}.xlsx`);
    }

    // ==========================================
    // 7. 周期平均上网率分析
    // ==========================================
    function renderPeriodTable() {
        const dimKey = document.getElementById('sr-period-dim').value;
        const dataObj = periodReportData[dimKey];
        const tbody = document.querySelector('#sr-period-table tbody');

        // 更新全局 Badge
        const summary = periodReportData.summary;
        const calcAvg = (val, tot) => tot === 0 ? "0.00%" : `${((val / tot) * 100).toFixed(2)}%`;
        document.getElementById('pb-total').innerText = summary.total;
        document.getElementById('pb-24').innerText = calcAvg(summary.in24, summary.total);
        document.getElementById('pb-48').innerText = calcAvg(summary.in48, summary.total);
        document.getElementById('pb-72').innerText = calcAvg(summary.in72, summary.total);

        tbody.innerHTML = '';
        const keys = Object.keys(dataObj).sort((a, b) => dataObj[b].total - dataObj[a].total);
        if (keys.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5">该周期内无符合条件的数据</td></tr>';
            return;
        }

        keys.forEach(key => {
            const row = dataObj[key];
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align:left; font-weight:bold;">${key}</td>
                <td>${row.total}</td>
                <td><span style="color:#1890ff;font-weight:bold">${calcAvg(row.in24, row.total)}</span><br/><span style="color:#aaa;font-size:11px;">(${row.in24}/${row.total})</span></td>
                <td><span style="color:#52c41a;font-weight:bold">${calcAvg(row.in48, row.total)}</span><br/><span style="color:#aaa;font-size:11px;">(${row.in48}/${row.total})</span></td>
                <td><span style="color:#faad14;font-weight:bold">${calcAvg(row.in72, row.total)}</span><br/><span style="color:#aaa;font-size:11px;">(${row.in72}/${row.total})</span></td>
            `;
            tbody.appendChild(tr);
        });
    }

    async function startPeriodProcess() {
        const btn = document.getElementById('sr-period-btn');
        const exportBtn = document.getElementById('sr-period-export');
        const startStr = document.getElementById('sr-period-start').value;
        const endStr = document.getElementById('sr-period-end').value;
        const skipWeekends = document.getElementById('sr-period-skip').checked;

        if (!startStr || !endStr) return alert('请先选择完整的周期起止时间！');
        if (new Date(startStr) > new Date(endStr)) return alert('开始日期不能晚于结束日期！');

        btn.disabled = true; exportBtn.disabled = true;
        btn.innerText = "数据拉取计算中...";

        // 初始化存储
        periodReportData = { channel: {}, warehouse: {}, customer: {}, summary: { total: 0, in24: 0, in48: 0, in72: 0 } };

        try {
            await forEachPeriodOrder(
                startStr,
                endStr,
                skipWeekends,
                (r) => {
                    if (!shouldIncludeInOnlineRateStats(r)) return;

                    const channel = r.logisticsChannel || '未知渠道';
                    const whName = r.whCodeName || '未知仓';
                    const customer = r.customerName || '未知客户';

                    const hoursDiff = getHoursDiff(r.createTime, r.receiptTime);
                    const is24 = hoursDiff <= 24;
                    const is48 = hoursDiff <= 48;
                    const is72 = hoursDiff <= 72;

                    const addData = (obj, key) => {
                        if (!obj[key]) obj[key] = { total: 0, in24: 0, in48: 0, in72: 0 };
                        obj[key].total++;
                        if (is24) obj[key].in24++;
                        if (is48) obj[key].in48++;
                        if (is72) obj[key].in72++;
                    };

                    addData(periodReportData.channel, channel);
                    addData(periodReportData.warehouse, whName);
                    addData(periodReportData.customer, customer);

                    periodReportData.summary.total++;
                    if (is24) periodReportData.summary.in24++;
                    if (is48) periodReportData.summary.in48++;
                    if (is72) periodReportData.summary.in72++;
                },
                (current, pages) => {
                    btn.innerText = `数据拉取计算中... (第${current}/${pages}页)`;
                }
            );

            renderPeriodTable();
            exportBtn.disabled = false;
            btn.innerText = "🚀 周期计算完成！";
            setTimeout(() => { btn.innerText = "🚀 开始计算周期平均"; }, 3000);

        } catch (error) {
            alert(`获取周期数据失败: ${error.message}`);
            btn.innerText = "🚀 开始计算周期平均";
        } finally {
            btn.disabled = false;
        }
    }

    function exportPeriodExcel() {
        if (!window.XLSX) return alert('Excel组件加载中...');
        const startStr = document.getElementById('sr-period-start').value;
        const endStr = document.getElementById('sr-period-end').value;
        const wb = XLSX.utils.book_new();

        const calcRate = (val, tot) => tot === 0 ? "0.00%" : `${((val / tot) * 100).toFixed(2)}%`;

        const genPeriodSheet = (obj, dimTitle) => Object.keys(obj).sort((a,b)=> obj[b].total - obj[a].total).map(k => {
            const r = obj[k];
            return {
                [dimTitle]: k,
                "周期总发货单量": r.total,
                "24H 平均达标率": calcRate(r.in24, r.total),
                "24H 达标单量": r.in24,
                "48H 平均达标率": calcRate(r.in48, r.total),
                "48H 达标单量": r.in48,
                "72H 平均达标率": calcRate(r.in72, r.total),
                "72H 达标单量": r.in72
            };
        });

        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genPeriodSheet(periodReportData.channel, "物流渠道")), "按渠道平均达标");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genPeriodSheet(periodReportData.warehouse, "发货仓库")), "按仓库平均达标");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(genPeriodSheet(periodReportData.customer, "客户名称")), "按客户平均达标");
        XLSX.writeFile(wb, `物流周期平均上网率_${startStr}至${endStr}.xlsx`);
    }

    function renderRegionDistributionTable() {
        const thead = document.querySelector('#sr-region-table thead');
        const tbody = document.querySelector('#sr-region-table tbody');
        const dim = document.getElementById('sr-region-dim').value;
        const rows = dim === 'warehouse' ? regionReportData.detailRows : regionReportData.customerRows;

        if (dim === 'warehouse') {
            thead.innerHTML = `<tr><th>客户名称</th><th>发货仓库</th><th>总订单量</th>${REGION_COLUMNS.map((region) => `<th>${region}</th>`).join('')}<th>样例邮编</th></tr>`;
        } else {
            thead.innerHTML = `<tr><th>客户名称</th><th>总订单量</th>${REGION_COLUMNS.map((region) => `<th>${region}</th>`).join('')}</tr>`;
        }

        tbody.innerHTML = '';
        if (rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="${dim === 'warehouse' ? REGION_COLUMNS.length + 4 : REGION_COLUMNS.length + 2}" style="color:#999; padding:30px 0;">暂无地区统计结果</td></tr>`;
            return;
        }

        rows.forEach((row) => {
            const regionCells = REGION_COLUMNS.map((region) => `<td>${row[region] || 0}</td>`).join('');
            const tr = document.createElement('tr');
            tr.innerHTML = dim === 'warehouse'
                ? `<td style="text-align:left; font-weight:bold;">${row["客户名称"]}</td><td>${row["发货仓库"]}</td><td>${row["总订单量"]}</td>${regionCells}<td>${row["样例邮编"] || '-'}</td>`
                : `<td style="text-align:left; font-weight:bold;">${row["客户名称"]}</td><td>${row["总订单量"]}</td>${regionCells}`;
            tbody.appendChild(tr);
        });
    }

    function renderRegionDistributionCharts() {
        if (!window.echarts || !regionReportData) return;

        const pieData = REGION_COLUMNS
            .map((region) => ({ name: region, value: regionReportData.regionTotals[region] || 0 }))
            .filter((item) => item.value > 0);

        if (!charts.regionPie) charts.regionPie = echarts.init(document.getElementById('chart-region-pie'));
        charts.regionPie.setOption({
            title: { text: '整体地区占比', left: 'center' },
            tooltip: { trigger: 'item' },
            series: [{
                type: 'pie',
                radius: ['35%', '68%'],
                center: ['50%', '56%'],
                data: pieData,
                label: { formatter: '{b}\n{c}单' }
            }]
        }, true);

        const dim = document.getElementById('sr-region-dim').value;
        const sourceRows = (dim === 'warehouse' ? regionReportData.detailRows : regionReportData.customerRows).slice(0, 8);
        const labels = sourceRows.map((row) => dim === 'warehouse' ? `${row["客户名称"]} / ${row["发货仓库"]}` : row["客户名称"]);

        if (!charts.regionBar) charts.regionBar = echarts.init(document.getElementById('chart-region-bar'));
        charts.regionBar.setOption({
            title: { text: dim === 'warehouse' ? 'Top 8 客户分仓地区结构' : 'Top 8 客户地区结构', left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            legend: { top: 28, type: 'scroll' },
            grid: { left: '3%', right: '3%', bottom: '5%', containLabel: true },
            xAxis: { type: 'value', name: '订单量' },
            yAxis: { type: 'category', data: labels },
            series: REGION_COLUMNS.map((region) => ({
                name: region,
                type: 'bar',
                stack: 'total',
                emphasis: { focus: 'series' },
                data: sourceRows.map((row) => Number(row[region] || 0))
            }))
        }, true);
    }

    function renderRegionDetailPies() {
        const container = document.getElementById('sr-region-pies');
        container.innerHTML = '';
        charts.regionDetailPies.forEach((chart) => { if (chart) chart.dispose(); });
        charts.regionDetailPies = [];

        if (!window.echarts || !regionReportData) return;

        const rows = regionReportData.detailRows.slice(0, 12);
        if (rows.length === 0) {
            container.innerHTML = '<div style="color:#999; padding:12px 4px;">暂无客户分仓饼图数据</div>';
            return;
        }

        const renderQueue = [];
        rows.forEach((row, index) => {
            const card = document.createElement('div');
            card.className = 'sr-region-pie-card';
            const chartId = `sr-region-pie-card-${index}`;
            card.innerHTML = `
                <div class="sr-region-pie-title">${row["客户名称"]}</div>
                <div class="sr-region-pie-meta">${row["发货仓库"]} ｜ 总订单量 ${row["总订单量"]}</div>
                <div id="${chartId}" class="sr-region-pie-box"></div>
            `;
            container.appendChild(card);
            renderQueue.push({ chartId, row });
        });

        requestAnimationFrame(() => {
            renderQueue.forEach(({ chartId, row }) => {
                const el = document.getElementById(chartId);
                if (!el) return;
                const pieData = REGION_COLUMNS
                    .map((region) => ({ name: region, value: Number(row[region] || 0) }))
                    .filter((item) => item.value > 0);
                const chart = echarts.init(el);
                chart.setOption({
                    title: { text: '地区占比', left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'item' },
                    series: [{
                        type: 'pie',
                        radius: ['30%', '60%'],
                        center: ['50%', '58%'],
                        data: pieData,
                        label: { formatter: '{b}\n{d}%' }
                    }]
                }, true);
                chart.resize();
                charts.regionDetailPies.push(chart);
            });
        });
    }

    function renderRegionDistributionView() {
        if (!regionReportData) return;

        document.getElementById('rb-total').innerText = regionReportData.totalOrders;
        document.getElementById('rb-customers').innerText = regionReportData.customerCount;
        document.getElementById('rb-warehouses').innerText = regionReportData.warehouseCount;
        document.getElementById('rb-unknown').innerText = regionReportData.excludedUnknown || 0;

        const failText = regionReportData.detailFailed > 0 ? `，详情失败 ${regionReportData.detailFailed} 单` : '';
        const excludedText = regionReportData.excludedUnknown > 0 ? `，未纳入统计 ${regionReportData.excludedUnknown} 单` : '';
        document.getElementById('sr-region-summary').innerText = `${regionReportData.startStr} 至 ${regionReportData.endStr}，已纳入统计 ${regionReportData.totalOrders} 单，主力地区：${regionReportData.topRegion.region}（${regionReportData.topRegion.value} 单）${failText}${excludedText}`;

        renderRegionDistributionTable();
        renderRegionDistributionCharts();
        renderRegionDetailPies();
    }

    function resetRegionDistributionView() {
        document.getElementById('rb-total').innerText = '-';
        document.getElementById('rb-customers').innerText = '-';
        document.getElementById('rb-warehouses').innerText = '-';
        document.getElementById('rb-unknown').innerText = '-';
        document.getElementById('sr-region-summary').innerText = '请先点击【开始统计并展示】。';
        document.querySelector('#sr-region-table thead').innerHTML = '';
        document.querySelector('#sr-region-table tbody').innerHTML = '<tr><td style="color:#999; padding:30px 0;">暂无地区统计结果</td></tr>';
        document.getElementById('sr-region-pies').innerHTML = '';
        if (charts.regionPie) charts.regionPie.clear();
        if (charts.regionBar) charts.regionBar.clear();
        charts.regionDetailPies.forEach((chart) => { if (chart) chart.dispose(); });
        charts.regionDetailPies = [];
    }

    async function computeRegionDistributionReport(startStr, endStr, skipWeekends, progress) {
        const orders = [];
        const regionStats = {};
        let detailFailed = 0;
        let excludedUnknown = 0;

        const matchedCount = await forEachPeriodOrder(
            startStr,
            endStr,
            skipWeekends,
            (r) => {
                orders.push({
                    deliveryNo: r.deliveryNo || '',
                    customerCode: r.customerCode || '',
                    customerName: r.customerName || '未知客户',
                    whCode: r.whCode || '',
                    whName: r.whCodeName || '未知仓'
                });
            },
            (current, pages, count) => {
                if (progress) progress({ stage: 'collect', current, pages, count, total: orders.length });
            }
        );

        if (matchedCount === 0 || orders.length === 0) {
            throw new Error('所选周期内没有可统计的订单');
        }

        await runWithConcurrency(orders, 10, async (order, index) => {
            try {
                const addressInfo = await fetchOrderAddressInfo(order.deliveryNo);
                const postCode = addressInfo.postCode || '';
                const region = getRegionByPostCode(postCode);
                if (!upsertRegionStats(regionStats, order, region, postCode)) {
                    excludedUnknown++;
                    logger.warn(`订单邮编未知，已跳过地区统计: ${order.deliveryNo}`, { postCode });
                }
            } catch (error) {
                detailFailed++;
                excludedUnknown++;
                logger.warn(`订单地址信息获取失败，已跳过地区统计: ${order.deliveryNo}`, error);
            }

            if (progress) progress({ stage: 'detail', current: index + 1, total: orders.length, count: matchedCount });
            await sleep(80);
        });

        return buildRegionReport(regionStats, detailFailed, orders.length, startStr, endStr, excludedUnknown);
    }

    async function startRegionDistributionProcess() {
        const btn = document.getElementById('sr-region-start-btn');
        const exportBtn = document.getElementById('sr-region-export');
        const startStr = document.getElementById('sr-region-start').value;
        const endStr = document.getElementById('sr-region-end').value;
        const skipWeekends = document.getElementById('sr-region-skip').checked;
        const statusEl = document.getElementById('sr-status');

        if (!startStr || !endStr) return alert('请先选择完整的周期起止时间！');
        if (new Date(startStr) > new Date(endStr)) return alert('开始日期不能晚于结束日期！');

        btn.disabled = true;
        exportBtn.disabled = true;
        btn.innerText = '收集订单中...';
        resetRegionDistributionView();

        try {
            regionReportData = await computeRegionDistributionReport(startStr, endStr, skipWeekends, (meta) => {
                if (meta.stage === 'collect') {
                    btn.innerText = `收集订单中... (${meta.current}/${meta.pages}页)`;
                    statusEl.innerText = `正在收集地区统计订单... 已筛出 ${meta.count} 单`;
                } else {
                    btn.innerText = `统计地区中... (${meta.current}/${meta.total})`;
                    statusEl.innerText = `正在拉取订单详情并统计地区... ${meta.current}/${meta.total}`;
                }
            });

            renderRegionDistributionView();
            exportBtn.disabled = false;
            const failText = regionReportData.detailFailed > 0 ? `，其中 ${regionReportData.detailFailed} 单详情失败` : '';
            const excludedText = regionReportData.excludedUnknown > 0 ? `，${regionReportData.excludedUnknown} 单未纳入统计` : '';
            statusEl.innerText = `✅ 订单分布分析完成，共纳入统计 ${regionReportData.totalOrders} 单${failText}${excludedText}`;
            updateSalesReportReadyState();
        } catch (error) {
            regionReportData = null;
            exportBtn.disabled = true;
            resetRegionDistributionView();
            statusEl.innerText = `❌ 订单分布分析失败: ${error.message}`;
            updateSalesReportReadyState();
        } finally {
            btn.disabled = false;
            btn.innerText = '开始统计并展示';
        }
    }

    function exportRegionDistributionExcel() {
        if (!window.XLSX) return alert('Excel组件加载中...');
        if (!regionReportData) return alert('请先完成地区统计，再导出 Excel。');

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(regionReportData.detailRows), "客户分仓地区分布");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(regionReportData.customerRows), "按客户地区汇总");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(regionReportData.ruleRows), "邮编地区规则");
        XLSX.writeFile(wb, `客户分仓地区分布_${regionReportData.startStr}至${regionReportData.endStr}.xlsx`);
        document.getElementById('sr-status').innerText = `✅ 订单分布分析已导出，共纳入统计 ${regionReportData.totalOrders} 单`;
    }

    function legacyRenderInventoryAnalysisTable() {
        const tbody = document.querySelector('#sr-inventory-table tbody');
        tbody.innerHTML = '';

        if (!inventoryReportData || inventoryReportData.warehouseRows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" style="color:#999; padding:30px 0;">暂无库存分析结果</td></tr>';
            return;
        }

        inventoryReportData.warehouseRows.forEach((row) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align:left; font-weight:bold;">${row["仓库名称"]}</td>
                <td>${row["客户数"]}</td>
                <td>${row["期初库存"]}</td>
                <td>${row["期末库存"]}</td>
                <td>${row["出库预占"]}</td>
                <td>${row["库存周转率"]}%</td>
                <td>${row["库存周转天数"]}</td>
                <td>${row["库存销率"]}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function legacyRenderInventoryAnalysisCharts() {
        if (!window.echarts || !inventoryReportData) return;

        const pieData = inventoryReportData.warehouseRows
            .map((row) => ({ name: row["仓库名称"], value: Number(row["库存周转率"] || 0) }))
            .filter((item) => item.value > 0);
        if (!charts.inventoryPie) charts.inventoryPie = echarts.init(document.getElementById('chart-inventory-pie'));
        charts.inventoryPie.setOption({
            title: { text: '库存周转率饼图', left: 'center' },
            tooltip: { trigger: 'item', formatter: (params) => `${params.name}<br/>周转率: ${params.value}%<br/>占比: ${params.percent}%` },
            legend: { type: 'scroll', orient: 'vertical', right: 10, top: 30, bottom: 20 },
            series: [{
                type: 'pie',
                radius: ['35%', '68%'],
                center: ['38%', '55%'],
                data: pieData,
                label: { formatter: '{b}\n{c}%' }
            }]
        }, true);

        const barRows = inventoryReportData.warehouseRows.slice();
        if (!charts.inventoryBar) charts.inventoryBar = echarts.init(document.getElementById('chart-inventory-bar'));
        charts.inventoryBar.setOption({
            title: { text: '库存周转天数柱状图', left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            grid: { left: '3%', right: '3%', bottom: '8%', containLabel: true },
            xAxis: { type: 'category', data: barRows.map((row) => row["仓库名称"]), axisLabel: { interval: 0, rotate: 20 } },
            yAxis: { type: 'value', name: '天数' },
            series: [{
                name: '库存周转天数',
                type: 'bar',
                data: barRows.map((row) => Number(row["库存周转天数"] || 0)),
                itemStyle: { color: '#73c0de' }
            }]
        }, true);
    }

    function legacyRenderInventoryAnalysisView() {
        if (!inventoryReportData) return;

        document.getElementById('ib-warehouse-count').innerText = inventoryReportData.totalWarehouseCount;
        document.getElementById('ib-close-stock').innerText = inventoryReportData.totalCloseStockQty;
        document.getElementById('ib-turnover-rate').innerText = `${inventoryReportData.totalTurnoverRate}%`;
        document.getElementById('ib-turnover-days').innerText = inventoryReportData.totalTurnoverDays;

        const topWarehouseText = inventoryReportData.topTurnoverWarehouse
            ? `${inventoryReportData.topTurnoverWarehouse["仓库名称"]}（${inventoryReportData.topTurnoverWarehouse["库存周转率"]}%）`
            : '-';
        document.getElementById('sr-inventory-summary').innerText = `${inventoryReportData.startStr} 至 ${inventoryReportData.endStr}，刷新日期 ${inventoryReportData.refreshTime || '-'}，周转率最高仓库：${topWarehouseText}`;

        renderInventoryAnalysisCharts();
        renderInventoryAnalysisTable();
    }

    function legacyResetInventoryAnalysisView() {
        document.getElementById('ib-warehouse-count').innerText = '-';
        document.getElementById('ib-close-stock').innerText = '-';
        document.getElementById('ib-turnover-rate').innerText = '-';
        document.getElementById('ib-turnover-days').innerText = '-';
        document.getElementById('sr-inventory-summary').innerText = '请先点击【开始库存分析】。';
        document.querySelector('#sr-inventory-table tbody').innerHTML = '<tr><td colspan="8" style="color:#999; padding:30px 0;">暂无库存分析结果</td></tr>';
        if (charts.inventoryPie) charts.inventoryPie.clear();
        if (charts.inventoryBar) charts.inventoryBar.clear();
    }

    function getInventoryDimensionMeta() {
        const dim = document.getElementById('sr-inventory-dim')?.value || 'detail';
        const rows = !inventoryReportData
            ? []
            : (dim === 'customer' ? inventoryReportData.customerRows : inventoryReportData.detailRows);
        return { dim, rows };
    }

    function getInventoryRowLabel(row, dim) {
        return dim === 'customer'
            ? row["客户名称"]
            : `${row["客户名称"]} / ${row["发货仓库"]}`;
    }

    function renderInventoryAnalysisTable() {
        const thead = document.querySelector('#sr-inventory-table thead');
        const tbody = document.querySelector('#sr-inventory-table tbody');
        const { dim, rows } = getInventoryDimensionMeta();

        thead.innerHTML = dim === 'customer'
            ? '<tr><th>客户名称</th><th>分仓数</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>'
            : '<tr><th>客户名称</th><th>发货仓库</th><th>期初库存</th><th>期末库存</th><th>出库预占</th><th>库存周转率</th><th>库存周转天数</th><th>库存售罄率</th></tr>';
        tbody.innerHTML = '';

        if (!inventoryReportData || rows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" style="color:#999; padding:30px 0;">暂无库存分析结果</td></tr>';
            return;
        }

        rows.forEach((row) => {
            const tr = document.createElement('tr');
            tr.innerHTML = dim === 'customer'
                ? `
                    <td style="text-align:left; font-weight:bold;">${row["客户名称"]}</td>
                    <td>${row["分仓数"]}</td>
                    <td>${row["期初库存"]}</td>
                    <td>${row["期末库存"]}</td>
                    <td>${row["出库预占"]}</td>
                    <td>${row["库存周转率"]}%</td>
                    <td>${row["库存周转天数"]}</td>
                    <td>${row["库存售罄率"]}</td>
                `
                : `
                    <td style="text-align:left; font-weight:bold;">${row["客户名称"]}</td>
                    <td>${row["发货仓库"]}</td>
                    <td>${row["期初库存"]}</td>
                    <td>${row["期末库存"]}</td>
                    <td>${row["出库预占"]}</td>
                    <td>${row["库存周转率"]}%</td>
                    <td>${row["库存周转天数"]}</td>
                    <td>${row["库存售罄率"]}</td>
                `;
            tbody.appendChild(tr);
        });
    }

    function renderInventoryAnalysisCharts() {
        if (!window.echarts || !inventoryReportData) return;

        const { dim, rows } = getInventoryDimensionMeta();
        const rankedRows = rows
            .slice()
            .sort((a, b) => b["期末库存"] - a["期末库存"] || b["库存周转率"] - a["库存周转率"]);
        const pieData = rankedRows
            .slice(0, 10)
            .map((row) => ({ name: getInventoryRowLabel(row, dim), value: Number(row["期末库存"] || 0) }))
            .filter((item) => item.value > 0);

        if (!charts.inventoryPie) charts.inventoryPie = echarts.init(document.getElementById('chart-inventory-pie'));
        charts.inventoryPie.setOption({
            title: { text: dim === 'customer' ? '客户期末库存占比' : '客户分仓期末库存占比', left: 'center' },
            tooltip: { trigger: 'item', formatter: (params) => `${params.name}<br/>期末库存: ${params.value}<br/>占比: ${params.percent}%` },
            legend: { type: 'scroll', orient: 'vertical', right: 10, top: 30, bottom: 20 },
            series: [{
                type: 'pie',
                radius: ['35%', '68%'],
                center: ['38%', '55%'],
                data: pieData,
                label: { formatter: '{b}\n{d}%' }
            }]
        }, true);

        const barRows = rows
            .slice()
            .sort((a, b) => Number(b["库存周转天数"] || 0) - Number(a["库存周转天数"] || 0) || Number(b["期末库存"] || 0) - Number(a["期末库存"] || 0))
            .slice(0, 8);
        if (!charts.inventoryBar) charts.inventoryBar = echarts.init(document.getElementById('chart-inventory-bar'));
        charts.inventoryBar.setOption({
            title: { text: dim === 'customer' ? 'Top 8 客户库存表现' : 'Top 8 客户分仓库存表现', left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
            legend: { top: 28, data: ['库存周转率', '库存周转天数'] },
            grid: { left: '3%', right: '3%', bottom: '10%', containLabel: true },
            xAxis: {
                type: 'category',
                data: barRows.map((row) => getInventoryRowLabel(row, dim)),
                axisLabel: { interval: 0, rotate: 18 }
            },
            yAxis: [
                { type: 'value', name: '周转率(%)' },
                { type: 'value', name: '周转天数' }
            ],
            series: [
                {
                    name: '库存周转率',
                    type: 'bar',
                    data: barRows.map((row) => Number(row["库存周转率"] || 0)),
                    itemStyle: { color: '#73c0de' }
                },
                {
                    name: '库存周转天数',
                    type: 'line',
                    yAxisIndex: 1,
                    data: barRows.map((row) => Number(row["库存周转天数"] || 0)),
                    itemStyle: { color: '#91cc75' },
                    lineStyle: { width: 3 },
                    symbolSize: 8
                }
            ]
        }, true);
    }

    function renderInventoryCustomerCharts() {
        const container = document.getElementById('sr-inventory-customer-charts');
        if (!container) return;

        container.innerHTML = '';
        charts.inventoryCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.inventoryCustomerCharts = [];

        if (!window.echarts || !inventoryReportData) return;

        const detailGroupMap = {};
        inventoryReportData.detailRows.forEach((row) => {
            const customerKey = `${row["客户编码"] || ''}__${row["客户名称"] || ''}`;
            if (!detailGroupMap[customerKey]) {
                detailGroupMap[customerKey] = [];
            }
            detailGroupMap[customerKey].push(row);
        });

        const groups = inventoryReportData.customerRows
            .slice()
            .sort((a, b) =>
                a["客户名称"].localeCompare(b["客户名称"], 'zh-Hans-CN') ||
                (a["客户编码"] || '').localeCompare(b["客户编码"] || '')
            )
            .map((customerRow) => {
                const customerKey = `${customerRow["客户编码"] || ''}__${customerRow["客户名称"] || ''}`;
                const detailRows = (detailGroupMap[customerKey] || []).slice().sort((a, b) =>
                    b["期末库存"] - a["期末库存"] || b["库存周转率"] - a["库存周转率"]
                );
                return { customerRow, detailRows };
            })
            .filter((group) => group.detailRows.length > 0);

        if (groups.length === 0) {
            container.innerHTML = '<div style="color:#999; padding:12px 4px;">暂无客户库存周转图表数据</div>';
            return;
        }

        const renderQueue = [];
        groups.forEach(({ customerRow, detailRows }, index) => {
            const card = document.createElement('div');
            card.className = 'sr-inventory-card';
            const pieId = `sr-inventory-customer-pie-${index}`;
            const barId = `sr-inventory-customer-bar-${index}`;
            card.innerHTML = `
                <div class="sr-inventory-card-title">${customerRow["客户名称"]}</div>
                <div class="sr-inventory-card-meta">分仓数 ${customerRow["分仓数"]} ｜ 期末库存 ${customerRow["期末库存"]}</div>
                <div class="sr-inventory-card-charts">
                    <div id="${pieId}" class="sr-inventory-card-box-half"></div>
                    <div id="${barId}" class="sr-inventory-card-box-half"></div>
                </div>
            `;
            container.appendChild(card);
            renderQueue.push({ pieId, barId, detailRows });
        });

        requestAnimationFrame(() => {
            renderQueue.forEach(({ pieId, barId, detailRows }) => {
                const pieEl = document.getElementById(pieId);
                const barEl = document.getElementById(barId);
                if (!pieEl || !barEl) return;

                const pieData = detailRows
                    .map((row) => ({ name: row["发货仓库"], value: Number(row["库存周转率"] || 0) }))
                    .filter((item) => item.value > 0);
                const barLabels = detailRows.map((row) => row["发货仓库"]);
                const barData = detailRows.map((row) => Number(row["库存周转天数"] || 0));

                const pieChart = echarts.init(pieEl);
                pieChart.setOption({
                    title: { text: '库存周转率', left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'item', formatter: (params) => `${params.name}<br/>周转率: ${params.value}%<br/>占比: ${params.percent}%` },
                    legend: { type: 'scroll', orient: 'vertical', right: 6, top: 34, bottom: 10 },
                    series: [{
                        type: 'pie',
                        radius: ['35%', '65%'],
                        center: ['38%', '50%'],
                        data: pieData,
                        label: { formatter: '{b}\n{d}%' }
                    }]
                }, true);
                pieChart.resize();
                charts.inventoryCustomerCharts.push(pieChart);

                const barChart = echarts.init(barEl);
                barChart.setOption({
                    title: { text: '库存周转天数', left: 'center', top: 0, textStyle: { fontSize: 13 } },
                    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
                    grid: { left: '8%', right: '6%', bottom: 72, top: 48, containLabel: true },
                    xAxis: {
                        type: 'category',
                        data: barLabels,
                        axisLabel: { interval: 0, rotate: 28, margin: 14 }
                    },
                    yAxis: { type: 'value', name: '天数' },
                    series: [{
                        name: '库存周转天数',
                        type: 'bar',
                        data: barData,
                        itemStyle: { color: '#91cc75' },
                        barMaxWidth: 42,
                        label: {
                            show: true,
                            position: 'top',
                            formatter: ({ value }) => value == null ? '' : value
                        }
                    }]
                }, true);
                barChart.resize();
                charts.inventoryCustomerCharts.push(barChart);
            });
        });
    }

    function renderInventoryAnalysisView() {
        if (!inventoryReportData) return;

        document.getElementById('ib-customer-count').innerText = inventoryReportData.customerCount;
        document.getElementById('ib-detail-count').innerText = inventoryReportData.detailCount;
        document.getElementById('ib-close-stock').innerText = inventoryReportData.totalCloseStockQty;
        document.getElementById('ib-turnover-rate').innerText = `${inventoryReportData.totalTurnoverRate}%`;

        const topWarehouseText = inventoryReportData.topTurnoverWarehouse
            ? `${inventoryReportData.topTurnoverWarehouse["客户名称"]} / ${inventoryReportData.topTurnoverWarehouse["发货仓库"]}（${inventoryReportData.topTurnoverWarehouse["库存周转率"]}%）`
            : '-';
        const topCustomerText = inventoryReportData.topStockCustomer
            ? `${inventoryReportData.topStockCustomer["客户名称"]}（${inventoryReportData.topStockCustomer["期末库存"]}）`
            : '-';
        document.getElementById('sr-inventory-summary').innerText =
            `${inventoryReportData.startStr} 至 ${inventoryReportData.endStr}，刷新日期 ${inventoryReportData.refreshTime || '-'}，总周转天数 ${inventoryReportData.totalTurnoverDays}，库存最高客户：${topCustomerText}，周转率最高分仓：${topWarehouseText}`;

        renderInventoryAnalysisCharts();
        renderInventoryCustomerCharts();
        renderInventoryAnalysisTable();
    }

    function resetInventoryAnalysisView() {
        const thead = document.querySelector('#sr-inventory-table thead');
        const tbody = document.querySelector('#sr-inventory-table tbody');
        document.getElementById('ib-customer-count').innerText = '-';
        document.getElementById('ib-detail-count').innerText = '-';
        document.getElementById('ib-close-stock').innerText = '-';
        document.getElementById('ib-turnover-rate').innerText = '-';
        document.getElementById('sr-inventory-summary').innerText = '请先点击【开始库存分析】。';
        thead.innerHTML = '';
        tbody.innerHTML = '<tr><td colspan="8" style="color:#999; padding:30px 0;">暂无库存分析结果</td></tr>';
        if (charts.inventoryPie) charts.inventoryPie.clear();
        if (charts.inventoryBar) charts.inventoryBar.clear();
        const customerCharts = document.getElementById('sr-inventory-customer-charts');
        if (customerCharts) customerCharts.innerHTML = '';
        charts.inventoryCustomerCharts.forEach((chart) => { if (chart) chart.dispose(); });
        charts.inventoryCustomerCharts = [];
    }

    async function computeInventoryAnalysisReport(startStr, endStr, skipWeekends, progress) {
        const startTime = `${startStr} 00:00:00`;
        const endTime = `${endStr} 23:59:59`;
        const records = [];
        let current = 1;
        let pages = 1;
        let refreshTime = '';
        let totalData = null;

        while (current <= pages) {
            if (progress) progress(current, pages, records.length);
            const data = await fetchStockAnalysisPage(startTime, endTime, current);
            const pageData = data.page || {};
            const pageRecords = (pageData.records || []).filter((record) =>
                !skipWeekends || !isWeekendDate(record.statisticDate)
            );
            const pageSize = Number(pageData.size || 500);
            const pageTotal = Number(pageData.total || 0);
            records.push(...pageRecords);
            refreshTime = data.refreshTime || refreshTime;
            totalData = data.total || totalData;
            pages = pageData.pages
                || data.pages
                || (pageTotal > 0 ? Math.ceil(pageTotal / pageSize) : (pageRecords.length < pageSize ? current : current + 1));
            current++;
            await sleep(150);
        }

        return buildInventoryReport(records, totalData, refreshTime, startStr, endStr);
    }

    async function startInventoryAnalysisProcess() {
        const btn = document.getElementById('sr-inventory-start-btn');
        const exportBtn = document.getElementById('sr-inventory-export');
        const startStr = document.getElementById('sr-inventory-start').value;
        const endStr = document.getElementById('sr-inventory-end').value;
        const skipWeekends = document.getElementById('sr-inventory-skip-weekends').checked;
        const statusEl = document.getElementById('sr-status');

        if (!startStr || !endStr) return alert('请先选择完整的库存分析周期！');
        if (new Date(startStr) > new Date(endStr)) return alert('开始日期不能晚于结束日期！');

        btn.disabled = true;
        exportBtn.disabled = true;
        btn.innerText = '库存分析中...';
        resetInventoryAnalysisView();

        try {
            inventoryReportData = await computeInventoryAnalysisReport(startStr, endStr, skipWeekends, (current, pages, count) => {
                btn.innerText = `库存分析中... (${current}/${pages}页)`;
                statusEl.innerText = `正在拉取库存分析数据... 已累计 ${count} 条`;
            });

            renderInventoryAnalysisView();
            exportBtn.disabled = false;
            statusEl.innerText = `✅ 库存分析完成，共汇总 ${inventoryReportData.totalWarehouseCount} 个仓库`;
            updateSalesReportReadyState();
        } catch (error) {
            inventoryReportData = null;
            exportBtn.disabled = true;
            resetInventoryAnalysisView();
            statusEl.innerText = `❌ 库存分析失败: ${error.message}`;
            updateSalesReportReadyState();
        } finally {
            btn.disabled = false;
            btn.innerText = '开始库存分析';
        }
    }

    // ==========================================
    // 8. 出库发货率独立获取与导出 (WMS系统)
    // ==========================================
    function normalizeOutboundRateValue(value) {
        const numeric = Number(String(value ?? '0').replace('%', ''));
        return Number.isFinite(numeric) ? Number(numeric.toFixed(2)) : 0;
    }

    function formatOutboundRate(value) {
        return `${normalizeOutboundRateValue(value).toFixed(2)}%`;
    }

    function buildOutboundSummary(rows) {
        const warehouseCount = rows.length;
        const avg = (key) => warehouseCount === 0
            ? 0
            : Number((rows.reduce((sum, row) => sum + Number(row[key] || 0), 0) / warehouseCount).toFixed(2));
        return {
            warehouseCount,
            avg24: avg('rate24'),
            avg48: avg('rate48'),
            avg72: avg('rate72')
        };
    }

    async function fetchOutboundRateWithWeekendOption(warehouse, startDateStr, endDateStr, skipWeekends) {
        if (!skipWeekends) {
            const rateData = await fetchOutboundRate(warehouse.whCode, warehouse.tenantCode, startDateStr, endDateStr);
            return {
                rate24: normalizeOutboundRateValue(rateData.dropShippingDeliveryRateOf24),
                rate48: normalizeOutboundRateValue(rateData.dropShippingDeliveryRateOf48),
                rate72: normalizeOutboundRateValue(rateData.dropShippingDeliveryRateOf72),
                includedDays: getDateRangeList(startDateStr, endDateStr, false).length
            };
        }

        const workdays = getDateRangeList(startDateStr, endDateStr, true);
        if (workdays.length === 0) throw new Error('所选周期排除周末后没有可统计的工作日');

        const totals = { rate24: 0, rate48: 0, rate72: 0, count: 0 };
        for (const dateStr of workdays) {
            const rateData = await fetchOutboundRate(warehouse.whCode, warehouse.tenantCode, dateStr, dateStr);
            totals.rate24 += normalizeOutboundRateValue(rateData.dropShippingDeliveryRateOf24);
            totals.rate48 += normalizeOutboundRateValue(rateData.dropShippingDeliveryRateOf48);
            totals.rate72 += normalizeOutboundRateValue(rateData.dropShippingDeliveryRateOf72);
            totals.count++;
            await sleep(80);
        }

        return {
            rate24: Number((totals.rate24 / totals.count).toFixed(2)),
            rate48: Number((totals.rate48 / totals.count).toFixed(2)),
            rate72: Number((totals.rate72 / totals.count).toFixed(2)),
            includedDays: totals.count
        };
    }

    function renderOutboundRateTable() {
        const tbody = document.querySelector('#sr-outbound-table tbody');
        if (!tbody) return;
        const rows = outboundReportData.rows || [];
        tbody.innerHTML = '';
        if (rows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="color:#999; padding:30px 0;">请选择周期并点击【抓取并展示】</td></tr>';
            return;
        }

        rows.forEach((row) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align:left; font-weight:bold;">${row.whName}</td>
                <td>${row.whCode}</td>
                <td>${formatOutboundRate(row.rate24)}</td>
                <td>${formatOutboundRate(row.rate48)}</td>
                <td>${formatOutboundRate(row.rate72)}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function renderOutboundRateCharts() {
        if (!window.echarts) return;
        const rows = outboundReportData.rows || [];
        const summary = outboundReportData.summary || buildOutboundSummary(rows);

        if (!charts.outboundLine) charts.outboundLine = echarts.init(document.getElementById('chart-outbound-line'));
        charts.outboundLine.setOption({
            title: { text: '整体平均出库发货率', left: 'center' },
            tooltip: { trigger: 'axis', valueFormatter: (value) => `${Number(value || 0).toFixed(2)}%` },
            grid: { left: '4%', right: '4%', bottom: 48, top: 58, containLabel: true },
            xAxis: { type: 'category', data: ['24H', '48H', '72H'] },
            yAxis: { type: 'value', name: '出库发货率(%)', min: 0, max: 100 },
            series: [{
                name: '平均出库发货率',
                type: 'line',
                data: [summary.avg24, summary.avg48, summary.avg72],
                itemStyle: { color: '#faad14' },
                lineStyle: { width: 3 },
                symbolSize: 8,
                label: { show: true, position: 'top', formatter: ({ value }) => `${Number(value || 0).toFixed(2)}%` }
            }]
        }, true);

        const topRows = rows
            .slice()
            .sort((a, b) => Number(b.rate24 || 0) - Number(a.rate24 || 0) || String(a.whName).localeCompare(String(b.whName), 'zh-Hans-CN'))
            .slice(0, 12);
        if (!charts.outboundBar) charts.outboundBar = echarts.init(document.getElementById('chart-outbound-bar'));
        charts.outboundBar.setOption({
            title: { text: topRows.length >= 12 ? 'Top 12 仓库出库发货率' : '仓库出库发货率', left: 'center' },
            tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: (value) => `${Number(value || 0).toFixed(2)}%` },
            legend: { top: 28 },
            grid: { left: '3%', right: '3%', bottom: 62, top: 72, containLabel: true },
            xAxis: { type: 'category', data: topRows.map((row) => row.whName), axisLabel: { interval: 0, rotate: 18 } },
            yAxis: { type: 'value', name: '出库发货率(%)', min: 0, max: 100 },
            series: [
                { name: '24H', type: 'bar', data: topRows.map((row) => row.rate24), itemStyle: { color: '#faad14' }, barMaxWidth: 22 },
                { name: '48H', type: 'bar', data: topRows.map((row) => row.rate48), itemStyle: { color: '#52c41a' }, barMaxWidth: 22 },
                { name: '72H', type: 'bar', data: topRows.map((row) => row.rate72), itemStyle: { color: '#1890ff' }, barMaxWidth: 22 }
            ]
        }, true);
    }

    function renderOutboundRateView() {
        const rows = outboundReportData.rows || [];
        const summary = outboundReportData.summary || buildOutboundSummary(rows);
        const startInput = document.getElementById('sr-outbound-start');
        const endInput = document.getElementById('sr-outbound-end');
        const skipInput = document.getElementById('sr-outbound-skip');
        if (startInput && outboundReportData.startDate) startInput.value = outboundReportData.startDate;
        if (endInput && outboundReportData.endDate) endInput.value = outboundReportData.endDate;
        if (skipInput) skipInput.checked = Boolean(outboundReportData.skipWeekends);
        document.getElementById('ob-warehouse-count').innerText = summary.warehouseCount || 0;
        document.getElementById('ob-24').innerText = formatOutboundRate(summary.avg24);
        document.getElementById('ob-48').innerText = formatOutboundRate(summary.avg48);
        document.getElementById('ob-72').innerText = formatOutboundRate(summary.avg72);
        renderOutboundRateTable();
        renderOutboundRateCharts();
    }

    async function startOutboundRateProcess() {
        const startDateStr = document.getElementById('sr-outbound-start').value;
        const endDateStr = document.getElementById('sr-outbound-end').value;
        const skipWeekends = document.getElementById('sr-outbound-skip').checked;

        if (!startDateStr || !endDateStr) return alert("请先选择出库发货率统计的「开始日期」和「结束日期」！");
        if (new Date(startDateStr) > new Date(endDateStr)) return alert("开始日期不能晚于结束日期，请重新选择！");
        if (skipWeekends && getDateRangeList(startDateStr, endDateStr, true).length === 0) {
            return alert("所选周期排除周末后没有可统计的工作日，请重新选择！");
        }

        const btn = document.getElementById('sr-outbound-btn');
        const exportBtn = document.getElementById('sr-outbound-export');
        btn.disabled = true;
        exportBtn.disabled = true;
        btn.innerText = "数据抓取中...";
        resetOutboundRateView();

        try {
            document.getElementById('sr-status').innerText = `正在获取发货仓库配置列表...`;
            const warehouses = await fetchWarehouses();
            const rows = [];

            for (let i = 0; i < warehouses.length; i++) {
                const wh = warehouses[i];
                document.getElementById('sr-status').innerText = `抓取出库发货率: ${wh.whNameCn} (${i+1}/${warehouses.length})${skipWeekends ? '，已排除周末' : ''}`;

                try {
                    const rateData = await fetchOutboundRateWithWeekendOption(wh, startDateStr, endDateStr, skipWeekends);
                    rows.push({
                        whName: wh.whNameCn || wh.whCode,
                        whCode: wh.whCode,
                        rate24: rateData.rate24,
                        rate48: rateData.rate48,
                        rate72: rateData.rate72,
                        includedDays: rateData.includedDays,
                        skipWeekends,
                        startDate: startDateStr,
                        endDate: endDateStr
                    });
                } catch(err) {
                    logger.warn(`仓库 ${wh.whNameCn} 出库率获取失败`, err);
                }
                await new Promise(res => setTimeout(res, 250)); // 防封号停顿
            }

            if (rows.length === 0) throw new Error("未获取到有效的出库率数据");

            outboundReportData = {
                rows,
                startDate: startDateStr,
                endDate: endDateStr,
                skipWeekends,
                summary: buildOutboundSummary(rows)
            };
            renderOutboundRateView();
            exportBtn.disabled = false;
            document.getElementById('sr-status').innerText = `✅ 出库发货率抓取完成，共 ${rows.length} 个仓库${skipWeekends ? '，已排除周末' : ''}`;
            updateSalesReportReadyState();

        } catch (error) {
            outboundReportData = { rows: [], startDate: startDateStr, endDate: endDateStr, skipWeekends, summary: { warehouseCount: 0, avg24: 0, avg48: 0, avg72: 0 } };
            renderOutboundRateView();
            document.getElementById('sr-status').innerText = `❌ 出库发货率抓取失败: ${error.message}`;
            updateSalesReportReadyState();
        } finally {
            btn.disabled = false;
            btn.innerText = "抓取并展示";
        }
    }

    function exportOutboundExcel() {
        if (!window.XLSX) return alert('Excel组件加载中...');
        const rows = outboundReportData.rows || [];
        if (rows.length === 0) return alert('请先抓取出库发货率数据，再导出 Excel。');

        const exportData = rows.map((row) => ({
            "仓库名称": row.whName,
            "仓库代码": row.whCode,
            "出库发货率 24h (%)": row.rate24.toFixed(2),
            "出库发货率 48h (%)": row.rate48.toFixed(2),
            "出库发货率 72h (%)": row.rate72.toFixed(2),
            "统计起始日期": outboundReportData.startDate,
            "统计结束日期": outboundReportData.endDate,
            "是否排除周末": outboundReportData.skipWeekends ? "是" : "否",
            "纳入日期数": row.includedDays || ''
        }));
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "仓库出库发货率");
        XLSX.writeFile(wb, `全仓出库发货率_${outboundReportData.startDate}_至_${outboundReportData.endDate}.xlsx`);
        document.getElementById('sr-status').innerText = `✅ 出库发货率已成功导出！`;
    }

    function resetOutboundRateView() {
        outboundReportData = { rows: [], startDate: '', endDate: '', skipWeekends: false, summary: { warehouseCount: 0, avg24: 0, avg48: 0, avg72: 0 } };
        ['ob-warehouse-count', 'ob-24', 'ob-48', 'ob-72'].forEach((id) => {
            const el = document.getElementById(id);
            if (el) el.innerText = '-';
        });
        const tbody = document.querySelector('#sr-outbound-table tbody');
        if (tbody) tbody.innerHTML = '<tr><td colspan="5" style="color:#999; padding:30px 0;">请选择周期并点击【抓取并展示】</td></tr>';
        const exportBtn = document.getElementById('sr-outbound-export');
        if (exportBtn) exportBtn.disabled = true;
        if (charts.outboundLine) charts.outboundLine.clear();
        if (charts.outboundBar) charts.outboundBar.clear();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', injectUI);
    } else {
        injectUI();
    }

})();
