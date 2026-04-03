// ==UserScript==
// @name         Schedulingで受諾済みを計算
// @namespace    https://github.com/yuyna-amazon/DSP-Scheduling
// @version      8.6
// @description  Amazon Logistics DSP Scheduling
// @author       yuyna
// @icon         https://www.google.com/s2/favicons?sz=64&domain=amazon.com
// @match        https://logistics.amazon.co.jp/internal/scheduling/dsps*
// @updateURL    https://raw.githubusercontent.com/yuyna-amazon/DSP-Scheduling/main/DSP-Scheduling.user.js
// @downloadURL  https://raw.githubusercontent.com/yuyna-amazon/DSP-Scheduling/main/DSP-Scheduling.user.js
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==


(function() {
    'use strict';


    // === 状態管理 ===
    let currentSSDData = null;
    let currentTimeDataList = [];
    let currentTotals = { accepted: 0, required: 0, proDPAccepted: 0, proDPRequired: 0 };
    let isCalculating = false;
    let debounceTimer = null;
    let observer = null;
    let cachedTable = null;


    // === プリコンパイル正規表現 ===
    const TIME_REGEX = /(\d+):(\d+)\s*(午前|午後|am|pm)/i;
    const TIME_STRICT_REGEX = /^\d{1,2}:\d{2}\s*(午前|午後|am|pm)$/i;
    const DIGIT_REGEX = /^\d+$/;
    const DATE_REGEX = /(\d+)-(\d+)月/;
    const REQUIRED_REGEX = /text:\s*required\(\)/;
    const SCHEDULED_REGEX = /text:\s*scheduled\(\)/;


    // === 定数 ===
    const SSD_DEFAULTS = {
        'SSD_1': 1, 'SSD_C1': 1, 'SSD_1_B': 1, 'SSD_2': 1,
        'SSD_3': 1, 'SSD_C3': 1, 'SSD_3_B': 1, 'SSD_4': 1
    };
    const SSD_ADJUSTMENT_DEFAULTS = {
        'SSD_1': 0, 'SSD_C1': 0, 'SSD_1_B': 0, 'SSD_2': 0,
        'SSD_3': 0, 'SSD_C3': 0, 'SSD_3_B': 0, 'SSD_4': 0
    };
    const C1C3_OVERRIDE_DEFAULTS = { 'SSD_C1': 0, 'SSD_C3': 0 };


    const SSD_TIME_RANGES = [
        { key: 'SSD_1',   min: 0,    max: 420  },
        { key: 'SSD_C1',  min: 420,  max: 420  },
        { key: 'SSD_1_B', min: 420,  max: 600  },
        { key: 'SSD_2',   min: 600,  max: 840  },
        { key: 'SSD_3',   min: 840,  max: 1020 },
        { key: 'SSD_C3',  min: 1020, max: 1020 },
        { key: 'SSD_3_B', min: 1020, max: 1200 },
        { key: 'SSD_4',   min: 1200, max: Infinity }
    ];
    const SSD_LIST = ['SSD_1', 'SSD_C1', 'SSD_1_B', 'SSD_2', 'SSD_3', 'SSD_C3', 'SSD_3_B', 'SSD_4'];
    const SPR_LIST = ['SSD_1', 'SSD_1_B', 'SSD_2', 'SSD_3', 'SSD_3_B', 'SSD_4'];


    // === ユーティリティ ===
    const isServiceTypeName = (name) => name && name.includes('AmFlex');


    const parseTime = (timeStr) => {
        const match = TIME_REGEX.exec(timeStr);
        if (!match) return 0;
        let hour = +match[1];
        const minute = +match[2];
        const period = match[3].toLowerCase();
        if ((period === '午後' || period === 'pm') && hour !== 12) hour += 12;
        if ((period === '午前' || period === 'am') && hour === 12) hour = 0;
        return hour * 60 + minute;
    };


    const getSSDGroup = (timeMinutes) => {
        for (let i = 0; i < SSD_TIME_RANGES.length; i++) {
            const r = SSD_TIME_RANGES[i];
            if (timeMinutes >= r.min && timeMinutes < r.max) return r.key;
        }
        return null;
    };


    const getShortServiceType = (name) => {
        if (!name) return '-';
        const parts = name.split('_');
        return parts.length >= 2 ? parts.slice(1).join('_') : name;
    };


    // C1/C3はoverride値をSoft計算ベースに使う
    const getBaseAccepted = (sk, d, overrides) =>
        (sk === 'SSD_C1' || sk === 'SSD_C3') ? (overrides[sk] || 0) : d.accepted;


    const applyPct = (baseSoft, pct) => Math.round(baseSoft * (1 + pct / 100));


    // 1B/3B減算を適用したSoft値を返す
    const applySubtract = (sk, soft, overrides, subtract) => {
        if (!subtract) return soft;
        if (sk === 'SSD_1_B') return Math.max(0, soft - (overrides['SSD_C1'] || 0));
        if (sk === 'SSD_3_B') return Math.max(0, soft - (overrides['SSD_C3'] || 0));
        return soft;
    };


    // === ストレージ ===
    const getStorage = (key, defaultVal) => {
        try {
            const stored = localStorage.getItem(key);
            if (stored === null) return defaultVal;
            if (typeof defaultVal === 'boolean') return stored === 'true';
            if (typeof defaultVal === 'object') return { ...defaultVal, ...JSON.parse(stored) };
            return stored;
        } catch { return defaultVal; }
    };
    const setStorage = (key, val) =>
        localStorage.setItem(key, typeof val === 'object' ? JSON.stringify(val) : String(val));


    const getSSDMultipliers   = () => getStorage('dsp-ssd-multipliers',   SSD_DEFAULTS);
    const saveSSDMultipliers  = (m) => setStorage('dsp-ssd-multipliers',   m);
    const getSSDAdjustments   = () => getStorage('dsp-ssd-adjustments',    SSD_ADJUSTMENT_DEFAULTS);
    const saveSSDAdjustments  = (a) => setStorage('dsp-ssd-adjustments',   a);
    const getSoftPct          = () => Number(getStorage('dsp-soft-pct',    '0'));
    const saveSoftPct         = (v) => setStorage('dsp-soft-pct',          String(v));
    const getC1C3Overrides    = () => getStorage('dsp-c1c3-overrides',     C1C3_OVERRIDE_DEFAULTS);
    const saveC1C3Overrides   = (v) => setStorage('dsp-c1c3-overrides',    v);
    const getC1C3Subtract     = () => getStorage('dsp-c1c3-subtract',      false);
    const saveC1C3Subtract    = (v) => setStorage('dsp-c1c3-subtract',     v);


    // === 日付 ===
    const getDateForFileName = () => {
        const span = document.querySelector('li.selected .dateText');
        if (span) {
            const match = DATE_REGEX.exec(span.textContent);
            if (match) return match[2].padStart(2, '0') + match[1].padStart(2, '0');
        }
        const d = new Date();
        return String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
    };


    // === テーブル検索 ===
    const findTargetTable = () => {
        if (cachedTable && document.contains(cachedTable)) return cachedTable;
        const tables = document.getElementsByTagName('table');
        for (let i = 0; i < tables.length; i++) {
            const ths = tables[i].getElementsByTagName('th');
            for (let j = 0; j < ths.length; j++) {
                if (ths[j].textContent.includes('受諾済み') || ths[j].title.includes('承諾され')) {
                    cachedTable = tables[i];
                    return cachedTable;
                }
            }
        }
        return null;
    };


    // === 行データ抽出 ===
    const extractRowData = (row) => {
        let timeText = null, requiredValue = 0, acceptedValue = 0;
        const cells = row.cells;
        for (let i = 0; i < cells.length; i++) {
            const cell = cells[i];
            const db = cell.dataset.bind || cell.getAttribute('data-bind') || '';
            const text = cell.textContent.trim();
            if (!timeText && TIME_REGEX.test(text)) timeText = text;
            if (!requiredValue && REQUIRED_REGEX.test(db) && !db.includes('total') && DIGIT_REGEX.test(text)) requiredValue = +text;
            if (!acceptedValue && SCHEDULED_REGEX.test(db) && DIGIT_REGEX.test(text)) acceptedValue = +text;
        }
        if (!requiredValue || !acceptedValue) {
            const els = row.querySelectorAll('[data-bind]');
            for (let i = 0; i < els.length; i++) {
                const db = els[i].getAttribute('data-bind') || '';
                const text = els[i].textContent.trim();
                if (!requiredValue && REQUIRED_REGEX.test(db) && !db.includes('total') && DIGIT_REGEX.test(text)) requiredValue = +text;
                if (!acceptedValue && SCHEDULED_REGEX.test(db) && DIGIT_REGEX.test(text)) acceptedValue = +text;
            }
        }
        if (!timeText) {
            for (let i = 0; i < cells.length; i++) {
                const text = cells[i].textContent.trim();
                if (TIME_STRICT_REGEX.test(text)) { timeText = text; break; }
            }
        }
        return { timeText, requiredValue, acceptedValue };
    };


    // =====================================================
    // === テーブルからデータ読み取り → 全UI再構築
    // =====================================================
    const calculateAndDisplay = () => {
        if (isCalculating) return;
        isCalculating = true;
        try {
            const targetTable = findTargetTable();
            if (!targetTable) { showError('対象テーブルが見つかりません'); return; }
            const tbody = targetTable.tBodies[0];
            if (!tbody) return;


            const rows = tbody.rows;
            const rowCount = rows.length;
            const timeDataList = [];
            const ssdData = {};
            SSD_LIST.forEach(k => ssdData[k] = { required: 0, accepted: 0 });


            let totalAccepted = 0, totalRequired = 0, proDPAccepted = 0, proDPRequired = 0;


            const serviceTypeIndices = [];
            for (let i = 0; i < rowCount; i++) {
                const span = rows[i].querySelector('span.expandable');
                if (span) {
                    const name = span.textContent.trim();
                    if (isServiceTypeName(name))
                        serviceTypeIndices.push({ index: i, name, isProDP: name.includes('ProDP') });
                }
            }


            let currentServiceType = null, serviceTypeIdx = 0;
            for (let i = 0; i < rowCount; i++) {
                while (serviceTypeIdx < serviceTypeIndices.length && serviceTypeIndices[serviceTypeIdx].index <= i)
                    currentServiceType = serviceTypeIndices[serviceTypeIdx++];


                const { timeText, requiredValue, acceptedValue } = extractRowData(rows[i]);
                if (!timeText) continue;


                if (currentServiceType?.isProDP) { proDPAccepted += acceptedValue; proDPRequired += requiredValue; continue; }


                const serviceTypeName = currentServiceType?.name || '-';
                const timeMinutes = parseTime(timeText);
                timeDataList.push({ time: timeText, timeMinutes, serviceType: serviceTypeName, required: requiredValue, accepted: acceptedValue });


                const ssdGroup = getSSDGroup(timeMinutes);
                if (ssdGroup) { ssdData[ssdGroup].required += requiredValue; ssdData[ssdGroup].accepted += acceptedValue; }
                totalAccepted += acceptedValue;
                totalRequired += requiredValue;
            }


            currentSSDData      = ssdData;
            currentTimeDataList = timeDataList;
            currentTotals       = { accepted: totalAccepted, required: totalRequired, proDPAccepted, proDPRequired };


            renderUI();
        } finally {
            isCalculating = false;
        }
    };


    // =====================================================
    // === 設定変更時：Cycleセルだけ in-place 更新
    // =====================================================
    const refreshCycleValues = () => {
        if (!currentSSDData) return;
        if (!document.getElementById('dsp-main-box')) { renderUI(); return; }


        const multipliers = getSSDMultipliers();
        const adjustments = getSSDAdjustments();
        const overrides   = getC1C3Overrides();
        const pct         = getSoftPct();
        const subtract    = getC1C3Subtract();


        for (var ci = 0; ci < SSD_LIST.length; ci++) {
            var sk  = SSD_LIST[ci];
            var d   = currentSSDData[sk];
            var m   = multipliers[sk] || 1;
            var adj = adjustments[sk] || 0;
            var baseAcc  = getBaseAccepted(sk, d, overrides);
            var baseSoft = (baseAcc + adj) * m;
            var soft     = applySubtract(sk, applyPct(baseSoft, pct), overrides, subtract);
            var hard     = Math.round(soft * 1.1);


            var accEl  = document.getElementById('cell-acc-'  + sk);
            var softEl = document.getElementById('cell-soft-' + sk);
            var hardEl = document.getElementById('cell-hard-' + sk);
            if (accEl)  accEl.innerHTML    = d.accepted + formatAdjustment(adj);
            if (softEl) softEl.textContent = soft;
            if (hardEl) hardEl.textContent = hard;
        }


        // pctラベル更新
        var pctEl = document.getElementById('pct-label');
        if (pctEl) {
            pctEl.textContent = pct === 0 ? '' : ' ' + (pct > 0 ? '+' : '') + pct + '%';
            pctEl.style.color = pct > 0 ? '#2196F3' : '#f44336';
        }


        // トグルボタン外観更新
        updateToggleBtn(subtract);
    };


    // トグルボタン外観を同期
    const updateToggleBtn = (subtract) => {
        var btn = document.getElementById('subtract-toggle-btn');
        if (!btn) return;
        btn.textContent  = subtract ? 'ON' : 'OFF';
        btn.style.background = subtract ? '#9C27B0' : '#e0e0e0';
        btn.style.color      = subtract ? 'white'   : '#555';
    };


    // === debounce ===
    const debouncedCalculate = (delay = 200) => {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(calculateAndDisplay, delay);
    };


    // === Observer ===
    const startObserver = () => {
        observer?.disconnect();
        observer = new MutationObserver((mutations) => {
            for (let i = 0; i < mutations.length; i++) {
                const m = mutations[i];
                const target = m.target;
                if (m.type === 'childList') {
                    if (target.tagName === 'TBODY' || target.tagName === 'TABLE' || target.closest?.('table')) {
                        debouncedCalculate(300); return;
                    }
                }
                if (m.type === 'attributes' && m.attributeName === 'class' && target.tagName === 'LI') {
                    debouncedCalculate(400); return;
                }
            }
        });
        observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class'] });
    };


    // === Excel ===
    const downloadExcel = () => {
        if (!currentSSDData) { alert('データがありません'); return; }
        const multipliers = getSSDMultipliers();
        const adjustments = getSSDAdjustments();
        const overrides   = getC1C3Overrides();
        const pct         = getSoftPct();
        const subtract    = getC1C3Subtract();
        const wsData = [['Station', 'Cycle', 'Soft Caps', 'Hard Caps']];
        for (const ssd of SSD_LIST) {
            const d      = currentSSDData[ssd];
            const m      = multipliers[ssd] || 1;
            const adj    = adjustments[ssd] || 0;
            const baseAcc  = getBaseAccepted(ssd, d, overrides);
            const baseSoft = (baseAcc + adj) * m;
            const soft     = applySubtract(ssd, applyPct(baseSoft, pct), overrides, subtract);
            wsData.push(['VFK1', ssd, soft, Math.round(soft * 1.1)]);
        }
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const fileName = 'SSD_caps_' + getDateForFileName() + '.xlsx';
        XLSX.writeFile(wb, fileName);
        showDownloadNotification(fileName);
    };


    const showDownloadNotification = (fileName) => {
        const n = document.createElement('div');
        n.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);padding:20px 30px;background:#4CAF50;color:white;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,0.3);z-index:10001;font-size:16px;text-align:center;';
        n.innerHTML = '<div style="font-size:24px;margin-bottom:10px;">&#10003;</div><div style="font-weight:bold;">' + fileName + '</div>';
        document.body.appendChild(n);
        setTimeout(() => { n.style.opacity = '0'; n.style.transition = 'opacity 0.3s'; setTimeout(() => n.remove(), 300); }, 1500);
    };


    const showError = (message) => {
        document.getElementById('dsp-main-box')?.remove();
        const box = document.createElement('div');
        box.id = 'dsp-main-box';
        box.style.cssText = 'position:fixed;bottom:20px;right:5px;padding:15px 20px;background:#fff;border:2px solid #f44336;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);z-index:9999;font-size:14px;width:400px;';
        box.innerHTML = '<div style="font-weight:bold;color:#f44336;">エラー</div><div style="margin-top:5px;">' + message + '</div>';
        document.body.appendChild(box);
    };


    const formatAdjustment = (adj) => {
        if (adj === 0) return '';
        return '<span style="color:#f44336;font-weight:bold;margin-left:2px;">' + (adj > 0 ? '+' : '') + adj + '</span>';
    };


    const sectionTitle = (text, color) =>
        '<div style="font-weight:bold;color:' + color + ';padding:4px 2px;margin-bottom:6px;font-size:13px;border-bottom:1px solid #eee;">' + text + '</div>';


    // =====================================================
    // === 全UI描画
    // =====================================================
    const renderUI = () => {
        document.getElementById('dsp-main-box')?.remove();


        const multipliers = getSSDMultipliers();
        const adjustments = getSSDAdjustments();
        const overrides   = getC1C3Overrides();
        const pct         = getSoftPct();
        const subtract    = getC1C3Subtract();
        const { accepted: totalAccepted, required: totalRequired, proDPAccepted, proDPRequired } = currentTotals;


        // ---- 左パネル：SPR入力行 ----
        var sprInputRows = '';
        for (var si = 0; si < SPR_LIST.length; si++) {
            var sk2 = SPR_LIST[si];
            sprInputRows +=
                '<div style="display:grid;grid-template-columns:70px 1fr 1fr;gap:6px;align-items:center;margin:4px 0;">' +
                '<label style="font-size:11px;color:#555;font-weight:bold;">' + sk2 + '</label>' +
                '<input type="number" id="mult-' + sk2 + '" value="' + (multipliers[sk2] || 1) + '" step="1" min="0" max="100"' +
                ' style="width:100%;padding:4px;border:1px solid #ccc;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;" />' +
                '<input type="number" id="adj-' + sk2 + '" value="' + (adjustments[sk2] || 0) + '" step="1" min="-999" max="999"' +
                ' style="width:100%;padding:4px;border:1px solid #ffcdd2;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#f44336;font-weight:bold;" />' +
                '</div>';
        }


        // トグルボタン初期スタイル
        var toggleBg    = subtract ? '#9C27B0' : '#e0e0e0';
        var toggleColor = subtract ? 'white'   : '#555';
        var toggleText  = subtract ? 'ON'      : 'OFF';


        // ---- 中パネル：Cycle別行 ----
        var ssdRowsHtml = '';
        for (var ci = 0; ci < SSD_LIST.length; ci++) {
            var sk  = SSD_LIST[ci];
            var d   = currentSSDData[sk];
            var m   = multipliers[sk] || 1;
            var adj = adjustments[sk] || 0;
            var baseAcc  = getBaseAccepted(sk, d, overrides);
            var baseSoft = (baseAcc + adj) * m;
            var soft     = applySubtract(sk, applyPct(baseSoft, pct), overrides, subtract);
            var hard     = Math.round(soft * 1.1);
            ssdRowsHtml +=
                '<div style="display:grid;grid-template-columns:80px 60px 70px 70px 70px;gap:6px;margin:3px 0;padding:8px;background:#e3f2fd;border-radius:3px;align-items:center;">' +
                '<span style="font-weight:bold;font-size:11px;">' + sk + '</span>' +
                '<span style="color:#FF9800;font-weight:bold;text-align:center;">' + d.required + '</span>' +
                '<span id="cell-acc-'  + sk + '" style="color:#4CAF50;font-weight:bold;text-align:center;">' + d.accepted + formatAdjustment(adj) + '</span>' +
                '<span id="cell-soft-' + sk + '" style="color:#2196F3;font-weight:bold;text-align:center;">' + soft + '</span>' +
                '<span id="cell-hard-' + sk + '" style="color:#9C27B0;font-weight:bold;text-align:center;">' + hard + '</span>' +
                '</div>';
        }


        // ---- 右パネル：開始時刻別 ----
        var sortedTimeData = currentTimeDataList.slice().sort(function(a, b) {
            return a.timeMinutes !== b.timeMinutes ? a.timeMinutes - b.timeMinutes : a.serviceType.localeCompare(b.serviceType);
        });
        var timeRowsHtml = sortedTimeData.length === 0 ? '<div style="color:#666;padding:10px;">データなし</div>' : '';
        for (var ti = 0; ti < sortedTimeData.length; ti++) {
            var td = sortedTimeData[ti];
            timeRowsHtml +=
                '<div style="display:grid;grid-template-columns:90px 130px 50px 50px;gap:8px;margin:3px 0;padding:6px 8px;background:#f5f5f5;border-radius:3px;align-items:center;">' +
                '<span style="font-weight:bold;font-size:11px;">' + td.time + '</span>' +
                '<span style="font-size:10px;color:#666;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + td.serviceType + '">' + getShortServiceType(td.serviceType) + '</span>' +
                '<span style="color:#FF9800;font-weight:bold;text-align:center;">' + td.required + '</span>' +
                '<span style="color:#4CAF50;font-weight:bold;text-align:center;">' + td.accepted + '</span>' +
                '</div>';
        }


        // ---- サマリー ----
        var diff = totalAccepted - totalRequired;
        var diffColor = diff >= 0 ? '#4CAF50' : '#f44336';
        var proDPHtml = (proDPAccepted || proDPRequired) ?
            '<div style="margin:5px 0;display:flex;justify-content:space-between;font-size:11px;color:#9e9e9e;padding-top:5px;border-top:1px dashed #e0e0e0;"><span>※ProDP除外:</span><span>受諾 ' + proDPAccepted + '</span></div>' : '';
        var pctLabelText  = pct === 0 ? '' : ' ' + (pct > 0 ? '+' : '') + pct + '%';
        var pctLabelColor = pct > 0 ? '#2196F3' : '#f44336';


        // ---- DOM組み立て ----
        var box = document.createElement('div');
        box.id = 'dsp-main-box';
        box.style.cssText = 'position:fixed;bottom:20px;right:5px;display:flex;align-items:flex-start;background:#fff;border:2px solid #4CAF50;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);z-index:9999;font-size:12px;';


        // 左パネル
        var leftPanel = document.createElement('div');
        leftPanel.style.cssText = 'width:270px;min-width:270px;padding:6px 8px;border-right:2px solid #e3f2fd;max-height:600px;overflow-y:auto;overflow-x:hidden;box-sizing:border-box;';
        leftPanel.innerHTML =
            '<div style="font-weight:bold;color:#2196F3;margin-bottom:8px;font-size:18px;">SPR設定</div>' +
            '<div style="display:grid;grid-template-columns:70px 1fr 1fr;gap:6px;margin-bottom:6px;padding:0 2px;font-size:14px;color:#999;font-weight:bold;">' +
            '<span></span><span style="text-align:center;">SPR</span><span style="text-align:center;color:#f44336;">Buffer</span>' +
            '</div>' +
            sprInputRows +
            '<div style="margin-top:8px;">' +
            '<button id="adj-reset-btn" style="width:100%;padding:4px;background:#ff8a80;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;font-weight:bold;">Bufferリセット</button>' +
            '</div>' +
            // C1/C3 受諾設定
            '<div style="margin-top:12px;padding-top:10px;border-top:2px solid #e3f2fd;">' +
            '<div style="display:grid;grid-template-columns:70px 1fr;gap:6px;align-items:center;margin:4px 0;">' +
            '<label style="font-size:11px;color:#555;font-weight:bold;">SSD_C1</label>' +
            '<input type="number" id="c1c3-SSD_C1" value="' + (overrides['SSD_C1'] || 0) + '" step="1" min="0" max="9999"' +
            ' style="width:100%;padding:4px;border:2px solid #CE93D8;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#9C27B0;font-weight:bold;" />' +
            '</div>' +
            '<div style="display:grid;grid-template-columns:70px 1fr;gap:6px;align-items:center;margin:4px 0;">' +
            '<label style="font-size:11px;color:#555;font-weight:bold;">SSD_C3</label>' +
            '<input type="number" id="c1c3-SSD_C3" value="' + (overrides['SSD_C3'] || 0) + '" step="1" min="0" max="9999"' +
            ' style="width:100%;padding:4px;border:2px solid #CE93D8;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#9C27B0;font-weight:bold;" />' +
            '</div>' +
            // 1B/3B 減算トグル
            '<div style="margin-top:8px;display:flex;align-items:center;gap:8px;">' +
            '<button id="subtract-toggle-btn"' +
            ' style="padding:4px 14px;background:' + toggleBg + ';color:' + toggleColor + ';border:none;border-radius:20px;cursor:pointer;font-size:12px;font-weight:bold;transition:all 0.2s;">' +
            toggleText +
            '</button>' +
            '<span style="font-size:11px;color:#555;">1B / 3B から減算</span>' +
            '</div>' +
            '</div>' +
            // 全Soft調整
            '<div style="margin-top:12px;padding-top:10px;border-top:2px solid #e3f2fd;">' +
            '<div style="font-weight:bold;color:#2196F3;font-size:13px;margin-bottom:6px;">全Soft調整</div>' +
            '<div style="display:flex;align-items:center;gap:6px;">' +
            '<input type="number" id="soft-pct-input" value="' + pct + '" step="1" min="-100" max="200"' +
            ' style="width:80px;padding:5px;border:2px solid #2196F3;border-radius:4px;text-align:center;font-size:14px;font-weight:bold;box-sizing:border-box;" />' +
            '<span style="font-size:14px;font-weight:bold;color:#555;">%</span>' +
            '<button id="soft-pct-reset" style="padding:4px 8px;background:#e0e0e0;color:#555;border:none;border-radius:4px;cursor:pointer;font-size:11px;">リセット</button>' +
            '</div>' +
            '<div style="font-size:10px;color:#999;margin-top:4px;">Soft = (受諾+Buffer) × SPR × (1+%/100)</div>' +
            '</div>';


        // 中パネル
        var midPanel = document.createElement('div');
        midPanel.style.cssText = 'width:390px;min-width:390px;padding:6px 8px;border-right:2px solid #e3f2fd;max-height:600px;overflow-y:auto;overflow-x:hidden;box-sizing:border-box;';
        midPanel.innerHTML =
            '<div style="background:#e8f5e9;padding:10px;border-radius:5px;margin-bottom:10px;">' +
            '<div style="margin:5px 0;display:flex;justify-content:space-between;"><span>必須合計:</span><strong style="color:#F57C00;">' + totalRequired + '</strong></div>' +
            '<div style="margin:5px 0;display:flex;justify-content:space-between;"><span>受諾済み:</span><strong style="color:#2e7d32;">' + totalAccepted + '</strong></div>' +
            '<div style="margin:5px 0;display:flex;justify-content:space-between;padding-top:5px;border-top:1px solid #c8e6c9;"><span>Gap:</span><strong style="color:' + diffColor + ';">' + (diff >= 0 ? '+' : '') + diff + '</strong></div>' +
            proDPHtml +
            '</div>' +
            '<div style="display:grid;grid-template-columns:80px 60px 70px 70px 70px;gap:6px;margin-bottom:5px;padding:3px 5px;font-weight:bold;color:#666;font-size:10px;">' +
            '<span>Cycle</span><span style="text-align:center;">必須</span>' +
            '<span style="text-align:center;">受諾</span>' +
            '<span style="text-align:center;">Soft<span id="pct-label" style="font-size:9px;color:' + pctLabelColor + ';">' + pctLabelText + '</span></span>' +
            '<span style="text-align:center;">Hard</span>' +
            '</div>' +
            ssdRowsHtml +
            '<div style="padding-top:10px;border-top:1px solid #ddd;">' +
            '<button id="dl-btn" style="width:100%;padding:5px;background:#4CAF50;color:white;border:none;border-radius:5px;cursor:pointer;font-size:12px;font-weight:bold;">Excel download</button>' +
            '</div>';


        // 右パネル
        var rightPanel = document.createElement('div');
        rightPanel.style.cssText = 'width:350px;min-width:350px;padding:6px 8px;overflow-y:auto;overflow-x:hidden;box-sizing:border-box;';
        rightPanel.innerHTML =
            sectionTitle('開始時刻別', '#4CAF50') +
            '<div style="display:grid;grid-template-columns:90px 130px 50px 50px;gap:8px;margin-bottom:5px;padding:3px 5px;font-weight:bold;color:#666;font-size:10px;">' +
            '<span>開始時刻</span><span>サービスタイプ</span><span style="text-align:center;">必須</span><span style="text-align:center;">受諾</span>' +
            '</div>' +
            timeRowsHtml;


        box.appendChild(leftPanel);
        box.appendChild(midPanel);
        box.appendChild(rightPanel);
        document.body.appendChild(box);


        // 右パネルの高さを中パネルに合わせる
        setTimeout(function() {
            var h = midPanel.offsetHeight;
            if (h > 0) { rightPanel.style.height = h + 'px'; rightPanel.style.maxHeight = h + 'px'; }
        }, 0);


        // ---- イベント登録 ----


        // SPR / Buffer
        SPR_LIST.forEach(function(ssd) {
            var multEl = document.getElementById('mult-' + ssd);
            var adjEl  = document.getElementById('adj-'  + ssd);
            if (multEl) multEl.addEventListener('change', function() {
                var mv = getSSDMultipliers(); mv[ssd] = +this.value || 1; saveSSDMultipliers(mv); refreshCycleValues();
            });
            if (adjEl) adjEl.addEventListener('change', function() {
                var av = getSSDAdjustments(); av[ssd] = +this.value || 0; saveSSDAdjustments(av); refreshCycleValues();
            });
        });


        // Buffer リセット
        document.getElementById('adj-reset-btn')?.addEventListener('click', function() {
            var reset = Object.assign({}, SSD_ADJUSTMENT_DEFAULTS);
            saveSSDAdjustments(reset);
            SSD_LIST.forEach(function(ssd) {
                var el = document.getElementById('adj-' + ssd);
                if (el) el.value = 0;
            });
            refreshCycleValues();
        });


        // C1/C3 受諾オーバーライド
        ['SSD_C1', 'SSD_C3'].forEach(function(key) {
            var el = document.getElementById('c1c3-' + key);
            if (el) el.addEventListener('change', function() {
                var ov = getC1C3Overrides();
                ov[key] = +this.value || 0;
                saveC1C3Overrides(ov);
                refreshCycleValues();
            });
        });


        // 1B/3B 減算トグル
        document.getElementById('subtract-toggle-btn')?.addEventListener('click', function() {
            var next = !getC1C3Subtract();
            saveC1C3Subtract(next);
            refreshCycleValues();
        });


        // 全Soft調整%
        document.getElementById('soft-pct-input')?.addEventListener('change', function() {
            saveSoftPct(Number(this.value) || 0); refreshCycleValues();
        });
        document.getElementById('soft-pct-reset')?.addEventListener('click', function() {
            saveSoftPct(0);
            var el = document.getElementById('soft-pct-input');
            if (el) el.value = 0;
            refreshCycleValues();
        });


        // Excel download
        document.getElementById('dl-btn')?.addEventListener('click', downloadExcel);
    };


    // === 初期化 ===
    const init = () => {
        setTimeout(() => {
            calculateAndDisplay();
            startObserver();
            console.log('[DSP Counter v9.6] 起動完了');
        }, 1000);
    };


    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();

