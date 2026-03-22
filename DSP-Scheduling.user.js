// ==UserScript==
// @name         Schedulingで受諾済みを計算
// @namespace    https://github.com/yuyna-amazon/DSP-Scheduling
// @version      8.2
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
    let currentMultipliers = null;
    let isCalculating = false;
    let debounceTimer = null;
    let lastCalculatedDate = null;
    let observer = null;
    let cachedTable = null;
    let lastTableHash = '';

    // === プリコンパイル正規表現 ===
    const TIME_REGEX = /(\d+):(\d+)\s*(午前|午後|am|pm)/i;
    const TIME_STRICT_REGEX = /^\d{1,2}:\d{2}\s*(午前|午後|am|pm)$/i;
    const DIGIT_REGEX = /^\d+$/;
    const DATE_REGEX = /(\d+)-(\d+)月/;
    const REQUIRED_REGEX = /text:\s*required\(\)/;
    const SCHEDULED_REGEX = /text:\s*scheduled\(\)/;

    // === 定数 ===
    const SSD_DEFAULTS = {
        'SSD_1': 1, 'SSD_1_B': 1, 'SSD_2': 1,
        'SSD_3': 1, 'SSD_3_B': 1, 'SSD_4': 1
    };

    const SSD_TIME_RANGES = [
        { key: 'SSD_1', min: 0, max: 420 },
        { key: 'SSD_C1', min: 420, max: 420 },
        { key: 'SSD_1_B', min: 420, max: 600 },
        { key: 'SSD_2', min: 600, max: 840 },
        { key: 'SSD_3', min: 840, max: 1020 },
        { key: 'SSD_C3', min: 1020, max: 1020 },
        { key: 'SSD_3_B', min: 1020, max: 1200 },
        { key: 'SSD_4', min: 1200, max: Infinity }
    ];

    // === ユーティリティ関数 ===
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
            const range = SSD_TIME_RANGES[i];
            if (timeMinutes >= range.min && timeMinutes < range.max) {
                return range.key;
            }
        }
        return null;
    };

    const getShortServiceType = (name) => {
        if (!name) return '-';
        const parts = name.split('_');
        if (parts.length >= 2) {
            return parts.slice(1).join('_');
        }
        return name;
    };

    // === ストレージ関数 ===
    const getStorage = (key, defaultVal) => {
        try {
            const stored = localStorage.getItem(key);
            if (stored === null) return defaultVal;
            if (typeof defaultVal === 'boolean') return stored === 'true';
            if (typeof defaultVal === 'object') return { ...defaultVal, ...JSON.parse(stored) };
            return stored;
        } catch { return defaultVal; }
    };

    const setStorage = (key, val) => {
        localStorage.setItem(key, typeof val === 'object' ? JSON.stringify(val) : String(val));
    };

    const getSSDMultipliers = () => getStorage('dsp-ssd-multipliers', SSD_DEFAULTS);
    const saveSSDMultipliers = (m) => setStorage('dsp-ssd-multipliers', m);
    const getSPRBoxVisible = () => getStorage('dsp-spr-box-visible', true);
    const saveSPRBoxVisible = (v) => setStorage('dsp-spr-box-visible', v);
    const getSSDSectionVisible = () => getStorage('dsp-ssd-section-visible', false);
    const saveSSDSectionVisible = (v) => setStorage('dsp-ssd-section-visible', v);
    const getTimeSectionVisible = () => getStorage('dsp-time-section-visible', false);
    const saveTimeSectionVisible = (v) => setStorage('dsp-time-section-visible', v);

    // === 日付取得 ===
    const getCurrentSelectedDate = () => {
        const span = document.querySelector('li.selected .dateText');
        return span?.textContent.trim() || null;
    };

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
        if (cachedTable && document.contains(cachedTable)) {
            return cachedTable;
        }
        const tables = document.getElementsByTagName('table');
        for (let i = 0; i < tables.length; i++) {
            const ths = tables[i].getElementsByTagName('th');
            for (let j = 0; j < ths.length; j++) {
                const text = ths[j].textContent;
                if (text.includes('受諾済み') || ths[j].title.includes('承諾され')) {
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
            const dataBind = cell.dataset.bind || cell.getAttribute('data-bind') || '';
            const text = cell.textContent.trim();

            if (!timeText && TIME_REGEX.test(text)) {
                timeText = text;
            }
            if (!requiredValue && REQUIRED_REGEX.test(dataBind) && !dataBind.includes('total') && DIGIT_REGEX.test(text)) {
                requiredValue = +text;
            }
            if (!acceptedValue && SCHEDULED_REGEX.test(dataBind) && DIGIT_REGEX.test(text)) {
                acceptedValue = +text;
            }
        }

        if (!requiredValue || !acceptedValue) {
            const elements = row.querySelectorAll('[data-bind]');
            for (let i = 0; i < elements.length; i++) {
                const el = elements[i];
                const dataBind = el.getAttribute('data-bind') || '';
                const text = el.textContent.trim();
                if (!requiredValue && REQUIRED_REGEX.test(dataBind) && !dataBind.includes('total') && DIGIT_REGEX.test(text)) {
                    requiredValue = +text;
                }
                if (!acceptedValue && SCHEDULED_REGEX.test(dataBind) && DIGIT_REGEX.test(text)) {
                    acceptedValue = +text;
                }
            }
        }

        if (!timeText) {
            for (let i = 0; i < cells.length; i++) {
                const text = cells[i].textContent.trim();
                if (TIME_STRICT_REGEX.test(text)) {
                    timeText = text;
                    break;
                }
            }
        }

        return { timeText, requiredValue, acceptedValue };
    };

    // === メイン計算関数 ===
    const calculateAndDisplay = () => {
        if (isCalculating) return;
        isCalculating = true;

        try {
            const targetTable = findTargetTable();
            if (!targetTable) {
                showError('対象テーブルが見つかりません');
                return;
            }

            const tbody = targetTable.tBodies[0];
            if (!tbody) return;

            const rows = tbody.rows;
            const rowCount = rows.length;

            const timeDataList = [];
            const ssdData = {
                'SSD_1': { required: 0, accepted: 0 },
                'SSD_C1': { required: 0, accepted: 0 },
                'SSD_1_B': { required: 0, accepted: 0 },
                'SSD_2': { required: 0, accepted: 0 },
                'SSD_3': { required: 0, accepted: 0 },
                'SSD_C3': { required: 0, accepted: 0 },
                'SSD_3_B': { required: 0, accepted: 0 },
                'SSD_4': { required: 0, accepted: 0 }
            };

            let totalAccepted = 0, totalRequired = 0;
            let proDPAccepted = 0, proDPRequired = 0;

            const serviceTypeIndices = [];

            for (let i = 0; i < rowCount; i++) {
                const span = rows[i].querySelector('span.expandable');
                if (span) {
                    const name = span.textContent.trim();
                    if (isServiceTypeName(name)) {
                        serviceTypeIndices.push({ index: i, name, isProDP: name.includes('ProDP') });
                    }
                }
            }

            let currentServiceType = null;
            let serviceTypeIdx = 0;

            for (let i = 0; i < rowCount; i++) {
                while (serviceTypeIdx < serviceTypeIndices.length && serviceTypeIndices[serviceTypeIdx].index <= i) {
                    currentServiceType = serviceTypeIndices[serviceTypeIdx];
                    serviceTypeIdx++;
                }

                const { timeText, requiredValue, acceptedValue } = extractRowData(rows[i]);
                if (!timeText) continue;

                if (currentServiceType?.isProDP) {
                    proDPAccepted += acceptedValue;
                    proDPRequired += requiredValue;
                    continue;
                }

                const serviceTypeName = currentServiceType?.name || '-';
                timeDataList.push({
                    time: timeText,
                    timeMinutes: parseTime(timeText),
                    serviceType: serviceTypeName,
                    required: requiredValue,
                    accepted: acceptedValue
                });

                const timeMinutes = parseTime(timeText);
                const ssdGroup = getSSDGroup(timeMinutes);
                if (ssdGroup) {
                    ssdData[ssdGroup].required += requiredValue;
                    ssdData[ssdGroup].accepted += acceptedValue;
                }

                totalAccepted += acceptedValue;
                totalRequired += requiredValue;
            }

            currentSSDData = ssdData;
            currentMultipliers = getSSDMultipliers();
            lastCalculatedDate = getCurrentSelectedDate();

            showResult(timeDataList, ssdData, totalAccepted, totalRequired, proDPAccepted, proDPRequired);

        } finally {
            isCalculating = false;
        }
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
                    if (target.tagName === 'TBODY' || target.tagName === 'TABLE' ||
                        target.closest?.('table')) {
                        debouncedCalculate(300);
                        return;
                    }
                }
                if (m.type === 'attributes' && m.attributeName === 'class' && target.tagName === 'LI') {
                    debouncedCalculate(400);
                    return;
                }
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true,
            attributes: true,
            attributeFilter: ['class']
        });
    };

    // === Excel ダウンロード ===
    const downloadExcel = () => {
        if (!currentSSDData || !currentMultipliers) {
            alert('データがありません');
            return;
        }

        const wsData = [['Station', 'Cycle', 'Soft Caps', 'Hard Caps']];
        const ssdMapping = [
            { key: 'SSD_1', cycle: 'SSD_1' },
            { key: 'SSD_C1', cycle: 'SSD_C1' },
            { key: 'SSD_1_B', cycle: 'SSD_1_B' },
            { key: 'SSD_2', cycle: 'SSD_2' },
            { key: 'SSD_3', cycle: 'SSD_3' },
            { key: 'SSD_C3', cycle: 'SSD_C3' },
            { key: 'SSD_3_B', cycle: 'SSD_3_B' },
            { key: 'SSD_4', cycle: 'SSD_4' }
        ];

        for (const ssd of ssdMapping) {
            const data = currentSSDData[ssd.key];
            const multiplier = currentMultipliers[ssd.key] || 1;
            const softCap = data.accepted * multiplier;
            wsData.push(['VFK1', ssd.cycle, softCap, Math.round(softCap * 1.1)]);
        }

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

        const fileName = `SSD_caps_${getDateForFileName()}.xlsx`;
        XLSX.writeFile(wb, fileName);
        showDownloadNotification(fileName);
    };

    const showDownloadNotification = (fileName) => {
        const notification = document.createElement('div');
        notification.style.cssText = `
            position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);
            padding:20px 30px;background:#4CAF50;color:white;border-radius:8px;
            box-shadow:0 4px 20px rgba(0,0,0,0.3);z-index:10001;font-size:16px;text-align:center;
        `;
        notification.innerHTML = `<div style="font-size:24px;margin-bottom:10px;">✓</div>
            <div style="font-weight:bold;">${fileName}</div>`;
        document.body.appendChild(notification);
        setTimeout(() => {
            notification.style.opacity = '0';
            notification.style.transition = 'opacity 0.3s';
            setTimeout(() => notification.remove(), 300);
        }, 1500);
    };

    const showError = (message) => {
        const existing = document.getElementById('dsp-accepted-counter-box');
        existing?.remove();

        const box = document.createElement('div');
        box.id = 'dsp-accepted-counter-box';
        box.style.cssText = `
            position:fixed;bottom:20px;right:20px;padding:15px 20px;background:#fff;
            border:2px solid #f44336;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);
            z-index:9999;font-size:14px;width:400px;
        `;
        box.innerHTML = `<div style="font-weight:bold;color:#f44336;">エラー</div>
            <div style="margin-top:5px;">${message}</div>`;
        document.body.appendChild(box);
    };

    // === 結果表示 ===
    const showResult = (timeDataList, ssdData, totalAccepted, totalRequired, proDPAccepted, proDPRequired) => {
        document.getElementById('dsp-accepted-counter-box')?.remove();

        const box = document.createElement('div');
        box.id = 'dsp-accepted-counter-box';
        box.style.cssText = `
            position:fixed;bottom:20px;right:5px;padding:5px;background:#fff;
            border:2px solid #4CAF50;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);
            z-index:9999;font-size:12px;width:400px;max-height:600px;overflow-y:auto;
        `;

        const multipliers = getSSDMultipliers();
        const ssdOrder = ['SSD_1', 'SSD_C1', 'SSD_1_B', 'SSD_2', 'SSD_3', 'SSD_C3', 'SSD_3_B', 'SSD_4'];
        const ssdVisible = getSSDSectionVisible();
        const timeVisible = getTimeSectionVisible();

        const ssdListHtml = ssdOrder.map(ssd => {
            const d = ssdData[ssd];
            if (!d.required && !d.accepted) return '';
            const m = multipliers[ssd] || 1;
            const soft = d.accepted * m;
            const hard = Math.round(soft * 1.1);
            return `<div style="display:grid;grid-template-columns:80px 60px 60px 70px 70px;gap:8px;margin:3px 0;padding:8px;background:#e3f2fd;border-radius:3px;align-items:center;">
                <span style="font-weight:bold;">${ssd}</span>
                <span style="color:#FF9800;font-weight:bold;text-align:center;">${d.required}</span>
                <span style="color:#4CAF50;font-weight:bold;text-align:center;">${d.accepted}</span>
                <span style="color:#2196F3;font-weight:bold;text-align:center;">${soft}</span>
                <span style="color:#9C27B0;font-weight:bold;text-align:center;">${hard}</span>
            </div>`;
        }).join('');

        const sortedTimeData = timeDataList.sort((a, b) => {
            if (a.timeMinutes !== b.timeMinutes) return a.timeMinutes - b.timeMinutes;
            return a.serviceType.localeCompare(b.serviceType);
        });

        const timeListHtml = sortedTimeData.length ? sortedTimeData.map(d =>
            `<div style="display:grid;grid-template-columns:90px 130px 50px 50px;gap:8px;margin:3px 0;padding:6px 8px;background:#f5f5f5;border-radius:3px;align-items:center;">
                <span style="font-weight:bold;font-size:11px;">${d.time}</span>
                <span style="font-size:10px;color:#666;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${d.serviceType}">${getShortServiceType(d.serviceType)}</span>
                <span style="color:#FF9800;font-weight:bold;text-align:center;">${d.required}</span>
                <span style="color:#4CAF50;font-weight:bold;text-align:center;">${d.accepted}</span>
            </div>`
        ).join('') : '<div style="color:#666;padding:10px;">データなし</div>';

        const diff = totalAccepted - totalRequired;
        const diffColor = diff >= 0 ? '#4CAF50' : '#f44336';
        const proDPHtml = (proDPAccepted || proDPRequired) ?
            `<div style="margin:5px 0;display:flex;justify-content:space-between;font-size:11px;color:#9e9e9e;padding-top:5px;border-top:1px dashed #e0e0e0;">
                <span>※ProDP除外:</span><span>受諾 ${proDPAccepted}</span>
            </div>` : '';

        box.innerHTML = `
            <div style="background:#e8f5e9;padding:10px;border-radius:5px;margin-bottom:10px;">
                <div style="margin:5px 0;display:flex;justify-content:space-between;">
                    <span>必須合計:</span><strong style="color:#F57C00;">${totalRequired}</strong>
                </div>
                <div style="margin:5px 0;display:flex;justify-content:space-between;">
                    <span>受諾済み:</span><strong style="color:#2e7d32;">${totalAccepted}</strong>
                </div>
                <div style="margin:5px 0;display:flex;justify-content:space-between;padding-top:5px;border-top:1px solid #c8e6c9;">
                    <span>Gap:</span><strong style="color:${diffColor};">${diff >= 0 ? '+' : ''}${diff}</strong>
                </div>
                ${proDPHtml}
            </div>
            <div id="ssd-header" style="font-weight:bold;margin-bottom:8px;color:#2196F3;cursor:pointer;padding:5px;user-select:none;display:flex;align-items:center;">
                <span id="ssd-icon">${ssdVisible ? '▼' : '▶'}</span>&nbsp;Cycle別
            </div>
            <div id="ssd-content" style="display:${ssdVisible ? 'block' : 'none'};">
                <div style="display:grid;grid-template-columns:80px 60px 60px 70px 70px;gap:8px;margin-bottom:5px;padding:5px;font-weight:bold;color:#666;font-size:11px;">
                    <span>Cycle</span><span style="text-align:center;">必須</span><span style="text-align:center;">受諾</span>
                    <span style="text-align:center;">Soft</span><span style="text-align:center;">Hard</span>
                </div>
                ${ssdListHtml}
            </div>
            <div id="time-header" style="font-weight:bold;margin-top:15px;margin-bottom:8px;color:#4CAF50;cursor:pointer;padding:5px;user-select:none;display:flex;align-items:center;">
                <span id="time-icon">${timeVisible ? '▼' : '▶'}</span>&nbsp;開始時刻別
            </div>
            <div id="time-content" style="display:${timeVisible ? 'block' : 'none'};">
                <div style="display:grid;grid-template-columns:90px 130px 50px 50px;gap:8px;margin-bottom:5px;padding:5px;font-weight:bold;color:#666;font-size:11px;">
                    <span>開始時刻</span><span>サービスタイプ</span><span style="text-align:center;">必須</span><span style="text-align:center;">受諾</span>
                </div>
                ${timeListHtml}
            </div>
            <div style="margin-top:15px;padding-top:10px;border-top:1px solid #ddd;">
                <button id="dl-btn" style="width:100%;padding:5px;background:#4CAF50;color:white;border:none;border-radius:5px;cursor:pointer;font-size:12px;font-weight:bold;">Excel download</button>
            </div>
        `;

        document.body.appendChild(box);

        setupToggle('ssd-header', 'ssd-content', 'ssd-icon', saveSSDSectionVisible);
        setupToggle('time-header', 'time-content', 'time-icon', saveTimeSectionVisible);
        document.getElementById('dl-btn')?.addEventListener('click', downloadExcel);
    };

    const setupToggle = (headerId, contentId, iconId, saveFunc) => {
        const header = document.getElementById(headerId);
        const content = document.getElementById(contentId);
        const icon = document.getElementById(iconId);
        if (!header || !content || !icon) return;

        header.onclick = () => {
            const visible = content.style.display === 'none';
            content.style.display = visible ? 'block' : 'none';
            icon.textContent = visible ? '▼' : '▶';
            saveFunc(visible);
        };
    };

    // === 乗数設定ボックス ===
    const createMultiplierBox = () => {
        document.getElementById('dsp-multiplier-box')?.remove();

        const box = document.createElement('div');
        box.id = 'dsp-multiplier-box';
        box.style.cssText = `
            position:fixed;top:10px;right:5px;padding:1px 5px;background:#fff;
            border:1px solid #2196F3;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);
            z-index:9999;font-size:12px;width:200px;
        `;

        const multipliers = getSSDMultipliers();
        const visible = getSPRBoxVisible();
        const ssdList = ['SSD_1', 'SSD_C1', 'SSD_1_B', 'SSD_2', 'SSD_3', 'SSD_C3', 'SSD_3_B', 'SSD_4'];

        const inputsHtml = ssdList.map(ssd =>
            `<div style="display:flex;align-items:center;justify-content:space-between;margin:6px 0;">
                <label for="mult-${ssd}" style="font-size:12px;color:#555;">${ssd}:</label>
                <input type="number" id="mult-${ssd}" value="${multipliers[ssd]}" step="1" min="0" max="100"
                    style="width:60px;padding:4px 6px;border:1px solid #ccc;border-radius:4px;text-align:center;font-size:13px;">
            </div>`
        ).join('');

        box.innerHTML = `
            <div id="spr-header" style="font-weight:bold;color:#2196F3;margin-bottom:8px;font-size:14px;cursor:pointer;user-select:none;display:flex;align-items:center;">
                <span id="spr-icon">${visible ? '▼' : '▶'}</span>&nbsp;SPR設定
            </div>
            <div id="spr-content" style="display:${visible ? 'block' : 'none'};">${inputsHtml}</div>
        `;

        document.body.appendChild(box);

        setupToggle('spr-header', 'spr-content', 'spr-icon', saveSPRBoxVisible);

        ssdList.forEach(ssd => {
            document.getElementById(`mult-${ssd}`)?.addEventListener('change', function() {
                const m = getSSDMultipliers();
                m[ssd] = +this.value || 1;
                saveSSDMultipliers(m);
                currentMultipliers = m;
                calculateAndDisplay();
            });
        });
    };

    // === 初期化 ===
    const init = () => {
        setTimeout(() => {
            calculateAndDisplay();
            createMultiplierBox();
            startObserver();
            console.log('[DSP Counter v8.0] 起動完了');
        }, 1000);
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
