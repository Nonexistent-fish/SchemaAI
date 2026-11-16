// ========== Excel 自定义模板专属模块 (跨月图例过滤 + Ctrl不连续选区捕获 + Skills技能引擎 + 连贯色带) ==========
let excelControlsConfig = {};
let isExcelCustomRefreshing = false;
let activePickingTag = null; // 当前处于拾取模式的控件 Tag

// 8 种高对比度清爽调色板
const DISTINCT_PALETTE = [
    { solid: "#2563EB", bg: "rgba(37, 99, 235, 0.22)", border: "#2563EB", text: "#1D4ED8" },
    { solid: "#16A34A", bg: "rgba(22, 163, 74, 0.22)", border: "#16A34A", text: "#15803D" },
    { solid: "#EA580C", bg: "rgba(234, 88, 12, 0.22)", border: "#EA580C", text: "#C2410C" },
    { solid: "#9333EA", bg: "rgba(147, 51, 234, 0.22)", border: "#9333EA", text: "#7E22CE" },
    { solid: "#0891B2", bg: "rgba(8, 145, 178, 0.22)", border: "#0891B2", text: "#0E7490" },
    { solid: "#E11D48", bg: "rgba(225, 29, 72, 0.22)", border: "#E11D48", text: "#BE123C" },
    { solid: "#4F46E5", bg: "rgba(79, 70, 229, 0.22)", border: "#4F46E5", text: "#4338CA" },
    { solid: "#CA8A04", bg: "rgba(202, 138, 4, 0.22)", border: "#CA8A04", text: "#A16207" }
];

// 内置专业教学/办公 Skills (技能预设库)
const SKILL_PRESETS = [
    { id: "custom", name: "自定义提示词", prompt: "" },
    { id: "lesson_progress", name: "【技能】教学进度梯度推演 (循序渐进)", prompt: "根据课题与日期生成前后连贯、由浅入深的授课知识点，严格控制在20字内，不带标点。" },
    { id: "practical_skill", name: "【技能】理实一体拆装实训要点", prompt: "围绕汽车/机械核心部件的规范拆装、结构认知与工位实操生成具体实训动作要点。" },
    { id: "safety_inspection", name: "【技能】安全文明实操与维保规范", prompt: "针对该阶段的实训设备生成规范的防护检查、工具使用注意与安全规范要点。" },
    { id: "concise_summary", name: "【技能】精炼核心概念提取", prompt: "直接提炼最核心的汽车机械构造术语或知识点，10-15字，简明扼要。" }
];

Office.onReady(() => {
    if (window.currentHost !== Office.HostType.Excel) return;

    injectExcelCustomStyles();

    // 注册 Excel 鼠标选区监听：支持 Ctrl 多重选区 (RangeAreas)
    try {
        Excel.run(async (context) => {
            context.workbook.onSelectionChanged.add(async () => {
                if (!activePickingTag) return;

                await Excel.run(async (ctx) => {
                    const selected = ctx.workbook.getSelectedRanges ? ctx.workbook.getSelectedRanges() : ctx.workbook.getSelectedRange();
                    const worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                    selected.load("address");
                    worksheet.load("name");
                    await ctx.sync();

                    const rawAddrs = selected.address.split(",").map(a => a.trim());
                    const formattedAddrs = rawAddrs.map(addr => {
                        const clean = addr.includes("!") ? addr.split("!").pop() : addr;
                        return `${worksheet.name}!${clean}`;
                    }).join(", ");

                    const inputEl = document.querySelector(`.range-input[data-tag="${activePickingTag}"]`);
                    const tagSpan = document.querySelector(`.control-item[data-tag="${activePickingTag}"] .control-tag`);

                    if (inputEl) {
                        inputEl.value = formattedAddrs;
                        inputEl.dispatchEvent(new Event('input'));
                        inputEl.style.borderColor = "var(--success-color)";
                        setTimeout(() => { inputEl.style.borderColor = ""; }, 500);
                    }
                    if (tagSpan) tagSpan.textContent = formattedAddrs || activePickingTag;
                });
            });
            await context.sync();
        });
    } catch (e) {
        console.warn("Excel 选区监听注册失败", e);
    }

    const saveBtn = document.getElementById("save-template-btn");
    if (saveBtn) {
        saveBtn.addEventListener("click", () => {
            if (window.currentHost !== Office.HostType.Excel) return;
            const titleInput = document.getElementById("custom-template-title");
            const descInput = document.getElementById("custom-template-desc");
            if (!titleInput) return;

            const title = titleInput.value.trim();
            const desc = descInput ? descInput.value.trim() : "";

            document.querySelectorAll('.error-flash').forEach(el => el.classList.remove('error-flash'));
            let hasError = false;

            if (!title) {
                const group = titleInput.closest('.input-group');
                if (group) {
                    group.classList.add('error-flash');
                    hasError = true;
                    setTimeout(() => group.classList.remove('error-flash'), 400);
                }
            }

            document.querySelectorAll('.control-item').forEach(item => {
                const ctrlTitleInput = item.querySelector('.control-title-input');
                if (ctrlTitleInput && !ctrlTitleInput.value.trim()) {
                    const headerLeft = ctrlTitleInput.closest('.control-header-left');
                    if (headerLeft) {
                        headerLeft.classList.add('error-flash');
                        hasError = true;
                        setTimeout(() => headerLeft.classList.remove('error-flash'), 400);
                    }
                }
            });

            if (hasError) return;

            CONFIG.customTemplateTitle = title;
            CONFIG.customTemplateDesc = desc;
            const cleanConfig = {};
            const titles = {};
            Object.keys(excelControlsConfig).forEach(tag => {
                cleanConfig[tag] = sanitizeExcelConfig(excelControlsConfig[tag]);
                const itemDiv = document.querySelector(`.control-item[data-tag="${tag}"]`);
                if (itemDiv) {
                    const ti = itemDiv.querySelector('.control-title-input');
                    titles[tag] = ti ? ti.value.trim() : '';
                }
            });
            CONFIG.customTemplateData = { title, desc, controlsConfig: cleanConfig, titles, host: "Excel" };
            saveStorage();

            showStatus("Excel 模板已保存", "success");
            switchView("view-home");
        });
    }

    const refreshBtn = document.getElementById("refresh-custom-btn");
    if (refreshBtn) {
        refreshBtn.addEventListener("click", () => {
            if (window.currentHost === Office.HostType.Excel) refreshExcelCustomView();
        });
    }
});

function injectExcelCustomStyles() {
    if (document.getElementById("schemaai-excel-custom-styles")) return;
    const style = document.createElement("style");
    style.id = "schemaai-excel-custom-styles";
    style.innerHTML = `
        /* 模式切换图文按钮 */
        .mode-switch-badge-btn {
            display: inline-flex; align-items: center; gap: 6px; padding: 5px 12px;
            border-radius: 8px; border: 1px solid var(--border-color);
            background: var(--card-bg); color: var(--accent-color); cursor: pointer;
            font-size: 12px; font-weight: 600; transition: all 0.2s cubic-bezier(0.25, 0.1, 0.25, 1);
            user-select: none;
        }
        .mode-switch-badge-btn:hover { background: #E3F2FD; border-color: var(--accent-color); transform: translateY(-1px); }

        /* 选区拾取瞄准按钮 */
        .pick-range-btn {
            padding: 0 12px; height: 42px; border-radius: 10px; border: 1.5px solid var(--border-color);
            background: var(--card-bg); color: var(--text-main); font-weight: 600; font-size: 13px;
            cursor: pointer; display: inline-flex; align-items: center; gap: 4px; transition: 0.2s; flex-shrink: 0;
        }
        .pick-range-btn:hover { border-color: var(--accent-color); color: var(--accent-color); }
        .pick-range-btn.picking {
            background: #E8F5E9; border-color: var(--success-color); color: var(--success-color); font-weight: 700;
            box-shadow: 0 0 0 3px rgba(52, 199, 89, 0.2);
        }

        .span-item-row {
            display: flex; align-items: center; gap: 8px; margin-bottom: 8px;
            background: var(--card-bg); padding: 8px 12px; border-radius: 10px; border: 1px solid var(--border-color);
        }
        .span-item-row select.span-select-wide {
            min-width: 100px !important; width: 100px !important; font-size: 13px !important;
            padding: 6px 8px !important; background: #F2F2F7; border: 1px solid transparent; border-radius: 8px;
            color: var(--text-main); font-weight: 600; outline: none;
        }
        .span-item-row select.span-select-wide:focus { background: #FFF; border-color: var(--accent-color); }

        .date-chip-btn {
            display: inline-flex; align-items: center; gap: 4px; padding: 6px 12px;
            border-radius: 8px; font-size: 12px; font-weight: 600; cursor: pointer;
            border: 1px solid var(--border-color); background: var(--card-bg); color: var(--text-main); transition: 0.2s;
            user-select: none;
        }
        .date-chip-btn.active {
            background: #E8F5E9; border-color: var(--success-color); color: var(--success-color); font-weight: 700;
        }

        /* 宽阔日历弹窗 */
        .cal-modal-overlay {
            position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 1000;
            display: none; align-items: center; justify-content: center; backdrop-filter: blur(3px);
        }
        .cal-modal-overlay.open { display: flex; }
        .cal-modal-box {
            background: var(--card-bg); border-radius: 20px; width: 94%; max-width: 480px;
            max-height: 90vh; overflow-y: auto; padding: 22px 24px; box-shadow: 0 20px 50px rgba(0,0,0,0.22);
        }
        .cal-grid {
            display: grid; grid-template-columns: repeat(7, 1fr); gap: 0; margin-top: 6px; user-select: none;
        }
        .cal-cell-header { font-size: 12px; font-weight: 700; text-align: center; color: var(--text-sub); padding: 6px 0; }
        .cal-day-cell {
            height: 48px; position: relative; background: transparent; cursor: pointer; transition: 0.15s;
            display: flex; align-items: center; justify-content: center; font-size: 13px; font-weight: 600;
            color: var(--text-main); padding: 0; margin: 2px 0;
        }
        .cal-day-cell.other-month { opacity: 0.15; cursor: default; }
        
        .cal-day-base-box {
            position: absolute; inset: 2px; border-radius: 8px; border: 1px solid var(--border-color);
            background: #FFF; z-index: 1; transition: 0.15s;
        }
        .cal-day-cell:hover:not(.other-month) .cal-day-base-box { border-color: var(--accent-color); }
        .cal-day-cell.selecting .cal-day-base-box { background: #E3F2FD !important; border-color: var(--accent-color) !important; }

        /* 连贯无缝覆盖色带层 */
        .cal-ribbon-layer {
            position: absolute; top: 4px; bottom: 4px; left: 0; right: 0; z-index: 2;
            pointer-events: none; transition: opacity 0.2s, transform 0.2s;
        }
        .cal-ribbon-layer.ribbon-start { left: 4px; border-top-left-radius: 8px; border-bottom-left-radius: 8px; }
        .cal-ribbon-layer.ribbon-end { right: 4px; border-top-right-radius: 8px; border-bottom-right-radius: 8px; }
        .cal-ribbon-layer.ribbon-single { left: 4px; right: 4px; border-radius: 8px; }
        .cal-ribbon-layer.dimmed { opacity: 0.05 !important; }
        .cal-ribbon-layer.highlighted { opacity: 0.92 !important; transform: scaleY(1.08); z-index: 3; }

        /* 宽阔图例胶囊区域 */
        .cal-legend-container {
            display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; max-height: 140px;
            overflow-y: auto; padding: 4px 2px;
        }
        .legend-capsule {
            display: inline-flex; align-items: center; gap: 8px; padding: 6px 14px;
            border-radius: 16px; font-size: 12px; font-weight: 600; background: #F2F2F7;
            border: 1.5px solid var(--border-color); cursor: pointer; transition: all 0.2s cubic-bezier(0.25, 0.1, 0.25, 1);
            user-select: none; min-height: 32px;
        }
        .legend-capsule:hover { transform: translateY(-1px); border-color: var(--accent-color); background: #FFF; }
        .legend-capsule.active {
            background: #1C1C1E !important; color: #FFF !important; border-color: #1C1C1E !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }
        .legend-color-box { width: 10px; height: 10px; border-radius: 3px; flex-shrink: 0; }
    `;
    document.head.appendChild(style);
}

// ========== 数据清洗与默认配置 ==========
function sanitizeExcelConfig(cfg) {
    const today = new Date();
    const defaultStart = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`;
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 2, 0);
    const defaultEnd = `${nextMonth.getFullYear()}-${String(nextMonth.getMonth() + 1).padStart(2, '0')}-${String(nextMonth.getDate()).padStart(2, '0')}`;

    const def = {
        inputType: 'none',
        autoMode: 'fixed',
        autoInputText: '',
        dateSeq: {
            startDate: defaultStart,
            endDate: defaultEnd,
            calcMode: 'weekday',
            weekdaySpans: [
                { start: 1, end: 2 },
                { start: 2, end: 3 }
            ],
            durationDays: 2,
            intervalDays: 1,
            skipSaturday: true,
            skipSunday: true,
            format: 'M.D-M.D',
            customRanges: []
        },
        manualSubType: 'text',
        aiEnabled: false,
        aiSkillId: 'lesson_progress',
        aiPrompt: '',
        thinkingMode: false,
        displayText: '',
        booleanDefault: 'none',
        falseAction: 'none',
        falseActionFixedText: '',
        trueAction: 'none',
        trueActionFixedText: '',
        choiceType: 'single',
        choiceOptions: ['选项1', '选项2'],
        choiceDefaultIndex: 0,
        choiceDefaultIndexes: [],
        targetRange: '',
        sheetName: 'Sheet1'
    };

    const clean = {};
    Object.keys(def).forEach(key => {
        let value = cfg[key];
        if (value === undefined || value === null) { clean[key] = def[key]; return; }
        if (key === 'dateSeq') {
            clean.dateSeq = Object.assign({}, def.dateSeq, value);
            if (!Array.isArray(clean.dateSeq.weekdaySpans) || clean.dateSeq.weekdaySpans.length === 0) {
                clean.dateSeq.weekdaySpans = [
                    { start: 1, end: 2 },
                    { start: 2, end: 3 }
                ];
            }
        } else if (Array.isArray(def[key])) {
            clean[key] = Array.isArray(value) ? value : def[key];
        } else if (typeof def[key] === 'boolean') {
            clean[key] = !!value;
        } else if (typeof def[key] === 'number') {
            clean[key] = isNaN(Number(value)) ? def[key] : Number(value);
        } else {
            clean[key] = value;
        }
    });
    return clean;
}

function formatDateItem(d1, d2, formatStr, weekIdx) {
    const pad = (n) => String(n).padStart(2, '0');
    const getParts = (d) => ({
        YYYY: d.getFullYear(),
        YY: String(d.getFullYear()).slice(-2),
        M: d.getMonth() + 1,
        MM: pad(d.getMonth() + 1),
        D: d.getDate(),
        DD: pad(d.getDate())
    });

    const p1 = getParts(d1);
    const p2 = d2 ? getParts(d2) : null;

    let res = formatStr || "M.D";
    res = res.replace(/第W周/g, `第${weekIdx + 1}周`).replace(/W/g, String(weekIdx + 1));

    if (p2 && res.includes("-")) {
        const segs = res.split("-");
        const formatSingle = (fmt, p) => {
            return fmt.replace(/YYYY/g, p.YYYY)
                      .replace(/YY/g, p.YY)
                      .replace(/MM/g, p.MM)
                      .replace(/M/g, p.M)
                      .replace(/DD/g, p.DD)
                      .replace(/D/g, p.D);
        };
        return `${formatSingle(segs[0], p1)}-${formatSingle(segs[1], p2)}`;
    }

    return res.replace(/YYYY/g, p1.YYYY)
              .replace(/YY/g, p1.YY)
              .replace(/MM/g, p1.MM)
              .replace(/M/g, p1.M)
              .replace(/DD/g, p1.DD)
              .replace(/D/g, p1.D);
}

function computeDateRanges(dCfg) {
    if (dCfg.customRanges && dCfg.customRanges.length > 0) {
        return dCfg.customRanges.map((r, i) => ({
            start: new Date(r.start),
            end: new Date(r.end),
            label: r.label || `${r.start.slice(5)}-${r.end.slice(5)}`,
            groupIndex: i
        }));
    }

    const startDate = new Date(dCfg.startDate ? dCfg.startDate.replace(/-/g, '/') : new Date());
    const endDate = new Date(dCfg.endDate ? dCfg.endDate.replace(/-/g, '/') : new Date());
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || startDate > endDate) return [];

    const isDayValid = (d) => {
        const day = d.getDay();
        if (day === 6 && dCfg.skipSaturday) return false;
        if (day === 0 && dCfg.skipSunday) return false;
        return true;
    };

    const ranges = [];

    if (dCfg.calcMode === 'weekday') {
        const spans = dCfg.weekdaySpans || [{ start: 1, end: 2 }];
        let currWeekMon = new Date(startDate);
        const dayOfWeek = currWeekMon.getDay() === 0 ? 7 : currWeekMon.getDay();
        currWeekMon.setDate(currWeekMon.getDate() - (dayOfWeek - 1));

        let weekIndex = 0;
        while (currWeekMon <= endDate) {
            for (let s = 0; s < spans.length; s++) {
                const span = spans[s];
                const d1 = new Date(currWeekMon);
                d1.setDate(currWeekMon.getDate() + (span.start - 1));
                const d2 = new Date(currWeekMon);
                d2.setDate(currWeekMon.getDate() + (span.end - 1));

                if (d2 < startDate) continue;
                if (d1 > endDate) break;
                if (!isDayValid(d1) && !isDayValid(d2)) continue;

                const actualStart = new Date(Math.max(d1, startDate));
                const actualEnd = new Date(Math.min(d2, endDate));

                ranges.push({
                    start: actualStart,
                    end: actualEnd,
                    weekIdx: weekIndex,
                    groupIndex: ranges.length,
                    label: `${actualStart.getMonth() + 1}.${actualStart.getDate()}-${actualEnd.getMonth() + 1}.${actualEnd.getDate()}`
                });
            }
            currWeekMon.setDate(currWeekMon.getDate() + 7);
            weekIndex++;
        }
    } else {
        const duration = Math.max(1, parseInt(dCfg.durationDays) || 1);
        const interval = Math.max(0, parseInt(dCfg.intervalDays) || 0);

        let curr = new Date(startDate);
        let count = 0;

        while (curr <= endDate) {
            while (curr <= endDate && !isDayValid(curr)) {
                curr.setDate(curr.getDate() + 1);
            }
            if (curr > endDate) break;

            const d1 = new Date(curr);
            let validCount = 1;
            let d2 = new Date(curr);

            while (validCount < duration && d2 <= endDate) {
                d2.setDate(d2.getDate() + 1);
                if (isDayValid(d2)) validCount++;
            }
            if (d2 > endDate) d2 = new Date(endDate);

            ranges.push({
                start: d1,
                end: d2,
                weekIdx: count,
                groupIndex: count,
                label: `${d1.getMonth() + 1}.${d1.getDate()}-${d2.getMonth() + 1}.${d2.getDate()}`
            });
            count++;

            curr = new Date(d2);
            curr.setDate(curr.getDate() + 1 + interval);
        }
    }

    return ranges;
}

function generateDateSequence(cfg, maxCount = 100) {
    const ranges = computeDateRanges(cfg.dateSeq || {});
    const fmt = (cfg.dateSeq && cfg.dateSeq.format) || 'M.D-M.D';
    return ranges.slice(0, maxCount).map((r, i) => formatDateItem(r.start, r.end, fmt, r.weekIdx || i));
}

// ========== 扫描全表与卡片构建 ==========
async function refreshExcelCustomView() {
    if (isExcelCustomRefreshing || window.currentHost !== Office.HostType.Excel) return;
    isExcelCustomRefreshing = true;

    const listContainer = document.getElementById("custom-controls-list");
    const emptyState = document.getElementById("custom-empty-state");
    if (!listContainer || !emptyState) { isExcelCustomRefreshing = false; return; }

    listContainer.innerHTML = '';
    listContainer.style.display = 'none';
    emptyState.style.display = 'none';

    try {
        const groups = {};

        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();

            for (const sheet of worksheets.items) {
                const usedRange = sheet.getUsedRangeOrNullObject();
                usedRange.load("values");
                await context.sync();

                if (!usedRange.isNullObject && usedRange.values) {
                    const vals = usedRange.values;
                    const regex = /\{\{(.+?)\}\}/g;
                    for (let r = 0; r < vals.length; r++) {
                        for (let c = 0; c < vals[r].length; c++) {
                            const text = String(vals[r][c] || "");
                            let match;
                            while ((match = regex.exec(text)) !== null) {
                                const tag = match[1].trim();
                                if (!groups[tag]) {
                                    groups[tag] = {
                                        tag: tag,
                                        title: tag,
                                        sheetName: sheet.name,
                                        targetRange: `${sheet.name}!${getColumnLetter(c + 1)}${r + 1}`
                                    };
                                }
                            }
                        }
                    }
                }
            }
        });

        const addRangeBtn = document.createElement("button");
        addRangeBtn.className = "primary-btn";
        addRangeBtn.style.marginBottom = "16px";
        addRangeBtn.style.background = "var(--card-bg)";
        addRangeBtn.style.border = "1px dashed var(--accent-color)";
        addRangeBtn.style.color = "var(--accent-color)";
        addRangeBtn.textContent = "+ 手动拾取当前鼠标选区 (支持 Ctrl 多选)";
        addRangeBtn.onclick = () => captureCurrentSelectionToNewCard(groups);
        listContainer.appendChild(addRangeBtn);

        if (Object.keys(groups).length === 0) {
            emptyState.style.display = 'block';
            emptyState.innerHTML = `
                <svg viewBox="0 0 24 24" width="56" height="56" stroke="#C7C7CC" stroke-width="1.5" fill="none"><polyline points="23 4 23 10 17 10"></polyline><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"></path></svg>
                <p style="color: var(--text-sub); font-size: 14px; margin-top: 16px;">未在任何 Sheet 中检测到 {{占位符}}<br><span style="font-size:12px;">可在单元格输入 {{标题}}，或点击上方直接拾取选区</span></p>
            `;
            listContainer.style.display = 'block';
            isExcelCustomRefreshing = false;
            return;
        }

        listContainer.style.display = 'block';
        renderExcelControlCards(groups, listContainer);

    } catch (error) {
        console.error("Excel 扫描异常:", error);
        emptyState.style.display = 'block';
        listContainer.style.display = 'none';
    } finally {
        isExcelCustomRefreshing = false;
    }
}

function getColumnLetter(colNum) {
    let letter = "";
    while (colNum > 0) {
        let mod = (colNum - 1) % 26;
        letter = String.fromCharCode(65 + mod) + letter;
        colNum = Math.floor((colNum - mod) / 26);
    }
    return letter;
}

async function captureCurrentSelectionToNewCard(groups) {
    try {
        await Excel.run(async (ctx) => {
            const selected = ctx.workbook.getSelectedRanges ? ctx.workbook.getSelectedRanges() : ctx.workbook.getSelectedRange();
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            selected.load("address");
            sheet.load("name");
            await ctx.sync();

            const rawAddrs = selected.address.split(",").map(a => a.trim());
            const fullAddress = rawAddrs.map(addr => {
                const clean = addr.includes("!") ? addr.split("!").pop() : addr;
                return `${sheet.name}!${clean}`;
            }).join(", ");

            const customTag = `cell_${Math.random().toString(36).substr(2, 4)}`;

            groups[customTag] = {
                tag: customTag,
                title: `选区 ${rawAddrs[0]}`,
                sheetName: sheet.name,
                targetRange: fullAddress
            };

            const listContainer = document.getElementById("custom-controls-list");
            if (listContainer) {
                listContainer.innerHTML = '';
                const addRangeBtn = document.createElement("button");
                addRangeBtn.className = "primary-btn";
                addRangeBtn.style.marginBottom = "16px";
                addRangeBtn.style.background = "var(--card-bg)";
                addRangeBtn.style.border = "1px dashed var(--accent-color)";
                addRangeBtn.style.color = "var(--accent-color)";
                addRangeBtn.textContent = "+ 手动拾取当前鼠标选区 (支持 Ctrl 多选)";
                addRangeBtn.onclick = () => captureCurrentSelectionToNewCard(groups);
                listContainer.appendChild(addRangeBtn);
                renderExcelControlCards(groups, listContainer);
            }
        });
    } catch (e) {
        showStatus("获取选区失败", "error");
    }
}

function renderExcelControlCards(groups, listContainer) {
    Object.keys(groups).forEach(tag => {
        const group = groups[tag];
        excelControlsConfig[tag] = sanitizeExcelConfig(excelControlsConfig[tag] || {
            targetRange: group.targetRange || '',
            sheetName: group.sheetName || 'Sheet1'
        });
        const cfg = excelControlsConfig[tag];

        const itemDiv = document.createElement('div');
        itemDiv.className = 'control-item';
        itemDiv.setAttribute('data-tag', tag);

        const headerDiv = document.createElement('div');
        headerDiv.className = 'control-header';
        const headerLeft = document.createElement('div');
        headerLeft.className = 'control-header-left';
        const titleInput = document.createElement('input');
        titleInput.type = 'text';
        titleInput.value = group.title;
        titleInput.placeholder = '标题';
        titleInput.className = 'control-title-input';
        titleInput.addEventListener('input', () => { group.title = titleInput.value; });
        const tagSpan = document.createElement('span');
        tagSpan.className = 'control-tag';
        tagSpan.textContent = cfg.targetRange || tag;
        headerLeft.appendChild(titleInput);
        headerLeft.appendChild(tagSpan);

        const arrowSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        arrowSvg.setAttribute('viewBox', '0 0 24 24');
        arrowSvg.setAttribute('stroke', 'currentColor');
        arrowSvg.setAttribute('stroke-width', '2');
        arrowSvg.setAttribute('fill', 'none');
        arrowSvg.classList.add('control-arrow');
        const polyline = document.createElementNS('http://www.w3.org/2000/svg', 'polyline');
        polyline.setAttribute('points', '6 9 12 15 18 9');
        arrowSvg.appendChild(polyline);
        headerDiv.appendChild(headerLeft);
        headerDiv.appendChild(arrowSvg);
        headerDiv.addEventListener('click', (e) => {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.closest('button')) return;
            itemDiv.classList.toggle('expanded');
        });

        const bodyDiv = document.createElement('div');
        bodyDiv.className = 'control-body';

        // 目标单元格/范围输入框（配备专属拾取瞄准按钮）
        const rangeGroup = document.createElement('div');
        rangeGroup.className = 'input-group';
        rangeGroup.style.marginBottom = '12px';
        const rangeLabel = document.createElement('label');
        rangeLabel.innerHTML = '目标单元格/范围 <span style="font-size:11px;color:var(--text-sub);">(支持按住 Ctrl 多重框选)</span>';

        const rangeWrapper = document.createElement('div');
        rangeWrapper.style.display = 'flex';
        rangeWrapper.style.gap = '8px';
        rangeWrapper.style.alignItems = 'center';

        const rangeInput = document.createElement('input');
        rangeInput.type = 'text';
        rangeInput.className = 'range-input';
        rangeInput.dataset.tag = tag;
        rangeInput.value = cfg.targetRange || '';
        rangeInput.placeholder = '如 Sheet1!A4:A20 或多选区 A1:A3, C1:C3';
        rangeInput.style.flex = '1';
        rangeInput.addEventListener('input', () => {
            cfg.targetRange = rangeInput.value.trim();
            tagSpan.textContent = cfg.targetRange || tag;
        });

        const pickBtn = document.createElement('button');
        pickBtn.type = 'button';
        pickBtn.className = 'pick-range-btn';
        pickBtn.innerHTML = `
            <svg viewBox="0 0 24 24" width="15" height="15" stroke="currentColor" stroke-width="2" fill="none"><circle cx="12" cy="12" r="10"></circle><line x1="22" y1="12" x2="18" y2="12"></line><line x1="6" y1="12" x2="2" y2="12"></line><line x1="12" y1="6" x2="12" y2="2"></line><line x1="12" y1="22" x2="12" y2="18"></line></svg>
            <span>拾取</span>
        `;

        pickBtn.onclick = () => {
            if (activePickingTag === tag) {
                activePickingTag = null;
                pickBtn.classList.remove('picking');
                pickBtn.querySelector('span').textContent = '拾取';
            } else {
                document.querySelectorAll('.pick-range-btn').forEach(b => {
                    b.classList.remove('picking');
                    b.querySelector('span').textContent = '拾取';
                });
                activePickingTag = tag;
                pickBtn.classList.add('picking');
                pickBtn.querySelector('span').textContent = '采集中...';
                showStatus("请在 Excel 页面中用鼠标框选单元格 (支持 Ctrl 多选)", "info");
            }
        };

        rangeWrapper.appendChild(rangeInput);
        rangeWrapper.appendChild(pickBtn);
        rangeGroup.appendChild(rangeLabel);
        rangeGroup.appendChild(rangeWrapper);
        bodyDiv.appendChild(rangeGroup);

        // 模式切换行
        const rowDiv = document.createElement('div');
        rowDiv.className = 'control-section-row';
        const inputLabel = document.createElement('label');
        inputLabel.textContent = '输入框内容';
        const inputTypeSelect = document.createElement('select');
        inputTypeSelect.className = 'custom-select-mini input-type-select';
        inputTypeSelect.innerHTML = '<option value="none">无 (自动填充/生成)</option><option value="manual">手动输入</option>';
        inputTypeSelect.value = cfg.inputType;

        const manualTypeLabel = document.createElement('label');
        manualTypeLabel.textContent = '类型';
        const manualTypeSelect = document.createElement('select');
        manualTypeSelect.className = 'custom-select-mini manual-sub-type-select';
        manualTypeSelect.innerHTML = `<option value="text">文本</option><option value="boolean">布尔值</option><option value="choice">选择</option>`;
        manualTypeSelect.value = cfg.manualSubType;

        rowDiv.appendChild(inputLabel);
        rowDiv.appendChild(inputTypeSelect);
        rowDiv.appendChild(manualTypeLabel);
        rowDiv.appendChild(manualTypeSelect);

        const inputSectionTitle = document.createElement('div');
        inputSectionTitle.className = 'control-section-title';
        inputSectionTitle.textContent = '输入与填充规则配置';
        bodyDiv.appendChild(inputSectionTitle);

        const inputSubContainer = document.createElement('div');
        inputSubContainer.className = 'control-sub-item';
        inputSubContainer.appendChild(rowDiv);

        // 自动填充容器
        const autoFillContainer = document.createElement('div');
        autoFillContainer.style.marginTop = '12px';

        const autoModeRow = document.createElement('div');
        autoModeRow.className = 'control-section-row';
        autoModeRow.style.marginBottom = '8px';
        const autoModeLabel = document.createElement('label');
        autoModeLabel.textContent = '填充模式';
        const autoModeSelect = document.createElement('select');
        autoModeSelect.className = 'custom-select-mini';
        autoModeSelect.innerHTML = `<option value="fixed">固定文本</option><option value="date_seq">日期序列推演</option>`;
        autoModeSelect.value = cfg.autoMode || 'fixed';
        autoModeRow.appendChild(autoModeLabel);
        autoModeRow.appendChild(autoModeSelect);
        autoFillContainer.appendChild(autoModeRow);

        // 1. 固定文本
        const fixedTextDiv = document.createElement('div');
        fixedTextDiv.className = 'input-group';
        fixedTextDiv.innerHTML = `<label>自动输入固定文本</label><input type="text" class="auto-input-text" value="${cfg.autoInputText || ''}" placeholder="填入要写入的固定内容...">`;
        const fixedInput = fixedTextDiv.querySelector('input');
        fixedInput.addEventListener('input', () => { cfg.autoInputText = fixedInput.value; });
        autoFillContainer.appendChild(fixedTextDiv);

        // 2. 日期序列推演面板
        const dateSeqDiv = document.createElement('div');
        dateSeqDiv.className = 'date-seq-container';
        dateSeqDiv.style.background = 'var(--bg-color)';
        dateSeqDiv.style.borderRadius = '12px';
        dateSeqDiv.style.padding = '12px';
        dateSeqDiv.style.marginTop = '8px';

        const dCfg = cfg.dateSeq;

        dateSeqDiv.innerHTML = `
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:12px;">
                <label style="margin:0; font-size:13px; font-weight:700;">日期范围与推演模式</label>
                <div style="display:flex; gap:8px; align-items:center;">
                    <button class="mode-switch-badge-btn toggle-calc-mode-btn" type="button" title="点击切换：星期模式 / 日期步长模式">
                        <svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
                        <span class="mode-btn-text">${dCfg.calcMode === 'weekday' ? '星期模式' : '日期模式'}</span>
                    </button>
                    <button class="icon-btn small open-cal-modal-btn" type="button" title="打开月份日历图表" style="color:var(--success-color); background:var(--card-bg); border:1px solid var(--border-color); width:32px; height:32px; border-radius:8px;">
                        <svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="2" fill="none"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>
                    </button>
                </div>
            </div>

            <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                <input type="date" class="date-range-start" value="${dCfg.startDate}" style="padding:6px 8px; width:135px; font-size:13px; border-radius:8px; border:1px solid var(--border-color);">
                <span style="font-size:13px; color:var(--text-sub);">至</span>
                <input type="date" class="date-range-end" value="${dCfg.endDate}" style="padding:6px 8px; width:135px; font-size:13px; border-radius:8px; border:1px solid var(--border-color);">
            </div>

            <!-- 星期模式 -->
            <div class="weekday-mode-panel" style="margin-top:12px; display:${dCfg.calcMode === 'weekday' ? 'block' : 'none'};">
                <div class="weekday-spans-list"></div>
                <button class="add-span-btn" type="button" style="background:none; border:1px dashed var(--border-color); border-radius:8px; padding:6px 12px; font-size:12px; color:var(--accent-color); font-weight:600; cursor:pointer; width:100%; text-align:center;">+ 添加分段跨度</button>
            </div>

            <!-- 日期模式 -->
            <div class="daystep-mode-panel" style="margin-top:12px; display:${dCfg.calcMode === 'day_step' ? 'block' : 'none'};">
                <div style="display:flex; align-items:center; gap:12px;">
                    <div style="display:flex; align-items:center; gap:6px;">
                        <span style="font-size:13px; color:var(--text-sub);">每组持续</span>
                        <input type="number" class="step-duration-input" value="${dCfg.durationDays || 2}" min="1" style="width:60px; padding:4px 8px; font-size:13px; border-radius:8px; border:1px solid var(--border-color);">
                        <span style="font-size:13px;">天</span>
                    </div>
                    <div style="display:flex; align-items:center; gap:6px;">
                        <span style="font-size:13px; color:var(--text-sub);">间隔</span>
                        <input type="number" class="step-interval-input" value="${dCfg.intervalDays || 1}" min="0" style="width:60px; padding:4px 8px; font-size:13px; border-radius:8px; border:1px solid var(--border-color);">
                        <span style="font-size:13px;">天</span>
                    </div>
                </div>
            </div>

            <!-- 过滤与格式 -->
            <div style="display:flex; align-items:center; gap:8px; margin-top:12px; border-top:1px dashed var(--border-color); padding-top:10px;">
                <div class="date-chip-btn skip-sat-btn ${dCfg.skipSaturday ? 'active' : ''}">跳过周六</div>
                <div class="date-chip-btn skip-sun-btn ${dCfg.skipSunday ? 'active' : ''}">跳过周日</div>
                
                <select class="custom-select-mini date-seq-format" style="margin-left:auto; width:140px;">
                    <option value="M.D-M.D">M.D-M.D (如 6.29-7.1)</option>
                    <option value="M.D">M.D (如 9.1)</option>
                    <option value="M月D日-M月D日">M月D日-M月D日</option>
                    <option value="YYYY.MM.DD">YYYY.MM.DD</option>
                    <option value="YYYY-MM-DD">YYYY-MM-DD</option>
                    <option value="第W周 (M.D-M.D)">第W周 (M.D-M.D)</option>
                </select>
            </div>
        `;

        autoFillContainer.appendChild(dateSeqDiv);

        const startInput = dateSeqDiv.querySelector('.date-range-start');
        const endInput = dateSeqDiv.querySelector('.date-range-end');
        const toggleModeBtn = dateSeqDiv.querySelector('.toggle-calc-mode-btn');
        const modeBtnText = dateSeqDiv.querySelector('.mode-btn-text');
        const openCalModalBtn = dateSeqDiv.querySelector('.open-cal-modal-btn');
        const weekdayPanel = dateSeqDiv.querySelector('.weekday-mode-panel');
        const daystepPanel = dateSeqDiv.querySelector('.daystep-mode-panel');
        const spansListDiv = dateSeqDiv.querySelector('.weekday-spans-list');
        const addSpanBtn = dateSeqDiv.querySelector('.add-span-btn');
        const durationInput = dateSeqDiv.querySelector('.step-duration-input');
        const intervalInput = dateSeqDiv.querySelector('.step-interval-input');
        const satBtn = dateSeqDiv.querySelector('.skip-sat-btn');
        const sunBtn = dateSeqDiv.querySelector('.skip-sun-btn');
        const formatSelect = dateSeqDiv.querySelector('.date-seq-format');

        formatSelect.value = dCfg.format;

        const renderSpansList = () => {
            spansListDiv.innerHTML = '';
            dCfg.weekdaySpans.forEach((sp, idx) => {
                const row = document.createElement('div');
                row.className = 'span-item-row';
                row.innerHTML = `
                    <span style="font-size:13px; color:var(--text-main); font-weight:600;">从</span>
                    <select class="span-select-wide span-start-sel">
                        <option value="1" ${sp.start == 1 ? 'selected' : ''}>周一</option>
                        <option value="2" ${sp.start == 2 ? 'selected' : ''}>周二</option>
                        <option value="3" ${sp.start == 3 ? 'selected' : ''}>周三</option>
                        <option value="4" ${sp.start == 4 ? 'selected' : ''}>周四</option>
                        <option value="5" ${sp.start == 5 ? 'selected' : ''}>周五</option>
                        <option value="6" ${sp.start == 6 ? 'selected' : ''}>周六</option>
                        <option value="7" ${sp.start == 7 ? 'selected' : ''}>周日</option>
                    </select>
                    <span style="font-size:13px; color:var(--text-main); font-weight:600;">至</span>
                    <select class="span-select-wide span-end-sel">
                        <option value="1" ${sp.end == 1 ? 'selected' : ''}>周一</option>
                        <option value="2" ${sp.end == 2 ? 'selected' : ''}>周二</option>
                        <option value="3" ${sp.end == 3 ? 'selected' : ''}>周三</option>
                        <option value="4" ${sp.end == 4 ? 'selected' : ''}>周四</option>
                        <option value="5" ${sp.end == 5 ? 'selected' : ''}>周五</option>
                        <option value="6" ${sp.end == 6 ? 'selected' : ''}>周六</option>
                        <option value="7" ${sp.end == 7 ? 'selected' : ''}>周日</option>
                    </select>
                    <button class="icon-btn small del-span-btn" type="button" style="color:#FF3B30; margin-left:auto; width:26px; height:26px; border-radius:50%;">
                        <svg viewBox="0 0 24 24" width="13" height="13"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                    </button>
                `;

                row.querySelector('.span-start-sel').addEventListener('change', (e) => {
                    sp.start = parseInt(e.target.value);
                    dCfg.customRanges = [];
                });
                row.querySelector('.span-end-sel').addEventListener('change', (e) => {
                    sp.end = parseInt(e.target.value);
                    dCfg.customRanges = [];
                });
                row.querySelector('.del-span-btn').addEventListener('click', () => {
                    if (dCfg.weekdaySpans.length <= 1) return;
                    dCfg.weekdaySpans.splice(idx, 1);
                    dCfg.customRanges = [];
                    renderSpansList();
                });
                spansListDiv.appendChild(row);
            });
        };

        renderSpansList();

        addSpanBtn.addEventListener('click', () => {
            const last = dCfg.weekdaySpans[dCfg.weekdaySpans.length - 1] || { start: 1, end: 2 };
            const nextStart = last.end;
            const nextEnd = Math.min(7, nextStart + 1);
            dCfg.weekdaySpans.push({ start: nextStart, end: nextEnd });
            dCfg.customRanges = [];
            renderSpansList();
        });

        toggleModeBtn.addEventListener('click', () => {
            dCfg.calcMode = dCfg.calcMode === 'weekday' ? 'day_step' : 'weekday';
            dCfg.customRanges = [];
            modeBtnText.textContent = dCfg.calcMode === 'weekday' ? '星期模式' : '日期模式';
            weekdayPanel.style.display = dCfg.calcMode === 'weekday' ? 'block' : 'none';
            daystepPanel.style.display = dCfg.calcMode === 'day_step' ? 'block' : 'none';
        });

        startInput.addEventListener('input', () => { dCfg.startDate = startInput.value; dCfg.customRanges = []; });
        endInput.addEventListener('input', () => { dCfg.endDate = endInput.value; dCfg.customRanges = []; });
        durationInput.addEventListener('input', () => { dCfg.durationDays = parseInt(durationInput.value) || 1; dCfg.customRanges = []; });
        intervalInput.addEventListener('input', () => { dCfg.intervalDays = parseInt(intervalInput.value) || 0; dCfg.customRanges = []; });

        satBtn.addEventListener('click', () => {
            dCfg.skipSaturday = !dCfg.skipSaturday;
            satBtn.classList.toggle('active', dCfg.skipSaturday);
            dCfg.customRanges = [];
        });

        sunBtn.addEventListener('click', () => {
            dCfg.skipSunday = !dCfg.skipSunday;
            sunBtn.classList.toggle('active', dCfg.skipSunday);
            dCfg.customRanges = [];
        });

        formatSelect.addEventListener('change', () => {
            dCfg.format = formatSelect.value;
        });

        autoModeSelect.addEventListener('change', () => {
            cfg.autoMode = autoModeSelect.value;
            fixedTextDiv.style.display = cfg.autoMode === 'fixed' ? 'block' : 'none';
            dateSeqDiv.style.display = cfg.autoMode === 'date_seq' ? 'block' : 'none';
        });

        fixedTextDiv.style.display = cfg.autoMode === 'fixed' ? 'block' : 'none';
        dateSeqDiv.style.display = cfg.autoMode === 'date_seq' ? 'block' : 'none';

        openCalModalBtn.addEventListener('click', () => {
            openInteractiveCalendarModal(cfg);
        });

        inputSubContainer.appendChild(autoFillContainer);

        const manualOptionsContainer = document.createElement('div');
        manualOptionsContainer.className = 'manual-options-container';
        inputSubContainer.appendChild(manualOptionsContainer);

        bodyDiv.appendChild(inputSubContainer);

        // AI 设置面板：引入 Skills 技能预设引擎
        const aiSectionTitle = document.createElement('div');
        aiSectionTitle.className = 'control-section-title';
        aiSectionTitle.style.display = 'flex';
        aiSectionTitle.style.alignItems = 'center';
        aiSectionTitle.style.justifyContent = 'space-between';
        aiSectionTitle.innerHTML = `<span>AI 设置 (Skills 赋能)</span><label class="switch"><input type="checkbox" class="ai-checkbox" ${cfg.aiEnabled ? 'checked' : ''}><span class="slider round"></span></label>`;
        bodyDiv.appendChild(aiSectionTitle);

        const aiSubContainer = document.createElement('div');
        aiSubContainer.className = 'control-sub-item';

        const aiSettingsDiv = document.createElement('div');
        aiSettingsDiv.className = 'ai-settings';
        aiSettingsDiv.style.display = cfg.aiEnabled ? 'block' : 'none';
        aiSettingsDiv.innerHTML = `
            <div class="input-group" style="margin-top:8px;">
                <label>技能预设 (Skill)</label>
                <select class="custom-select-mini ai-skill-select" style="width:100%;">
                    ${SKILL_PRESETS.map(s => `<option value="${s.id}" ${cfg.aiSkillId === s.id ? 'selected' : ''}>${s.name}</option>`).join("")}
                </select>
            </div>
            <div class="input-group" style="margin-top:8px;">
                <label>补充微调要求 (可选)</label>
                <input type="text" class="ai-prompt-input" placeholder="输入个性化补充要求..." value="${cfg.aiPrompt || ''}">
            </div>
            <div style="display:flex; align-items:center; gap:8px; margin-top:12px;">
                <label style="margin:0;">思考模式</label>
                <label class="switch"><input type="checkbox" class="thinking-checkbox" ${cfg.thinkingMode ? 'checked' : ''}><span class="slider round"></span></label>
            </div>
        `;

        const skillSelect = aiSettingsDiv.querySelector('.ai-skill-select');
        skillSelect.addEventListener('change', () => { cfg.aiSkillId = skillSelect.value; });

        const aiPromptInput = aiSettingsDiv.querySelector('.ai-prompt-input');
        aiPromptInput.addEventListener('input', () => { cfg.aiPrompt = aiPromptInput.value; });

        const thinkCheckbox = aiSettingsDiv.querySelector('.thinking-checkbox');
        thinkCheckbox.addEventListener('change', () => { cfg.thinkingMode = thinkCheckbox.checked; });

        aiSubContainer.appendChild(aiSettingsDiv);
        bodyDiv.appendChild(aiSubContainer);

        const aiCheckbox = aiSectionTitle.querySelector('.ai-checkbox');
        aiCheckbox.addEventListener('change', () => {
            cfg.aiEnabled = aiCheckbox.checked;
            aiSettingsDiv.style.display = cfg.aiEnabled ? 'block' : 'none';
        });

        const updateInputTypeUI = () => {
            const isManual = cfg.inputType === 'manual';
            manualTypeLabel.style.display = isManual ? '' : 'none';
            manualTypeSelect.style.display = isManual ? '' : 'none';
            autoFillContainer.style.display = isManual ? 'none' : 'block';
            manualOptionsContainer.style.display = isManual ? 'block' : 'none';
        };

        inputTypeSelect.addEventListener('change', () => {
            cfg.inputType = inputTypeSelect.value;
            updateInputTypeUI();
        });
        manualTypeSelect.addEventListener('change', () => {
            cfg.manualSubType = manualTypeSelect.value;
        });

        updateInputTypeUI();

        itemDiv.appendChild(headerDiv);
        itemDiv.appendChild(bodyDiv);
        listContainer.appendChild(itemDiv);
    });
}

// ========== 可视化交互日历弹窗核心 (仅展示当月图例 + 连贯色带 + 选中修改不新建) ==========
function openInteractiveCalendarModal(cfg) {
    let modal = document.getElementById("excel-calendar-modal-overlay");
    if (!modal) {
        modal = document.createElement("div");
        modal.id = "excel-calendar-modal-overlay";
        modal.className = "cal-modal-overlay";
        document.body.appendChild(modal);
    }

    const dCfg = cfg.dateSeq;
    let viewDate = new Date(dCfg.startDate ? dCfg.startDate.replace(/-/g, '/') : new Date());
    if (isNaN(viewDate.getTime())) viewDate = new Date();

    let currentRanges = computeDateRanges(dCfg).map((r, i) => ({
        start: new Date(r.start),
        end: new Date(r.end),
        label: r.label || `${r.start.getMonth() + 1}.${r.start.getDate()}-${r.end.getMonth() + 1}.${r.end.getDate()}`,
        colorIdx: i % DISTINCT_PALETTE.length
    }));

    let isSelecting = false;
    let dragStartDate = null;
    let activeHighlightIdx = null;

    const renderCalendar = () => {
        const year = viewDate.getFullYear();
        const month = viewDate.getMonth();

        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        let firstDayWeek = firstDay.getDay() === 0 ? 7 : firstDay.getDay();

        // 仅提取在“当前查看月份”内有交集的分组图例
        const monthStart = new Date(year, month, 1);
        const monthEnd = new Date(year, month + 1, 0, 23, 59, 59);

        const visibleRangesWithIdx = [];
        currentRanges.forEach((r, idx) => {
            if (r.end >= monthStart && r.start <= monthEnd) {
                visibleRangesWithIdx.push({ range: r, realIdx: idx });
            }
        });

        let statusHint = `<span style="color:var(--text-sub);">滑动框选可新建；点击下方图例选中后滑动即可<b>直接修改</b></span>`;
        if (activeHighlightIdx !== null && currentRanges[activeHighlightIdx]) {
            const cur = currentRanges[activeHighlightIdx];
            statusHint = `<span style="color:var(--accent-color); font-weight:700;">正在修改【${cur.label}】：请在日历上滑动重新框选</span>`;
        }

        modal.innerHTML = `
            <div class="cal-modal-box">
                <!-- 顶部栏：左侧返回X -->
                <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        <button class="icon-btn small btn-close-x" type="button" title="不保存直接返回" style="width:28px; height:28px; border-radius:50%;">
                            <svg viewBox="0 0 24 24" width="14" height="14"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                        </button>
                        <span style="font-weight:700; font-size:15px; color:var(--text-main);">${year}年 ${month + 1}月</span>
                    </div>
                    <div style="display:flex; gap:6px;">
                        <button class="icon-btn small prev-month-btn" type="button" style="width:28px; height:28px;"><svg viewBox="0 0 24 24" width="14" height="14"><polyline points="15 18 9 12 15 6"></polyline></svg></button>
                        <button class="icon-btn small next-month-btn" type="button" style="width:28px; height:28px;"><svg viewBox="0 0 24 24" width="14" height="14"><polyline points="9 18 15 12 9 6"></polyline></svg></button>
                    </div>
                </div>

                <div style="font-size:12px; margin-bottom:10px;">${statusHint}</div>

                <div class="cal-grid">
                    <div class="cal-cell-header">一</div><div class="cal-cell-header">二</div><div class="cal-cell-header">三</div>
                    <div class="cal-cell-header">四</div><div class="cal-cell-header">五</div><div class="cal-cell-header">六</div><div class="cal-cell-header">日</div>
                </div>

                <div class="cal-grid cal-days-container"></div>

                <!-- 宽阔图例胶囊区域：仅展示本月图例 -->
                <div style="margin-top:14px; border-top:1px solid var(--border-color); padding-top:10px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                        <span style="font-size:12px; font-weight:700; color:var(--text-main);">${month + 1}月份分组图例</span>
                        ${activeHighlightIdx !== null ? '<span class="clear-selection-btn" style="font-size:11px; color:var(--accent-color); cursor:pointer; font-weight:600;">取消选中</span>' : ''}
                    </div>
                    <div class="cal-legend-container"></div>
                </div>

                <div style="display:flex; gap:10px; margin-top:16px;">
                    <button class="primary-btn btn-reset-cal" style="background:#E5E5EA; color:var(--text-main); margin-top:0; padding:11px;">重置</button>
                    <button class="primary-btn btn-save-cal" style="margin-top:0; padding:11px;">应用并同步</button>
                </div>
            </div>
        `;

        const daysContainer = modal.querySelector(".cal-days-container");
        const legendContainer = modal.querySelector(".cal-legend-container");

        for (let i = 1; i < firstDayWeek; i++) {
            const emptyCell = document.createElement("div");
            emptyCell.className = "cal-day-cell other-month";
            daysContainer.appendChild(emptyCell);
        }

        for (let d = 1; d <= lastDay.getDate(); d++) {
            const cellDate = new Date(year, month, d);
            const cell = document.createElement("div");
            cell.className = "cal-day-cell";
            cell.dataset.date = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;

            const baseBox = document.createElement("div");
            baseBox.className = "cal-day-base-box";
            cell.appendChild(baseBox);

            const numSpan = document.createElement("span");
            numSpan.style.position = "relative";
            numSpan.style.zIndex = "4";
            numSpan.textContent = d;
            cell.appendChild(numSpan);

            // 绘制连贯色带
            currentRanges.forEach((r, idx) => {
                const s = new Date(r.start.getFullYear(), r.start.getMonth(), r.start.getDate());
                const e = new Date(r.end.getFullYear(), r.end.getMonth(), r.end.getDate());

                if (cellDate >= s && cellDate <= e) {
                    const col = DISTINCT_PALETTE[r.colorIdx % DISTINCT_PALETTE.length];
                    const layer = document.createElement("div");
                    layer.className = "cal-ribbon-layer";
                    layer.style.backgroundColor = col.bg;
                    layer.style.color = col.solid;

                    const isStart = cellDate.getTime() === s.getTime() || cellDate.getDay() === 1;
                    const isEnd = cellDate.getTime() === e.getTime() || cellDate.getDay() === 0;

                    if (isStart && isEnd) {
                        layer.classList.add("ribbon-single");
                        layer.style.border = `2px solid ${col.border}`;
                    } else if (isStart) {
                        layer.classList.add("ribbon-start");
                        layer.style.borderTop = `2px solid ${col.border}`;
                        layer.style.borderBottom = `2px solid ${col.border}`;
                        layer.style.borderLeft = `2px solid ${col.border}`;
                    } else if (isEnd) {
                        layer.classList.add("ribbon-end");
                        layer.style.borderTop = `2px solid ${col.border}`;
                        layer.style.borderBottom = `2px solid ${col.border}`;
                        layer.style.borderRight = `2px solid ${col.border}`;
                    } else {
                        layer.style.borderTop = `2px solid ${col.border}`;
                        layer.style.borderBottom = `2px solid ${col.border}`;
                    }

                    if (activeHighlightIdx !== null) {
                        if (activeHighlightIdx === idx) {
                            layer.classList.add("highlighted");
                        } else {
                            layer.classList.add("dimmed");
                        }
                    }

                    cell.appendChild(layer);
                }
            });

            cell.addEventListener("mousedown", () => {
                isSelecting = true;
                dragStartDate = new Date(cellDate);
                document.querySelectorAll(".cal-day-cell").forEach(c => c.classList.remove("selecting"));
                cell.classList.add("selecting");
            });

            cell.addEventListener("mouseenter", () => {
                if (!isSelecting || !dragStartDate) return;
                const currHoverDate = new Date(cellDate);
                const minD = new Date(Math.min(dragStartDate, currHoverDate));
                const maxD = new Date(Math.max(dragStartDate, currHoverDate));

                document.querySelectorAll(".cal-day-cell[data-date]").forEach(c => {
                    const cd = new Date(c.dataset.date.replace(/-/g, '/'));
                    c.classList.toggle("selecting", cd >= minD && cd <= maxD);
                });
            });

            daysContainer.appendChild(cell);
        }

        // 仅渲染当前查看月份的图例胶囊
        if (visibleRangesWithIdx.length === 0) {
            legendContainer.innerHTML = `<span style="font-size:12px; color:var(--text-sub);">${month + 1} 月份暂无分组安排</span>`;
        } else {
            visibleRangesWithIdx.forEach(item => {
                const r = item.range;
                const realIdx = item.realIdx;
                const col = DISTINCT_PALETTE[r.colorIdx % DISTINCT_PALETTE.length];
                const capsule = document.createElement("div");
                capsule.className = `legend-capsule ${activeHighlightIdx === realIdx ? 'active' : ''}`;
                capsule.innerHTML = `
                    <div class="legend-color-box" style="background:${col.solid};"></div>
                    <span>${r.label || `组${realIdx + 1}`}</span>
                `;

                capsule.onclick = () => {
                    activeHighlightIdx = (activeHighlightIdx === realIdx) ? null : realIdx;
                    renderCalendar();
                };

                legendContainer.appendChild(capsule);
            });
        }

        const clearBtn = modal.querySelector(".clear-selection-btn");
        if (clearBtn) {
            clearBtn.onclick = () => {
                activeHighlightIdx = null;
                renderCalendar();
            };
        }

        modal.querySelector(".btn-close-x").onclick = () => {
            modal.classList.remove("open");
        };

        modal.querySelector(".prev-month-btn").onclick = () => {
            viewDate.setMonth(viewDate.getMonth() - 1);
            renderCalendar();
        };
        modal.querySelector(".next-month-btn").onclick = () => {
            viewDate.setMonth(viewDate.getMonth() + 1);
            renderCalendar();
        };

        modal.querySelector(".btn-reset-cal").onclick = () => {
            dCfg.customRanges = [];
            activeHighlightIdx = null;
            currentRanges = computeDateRanges(dCfg).map((r, i) => ({
                start: new Date(r.start),
                end: new Date(r.end),
                label: r.label,
                colorIdx: i % DISTINCT_PALETTE.length
            }));
            renderCalendar();
        };

        modal.querySelector(".btn-save-cal").onclick = () => {
            dCfg.customRanges = currentRanges.map(r => ({
                start: `${r.start.getFullYear()}-${String(r.start.getMonth() + 1).padStart(2, '0')}-${String(r.start.getDate()).padStart(2, '0')}`,
                end: `${r.end.getFullYear()}-${String(r.end.getMonth() + 1).padStart(2, '0')}-${String(r.end.getDate()).padStart(2, '0')}`,
                label: r.label
            }));
            modal.classList.remove("open");
        };
    };

    const handleMouseUp = (e) => {
        if (!isSelecting) return;
        isSelecting = false;
        const targetCell = e.target.closest(".cal-day-cell[data-date]");
        if (targetCell && dragStartDate) {
            const dragEndDate = new Date(targetCell.dataset.date.replace(/-/g, '/'));
            const sDate = new Date(Math.min(dragStartDate, dragEndDate));
            const eDate = new Date(Math.max(dragStartDate, dragEndDate));
            const newLabel = `${sDate.getMonth() + 1}.${sDate.getDate()}-${eDate.getMonth() + 1}.${eDate.getDate()}`;

            if (activeHighlightIdx !== null && currentRanges[activeHighlightIdx]) {
                currentRanges[activeHighlightIdx].start = sDate;
                currentRanges[activeHighlightIdx].end = eDate;
                currentRanges[activeHighlightIdx].label = newLabel;
            } else {
                const newIdx = currentRanges.length;
                currentRanges.push({
                    start: sDate,
                    end: eDate,
                    label: newLabel,
                    colorIdx: newIdx % DISTINCT_PALETTE.length
                });
            }
            renderCalendar();
        }
    };

    document.removeEventListener("mouseup", handleMouseUp);
    document.addEventListener("mouseup", handleMouseUp);

    renderCalendar();
    modal.classList.add("open");
}

// ========== Excel 多 Sheet + 不连续选区 (Ctrl选区) 写入引擎 ==========
async function writeExcelCustomData(template, aiData) {
    return Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();
        const config = template.controlsConfig || {};

        for (const [tag, cfg] of Object.entries(config)) {
            let valuesToWrite = [];

            if (cfg.inputType === 'none') {
                if (cfg.autoMode === 'date_seq') {
                    valuesToWrite = generateDateSequence(cfg, 100);
                } else {
                    valuesToWrite = [cfg.autoInputText || ""];
                }
            } else if (cfg.manualSubType === 'boolean') {
                const realValue = (aiData[tag] === 'true');
                if (realValue && cfg.trueAction === 'fixed') valuesToWrite = [cfg.trueActionFixedText || ''];
                else if (realValue && cfg.trueAction === 'ai') valuesToWrite = [aiData[tag] || ''];
                else if (!realValue && cfg.falseAction === 'fixed') valuesToWrite = [cfg.falseActionFixedText || ''];
                else if (!realValue && cfg.falseAction === 'ai') valuesToWrite = [aiData[tag] || ''];
            } else {
                valuesToWrite = [aiData[tag] || cfg.autoInputText || ""];
            }

            // 支持以逗号分割的多个不连续选区（Ctrl 框选出的多区域）
            if (cfg.targetRange && cfg.targetRange.includes("!")) {
                const rangeSegments = cfg.targetRange.split(",").map(s => s.trim()).filter(s => s);
                let valIdx = 0;

                for (const seg of rangeSegments) {
                    const parts = seg.split("!");
                    const sName = parts[0].trim();
                    const cellAddr = parts[1].trim();

                    try {
                        const targetSheet = context.workbook.worksheets.getItem(sName);
                        const targetRange = targetSheet.getRange(cellAddr);
                        targetRange.load(["rowCount", "columnCount"]);
                        await context.sync();

                        const rCount = targetRange.rowCount;
                        const cCount = targetRange.columnCount;
                        const gridValues = [];

                        for (let r = 0; r < rCount; r++) {
                            const row = [];
                            for (let c = 0; c < cCount; c++) {
                                row.push(valuesToWrite[valIdx] !== undefined ? valuesToWrite[valIdx] : (valuesToWrite[valuesToWrite.length - 1] || ""));
                                valIdx++;
                            }
                            gridValues.push(row);
                        }

                        targetRange.values = gridValues;
                    } catch (err) {
                        console.warn(`选区写入 ${seg} 失败:`, err);
                    }
                }
            }

            // 全表占位符 {{tag}} 全局替换
            const singleVal = valuesToWrite[0] || "";
            if (singleVal) {
                for (const sheet of worksheets.items) {
                    const usedRange = sheet.getUsedRangeOrNullObject();
                    usedRange.load("values");
                    await context.sync();

                    if (!usedRange.isNullObject && usedRange.values) {
                        let vals = usedRange.values;
                        let modified = false;
                        const placeholder = `{{${tag}}}`;

                        for (let r = 0; r < vals.length; r++) {
                            for (let c = 0; c < vals[r].length; c++) {
                                let text = String(vals[r][c] || "");
                                if (text.includes(placeholder)) {
                                    vals[r][c] = text.split(placeholder).join(singleVal);
                                    modified = true;
                                }
                            }
                        }
                        if (modified) usedRange.values = vals;
                    }
                }
            }
        }

        await context.sync();
    });
}