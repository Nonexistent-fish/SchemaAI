// ========== 1. 课程表 (6行×5列：一天六节课) 与人数配置交互 ==========
function updateTimetablePlaceholders() {
    var hasAnyData = false;
    document.querySelectorAll(".timetable-cell").forEach(function (cell) {
        var cInput = cell.querySelector(".tt-course");
        var clsInput = cell.querySelector(".tt-class");
        if (cInput && clsInput) {
            if (cInput.value.trim() !== "" || clsInput.value.trim() !== "") {
                hasAnyData = true;
            }
        }
    });

    var isFirstCell = true;
    document.querySelectorAll(".timetable-cell").forEach(function (cell) {
        var cInput = cell.querySelector(".tt-course");
        var clsInput = cell.querySelector(".tt-class");
        if (cInput && clsInput) {
            if (!hasAnyData && isFirstCell) {
                cInput.placeholder = "授课课程";
                clsInput.placeholder = "授课班级";
                isFirstCell = false;
            } else {
                cInput.placeholder = "";
                clsInput.placeholder = "";
            }
        }
    });
}

function renderTimetableUI() {
    var container = document.getElementById("timetable-container");
    if (!container) return;

    var days = ["周一", "周二", "周三", "周四", "周五"];
    var html = "<table class='timetable'><tr><th style='width: 50px;'></th>";
    days.forEach(function (d) { html += "<th>" + d + "</th>"; });
    html += "</tr>";

    for (var p = 0; p < 6; p++) {
        html += "<tr><th>第" + (p + 1) + "节</th>";
        for (var d = 0; d < 5; d++) {
            var cell = (CONFIG.timetable[d] && CONFIG.timetable[d][p]) ? CONFIG.timetable[d][p] : { course: "", classStr: "" };
            html += "<td><div class='timetable-cell' data-day='" + d + "' data-p='" + p + "' draggable='true'><input type='text' class='tt-course' placeholder='' value='" + (cell.course || "") + "'><input type='text' class='tt-class' placeholder='' value='" + (cell.classStr || "") + "'></div></td>";
        }
        html += "</tr>";
    }
    html += "</table>";
    container.innerHTML = html;

    updateTimetablePlaceholders();

    var hintZone = document.getElementById("timetable-drag-hint");
    var trashZone = document.getElementById("timetable-trash-zone");

    container.querySelectorAll(".timetable-cell").forEach(function (cell) {
        var inputs = cell.querySelectorAll("input");
        inputs.forEach(function (inp) {
            inp.addEventListener("blur", function () {
                updateTimetableConfig();
            });
        });

        cell.addEventListener("dragstart", function (e) {
            draggedSource = this;
            e.dataTransfer.effectAllowed = "copyMove";
            e.dataTransfer.setData("text/plain", "cell");
            if (hintZone) hintZone.style.display = "none";
            if (trashZone) trashZone.style.display = "flex";
        });
        cell.addEventListener("dragover", function (e) {
            e.preventDefault();
            e.dataTransfer.dropEffect = "copy";
            this.classList.add("drag-over");
        });
        cell.addEventListener("dragleave", function () {
            this.classList.remove("drag-over");
        });
        cell.addEventListener("drop", function (e) {
            e.preventDefault();
            this.classList.remove("drag-over");
            if (draggedSource && draggedSource !== this) {
                var sourceCourse = draggedSource.querySelector(".tt-course") ? draggedSource.querySelector(".tt-course").value : "";
                var sourceClass = draggedSource.querySelector(".tt-class") ? draggedSource.querySelector(".tt-class").value : "";
                var targetCourse = this.querySelector(".tt-course");
                var targetClass = this.querySelector(".tt-class");
                if (targetCourse) targetCourse.value = sourceCourse;
                if (targetClass) targetClass.value = sourceClass;
                updateTimetableConfig();
            }
        });
        cell.addEventListener("dragend", function () {
            if (hintZone) hintZone.style.display = "block";
            if (trashZone) trashZone.style.display = "none";
        });
    });

    if (trashZone) {
        trashZone.addEventListener("dragover", function (e) {
            e.preventDefault();
            e.dataTransfer.dropEffect = "move";
            trashZone.style.backgroundColor = "rgba(255, 59, 48, 0.15)";
        });
        trashZone.addEventListener("dragleave", function () {
            trashZone.style.backgroundColor = "rgba(255, 59, 48, 0.05)";
        });
        trashZone.addEventListener("drop", function (e) {
            e.preventDefault();
            trashZone.style.backgroundColor = "rgba(255, 59, 48, 0.05)";
            if (draggedSource) {
                var courseInput = draggedSource.querySelector(".tt-course");
                var classInput = draggedSource.querySelector(".tt-class");
                if (courseInput) courseInput.value = "";
                if (classInput) classInput.value = "";
                updateTimetableConfig();
            }
        });
    }
}

function updateTimetableConfig() {
    var uniqueClasses = new Set();
    document.querySelectorAll(".timetable-cell").forEach(function (cell) {
        var day = parseInt(cell.dataset.day);
        var p = parseInt(cell.dataset.p);
        var cInput = cell.querySelector(".tt-course");
        var clsInput = cell.querySelector(".tt-class");
        var course = cInput ? cInput.value.trim() : "";
        var cls = clsInput ? clsInput.value.trim() : "";
        if (!CONFIG.timetable[day]) CONFIG.timetable[day] = [];
        CONFIG.timetable[day][p] = { course: course, classStr: cls };
        if (cls) uniqueClasses.add(cls);
    });

    var newExpected = {};
    uniqueClasses.forEach(function (cls) {
        newExpected[cls] = (CONFIG.expectedStudents[cls] !== undefined && CONFIG.expectedStudents[cls] !== "") ? CONFIG.expectedStudents[cls] : "";
    });
    CONFIG.expectedStudents = newExpected;

    saveStorage();
    renderExpectedStudentsUI();
    updateTimetablePlaceholders();
}

function renderExpectedStudentsUI() {
    var container = document.getElementById("expected-students-container");
    if (!container) return;
    container.innerHTML = "";
    var classes = Object.keys(CONFIG.expectedStudents);

    if (classes.length === 0) return;

    classes.forEach(function (cls) {
        var item = document.createElement("div");
        item.className = "expected-item";

        var label = document.createElement("label");
        label.textContent = cls;
        label.title = cls;

        var input = document.createElement("input");
        input.type = "number";
        input.placeholder = "";
        input.value = CONFIG.expectedStudents[cls];
        input.addEventListener("blur", function () {
            var v = parseInt(input.value);
            CONFIG.expectedStudents[cls] = isNaN(v) ? "" : v;
            saveStorage();
        });

        item.appendChild(label);
        item.appendChild(input);
        container.appendChild(item);
    });
}

// ========== 2. 授课日志模块 ==========
function renderExcelDailyForms() {
    var container = document.getElementById("excel-daily-forms");
    if (!container) return;
    container.innerHTML = "";
    var days = ["周一", "周二", "周三", "周四", "周五"];

    if (!CONFIG.excelDailySettingsMap) CONFIG.excelDailySettingsMap = {};
    if (!CONFIG.masterContentMap) CONFIG.masterContentMap = {};

    for (var d = 0; d < 5; d++) {
        var uniqueLessons = [];
        var seen = new Set();

        (CONFIG.timetable[d] || []).forEach(function (cell) {
            if (cell && cell.course && cell.course.trim() !== "" && cell.classStr && cell.classStr.trim() !== "") {
                var cName = cell.course.trim();
                var clsName = cell.classStr.trim();
                var key = cName + "_" + clsName;
                if (!seen.has(key)) {
                    seen.add(key);
                    uniqueLessons.push({ course: cName, classStr: clsName, count: 1 });
                } else {
                    var existing = uniqueLessons.find(function (l) { return l.course === cName && l.classStr === clsName; });
                    if (existing) existing.count++;
                }
            }
        });

        if (uniqueLessons.length === 0) {
            var emptyCard = document.createElement("div");
            emptyCard.className = "day-card";
            emptyCard.innerHTML = "<div class='day-card-header' style='margin:0;'><div class='day-card-title' style='color:var(--text-sub);'>" + days[d] + "</div><div class='day-card-subtitle' style='color:var(--border-color);'>无排课安排</div></div>";
            container.appendChild(emptyCard);
            continue;
        }

        uniqueLessons.forEach(function (lesson) {
            var card = document.createElement("div");
            card.className = "day-card";

            var settingsKey = d + "_" + lesson.course + "_" + lesson.classStr;
            var saved = CONFIG.excelDailySettingsMap[settingsKey] || { location: "", actual: "", modifier: 0, content: "" };

            var masterKey = lesson.course + "_" + lesson.classStr;
            if (CONFIG.masterContentMap[masterKey] && !saved.content) {
                saved.content = CONFIG.masterContentMap[masterKey];
            }

            var expectedCount = CONFIG.expectedStudents[lesson.classStr] || "";
            var expectedHint = expectedCount !== "" ? "应到: " + expectedCount + "人" : "人数未配置";

            var headerHtml = "<div class='day-card-header'><div class='day-card-title'>" + days[d] + "</div><div class='day-card-subtitle'>" + lesson.course + " · " + lesson.classStr + " (" + lesson.count + "节课)</div></div>";

            var safeLoc = (saved.location || "").replace(/"/g, '&quot;');
            var safeContent = (saved.content || "").replace(/"/g, '&quot;');
            var safeActual = saved.actual !== "" ? saved.actual : expectedCount;

            var bodyHtml = "<div class='day-card-row'><div class='input-modern-wrapper' style='width: 100%;'><span class='modern-label'>授课地点</span><input type='text' class='loc-input' value=\"" + safeLoc + "\" placeholder=''></div></div>" +
                "<div class='day-card-row' style='justify-content: space-between;'>" +
                "    <div class='input-modern-wrapper' style='width: 190px;'><span class='modern-label'>实到</span><input type='number' class='act-input' value=\"" + safeActual + "\"><span class='modern-label' style='margin-left: 2px; margin-right: 0;'>人</span><span class='modern-hint' style='font-size:11px; margin-left:8px;'>" + expectedHint + "</span></div>" +
                "    <div class='input-modern-wrapper mod-wrapper' style='width: 120px;'><span class='modern-label'>修正</span><button class='mod-btn dec-btn'>-</button><span class='mod-val'>" + saved.modifier + "</span><button class='mod-btn inc-btn'>+</button></div>" +
                "</div>" +
                "<div class='day-card-row'><div class='input-modern-wrapper' style='width: 100%;'><span class='modern-label'>内容</span><input type='text' class='cont-input' data-master=\"" + masterKey.replace(/"/g, '&quot;') + "\" value=\"" + safeContent + "\" placeholder=''></div></div>";
            card.innerHTML = headerHtml + bodyHtml;

            var locInput = card.querySelector(".loc-input");
            var actInput = card.querySelector(".act-input");
            var contInput = card.querySelector(".cont-input");
            var decBtn = card.querySelector(".dec-btn");
            var incBtn = card.querySelector(".inc-btn");
            var modValSpan = card.querySelector(".mod-val");

            var saveFields = function () {
                if (!CONFIG.excelDailySettingsMap[settingsKey]) CONFIG.excelDailySettingsMap[settingsKey] = { location: "", actual: "", modifier: 0, content: "" };
                CONFIG.excelDailySettingsMap[settingsKey].location = locInput.value;
                CONFIG.excelDailySettingsMap[settingsKey].actual = actInput.value;
                CONFIG.excelDailySettingsMap[settingsKey].content = contInput.value;
                saveStorage();
            };

            locInput.addEventListener("blur", saveFields);
            actInput.addEventListener("blur", saveFields);

            contInput.addEventListener("blur", function () {
                var val = contInput.value.trim();
                CONFIG.masterContentMap[masterKey] = val;

                var safeQ = masterKey.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
                document.querySelectorAll(".cont-input[data-master=\"" + safeQ + "\"]").forEach(function (input) {
                    input.value = val;
                });

                for (var k in CONFIG.excelDailySettingsMap) {
                    if (k.endsWith("_" + masterKey)) {
                        CONFIG.excelDailySettingsMap[k].content = val;
                    }
                }
                saveFields();
            });

            decBtn.addEventListener("click", function () {
                if (!CONFIG.excelDailySettingsMap[settingsKey]) CONFIG.excelDailySettingsMap[settingsKey] = { location: "", actual: "", modifier: 0, content: "" };
                var m = parseInt(CONFIG.excelDailySettingsMap[settingsKey].modifier) || 0;
                m = Math.max(0, m - 1);
                CONFIG.excelDailySettingsMap[settingsKey].modifier = m;
                modValSpan.innerText = m;
                saveStorage();
            });

            incBtn.addEventListener("click", function () {
                if (!CONFIG.excelDailySettingsMap[settingsKey]) CONFIG.excelDailySettingsMap[settingsKey] = { location: "", actual: "", modifier: 0, content: "" };
                var m = parseInt(CONFIG.excelDailySettingsMap[settingsKey].modifier) || 0;
                m++;
                CONFIG.excelDailySettingsMap[settingsKey].modifier = m;
                modValSpan.innerText = m;
                saveStorage();
            });

            container.appendChild(card);
        });
    }
}

async function fetchExcelAISequence(course, className, lessonsCount, userInputs) {
    var apiUrl = normalizeApiUrl(CONFIG.url);
    if (!apiUrl) throw new Error("未配置 API");

    var majorStr = CONFIG.major ? "专业：" + CONFIG.major + "。" : "";
    var prompt = "身份：专业教师。" + majorStr + "课程：" + course + "。班级：" + className + "。\n本月该班级的该课程共有 " + lessonsCount + " 次课。\n任务：生成一个涵盖全月的、具有连贯性和递进性的授课内容序列。";

    if (userInputs && userInputs.length > 0) {
        prompt += "\n参考以下用户已填写的主题进行发散、提炼和逻辑顺延：\n" + userInputs.join("；");
    }

    prompt += "\n要求：只返回纯JSON对象，格式形如 {\"contents\": [\"第1次课内容\", \"第2次课内容\", ...]}。数组长度必须恰好为 " + lessonsCount + "！内容禁止包含标点符号，严格限制在20个汉字以内。直接输出核心知识点（如：自动变速箱分类与MT构造），拒绝废话！";

    var requestBody = {
        model: CONFIG.model || "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.5
    };

    var isDeepSeek = /deepseek/i.test(CONFIG.model);
    var isReasoner = /reasoner|r1/i.test(CONFIG.model);
    if (isDeepSeek && !isReasoner) requestBody.response_format = { type: "json_object" };
    if (/^deepseek(-v4)?-(flash|pro)$/i.test(CONFIG.model)) requestBody.thinking = { type: "disabled" };

    var headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = "Bearer " + CONFIG.key;

    var response = await fetch(apiUrl, { method: "POST", headers: headers, body: JSON.stringify(requestBody) });
    var resText = await response.text();
    if (!response.ok) throw new Error("HTTP " + response.status + ": " + resText.slice(0, 60));

    var resJson = JSON.parse(resText);
    var choice = resJson.choices && resJson.choices[0];
    var rawText = (choice && choice.message) ? (choice.message.content || choice.message.reasoning_content || "") : "";
    rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, "").trim().replace(/\x60{3}json|\x60{3}/gi, "").trim();

    var startIdx = rawText.indexOf("{");
    var endIdx = rawText.lastIndexOf("}");
    if (startIdx !== -1 && endIdx !== -1) {
        rawText = rawText.substring(startIdx, endIdx + 1);
    }

    var obj = JSON.parse(rawText);
    var contents = obj.contents || obj.topics || Object.values(obj)[0];
    if (Array.isArray(contents)) {
        return contents.map(function (c) { return String(c).replace(/[，。、！？；：“”‘’]/g, "").substring(0, 20); });
    } else {
        throw new Error("模型未返回有效数组");
    }
}

async function handleExcelGenerate() {
    var btn = document.getElementById("generate-excel-btn");
    var saveBtn = document.getElementById("save-excel-btn");
    if (btn) { btn.disabled = true; btn.innerText = "处理中..."; }
    if (saveBtn) saveBtn.classList.remove("show");

    var statusEl = document.getElementById("excel-status-msg");
    var safeShowStatus = function (txt, type) { if (statusEl) statusEl.innerHTML = "<span class='status-" + type + "'>" + txt + "</span>"; };

    var pbContainer = document.getElementById("excel-progress-container");
    var pbBar = document.getElementById("excel-progress-bar");
    var showProgress = function (w) { if (pbContainer) pbContainer.style.display = 'block'; if (pbBar) pbBar.style.width = w + '%'; };

    var modelNameDisplay = CONFIG.model || "AI模型";
    showProgress(10);

    try {
        var now = new Date();
        var y = now.getFullYear();
        var m = now.getMonth() + 1;
        var startDate = new Date(y, m - 1, 1);
        var endDate = new Date(y, m, 0);

        var holidayMap = await getHolidayConfig(y);
        if (!CONFIG.excelDailySettingsMap) CONFIG.excelDailySettingsMap = {};
        var monthSchedule = [];

        for (var currentD = new Date(startDate); currentD <= endDate; currentD.setDate(currentD.getDate() + 1)) {
            var wd = currentD.getDay();
            var dateStr = currentD.getFullYear() + "." + (currentD.getMonth() + 1) + "." + currentD.getDate();

            var isWorkday = false;
            if (holidayMap[dateStr] === "holiday") {
                isWorkday = false;
            } else if (holidayMap[dateStr] === "workday") {
                isWorkday = true;
            } else {
                if (wd >= 1 && wd <= 5) isWorkday = true;
            }

            if (!isWorkday) continue;

            var dayIndex = (wd === 0 || wd === 6) ? 4 : wd - 1;

            var uniqueDayLessons = [];
            var daySeen = new Set();
            (CONFIG.timetable[dayIndex] || []).forEach(function (cell) {
                if (cell && cell.course && cell.course.trim() !== "" && cell.classStr && cell.classStr.trim() !== "") {
                    var cName = cell.course.trim();
                    var clsName = cell.classStr.trim();
                    var key = cName + "_" + clsName;
                    if (!daySeen.has(key)) {
                        daySeen.add(key);
                        uniqueDayLessons.push({ course: cName, classStr: clsName, count: 1 });
                    } else {
                        var existing = uniqueDayLessons.find(function (l) { return l.course === cName && l.classStr === clsName; });
                        if (existing) existing.count++;
                    }
                }
            });

            uniqueDayLessons.forEach(function (lesson) {
                var settingsKey = dayIndex + "_" + lesson.course + "_" + lesson.classStr;
                var savedSet = CONFIG.excelDailySettingsMap[settingsKey] || {};

                var expectedCount = parseInt(CONFIG.expectedStudents[lesson.classStr]) || 0;
                var actBase = savedSet.actual !== "" && savedSet.actual !== undefined ? parseInt(savedSet.actual) : expectedCount;
                var mod = parseInt(savedSet.modifier) || 0;

                var randomActual = "";
                if (actBase > 0) {
                    var min = Math.max(0, actBase - mod);
                    var max = expectedCount > 0 ? Math.min(expectedCount, actBase + mod) : (actBase + mod);
                    if (min > max) min = max;
                    randomActual = Math.floor(Math.random() * (max - min + 1)) + min;
                }

                monthSchedule.push({
                    dateStr: dateStr,
                    course: lesson.course,
                    classStr: lesson.classStr,
                    count: lesson.count,
                    location: savedSet.location || "",
                    expectedCount: expectedCount,
                    randomActual: randomActual,
                    savedContent: savedSet.content || ""
                });
            });
        }

        var groups = {};
        monthSchedule.forEach(function (target) {
            var key = target.course + "|||" + target.classStr;
            if (!groups[key]) groups[key] = [];
            groups[key].push(target);
        });

        safeShowStatus("等待 " + modelNameDisplay + " 响应...", "info");
        showProgress(30);

        var fetchPromises = Object.values(groups).map(async function (groupArray) {
            var userInputs = Array.from(new Set(groupArray.map(function (t) { return t.savedContent; }).filter(function (c) { return c && c.trim() !== ""; })));
            try {
                var aiContents = await fetchExcelAISequence(groupArray[0].course, groupArray[0].classStr, groupArray.length, userInputs);
                groupArray.forEach(function (target, idx) {
                    target.aiContent = aiContents[idx] || target.course;
                });
            } catch (e) {
                groupArray.forEach(function (target) { target.aiContent = target.savedContent || target.course; });
            }
        });

        await Promise.all(fetchPromises);

        safeShowStatus("写入表格中...", "info");
        showProgress(80);

        await Excel.run(async function (context) {
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            var headerRange = sheet.getRange("A2");
            var now = new Date();
            var y = now.getFullYear();
            var m = now.getMonth() + 1;
            headerRange.values = [["授课专业：" + CONFIG.major + "                    授课教师：" + CONFIG.teacher + "                    " + y + " 年 " + m + " 月"]];

            var writeData = [];
            monthSchedule.forEach(function (item) {
                writeData.push([
                    item.dateStr,
                    item.count,
                    item.course,
                    item.aiContent,
                    item.classStr,
                    item.location,
                    item.expectedCount > 0 ? item.expectedCount : "",
                    item.randomActual,
                    "良好",
                    CONFIG.teacher
                ]);
            });

            if (writeData.length > 0) {
                var startRow = 4;
                var clearRange = sheet.getRange("A5:J505");
                clearRange.clear("Contents");

                var targetRange = sheet.getRangeByIndexes(startRow, 0, writeData.length, 10);
                targetRange.values = writeData;
                targetRange.format.horizontalAlignment = "Center";
            }
            await context.sync();
        });

        showProgress(100);
        safeShowStatus("表格写入成功", "success");
        setTimeout(function () {
            if (pbContainer) pbContainer.style.display = 'none';
            if (saveBtn) saveBtn.classList.add("show");
        }, 1000);

    } catch (error) {
        safeShowStatus("运行异常：" + error.message, "error");
        if (pbBar) pbBar.style.background = "var(--error-color)";
    } finally {
        if (btn) { btn.disabled = false; btn.innerText = "生成并写入表格"; }
        setTimeout(function () { if (pbBar) pbBar.style.background = ""; }, 2500);
    }
}

function handleExcelSaveDocument() {
    var validParts = [];
    for (var i = 0; i < 4; i++) {
        var val = CONFIG.excelNamingVals[i];
        var actualStr = "";
        if (val === '教师授课日志') actualStr = "教师授课日志";
        else if (val === '授课计划进度表') actualStr = "授课计划进度表";
        else if (val === '授课专业') actualStr = CONFIG.major || "";
        else if (val === '授课教师') actualStr = CONFIG.teacher || "";
        else if (val === '日期') actualStr = new Date().getFullYear() + "年" + String(new Date().getMonth() + 1).padStart(2, '0') + "月";
        else actualStr = val;

        if (actualStr && actualStr.trim() !== "") {
            validParts.push({ str: actualStr.trim(), originalIndex: i });
        }
    }

    var fileName = "";
    for (var j = 0; j < validParts.length; j++) {
        fileName += validParts[j].str;
        if (j < validParts.length - 1) {
            fileName += CONFIG.excelNamingSeps[validParts[j].originalIndex];
        }
    }
    if (!fileName) fileName = "Excel表格";
    fileName += ".xlsx";

    showStatus("准备保存表格...", "info");
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var file = result.value;
            var slicesReceived = 0, sliceCount = file.sliceCount;
            var docData = [];
            getSlice(file, 0);

            function getSlice(file, nextSlice) {
                file.getSliceAsync(nextSlice, function (sliceResult) {
                    if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                        docData = docData.concat(sliceResult.value.data);
                        slicesReceived++;
                        if (slicesReceived === sliceCount) {
                            file.closeAsync();
                            var u8Array = new Uint8Array(docData);
                            var blob = new Blob([u8Array], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

                            (async () => {
                                try {
                                    if (window.showSaveFilePicker) {
                                        const fileHandle = await window.showSaveFilePicker({
                                            suggestedName: fileName,
                                            types: [{ description: 'Excel 工作簿', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
                                        });
                                        const writable = await fileHandle.createWritable();
                                        await writable.write(blob);
                                        await writable.close();
                                        showStatus("保存成功", "success");
                                    } else {
                                        var url = window.URL.createObjectURL(blob);
                                        var a = document.createElement("a");
                                        a.href = url;
                                        a.download = fileName;
                                        document.body.appendChild(a);
                                        a.click();
                                        window.URL.revokeObjectURL(url);
                                        document.body.removeChild(a);
                                        showStatus("下载完成", "success");
                                    }
                                } catch (err) {
                                    if (err.name !== 'AbortError') showStatus("保存失败", "error");
                                }
                            })();
                        } else { getSlice(file, ++nextSlice); }
                    } else {
                        file.closeAsync();
                        showStatus("处理失败", "error");
                    }
                });
            }
        } else {
            showStatus("获取接口异常", "error");
        }
    });
}

// ========== 3. 🌟 授课计划模块（低可视度五色胶囊 + 多行引导授课内容 + 纯净摘要 + 自动续写） ==========

// 五个颜色各异的低可视度半透明小胶囊样式（对应：一、二、三、四、五）
var DAY_CAPSULE_STYLES = [
    { label: "一", bg: "rgba(0, 122, 255, 0.08)", color: "rgba(0, 122, 255, 1)", border: "rgba(0, 122, 255, 0.22)" },   // 周一：柔和蓝
    { label: "二", bg: "rgba(46, 125, 50, 0.08)", color: "rgba(46, 125, 50, 1)", border: "rgba(46, 125, 50, 0.22)" },   // 周二：柔和绿
    { label: "三", bg: "rgba(239, 108, 0, 0.08)", color: "rgba(239, 108, 0, 1)", border: "rgba(239, 108, 0, 0.22)" },   // 周三：柔和橙
    { label: "四", bg: "rgba(123, 31, 162, 0.08)", color: "rgba(123, 31, 162, 1)", border: "rgba(123, 31, 162, 0.22)" },// 周四：柔和紫
    { label: "五", bg: "rgba(194, 24, 91, 0.08)", color: "rgba(194, 24, 91, 1)", border: "rgba(194, 24, 91, 0.22)" }    // 周五：柔和粉
];

// 生成课程右侧的低透明度星期小胶囊 HTML
function buildDayCapsulesHtml(daysArr) {
    if (!Array.isArray(daysArr) || daysArr.length === 0) {
        return "<span style='font-size:11px; color:var(--text-sub); opacity:0.6;'>未排课</span>";
    }
    var pills = daysArr.map(function (dIdx) {
        var st = DAY_CAPSULE_STYLES[dIdx] || DAY_CAPSULE_STYLES[0];
        return "<span style='display:inline-flex; align-items:center; justify-content:center; padding:1px 7px; border-radius:999px; font-size:11px; font-weight:500; line-height:1.35; opacity:0.55; background:" + st.bg + "; color:" + st.color + "; border:1px solid " + st.border + ";'>" + st.label + "</span>";
    }).join("");
    return "<div style='display:inline-flex; align-items:center; gap:4px;'>" + pills + "</div>";
}

// 从用户填写的课程表中提取所有课程及其排课星期分布
function getTimetableCoursesMap() {
    var courseMap = {};

    for (var d = 0; d < 5; d++) {
        var daySlots = (CONFIG.timetable && CONFIG.timetable[d]) ? CONFIG.timetable[d] : [];
        for (var p = 0; p < 6; p++) {
            var cell = daySlots[p];
            if (cell && cell.course && cell.course.trim() !== "") {
                var cName = cell.course.trim();
                if (!courseMap[cName]) {
                    courseMap[cName] = new Set();
                }
                courseMap[cName].add(d);
            }
        }
    }

    (CONFIG.courses || []).forEach(function (c) {
        if (c && c.trim() && !courseMap[c.trim()]) {
            courseMap[c.trim()] = new Set();
        }
    });

    var list = [];
    Object.keys(courseMap).forEach(function (cName) {
        var daysArr = Array.from(courseMap[cName]).sort();
        list.push({
            course: cName,
            days: daysArr
        });
    });

    return list;
}

// 格式化日期为 YYYY.MM.DD
function formatPlanDateDot(dateObj) {
    var y = dateObj.getFullYear();
    var m = String(dateObj.getMonth() + 1).padStart(2, '0');
    var d = String(dateObj.getDate()).padStart(2, '0');
    return y + "." + m + "." + d;
}

// 根据起始日期、终止日期以及该课程在课程表中的排课星期，推算中间所有周次的授课时间 (YYYY.MM.DD-YYYY.MM.DD)
function computeWeeklyDateSpans(startDateStr, endDateStr, courseDays) {
    var startD = new Date(startDateStr.replace(/-/g, '/'));
    var endD = new Date(endDateStr.replace(/-/g, '/'));
    if (isNaN(startD.getTime()) || isNaN(endD.getTime()) || startD > endD) {
        throw new Error("请填写有效的起始日期与终止日期（起始日期不能晚于终止日期）");
    }

    var dStart = 0, dEnd = 2;
    if (Array.isArray(courseDays) && courseDays.length >= 2) {
        dStart = courseDays[0];
        dEnd = courseDays[courseDays.length - 1];
    } else if (Array.isArray(courseDays) && courseDays.length === 1) {
        if (courseDays[0] <= 2) { dStart = 0; dEnd = 2; }
        else { dStart = 2; dEnd = 4; }
    }

    var dayOfWeek = startD.getDay() === 0 ? 7 : startD.getDay();
    var currMon = new Date(startD);
    currMon.setDate(startD.getDate() - (dayOfWeek - 1));

    var spans = [];
    while (currMon <= endD) {
        var wStart = new Date(currMon);
        wStart.setDate(currMon.getDate() + dStart);
        var wEnd = new Date(currMon);
        wEnd.setDate(currMon.getDate() + dEnd);

        if (wEnd >= startD && wStart <= endD) {
            spans.push(formatPlanDateDot(wStart) + "-" + formatPlanDateDot(wEnd));
        }
        currMon.setDate(currMon.getDate() + 7);
    }

    return spans;
}

// 渲染「授课计划」页面：默认填充当月初到当月末，并渲染带低透明度五色胶囊的课程下拉框
function renderExcelPlanView() {
    var startInp = document.getElementById("plan-start-date");
    var endInp = document.getElementById("plan-end-date");
    var hiddenSelect = document.getElementById("plan-course-select");
    var displayContent = document.getElementById("plan-course-selected-content");
    var dropdown = document.getElementById("plan-course-select-dropdown");
    var displayBox = document.getElementById("plan-course-select-display");

    if (startInp && endInp && (!startInp.value || !endInp.value)) {
        var now = new Date();
        var y = now.getFullYear();
        var m = String(now.getMonth() + 1).padStart(2, '0');
        var lastDay = new Date(y, now.getMonth() + 1, 0).getDate();
        startInp.value = y + "-" + m + "-01";
        endInp.value = y + "-" + m + "-" + String(lastDay).padStart(2, '0');
    }

    if (!hiddenSelect || !displayContent || !dropdown) return;

    var courseList = getTimetableCoursesMap();
    dropdown.innerHTML = "";

    var selectCourseItem = function (item) {
        hiddenSelect.value = item.course;
        displayContent.innerHTML = "<span style='font-weight:600; color:var(--text-main);'>" + item.course + "</span>" + buildDayCapsulesHtml(item.days);
    };

    if (courseList.length === 0) {
        hiddenSelect.value = "";
        displayContent.innerHTML = "<span style='color:var(--text-sub); font-size:14px;'>请先在右上角「用户信息-课程表」中填写课程</span>";
    } else {
        var currentVal = hiddenSelect.value;
        var matched = courseList.find(function (x) { return x.course === currentVal; }) || courseList[0];
        selectCourseItem(matched);

        courseList.forEach(function (item) {
            var optDiv = document.createElement("div");
            optDiv.className = "custom-select-item";
            optDiv.style.cssText = "display:flex; align-items:center; justify-content:space-between; gap:10px; padding:12px 16px;";
            optDiv.innerHTML = "<span style='font-weight:600;'>" + item.course + "</span>" + buildDayCapsulesHtml(item.days);
            optDiv.onclick = function (e) {
                e.stopPropagation();
                selectCourseItem(item);
                dropdown.style.display = "none";
                if (displayBox) displayBox.classList.remove("active");
            };
            dropdown.appendChild(optDiv);
        });
    }
}

// 清理生成的章节内容摘要，去除任何 XXXXX-XXXXX 或“任务一：”前缀，只保留如“离合器的功用与组成”的纯净表述
function cleanConciseSummary(rawStr) {
    if (!rawStr) return "";
    var s = String(rawStr).trim();
    if (s.indexOf(" - ") !== -1) s = s.split(" - ").pop().trim();
    else if (s.indexOf("—") !== -1) s = s.split("—").pop().trim();
    else if (s.indexOf("-") !== -1 && /^[^\d]+-[^\d]+$/.test(s)) s = s.split("-").pop().trim();
    s = s.replace(/^任务[一二三四五六七八九十\d]+[、：:\s]*/g, "");
    s = s.replace(/^\d+[、．\.\s]*/g, "");
    s = s.replace(/[；;。]+$/g, "").trim();
    return s;
}

// 兜底匹配器：根据授课内容或摘要关键词自动匹配最接近的大纲行
function findClosestSyllabusRowByText(queryText, allRows) {
    if (!allRows || allRows.length === 0) return { task: "汽车构造基础", total: 6, theory: 2, practical: 4 };
    var q = String(queryText || "").trim();
    if (!q) return allRows[0];

    var bestRow = allRows[0];
    var bestScore = -1;

    allRows.forEach(function (r) {
        var score = 0;
        var fullStr = r.chapter + " " + r.task + " " + r.skill + " " + r.content;
        if (r.chapter && q.indexOf(r.chapter) !== -1) score += 12;
        if (r.task && q.indexOf(r.task) !== -1) score += 10;
        for (var i = 0; i < q.length - 1; i++) {
            var bi = q.substr(i, 2);
            if (fullStr.indexOf(bi) !== -1) score += 1;
        }
        if (score > bestScore) {
            bestScore = score;
            bestRow = r;
        }
    });

    return bestRow;
}

// 请求 AI：根据用户多行输入的「授课内容」（如 1. ... 2. ... 3. ...）和周次数，自动匹配大纲中的章节行并生成纯净摘要
async function fetchAIPlanAndRowIds(courseName, teachingContent, weeksCount) {
    var apiUrl = normalizeApiUrl(CONFIG.url);
    if (!apiUrl) throw new Error("请先在系统配置中填写 API 接口地址");

    var allRows = getAllSyllabusRows();
    var syllabusCatalog = allRows.map(function (r) {
        return "[" + r.id + "] " + r.chapter + "：" + r.task + "（" + r.content + "）";
    }).join("\n");

    var prompt = `身份：中等职业技术学校专业教师。
课程名称：${courseName}
需要生成的周次数量：${weeksCount} 周

用户输入的授课内容要点如下（可能按 1. 2. 3. 或多行顺序列出）：
${teachingContent}

以下是教学大纲的数据行列表（格式为 [行号id] 章节：任务与内容）：
${syllabusCatalog}

请严格按照以下要求生成 ${weeksCount} 周的授课计划：
1. 仔细解析用户输入的各条授课内容要点，严格遵循用户列出的先后顺序，在大纲数据行中匹配对应的章节与任务行，并将教学内容合理、均匀地分配到 ${weeksCount} 周中。
2. 对于每一周，从上述大纲数据行中选出最匹配的一行，记录其数字行号 \`row_id\`（程序将根据你选的 \`row_id\` 自动填入该行的大纲规定课时数、理论课和实训课课时）。
3. 为每一周编写简明扼要的「授课章节的内容摘要」字段 \`summary\`：
   - 严禁使用“XXXXX - XXXXX”、“任务一：”或“1.”之类的前缀格式！
   - 只需要直接写核心教学内容短语，例如：“离合器的功用与组成”、“曲柄连杆机构的构造与检修”、“气门间隙的检查与调整”，字数控制在 8~20 字以内。
4. 只返回纯 JSON 对象，禁止包含任何 Markdown 代码块或多余文字，格式严格如下：
{"items": [{"row_id": 6, "summary": "曲柄连杆机构的功用与组成"}, {"row_id": 7, "summary": "机体组的结构原理与分解装配"}]}
注意：\`items\` 数组长度必须恰好为 ${weeksCount}！`;

    var requestBody = {
        model: CONFIG.model || "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.4,
        max_tokens: 8192
    };

    var headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = "Bearer " + CONFIG.key;

    var response = await fetch(apiUrl, { method: "POST", headers: headers, body: JSON.stringify(requestBody) });
    var resText = await response.text();
    if (!response.ok) throw new Error("HTTP " + response.status + ": " + resText.slice(0, 60));

    var resJson;
    try {
        resJson = JSON.parse(resText);
    } catch (e) {
        throw new Error("接口返回了非 JSON 数据，请检查 API 地址是否正确");
    }

    var choice = resJson.choices && resJson.choices[0];
    var rawText = (choice && choice.message) ? (choice.message.content || choice.message.reasoning_content || "") : "";
    rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, "").trim().replace(/\x60{3}(?:json)?|\x60{3}/gi, "").trim();

    var startIdx = rawText.indexOf("{");
    var endIdx = rawText.lastIndexOf("}");
    if (startIdx !== -1 && endIdx !== -1 && endIdx > startIdx) {
        rawText = rawText.substring(startIdx, endIdx + 1);
    }

    var parsed = JSON.parse(rawText);
    var items = parsed.items || parsed.plan || parsed.weeks || Object.values(parsed).find(Array.isArray);
    if (!Array.isArray(items) || items.length === 0) {
        throw new Error("AI 未返回有效的周计划数组");
    }

    // 由程序根据 AI 选出的 row_id 填入大纲规定课时数、理论课、实训课
    var resultRows = [];
    for (var i = 0; i < weeksCount; i++) {
        var aiItem = items[i] || items[items.length - 1] || {};
        var rid = parseInt(aiItem.row_id !== undefined ? aiItem.row_id : aiItem.id);
        var matchedRow = (!isNaN(rid) && allRows[rid]) ? allRows[rid] : findClosestSyllabusRowByText(aiItem.summary || teachingContent, allRows);

        var cleanSum = cleanConciseSummary(aiItem.summary);
        if (!cleanSum) {
            cleanSum = cleanConciseSummary(matchedRow.task || matchedRow.content);
        }

        resultRows.push({
            summary: cleanSum,
            total: matchedRow.total,
            theory: matchedRow.theory,
            practical: matchedRow.practical
        });
    }

    return resultRows;
}

// 执行「生成授课计划」并写入表格（自动检测已有数据并往下顺延续写）
async function handleExcelPlanGenerate() {
    var btn = document.getElementById("generate-plan-btn");
    var saveBtn = document.getElementById("save-plan-btn");
    var startInp = document.getElementById("plan-start-date");
    var endInp = document.getElementById("plan-end-date");
    var courseSelect = document.getElementById("plan-course-select");
    var contentInp = document.getElementById("plan-teaching-content");

    var pbContainer = document.getElementById("plan-progress-container");
    var pbBar = document.getElementById("plan-progress-bar");
    var showProgress = function (w) { if (pbContainer) pbContainer.style.display = "block"; if (pbBar) pbBar.style.width = w + "%"; };

    clearStatus();
    if (saveBtn) saveBtn.classList.remove("show");

    var startDateStr = startInp ? startInp.value : "";
    var endDateStr = endInp ? endInp.value : "";
    var selectedCourse = courseSelect ? courseSelect.value.trim() : "";
    var teachingContent = contentInp ? contentInp.value.trim() : "";

    if (!startDateStr || !endDateStr) {
        showStatus("请选择起始日期和终止日期", "error");
        return;
    }
    if (!selectedCourse) {
        showStatus("请先在「用户信息-课程表」中填写课程并选择", "error");
        return;
    }
    if (!teachingContent) {
        showStatus("请填写授课内容", "error");
        return;
    }
    if (!normalizeApiUrl(CONFIG.url)) {
        showStatus("请先在系统配置中设置 API 接口地址", "error");
        return;
    }

    try {
        if (btn) { btn.disabled = true; btn.innerText = "生成中..."; }
        showProgress(15);

        // 1. 获取所选课程在课表中的排课日，计算起止日期之间的所有周次授课时间
        var courseList = getTimetableCoursesMap();
        var courseObj = courseList.find(function (c) { return c.course === selectedCourse; });
        var courseDays = courseObj ? courseObj.days : [0, 1, 2];
        var weeklyDateSpans = computeWeeklyDateSpans(startDateStr, endDateStr, courseDays);

        if (weeklyDateSpans.length === 0) {
            throw new Error("所选起止日期区间内没有匹配的授课周次");
        }

        // 2. 扫描 Excel 表格（第4行~第23行，即周次1~20），检测哪些周次已有数据，从第一个空白周次往下续写
        var firstEmptyWeekIdx = 0;
        await Excel.run(async function (context) {
            var sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            var targetSheet = sheets.items.find(function (s) { return s.name.indexOf("周计划") !== -1 || s.name.indexOf("进度表（二）") !== -1; })
                || context.workbook.worksheets.getActiveWorksheet();

            var checkRange = targetSheet.getRange("B4:F23");
            checkRange.load("values");
            await context.sync();

            var vals = checkRange.values || [];
            var lastOccupied = -1;
            for (var r = 0; r < vals.length; r++) {
                var rowHasData = vals[r].some(function (cellVal) {
                    return cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== "";
                });
                if (rowHasData) {
                    lastOccupied = r;
                }
            }
            firstEmptyWeekIdx = lastOccupied + 1;
        });

        if (firstEmptyWeekIdx >= 20) {
            throw new Error("当前表格的 1~20 周次均已有数据，无剩余空行可往下续写");
        }

        var availableSlots = 20 - firstEmptyWeekIdx;
        var targetDateSpans = weeklyDateSpans.slice(0, availableSlots);
        var weeksToGenerate = targetDateSpans.length;

        showStatus("正在请求 " + (CONFIG.model || "AI") + " 规划 " + weeksToGenerate + " 周授课计划...", "info");
        showProgress(40);

        // 3. 调用 AI 生成纯净章节摘要并匹配大纲行号，由程序提取大纲课时
        var generatedPlanRows = await fetchAIPlanAndRowIds(selectedCourse, teachingContent, weeksToGenerate);

        showStatus("正在写入授课计划进度表（第 " + (firstEmptyWeekIdx + 1) + " 周起）...", "info");
        showProgress(80);

        // 4. 写入 Excel（从 firstEmptyWeekIdx 开始往下写，绝不覆盖已有周次）
        await Excel.run(async function (context) {
            var sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            var sheet = sheets.items.find(function (s) { return s.name.indexOf("周计划") !== -1 || s.name.indexOf("进度表（二）") !== -1; })
                || context.workbook.worksheets.getActiveWorksheet();

            sheet.getRange("A1").values = [["《 " + selectedCourse + " 》授课计划进度表（二）"]];

            var writeMatrix = [];
            for (var i = 0; i < weeksToGenerate; i++) {
                var item = generatedPlanRows[i];
                writeMatrix.push([
                    targetDateSpans[i], // B列：授课时间 (YYYY.MM.DD-YYYY.MM.DD)
                    item.summary,       // C列：授课章节的内容摘要（如“离合器的功用与组成”）
                    item.total,         // D列：大纲规定课时数
                    item.theory,        // E列：理论课
                    item.practical,     // F列：实训课
                    0,                  // G列：考核常态填写 0
                    ""                  // H列：教具名称不填为空
                ]);
            }

            var startExcelRow = 4 + firstEmptyWeekIdx;
            var endExcelRow = startExcelRow + weeksToGenerate - 1;
            var targetRange = sheet.getRange("B" + startExcelRow + ":H" + endExcelRow);
            targetRange.values = writeMatrix;

            // 第 24 行总计：自动汇总 D4:G23
            sheet.getRange("D24:G24").formulas = [[
                "=SUM(D4:D23)",
                "=SUM(E4:E23)",
                "=SUM(F4:F23)",
                "=SUM(G4:G23)"
            ]];

            // 第 25 行：填入用户设置中的授课教师与系统当天年月日
            var teacherName = CONFIG.teacher || "";
            sheet.getRange("A25").values = [["授课教师："]];
            sheet.getRange("C25").values = [[teacherName]];

            var now = new Date();
            var todayStr = now.getFullYear() + " 年 " + (now.getMonth() + 1) + " 月 " + now.getDate() + " 日";
            sheet.getRange("D25").values = [[todayStr]];

            await context.sync();
        });

        showProgress(100);
        showStatus("成功写入第 " + (firstEmptyWeekIdx + 1) + " ~ " + (firstEmptyWeekIdx + weeksToGenerate) + " 周授课计划！", "success");
        setTimeout(function () {
            if (pbContainer) pbContainer.style.display = "none";
            if (saveBtn) saveBtn.classList.add("show");
        }, 1000);

    } catch (error) {
        if (pbBar) pbBar.style.background = "var(--error-color)";
        showStatus("生成失败：" + error.message, "error");
    } finally {
        if (btn) { btn.disabled = false; btn.innerText = "生成授课计划"; }
        setTimeout(function () { if (pbBar) pbBar.style.background = ""; }, 2500);
    }
}

// 解析后期导入的大纲文件 (.xlsx/.xls/.csv) 并扁平化追加到 customSyllabusRows
function parseImportedSyllabusToFlatRows(rows) {
    var flatList = [];
    var colMap = { proj: 0, task: 2, skill: 3, content: 5, total: 6, theory: 7, practical: 8 };
    var currChapter = "";
    var lastTask = "";

    for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        if (!Array.isArray(row) || row.length === 0) continue;

        var c0 = String(row[0] !== undefined && row[0] !== null ? row[0] : "").trim();
        var rowJoined = row.map(function (x) { return String(x || "").trim(); }).join(" ");

        if (c0.indexOf("教学大纲") !== -1 || c0 === "小计" || c0 === "总计" || c0.indexOf("本人对以上") !== -1) continue;

        if (c0 === "项目" || c0 === "章节" || (rowJoined.indexOf("理论") !== -1 && rowJoined.indexOf("实训") !== -1)) {
            row.forEach(function (cellVal, cIdx) {
                var s = String(cellVal || "").replace(/\s+/g, "");
                if (s === "项目" || s === "章节" || s === "模块") colMap.proj = cIdx;
                else if (s === "任务" || s === "课题") colMap.task = cIdx;
                else if (s.indexOf("知识点") !== -1 || s.indexOf("技能点") !== -1) colMap.skill = cIdx;
                else if (s.indexOf("教学内容") !== -1 || s.indexOf("内容摘要") !== -1) colMap.content = cIdx;
                else if (s === "课时" || s === "总课时" || s === "大纲规定课时数") colMap.total = cIdx;
                else if (s.indexOf("理论") !== -1) colMap.theory = cIdx;
                else if (s.indexOf("实训") !== -1 || s.indexOf("实操") !== -1) colMap.practical = cIdx;
            });
            continue;
        }

        var taskStr = String(row[colMap.task] || "").trim().replace(/\s+/g, " ");
        var skillStr = String(row[colMap.skill] || "").trim().replace(/\s+/g, " ");
        var contentStr = String(row[colMap.content] || "").trim().replace(/\s+/g, " ");
        if (!c0 && !taskStr && !contentStr) continue;

        var tot = parseFloat(row[colMap.total]);
        var the = parseFloat(row[colMap.theory]);
        var pra = parseFloat(row[colMap.practical]);
        tot = isNaN(tot) ? 0 : tot;
        the = isNaN(the) ? 0 : the;
        pra = isNaN(pra) ? 0 : pra;
        if (tot === 0 && (the > 0 || pra > 0)) tot = the + pra;
        if (tot === 0 && the === 0 && pra === 0) continue;

        if (c0 !== "") {
            currChapter = c0.replace(/[\r\n]+/g, "").replace(/[（\(]?\d+(?:\.\d+)?[）\)]?$/, "").trim();
            lastTask = "";
        }
        if (taskStr) lastTask = taskStr;

        flatList.push({
            chapter: currChapter || "导入章节",
            task: cleanConciseSummary(lastTask || currChapter || "常规任务"),
            skill: skillStr.replace(/[（\(][\d\.]+[）\)][A-Z]?/g, "").trim(),
            content: contentStr.replace(/[（\(][\d\.]+[）\)][A-Z]?/g, "").trim(),
            total: tot,
            theory: the,
            practical: pra
        });
    }
    return flatList;
}

// ========== 页面专属事件绑定 ==========
document.addEventListener("DOMContentLoaded", function () {
    var genExcelBtn = document.getElementById("generate-excel-btn");
    if (genExcelBtn) genExcelBtn.addEventListener("click", handleExcelGenerate);

    var saveExcelBtn = document.getElementById("save-excel-btn");
    if (saveExcelBtn) saveExcelBtn.addEventListener("click", handleExcelSaveDocument);

    var genPlanBtn = document.getElementById("generate-plan-btn");
    if (genPlanBtn) genPlanBtn.addEventListener("click", handleExcelPlanGenerate);

    var savePlanBtn = document.getElementById("save-plan-btn");
    if (savePlanBtn) savePlanBtn.addEventListener("click", handleExcelSaveDocument);

    // 绑定「授课计划 - 选择课程」自定义彩色胶囊下拉菜单展开/收起
    var planCourseDisplay = document.getElementById("plan-course-select-display");
    var planCourseDropdown = document.getElementById("plan-course-select-dropdown");
    if (planCourseDisplay && planCourseDropdown) {
        planCourseDisplay.addEventListener("click", function (e) {
            e.stopPropagation();
            var isOpen = planCourseDropdown.style.display === "block";
            document.querySelectorAll(".custom-select-dropdown").forEach(function (d) { d.style.display = "none"; });
            document.querySelectorAll(".custom-select-display").forEach(function (d) { d.classList.remove("active"); });
            if (!isOpen) {
                planCourseDropdown.style.display = "block";
                planCourseDisplay.classList.add("active");
            }
        });
    }

    document.addEventListener("click", function () {
        if (planCourseDropdown) planCourseDropdown.style.display = "none";
        if (planCourseDisplay) planCourseDisplay.classList.remove("active");
    });

    // 右上角导入新大纲拓展行库
    var importFileBtn = document.getElementById("import-syllabus-file-btn");
    var fileInput = document.getElementById("syllabus-file-input");
    if (importFileBtn && fileInput) {
        importFileBtn.addEventListener("click", function () {
            fileInput.click();
        });

        fileInput.addEventListener("change", function (e) {
            var file = e.target.files && e.target.files[0];
            if (!file) return;

            var reader = new FileReader();
            reader.onload = function (ev) {
                try {
                    if (typeof XLSX === "undefined") throw new Error("未加载表格解析组件");
                    var data = new Uint8Array(ev.target.result);
                    var workbook = XLSX.read(data, { type: "array" });
                    var addedCount = 0;

                    if (!Array.isArray(CONFIG.customSyllabusRows)) CONFIG.customSyllabusRows = [];

                    workbook.SheetNames.forEach(function (sName) {
                        var sheet = workbook.Sheets[sName];
                        var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
                        var extracted = parseImportedSyllabusToFlatRows(rows);
                        extracted.forEach(function (r) {
                            CONFIG.customSyllabusRows.push(r);
                            addedCount++;
                        });
                    });

                    if (addedCount === 0) throw new Error("未在文件中识别到有效的大纲课时行");
                    saveStorage();
                    showStatus("已成功拓展导入 " + addedCount + " 条大纲数据行", "success");
                } catch (err) {
                    showStatus("导入失败：" + err.message, "error");
                }
            };
            reader.readAsArrayBuffer(file);
            e.target.value = "";
        });
    }
});